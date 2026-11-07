import json
import math
import os
import re
import sys
from datetime import datetime, timezone

import pandas as pd
import requests
from bs4 import BeautifulSoup

DATAJOLT_BASE_URL = "https://datajolt.io/prices/"
TIMEOUT_SECONDS = 5
LOG_COLUMNS = ["Timestamp", "Ingredient", "Price"]

# Tier 2 Fallback: Standard wholesale container weights (used when page text omits package dimensions)
STANDARD_COMMODITY_WEIGHTS = {
    "Roma Tomatoes": 25.0,   # Standard commercial 25-lb box / carton
    "Tomatillos": 35.0,      # Standard 35-lb bushel basket or crate
    "Onions": 50.0,          # Standard wholesale 50-lb mesh sack
    "Jalapenos": 30.0,       # Standard 1 1/9 bushel box (~30 lbs)
    "Cilantro": 12.0,        # Industry Standard wholesale crate (~12 lbs net herbs)
    "Mangoes": 4.0,          # Standard commercial single-layer flat box (~4 lbs)
    "Pineapples": 5.0,       # Standard single fresh premium bulk unit (~5 lbs)
}

# Complete fallback estimates to protect pipeline execution against scraping/network errors
FALLBACK_ESTIMATES = {
    "Roma Tomatoes": 1.30,
    "Tomatillos": 1.50,
    "Onions": 0.65,
    "Jalapenos": 0.95,
    "Cilantro": 1.10,
    "Mangoes": 2.20,
    "Pineapples": 1.80,
}

SECONDARY_ESTIMATES = {
    "Roma Tomatoes": 1.25,
    "Tomatillos": 1.40,
    "Onions": 0.60,
    "Jalapenos": 1.30,
    "Cilantro": 1.15,
    "Mangoes": 2.10,
    "Pineapples": 1.90,
}

# Map internal ingredient labels to their URL slugs on Datajolt
SLUG_MAPPING = {
    "Roma Tomatoes": "tomatoes-plum-type",
    "Tomatillos": "tomatillos",
    "Onions": "onions-dry",
    "Jalapenos": "peppers-jalapeno",
    "Cilantro": "cilantro",
    "Mangoes": "mangoes",
    "Pineapples": "pineapples",
}


def _is_finite_number(value):
    """Return True for finite integers or floats, excluding booleans."""
    return (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(value)
    )


def fetch_datajolt_prices():
    """
    Scrape live prices and isolate per-pound metrics. 
    
    Uses a 3-tier approach:
    1. Parse numeric weights right out of the text (e.g., '25 lb carton').
    2. Fall back to standard commercial package weights (e.g., Cilantro = 12 lbs).
    3. Default to 1.0 to guarantee code logic never triggers a zero division error.
    """
    estimates = {}

    for ingredient, slug in SLUG_MAPPING.items():
        try:
            url = f"{DATAJOLT_BASE_URL}{slug}"
            response = requests.get(url, timeout=TIMEOUT_SECONDS)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")
            lead_element = soup.find("p", class_=lambda c: c and "text-lead" in c)

            if not lead_element:
                raise ValueError("Summary paragraph ('text-lead') not found on page")

            full_text = lead_element.get_text()
            print(f"DEBUG [{ingredient}]: Extracted text: '{full_text.strip()}'")

            price_match = re.search(r"\$\s*([\d,]+\.\d{2})", full_text)
            if not price_match:
                raise ValueError(f"Could not parse valid price pattern from text.")

            raw_pack_price = float(price_match.group(1).replace(",", ""))

            # --- TIERED WEIGHT DETERMINATION ENGINE ---
            weight_match = re.search(r"(\d+)\s*lb", full_text, re.IGNORECASE)
            
            if weight_match:
                conversion_divisor = float(weight_match.group(1))
                print(f"DEBUG [{ingredient}]: Dynamic text match caught container size: {conversion_divisor} lbs.")
            elif ingredient in STANDARD_COMMODITY_WEIGHTS:
                conversion_divisor = STANDARD_COMMODITY_WEIGHTS[ingredient]
                print(f"DEBUG [{ingredient}]: Packaging details omitted in text. Applied commodity fallback: {conversion_divisor} lbs.")
            else:
                print(f"WARNING [{ingredient}]: Unknown package size pattern. Setting conversion divisor to 1.0.", file=sys.stderr)
                conversion_divisor = 1.0

            if not _is_finite_number(raw_pack_price) or raw_pack_price < 0 or conversion_divisor <= 0:
                raise ValueError("Price and weight metrics must evaluate to valid non-negative scales.")

            price_per_pound = raw_pack_price / conversion_divisor
            estimates[ingredient] = round(price_per_pound, 2)
            print(f"SUCCESS [{ingredient}]: Evaluated and logged at ${estimates[ingredient]}/lb")

        except (requests.RequestException, ValueError, TypeError) as exc:
            print(
                f"CRITICAL NETWORK/PARSING ERROR for {ingredient}: {exc}",
                file=sys.stderr,
            )

    return estimates

def _resolve_log_path(data_dir):
    if data_dir is not None:
        directory = os.path.abspath(os.fspath(data_dir))
    else:
        script_path = globals().get("__file__")
        if script_path:
            directory = os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(script_path))),
                "data",
            )
        else:
            directory = os.path.join(os.getcwd(), "data")

    os.makedirs(directory, exist_ok=True)
    return os.path.join(directory, "historical_estimates.csv")


def _append_estimate_log(csv_path, estimates):
    """Append long-format estimates, rejecting files with an incompatible schema."""
    if not estimates:
        return

    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    rows = [
        {"Timestamp": timestamp, "Ingredient": ingredient, "Price": price}
        for ingredient, price in estimates.items()
    ]
    new_log = pd.DataFrame(rows, columns=LOG_COLUMNS)

    if os.path.exists(csv_path) and os.path.getsize(csv_path) > 0:
        try:
            existing_columns = list(pd.read_csv(csv_path, nrows=0).columns)
        except (OSError, pd.errors.ParserError, UnicodeDecodeError) as exc:
            raise RuntimeError(f"Cannot read existing log schema: {exc}") from exc

        if existing_columns != LOG_COLUMNS:
            raise RuntimeError(
                f"Incompatible log schema in {csv_path}: "
                f"expected {LOG_COLUMNS}, found {existing_columns}."
            )

        new_log.to_csv(csv_path, mode="a", header=False, index=False)
    else:
        new_log.to_csv(csv_path, index=False)


def run_reconciliation_pipeline(spec_path, data_dir=None):
    """
    Read SKU specifications, calculate estimated costs, and log estimates.
    """
    with open(spec_path, "r", encoding="utf-8") as spec_file:
        specs = json.load(spec_file)

    if not isinstance(specs, dict) or not isinstance(specs.get("skus"), dict):
        raise ValueError("Specification must be a JSON object containing a 'skus' object.")

    estimates = fetch_datajolt_prices()
    reconciled_estimates = {}
    market_prices_workbook_data = {}

    for ingredient in SLUG_MAPPING.keys():
        benchmark = SECONDARY_ESTIMATES[ingredient]
        
        # Branch A: Live Scraping Succeeded -> Evaluate Volatility Variance
        if ingredient in estimates:
            estimate = estimates[ingredient]
            
            variance = abs(estimate - benchmark) / max(abs(benchmark), 1e-5)

            if variance > 0.25:
                reconciled_price = round(benchmark, 2)
                print(f"NOTICE [{ingredient}]: High Volatility ({variance:.1%}). Clamped to baseline safety target: ${reconciled_price}/lb")
            else:
                reconciled_price = round((estimate + benchmark) / 2, 2)
                
            market_prices_workbook_data[ingredient] = {
                "unclamped": round(estimate, 2),
                "reconciled": reconciled_price
            }
            
        # Branch B: Web Scraper Failed -> Trigger System Fail-safe Bypassing Variance
        else:
            reconciled_price = FALLBACK_ESTIMATES[ingredient]
            print(
                f"WARNING [{ingredient}]: Web Scraper Offline/Faulty. Variance logic bypassed. "
                f"Deploying emergency baseline: ${reconciled_price}/lb",
                file=sys.stderr
            )
            
            market_prices_workbook_data[ingredient] = {
                "unclamped": "DOWN (Offline)",
                "reconciled": reconciled_price
            }
            
        reconciled_estimates[ingredient] = reconciled_price

    csv_path = _resolve_log_path(data_dir)
    try:
        _append_estimate_log(csv_path, reconciled_estimates)
        print(f"Estimate log updated: {csv_path}")
    except (OSError, RuntimeError, pd.errors.ParserError, ValueError) as exc:
        print(f"Estimate logging failed: {exc}", file=sys.stderr)

    sku_records = []

    for sku, details in specs["skus"].items():
        if not isinstance(details, dict):
            continue

        recipe = details.get("recipe_weights_lbs")
        if not isinstance(recipe, dict):
            continue

        yields = details.get("yield_efficiencies", {})
        if not isinstance(yields, dict):
            yields = {}

        material_cost = 0.0
        for ingredient, weight in recipe.items():
            if not _is_finite_number(weight) or weight < 0:
                weight = 0.0

            market_estimate = reconciled_estimates.get(ingredient, 1.00)
            yield_fraction = yields.get(ingredient, 0.70)

            if not _is_finite_number(yield_fraction) or yield_fraction <= 0 or yield_fraction > 1:
                yield_fraction = 0.70

            material_cost += ((market_estimate / yield_fraction) * weight) * 0.75

        base_freight = details.get("nj_to_de_freight_per_unit", 0.12)
        packaging_cost = details.get("packaging_cost", 0.45)
        
        try:
            base_freight = float(base_freight)
            packaging_cost = float(packaging_cost)
        except (TypeError, ValueError):
            base_freight = 0.12
            packaging_cost = 0.45

        labor_cost_per_unit = 17.00 / 25.0
        allocated_toll = 15.00 / 500.0  
        total_logistics_cost = base_freight + allocated_toll

        landed_cost = material_cost + labor_cost_per_unit + packaging_cost + total_logistics_cost

        sku_records.append(
            {
                "SKU": sku,
                "Landed Cost": round(landed_cost, 2),
                "Wholesale Price": round(landed_cost * 1.35, 2),
                "Subscription Price": round(landed_cost * 1.50, 2),
            }
        )

    return pd.DataFrame(sku_records), market_prices_workbook_data
