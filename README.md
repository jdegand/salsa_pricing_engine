# Salsa Pricing Engine

An automated agribusiness cost-accounting and B2B pricing model. It scrapes live agricultural commodity markets, applies data-cleansing filters, evaluates landed unit economics, and generates client-ready corporate financial workbooks.
The engine protects fresh-food margins by assessing material yield efficiencies, freight disruptions, and market volatility inside a multi-tab financial ecosystem.

## Business Context

Fresh, refrigerated salsa commands premium clean-label margins but exhibits heavy operational fragility compared to mass-market pasteurized jarred goods. This engine models the financial realities of a manufacturing hub in southern New Jersey, distributing to commercial buyers and distribution centers across Delaware.

## Repository Architecture

```text
.
├── data
│   ├── historical_estimates.csv      # Appended time-series of reconciled commodity prices
│   └── product_specifications.json   # Structural SKU definitions, BOM weights, and yields
├── output
│   └── sams_salsa_pricing_matrix.xlsx # Generated multi-tab dynamic corporate workbook
├── README.md                         # Project documentation and execution guide
├── requirements.txt                  # Locked Python dependency track
└── src
    ├── data_pipeline.py              # Ingestion, validation, and SKU cost calculations
    ├── main.py                       # Orchestration pipeline entry point
    └── workbook_generator.py         # openpyxl spreadsheet formatting and formula injection
```

## Core Operational Phases

1. Data Ingestion & Reconciliation Engine (`data_pipeline.py`)

    * 3-Tier Weight Extraction: Scrapes data from live text feeds. It extracts package parameters (e.g., "25 lb carton") to calculate a true per-pound cost. If text patterns omit metrics, it falls back to industry-standard commercial wholesale crate weights.
    * Volatility Variance Clamping: Evaluates live scraped price data against a secondary baseline. If market variance exceeds 25%, the engine flags high volatility and clamps the price to a baseline safety target to eliminate tracking distortions.
    * Fail-Safe Bypassing: Employs pre-baked baseline emergency overrides if web data drops or connections time out, preventing pipeline execution blocks.
    * Historical Auditing: Logs all validated records in a centralized time-series tracking ledger.

1. Financial Modeling Matrix

    * Yield Efficiency Calibration: Adjusts raw ingredient weights to offset material loss during processing (e.g., trimming tomatillos or coring pineapples).
    * Production Logic: Factor inputs scale by a 0.75 weight adjustment factor, layered with standard manufacturing labor rates ($17.00 per 25-unit batch).
    * Logistics Allocation: Combines specific interstate freight rates with commercial bridge tolls spread across a standard 500-unit shipment.

1. Corporate Workbook Generation (`workbook_generator.py`) Tabs

    * Global Assumptions: Centrally isolates commercial levers like labor, markups, and loss allowances.
    * Market Inputs: Tracks live unclamped costs alongside reconciled manufacturing values.
    * Product Costing Matrix: Contains itemized Bill of Materials (BOM) cost ledgers linked to market inputs.
    * B2B Client Pricing Tiers: Details volume pricing structures, spoilage buffers, and portfolio margins.
    * Risk Sensitivity Matrix: Simulates a 2D matrix modeling gross profit fluctuations against material inflation and freight spikes.

## Technical & Business Assumptions

1. Raw Sourcing & Package De-bulking

    * Undeclared Container Weights: When a scraped listing fails to mention package sizes, the engine overrides the record with standardized, wholesale terminal market container weights:
        * Roma Tomatoes: 25.0 lb carton
        * Tomatillos: 35.0 lb bushel basket
        * Onions: 50.0 lb mesh sack
        * Jalapenos: 30.0 lb (1 1/9 bushel box)
        * Cilantro: 12.0 lb crate
        * Mangoes: 4.0 lb flat box
        * Pineapples: 5.0 lb premium bulk unit

1. Yield Efficiencies & Recipe Standardization

    * Processing Net Weights: All recipe fractions are normalized to a 1-pound manufacturing base yield rather than finished retail fluid ounces. This prevents downstream rounding distortions when resizing consumer retail packaging (e.g., converting formulas from 12oz cups to 16oz deli tubs).
    * BOM Weight Inflation: Raw ingredient weights are divided by their specific yield efficiencies (e.g., 0.75 for Roma Tomatoes, 0.50 for Pineapples) to calculate the raw purchasing volume required to yield the target clean weight.
    * Scaling Adjustment Factor: Total raw material costs are multiplied by a standardized processing scalar of 0.75 to simulate industrial batch condensation, density variance, and moisture adjustments occurring during commercial blending.

1. Cost-Accounting & Freight Overhead Allocation

    * Labor Overhead: Labor costs are fixed at $17.00 per 25-unit batch ($0.68 per retail unit).
    * Logistics & Infrastructure: Bridge tolls and freight parameters are distributed evenly across an entire commercial shipment. The engine adds a base per-unit interstate freight fee (e.g., $0.12–$0.14 depending on SKU category) to a co-allocated share of the $15.00 Delaware River commercial bridge toll, amortized over a standard 500-unit freight pallet ($0.03 per unit overhead).

## Developer Retrospective & Analytical Thoughts

1. Human-AI Collaborative Architecture

    This engine was engineered through an iterative, rapid prototyping process using AI to scaffold data schemas and generate complex boilerplate spreadsheet logic. My historical experience with openpyxl was critical in structuring the workbook's object models. Rather than relying blindly on AI-generated layout structures, I directed the code generation to enforce financial workbook design paradigms: explicitly isolating user inputs, eliminating hardcoded results, and enforcing uppercase, cell-coordinated Excel formulas (e.g., `=SUM()`, `=AVERAGE()`) to guarantee a professional dynamic workbook.

1. Strategic Scraping vs. API Key Dependency

    During the discovery phase, I attempted to source direct, public, keyless APIs for agricultural market data. However, free endpoints proved highly unstable or severely throttled. I opted to write a resilient scraper around `datajolt`, which mirrors public USDA terminal market reports without requiring complex environment variable configurations or authentication keys.
    While a production deployment would ideally integrate directly with the official USDA Market News API using formal authorization headers, this lightweight web scraping approach allowed for zero-config environmental setups—ideal for localized execution, testing, or rapid sandbox evaluation.

1. Data Inconsistency & Defensive Engineering

    Agricultural pricing streams are notoriously messy. Terminal market listings frequently cycle through varying nomenclature and display highly erratic packaging dimensions. To combat this, I designed a defensive tiered resolution engine:
    1. Regular Expressions: Isolate explicit weight patterns directly from live context strings.
    2. Hardcoded Domain Mappings: Instantly step in with standard wholesale crate weights if text is ambiguous.
    3. Volatility & Variance Filters: Cross-examine scraped prices against a statistical baseline. If a market anomaly spikes above a 25% tolerance threshold, the application dampens the shock by clamping the price back to an institutional benchmark.
    This defensive configuration protects client invoices from flash market spikes and prevents fatal division-by-zero pipeline failures.

1. Overcoming Business Logic Complexities

    The primary hurdle of this project was not the software engineering layer, but interpreting and structuring the agricultural cost-accounting model. Information regarding industrial fresh food yields and toll-road logistics can be highly fragmented.
    Early iterations hit design roadblocks because I initially bound the calculation formulas directly to a 12oz finished container footprint. This architectural choice broke the linearity of the pricing math, skewing material costs when ingredient pricing fluctuated. Re-architecting the engine to calculate costs using a strict 1-pound recipe mass framework, and then down-scaling to the final retail package size (12oz), resolved these systemic logic flaws and created an extensible platform capable of modeling any consumer packaging volume.

## Possible Future Enhancments

* Adapt the `reconciliation_pipeline` to weigh prices based on a Seasonal Index Matrix rather than a flat 25% volatility clamp.

## Getting Started

### Prerequisites

* Python 3.10 or higher
* Active internet connectivity (for real-time market data ingestion)

### Installation

1. Clone this repository and navigate into the project root directory:

    ```bash
    cd salsa_pricing_engine
    ```

1. Initialize and activate a localized virtual environment:

    ```bash
    python3 -m venv venv
    source venv/bin/activate
    ```

1. Install the locked project dependencies:

    ```bash
    pip3 install -r requirements.txt
    ```

### Execution

Run the main orchestration script to execute the end-to-end pipeline:

```bash
python3 src/main.py
```

Upon a successful runtime cycle, the engine prints a terminal-friendly summary of live single-unit product margins, appends a record to the time-series historical CSV log, and exports a newly calculated corporate spreadsheet matrix directly into the `output/` folder.
