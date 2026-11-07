import json
import openpyxl
import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

class ExcelStyleConfig:
    """Encapsulates corporate typography and formatting properties."""
    def __init__(self):
        font_family = "Segoe UI"
        self.title_font = Font(name=font_family, size=14, bold=True, color="0F1E36")
        self.header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
        self.bold_font = Font(name=font_family, size=11, bold=True)
        self.reg_font = Font(name=font_family, size=11)
        self.kpi_title_font = Font(name=font_family, size=9, bold=True, color="595959")
        self.kpi_value_font = Font(name=font_family, size=18, bold=True, color="0F1E36")
        
        self.navy_fill = PatternFill(fill_type="solid", start_color="0F1E36", end_color="0F1E36")
        self.zebra_fill = PatternFill(fill_type="solid", start_color="F8F9FA", end_color="F8F9FA")
        self.input_fill = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")
        self.kpi_fill = PatternFill(fill_type="solid", start_color="F2F4F7", end_color="F2F4F7")
        
        thin_side = Side(style='thin', color='D9D9D9')
        double_side = Side(style='double', color='000000')
        self.thick_side = Side(style='medium', color='0F1E36')
        
        self.grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        self.total_border = Border(top=thin_side, bottom=double_side)
        self.thin_side = thin_side


def build_global_assumptions_tab(wb, styles):
    """TAB 1: Global Commercial Assumptions. Returns cell maps to prevent hardcoding."""
    ws = wb.active
    ws.title = "1. Global Assumptions"
    ws.views.sheetView[0].showGridLines = True
    ws.cell(row=2, column=1, value="B2B Commercial Drivers & Levers").font = styles.title_font
    
    headers = ["Parameter Variable", "Value Data Input"]
    for c_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c_idx, value=text)
        cell.font = styles.header_font
        cell.fill = styles.navy_fill
        
    drivers = [
        ("Labor Cost / Batch Unit", 17.00 / 25.0, "$#,##0.00", "labor"),
        ("Commercial Bridge Toll", 15.00, "$#,##0.00", "toll"),
        ("Logistics Units / Shipment", 500, "#,##0", "shipment_units"),
        ("Wholesale Markup Factor", 1.35, "0.00x", "wholesale_markup"),
        ("Subscription Markup Factor", 1.50, "0.00x", "subscription_markup"),
        ("Spoilage Loss Allowance", 0.03, "0.0%", "spoilage_allowance")
    ]
    
    assumptions_map = {}
    for idx, (label, val, num_fmt, key) in enumerate(drivers, start=5):
        ws.cell(row=idx, column=1, value=label).font = styles.reg_font
        v_cell = ws.cell(row=idx, column=2, value=val)
        v_cell.font = styles.reg_font
        v_cell.fill = styles.input_fill
        v_cell.number_format = num_fmt
        
        ws.cell(row=idx, column=1).border = styles.grid_border
        v_cell.border = styles.grid_border
        assumptions_map[key] = f"'{ws.title}'!$B${idx}"
        
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 24
    return assumptions_map


def build_market_inputs_tab(wb, market_prices, styles):
    """TAB 2: Market Inputs Data Layer with comprehensive type validation."""
    ws = wb.create_sheet(title="2. Market Inputs")
    ws.views.sheetView[0].showGridLines = True
    ws.cell(row=2, column=1, value="Reconciled Agribusiness Commodity Inputs").font = styles.title_font
    
    headers = ["Commodity Ingredient", "Unclamped Raw Cost/Lb", "Reconciled Margin Price/Lb"]
    for c_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c_idx, value=text)
        cell.font = styles.header_font
        cell.fill = styles.navy_fill
        
    ing_row_map = {}
    if not isinstance(market_prices, dict):
        raise TypeError("Market prices must be provided as a dictionary object.")
        
    for idx, (ing, price_dict) in enumerate(market_prices.items(), start=5):
        if not isinstance(price_dict, dict) or "unclamped" not in price_dict or "reconciled" not in price_dict:
            raise KeyError(f"Ingredient structure missing target price entries for: {ing}")
            
        ws.cell(row=idx, column=1, value=ing).font = styles.reg_font
        
        for col_idx, (key, fill) in enumerate([("unclamped", styles.zebra_fill), ("reconciled", styles.input_fill)], start=2):
            val = price_dict[key]
            if not isinstance(val, (int, float)) or val < 0 or not np.isfinite(val):
                raise ValueError(f"Invalid market price valuation input ({val}) processed for ingredient: {ing}")
                
            cell = ws.cell(row=idx, column=col_idx, value=val)
            cell.font = styles.reg_font
            cell.fill = fill
            cell.number_format = '$#,##0.00'
            cell.border = styles.grid_border
            
        ws.cell(row=idx, column=1).border = styles.grid_border
        ing_row_map[ing] = f"'{ws.title}'!$C${idx}"
        
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 26
    ws.column_dimensions['C'].width = 28
    return ing_row_map


def build_product_costing_tab(wb, df_pricing, specs_data, ing_row_map, assumptions_map, styles):
    """TAB 3: Product Costing Matrix Ledger with strict input assertions."""
    ws = wb.create_sheet(title="3. Product Costing Matrix")
    ws.views.sheetView[0].showGridLines = True
    ws.cell(row=2, column=1, value="Landed Cost Accounting Ledger (12oz Units)").font = styles.title_font
    
    headers = ["Salsa Product SKU", "Material BOM Cost ($)", "Direct Labor ($)", "Packaging BOM ($)", "Freight Logistics ($)", "Total Landed Cost ($)"]
    for c_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c_idx, value=text)
        cell.font = styles.header_font
        cell.fill = styles.navy_fill
        
    costing_row_map = {}
    seen_skus = set()
    
    for seq_idx, (idx, row) in enumerate(df_pricing.iterrows(), start=5):
        sku_name = row['SKU']
        
        # Validates that the SKU value is present, not null, and not empty string
        if pd.isna(sku_name) or str(sku_name).strip() == "":
            raise ValueError(f"Blank or null SKU value encountered at DataFrame row index {idx}.")
            
        sku_name = str(sku_name).strip()
        if sku_name in seen_skus:
            raise ValueError(f"Duplicate SKU '{sku_name}' discovered. SKUs must be unique keys.")
        seen_skus.add(sku_name)
        
        sku_spec = specs_data.get(sku_name, {})
        if not isinstance(sku_spec, dict):
            raise TypeError(f"Specification configurations for SKU '{sku_name}' must be structured as a dictionary.")
            
        ws.cell(row=seq_idx, column=1, value=sku_name).font = styles.reg_font
        
        formula_parts = []
        recipe = sku_spec.get("recipe_weights_lbs", {})
        yields = sku_spec.get("yield_efficiencies", {})
        
        # Verifies that structural recipe fields map to valid nested dictionaries
        if not isinstance(recipe, dict):
            raise TypeError(f"recipe_weights_lbs configuration for SKU '{sku_name}' must be structured as a dictionary.")
        if not isinstance(yields, dict):
            raise TypeError(f"yield_efficiencies configuration for SKU '{sku_name}' must be structured as a dictionary.")
            
        if not recipe:
            raise ValueError(f"BOM recipe details are absent or empty for target SKU: {sku_name}")
            
        for ing, weight in recipe.items():
            # Enforces pure integer/float checks, explicitly rejecting boolean variables
            if isinstance(weight, bool) or not isinstance(weight, (int, float)) or weight < 0 or not np.isfinite(weight):
                raise ValueError(f"Invalid production weight requirement ({weight}) identified for ingredient '{ing}' in SKU '{sku_name}'.")
                
            ing_cell_ref = ing_row_map.get(ing)
            if not ing_cell_ref:
                raise KeyError(f"Ingredient '{ing}' required by SKU '{sku_name}' is absent from Market Inputs mapping.")
                
            y_eff = yields.get(ing, 0.70)
            if isinstance(y_eff, bool) or not isinstance(y_eff, (int, float)) or y_eff <= 0 or y_eff > 1.0:
                raise ValueError(f"Invalid yield efficiency allocation ({y_eff}) detected for ingredient '{ing}' in SKU '{sku_name}'.")
                
            formula_parts.append(f"(({ing_cell_ref}/{y_eff}) * {weight} * 0.75)")
            
        pack_cost = sku_spec.get("packaging_cost", 0.45)
        if isinstance(pack_cost, bool) or not isinstance(pack_cost, (int, float)) or pack_cost < 0 or not np.isfinite(pack_cost):
            raise ValueError(f"Malformed structural cost data encountered on packaging metrics for SKU: {sku_name}")
            
        spec_freight = sku_spec.get("nj_to_de_freight_per_unit", 0.12)
        if isinstance(spec_freight, bool) or not isinstance(spec_freight, (int, float)) or spec_freight < 0 or not np.isfinite(spec_freight):
            raise ValueError(f"Malformed structural freight cost configuration verified for SKU: {sku_name}")
            
        ws.cell(row=seq_idx, column=2, value=f"={'+'.join(formula_parts)}" if formula_parts else 0).number_format = '$#,##0.00'
        ws.cell(row=seq_idx, column=3, value=f"={assumptions_map['labor']}").number_format = '$#,##0.00'
        ws.cell(row=seq_idx, column=4, value=pack_cost).number_format = '$#,##0.00'
        ws.cell(row=seq_idx, column=5, value=f"={spec_freight}+({assumptions_map['toll']}/{assumptions_map['shipment_units']})").number_format = '$#,##0.00'
        ws.cell(row=seq_idx, column=6, value=f"=SUM(B{seq_idx}:E{seq_idx})").number_format = '$#,##0.00'
        
        for c in range(1, 7):
            cell = ws.cell(row=seq_idx, column=c)
            cell.border = styles.grid_border
            cell.font = styles.reg_font
            
        costing_row_map[sku_name] = {
            "total_landed": f"'{ws.title}'!$F${seq_idx}",
            "material_bom": f"'{ws.title}'!$B${seq_idx}",
            "direct_labor": f"'{ws.title}'!$C${seq_idx}",
            "packaging_bom": f"'{ws.title}'!$D${seq_idx}",
            "freight_logistics": f"'{ws.title}'!$E${seq_idx}",
            "row_num": seq_idx
        }
        
    widths = [34, 24, 20, 22, 24, 24]
    for c_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = w
    return costing_row_map


def build_client_pricing_tab(wb, df_pricing, costing_row_map, assumptions_map, styles):
    """TAB 4: B2B Client Pricing Tiers Framework."""
    ws = wb.create_sheet(title="4. B2B Client Pricing Tiers")
    ws.views.sheetView[0].showGridLines = True
    ws.cell(row=2, column=1, value="Commercial Distribution Pricing Framework").font = styles.title_font
    
    # Updated Header from 'Spoilage Allowance' to 'Spoilage Cost ($)' to reflect dollar metrics
    headers = ["Product SKU", "Landed Base Anchor", "Grocery Whls Price", "Spoilage Cost ($)", "Grocery Net Gross Profit ($)", "Subscription Price", "Subscription Net Gross Profit ($)"]
    for c_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c_idx, value=text)
        cell.font = styles.header_font
        cell.fill = styles.navy_fill
        
    pricing_row_map = {}
    last_data_row = 4
    
    for seq_idx, (idx, row) in enumerate(df_pricing.iterrows(), start=5):
        last_data_row = seq_idx
        sku_name = row['SKU']
        landed_ref = costing_row_map[sku_name]["total_landed"]
        
        ws.cell(row=seq_idx, column=1, value=sku_name).font = styles.reg_font
        ws.cell(row=seq_idx, column=2, value=f"={landed_ref}").number_format = '$#,##0.00'
        ws.cell(row=seq_idx, column=3, value=f"=B{seq_idx} * {assumptions_map['wholesale_markup']}").number_format = '$#,##0.00'
        ws.cell(row=seq_idx, column=4, value=f"=C{seq_idx} * {assumptions_map['spoilage_allowance']}").number_format = '$#,##0.00'
        ws.cell(row=seq_idx, column=5, value=f"=C{seq_idx}-B{seq_idx}-D{seq_idx}").number_format = '$#,##0.00'
        ws.cell(row=seq_idx, column=6, value=f"=B{seq_idx} * {assumptions_map['subscription_markup']}").number_format = '$#,##0.00'
        ws.cell(row=seq_idx, column=7, value=f"=F{seq_idx}-B{seq_idx}-(F{seq_idx} * {assumptions_map['spoilage_allowance']})").number_format = '$#,##0.00'
        
        for c in range(1, 8):
            cell = ws.cell(row=seq_idx, column=c)
            cell.border = styles.grid_border
            cell.font = styles.reg_font
            
        pricing_row_map[sku_name] = {
            "grocery_revenue": f"'{ws.title}'!$C${seq_idx}",
            "subscription_revenue": f"'{ws.title}'!$F${seq_idx}",
            "grocery_profit": f"'{ws.title}'!$E${seq_idx}",
            "subscription_profit": f"'{ws.title}'!$G${seq_idx}",
            "row_num": seq_idx
        }
        
    tot_row = last_data_row + 1
    ws.cell(row=tot_row, column=1, value="Portfolio Average Summary").font = styles.bold_font
    ws.cell(row=tot_row, column=1).border = styles.total_border
    
    for c in range(2, 8):
        col_letter = get_column_letter(c)
        formula_str = f"=AVERAGE({col_letter}5:{col_letter}{last_data_row})" if last_data_row >= 5 else 0
        f_cell = ws.cell(row=tot_row, column=c, value=formula_str)
        f_cell.font = styles.bold_font
        f_cell.number_format = '$#,##0.00'
        f_cell.border = styles.total_border
        
    widths = [34, 22, 22, 24, 28, 22, 32]
    for c_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = w
    return pricing_row_map, tot_row


def build_risk_matrices_tabs(wb, df_pricing, costing_row_map, pricing_row_map, assumptions_map, is_subscription, styles):
    """TABS 5 & 6: 2D Risk Simulation models mapped accurately per individual SKU asset lines."""
    title_str = "2D SKU-Specific Gross Dollar Profit Disruption Modeling (Subscriptions)" if is_subscription else "2D SKU-Specific Gross Dollar Profit Disruption Modeling"
    sheet_title = "6. Subscription Risk Matrix" if is_subscription else "5. Risk Sensitivity Matrix"
    
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    ws.cell(row=2, column=1, value=title_str).font = styles.title_font
    
    inflation_shifts = [1.00, 1.15, 1.30, 1.45]
    freight_surges = [1.00, 1.50, 2.00, 2.50]
    current_matrix_row = 5
    
    for seq_idx, (df_idx, row) in enumerate(df_pricing.iterrows()):
        sku_name = row['SKU']
        target_price_cell = pricing_row_map[sku_name]["subscription_revenue"] if is_subscription else pricing_row_map[sku_name]["grocery_revenue"]
        
        material_ref = costing_row_map[sku_name]["material_bom"]
        labor_ref = costing_row_map[sku_name]["direct_labor"]
        packaging_ref = costing_row_map[sku_name]["packaging_bom"]
        freight_ref = costing_row_map[sku_name]["freight_logistics"]
        
        ws.cell(row=current_matrix_row, column=1, value=f"SKU Matrix: {sku_name} (12oz)").font = styles.bold_font
        ws.cell(row=current_matrix_row + 1, column=2, value="Logistics Cost Multiplier ->").font = styles.bold_font
        ws.cell(row=current_matrix_row + 2, column=1, value="Material Inflation Multiplier").font = styles.bold_font
        
        for col_i, f_mult in enumerate(freight_surges, start=2):
            cell = ws.cell(row=current_matrix_row + 2, column=col_i, value=f_mult)
            cell.font = styles.bold_font
            cell.number_format = '0.0x'
            cell.fill = styles.zebra_fill
            cell.border = styles.grid_border
            
        for r_offset, inf_mult in enumerate(inflation_shifts):
            cell = ws.cell(row=current_matrix_row + 3 + r_offset, column=1, value=inf_mult)
            cell.font = styles.bold_font
            cell.number_format = '0.0x'
            cell.fill = styles.zebra_fill
            cell.border = styles.grid_border
            
        for r_offset, inf_mult in enumerate(inflation_shifts):
            grid_row_num = current_matrix_row + 3 + r_offset
            for col_i, f_mult in enumerate(freight_surges, start=2):
                col_letter = get_column_letter(col_i)
                axis_y_coordinate = f"$A{grid_row_num}"
                axis_x_coordinate = f"{col_letter}${current_matrix_row + 2}"
                
                # Unmasked formula allows errors to surface cleanly
                sim_formula = (
                    f"=(({target_price_cell} * (1-{assumptions_map['spoilage_allowance']})) - "
                    f"({material_ref} * {axis_y_coordinate} + {labor_ref} + {packaging_ref} + "
                    f"({freight_ref} * {axis_x_coordinate})))"
                )
                
                m_cell = ws.cell(row=grid_row_num, column=col_i, value=sim_formula)
                m_cell.number_format = '$#,##0.00'
                m_cell.font = styles.reg_font
                m_cell.border = styles.grid_border
                
        current_matrix_row += 10
        
    ws.column_dimensions['A'].width = 30
    for c in range(2, 6):
        ws.column_dimensions[get_column_letter(c)].width = 24


def build_executive_summary_tab(wb, num_skus, pricing_avg_row, styles):
    """TAB 7: Executive Summary Dashboard."""
    ws = wb.create_sheet(title="Summary")
    ws.views.sheetView[0].showGridLines = True
    ws.cell(row=2, column=1, value="Salsa Pricing Engine: Executive Summary").font = styles.title_font
    
    kpis = [
        ("AVG LANDED COST", f"='4. B2B Client Pricing Tiers'!$B${pricing_avg_row}", 2),
        ("AVG GROCERY PROFIT", f"='4. B2B Client Pricing Tiers'!$E${pricing_avg_row}", 5),
        ("AVG SUBSCRIPTION PROFIT", f"='4. B2B Client Pricing Tiers'!$G${pricing_avg_row}", 8)
    ]
    
    for title, formula, col_start in kpis:
        ws.merge_cells(start_row=4, start_column=col_start, end_row=4, end_column=col_start+1)
        ws.merge_cells(start_row=5, start_column=col_start, end_row=5, end_column=col_start+1)
        
        ws.cell(row=4, column=col_start, value=title).font = styles.kpi_title_font
        val_cell = ws.cell(row=5, column=col_start, value=formula)
        val_cell.font = styles.kpi_value_font
        val_cell.number_format = '$#,##0.00'
        
        # Range(4, 6) wraps styling uniformly across both titles and values
        for r in range(4, 6):
            for c in range(col_start, col_start + 2):
                cell = ws.cell(row=r, column=c)
                cell.fill = styles.kpi_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
        ws.cell(row=4, column=col_start).border = Border(left=styles.thin_side, right=styles.thin_side, top=styles.thick_side)
        ws.cell(row=4, column=col_start+1).border = Border(left=styles.thin_side, right=styles.thin_side, top=styles.thick_side)
        ws.cell(row=5, column=col_start).border = Border(left=styles.thin_side, right=styles.thin_side, bottom=styles.thin_side)
        ws.cell(row=5, column=col_start+1).border = Border(left=styles.thin_side, right=styles.thin_side, bottom=styles.thin_side)
        
    if num_skus > 0:
        chart = BarChart()
        chart.type = "bar"
        chart.style = 10
        chart.grouping = "stacked"
        chart.overlap = 100
        chart.title = "Landed Cost Structural Breakdown"
        chart.height = 14
        chart.width = 18
        chart.add_data(Reference(wb["3. Product Costing Matrix"], min_col=2, min_row=4, max_col=5, max_row=4 + num_skus), titles_from_data=True)
        chart.set_categories(Reference(wb["3. Product Costing Matrix"], min_col=1, min_row=5, max_row=4 + num_skus))
        ws.add_chart(chart, "B8")
        
    ws.column_dimensions['A'].width = 4
    for col in ['B', 'C', 'E', 'F', 'H', 'I']:
        ws.column_dimensions[col].width = 16
    ws.column_dimensions['D'].width = 4
    ws.column_dimensions['G'].width = 4


def create_corporate_workbook(pipeline_tuple, output_path, spec_path):
    """Main execution block orchestrating compilation sequence."""
    if not isinstance(pipeline_tuple, tuple) or len(pipeline_tuple) != 2:
        raise TypeError("The pipeline execution argument must be a valid 2-element tuple.")
        
    df_pricing, market_prices = pipeline_tuple
    
    # Asserting type checking protections across Dataframe objects 
    if not isinstance(df_pricing, pd.DataFrame):
        raise TypeError("Input parameter df_pricing must be structured as a pandas DataFrame.")
        
    if df_pricing.empty:
        raise ValueError("The incoming df_pricing DataFrame contains no viable data records.")
        
    # Hardens model against missing or malformed unique identifier fields
    if "SKU" not in df_pricing.columns:
        raise ValueError("The incoming df_pricing DataFrame must contain an 'SKU' column.")
        
    # Early validation check to ensure SKU array has no completely blank entries
    if df_pricing["SKU"].isna().any() or (df_pricing["SKU"].astype(str).str.strip() == "").any():
        raise ValueError("The df_pricing 'SKU' column contains null, NaN, or completely blank rows.")

    wb = openpyxl.Workbook()
    try:
        # Encompasses file input stream and generation sequence within the resource block
        with open(spec_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            specs_data = data.get("skus")
            
            # Prevents downstream mapping exceptions if the JSON structure is incorrect
            if not isinstance(specs_data, dict):
                raise TypeError("The specification JSON document must contain a valid 'skus' dictionary mapping.")
            
        styles = ExcelStyleConfig()
        
        assumptions_map = build_global_assumptions_tab(wb, styles)
        ing_row_map = build_market_inputs_tab(wb, market_prices, styles)
        costing_row_map = build_product_costing_tab(wb, df_pricing, specs_data, ing_row_map, assumptions_map, styles)
        pricing_row_map, pricing_avg_row = build_client_pricing_tab(wb, df_pricing, costing_row_map, assumptions_map, styles)
        
        build_risk_matrices_tabs(wb, df_pricing, costing_row_map, pricing_row_map, assumptions_map, is_subscription=False, styles=styles)
        build_risk_matrices_tabs(wb, df_pricing, costing_row_map, pricing_row_map, assumptions_map, is_subscription=True, styles=styles)
        build_executive_summary_tab(wb, len(df_pricing), pricing_avg_row, styles)
        
        wb.save(output_path)
    finally:
        # Guarantees that system handles are recycled cleanly even if data compilation breaks
        wb.close()
