import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

def create_corporate_workbook(pipeline_tuple, output_path, spec_path):
    df_pricing, market_prices = pipeline_tuple
    wb = openpyxl.Workbook()
    
    with open(spec_path, "r", encoding="utf-8") as f:
        specs_data = json.load(f).get("skus", {})
        
    font_family = "Segoe UI"
    title_font = Font(name=font_family, size=14, bold=True, color="0F1E36")
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    bold_font = Font(name=font_family, size=11, bold=True)
    reg_font = Font(name=font_family, size=11)

    # KPI blocks styling
    kpi_title_font = Font(name=font_family, size=9, bold=True, color="595959")
    kpi_value_font = Font(name=font_family, size=18, bold=True, color="0F1E36")
    
    navy_fill = PatternFill(fill_type="solid", start_color="0F1E36", end_color="0F1E36")
    zebra_fill = PatternFill(fill_type="solid", start_color="F8F9FA", end_color="F8F9FA")
    input_fill = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")
    
    thin_side = Side(style='thin', color='D9D9D9')
    double_side = Side(style='double', color='000000')
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    total_border = Border(top=thin_side, bottom=double_side)

    # ---------------------------------------------------------
    # TAB 1: 1. Global Commercial Assumptions
    # ---------------------------------------------------------
    ws_assumptions = wb.active
    ws_assumptions.title = "1. Global Assumptions"
    ws_assumptions.views.sheetView[0].showGridLines = True
    
    ws_assumptions.cell(row=2, column=1, value="B2B Commercial Drivers & Levers").font = title_font
    
    assumptions_headers = ["Parameter Variable", "Value Data Input"]
    for c_idx, text in enumerate(assumptions_headers, start=1):
        cell = ws_assumptions.cell(row=4, column=c_idx, value=text)
        cell.font = header_font
        cell.fill = navy_fill
        
    drivers = [
        ("Labor Cost / Batch Unit", 17.00 / 25.0, "$#,##0.00"),
        ("Commercial Bridge Toll", 15.00, "$#,##0.00"),
        ("Logistics Units / Shipment", 500, "#,##0"),
        ("Wholesale Markup Factor", 1.35, "0.00x"),
        ("Subscription Markup Factor", 1.50, "0.00x"),
        ("Spoilage Loss Allowance", 0.03, "0.0%")
    ]
    
    for idx, (label, val, num_fmt) in enumerate(drivers, start=5):
        ws_assumptions.cell(row=idx, column=1, value=label).font = reg_font
        v_cell = ws_assumptions.cell(row=idx, column=2, value=val)
        v_cell.font = reg_font
        v_cell.fill = input_fill
        v_cell.number_format = num_fmt
        ws_assumptions.cell(row=idx, column=1).border = grid_border
        v_cell.border = grid_border

    ws_assumptions.column_dimensions['A'].width = 32
    ws_assumptions.column_dimensions['B'].width = 24

    # ---------------------------------------------------------
    # TAB 2: 2. Market Inputs (Updated with Unclamped Column)
    # ---------------------------------------------------------
    ws_inputs = wb.create_sheet(title="2. Market Inputs")
    ws_inputs.views.sheetView[0].showGridLines = True
    
    ws_inputs.cell(row=2, column=1, value="Reconciled Agribusiness Commodity Inputs").font = title_font
    
    # Three-column layout to clearly track raw vs smoothed data layers
    headers_in = ["Commodity Ingredient", "Unclamped Raw Cost/Lb", "Reconciled Margin Price/Lb"]
    for c_idx, text in enumerate(headers_in, start=1):
        cell = ws_inputs.cell(row=4, column=c_idx, value=text)
        cell.font = header_font
        cell.fill = navy_fill
    
    ing_row_map = {}
    for idx, (ing, price_dict) in enumerate(market_prices.items(), start=5):
        ws_inputs.cell(row=idx, column=1, value=ing).font = reg_font
        
        # Column B: Live Unclamped Scraper Data
        u_cell = ws_inputs.cell(row=idx, column=2, value=price_dict["unclamped"])
        u_cell.font = reg_font
        u_cell.fill = zebra_fill # Clear background color tint differentiation
        u_cell.number_format = '$#,##0.00'
        u_cell.border = grid_border
        
        # Column C: Cleaned Reconciled Price used for manufacturing formulas
        r_cell = ws_inputs.cell(row=idx, column=3, value=price_dict["reconciled"])
        r_cell.font = reg_font
        r_cell.fill = input_fill 
        r_cell.number_format = '$#,##0.00'
        r_cell.border = grid_border
        
        ws_inputs.cell(row=idx, column=1).border = grid_border
        ing_row_map[ing] = idx

    ws_inputs.column_dimensions['A'].width = 28
    ws_inputs.column_dimensions['B'].width = 26
    ws_inputs.column_dimensions['C'].width = 28

    # ---------------------------------------------------------
    # TAB 3: 3. Product Costing Matrix (Updated Cell Coordinate Links)
    # ---------------------------------------------------------
    ws_costing = wb.create_sheet(title="3. Product Costing Matrix")
    ws_costing.views.sheetView[0].showGridLines = True
    
    ws_costing.cell(row=2, column=1, value="Landed Cost Accounting Ledger (12oz Units)").font = title_font
    
    headers_cost = ["Salsa Product SKU", "Material BOM Cost ($)", "Direct Labor ($)", "Packaging BOM ($)", "Freight Logistics ($)", "Total Landed Cost ($)"]
    for c_idx, text in enumerate(headers_cost, start=1):
        cell = ws_costing.cell(row=4, column=c_idx, value=text)
        cell.font = header_font
        cell.fill = navy_fill

    for seq_idx, (idx, row) in enumerate(df_pricing.iterrows(), start=5):
        r_num = seq_idx
        sku_name = row['SKU']
        sku_spec = specs_data.get(sku_name, {})
        
        ws_costing.cell(row=r_num, column=1, value=sku_name).font = reg_font
        
        formula_parts = []
        recipe = sku_spec.get("recipe_weights_lbs", {})
        yields = sku_spec.get("yield_efficiencies", {})
        for ing, weight in recipe.items():
            ing_row = ing_row_map.get(ing)
            y_eff = yields.get(ing, 0.70)
            # UPDATED: Direct cross-sheet links mapped specifically to Column C (Reconciled)
            formula_parts.append(f"(('2. Market Inputs'!$C${ing_row}/{y_eff})*{weight}*0.75)")
        
        ws_costing.cell(row=r_num, column=2, value=f"={'+'.join(formula_parts)}").number_format = '$#,##0.00'
        ws_costing.cell(row=r_num, column=3, value="='1. Global Assumptions'!$B$5").number_format = '$#,##0.00'  
        ws_costing.cell(row=r_num, column=4, value=sku_spec.get("packaging_cost", 0.45)).number_format = '$#,##0.00'
        
        spec_freight = sku_spec.get("nj_to_de_freight_per_unit", 0.12)
        ws_costing.cell(row=r_num, column=5, value=f"={spec_freight}+('1. Global Assumptions'!$B$6/'1. Global Assumptions'!$B$7)").number_format = '$#,##0.00'
        ws_costing.cell(row=r_num, column=6, value=f"=SUM(B{r_num}:E{r_num})").number_format = '$#,##0.00'
        
        for c in range(1, 6 + 1):
            cell = ws_costing.cell(row=r_num, column=c)
            cell.border = grid_border
            cell.font = reg_font

    ws_costing.column_dimensions['A'].width = 34
    ws_costing.column_dimensions['B'].width = 24
    ws_costing.column_dimensions['C'].width = 20
    ws_costing.column_dimensions['D'].width = 22
    ws_costing.column_dimensions['E'].width = 24
    ws_costing.column_dimensions['F'].width = 24


    # ---------------------------------------------------------
    # TAB 4: 4. B2B Client Pricing Tiers
    # ---------------------------------------------------------
    ws_pricing = wb.create_sheet(title="4. B2B Client Pricing Tiers")
    ws_pricing.views.sheetView[0].showGridLines = True
    
    ws_pricing.cell(row=2, column=1, value="Commercial Distribution Pricing Framework").font = title_font
    
    headers_price = ["Product SKU", "Landed Base Anchor", "Grocery Whls Price", "Spoilage Allowance", "Grocery Net Gross Profit ($)", "Subscription Price", "Subscription Net Gross Profit ($)"]
    for c_idx, text in enumerate(headers_price, start=1):
        cell = ws_pricing.cell(row=4, column=c_idx, value=text)
        cell.font = header_font
        cell.fill = navy_fill

    last_data_row = 4
    for seq_idx, (idx, row) in enumerate(df_pricing.iterrows(), start=5):
        r_num = seq_idx
        last_data_row = r_num
        ws_pricing.cell(row=r_num, column=1, value=row['SKU']).font = reg_font
        
        ws_pricing.cell(row=r_num, column=2, value=f"='3. Product Costing Matrix'!F{r_num}").number_format = '$#,##0.00'
        ws_pricing.cell(row=r_num, column=3, value=f"=B{r_num}*'1. Global Assumptions'!$B$8").number_format = '$#,##0.00'
        ws_pricing.cell(row=r_num, column=4, value=f"=C{r_num}*'1. Global Assumptions'!$B$10").number_format = '$#,##0.00'
        ws_pricing.cell(row=r_num, column=5, value=f"=C{r_num}-B{r_num}-D{r_num}").number_format = '$#,##0.00'
        ws_pricing.cell(row=r_num, column=6, value=f"=B{r_num}*'1. Global Assumptions'!$B$9").number_format = '$#,##0.00'

        # FIXED: Subscription Net Gross Profit formula now subtracts the spoilage allocation (Price * Spoilage Rate)
        ws_pricing.cell(row=r_num, column=7, value=f"=F{r_num}-B{r_num}-(F{r_num}*'1. Global Assumptions'!$B$10)").number_format = '$#,##0.00'
        
        
        for c in range(1, 8):
            cell = ws_pricing.cell(row=r_num, column=c)
            cell.border = grid_border
            cell.font = reg_font

    tot_row = last_data_row + 1
    ws_pricing.cell(row=tot_row, column=1, value="Portfolio Average Summary").font = bold_font
    ws_pricing.cell(row=tot_row, column=1).border = total_border
    
    for c in range(2, 8):
        col_letter = get_column_letter(c)
        f_cell = ws_pricing.cell(row=tot_row, column=c, value=f"=AVERAGE({col_letter}5:{col_letter}{last_data_row})")
        f_cell.font = bold_font
        f_cell.number_format = '$#,##0.00'
        f_cell.border = total_border

    ws_pricing.column_dimensions['A'].width = 34
    ws_pricing.column_dimensions['B'].width = 22
    ws_pricing.column_dimensions['C'].width = 22
    ws_pricing.column_dimensions['D'].width = 24
    ws_pricing.column_dimensions['E'].width = 28
    ws_pricing.column_dimensions['F'].width = 22
    ws_pricing.column_dimensions['G'].width = 32

    # ---------------------------------------------------------
    # TAB 5: 5. Risk Sensitivity Matrix (Fully Corrected & Dynamic)
    # ---------------------------------------------------------
    ws_matrix = wb.create_sheet(title="5. Risk Sensitivity Matrix")
    ws_matrix.views.sheetView[0].showGridLines = True
    
    ws_matrix.cell(row=2, column=1, value="2D Portfolio Gross Profit Disruption Modeling").font = title_font
    
    yield_shifts = [1.00, 0.85, 0.70, 0.55] 
    freight_surges = [1.00, 1.50, 2.00, 2.50] 
    
    # We begin writing tables sequentially down the sheet starting at row 5
    current_matrix_row = 5

    # Iterate through every real SKU in your dataset to give each its own simulation grid
    for seq_idx, (df_idx, row) in enumerate(df_pricing.iterrows()):
        sku_name = row['SKU']
        
        # Calculate exactly which row this specific SKU occupies on Tabs 3 and 4
        # Since loop data rows on both tabs start at row 5, the math is: 5 + sequential_index
        target_data_row = 5 + seq_idx
        
        # Section Header for the specific SKU Matrix
        ws_matrix.cell(row=current_matrix_row, column=1, value=f"SKU Matrix: {sku_name} (12oz)").font = bold_font
        
        # Write Axis Labels
        ws_matrix.cell(row=current_matrix_row + 1, column=2, value="Logistics Cost Multiplier ->").font = bold_font
        ws_matrix.cell(row=current_matrix_row + 2, column=1, value="Material Inflation Multiplier").font = bold_font

        # Populate Horizontal X-Axis (Freight Surges)
        for col_i, f_mult in enumerate(freight_surges, start=2):
            cell = ws_matrix.cell(row=current_matrix_row + 2, column=col_i, value=f_mult)
            cell.font = bold_font
            cell.number_format = '0.0x'
            cell.fill = zebra_fill
            cell.border = grid_border

        # Populate Vertical Y-Axis (Yield Shifts)
        for r_offset, y_mult in enumerate(yield_shifts):
            cell = ws_matrix.cell(row=current_matrix_row + 3 + r_offset, column=1, value=y_mult)
            cell.font = bold_font
            cell.number_format = '0.0x'
            cell.fill = zebra_fill
            cell.border = grid_border

        # Build Nested Grid Formulation tracking back specifically to target_data_row coordinates
        for r_offset, y_mult in enumerate(yield_shifts):
            grid_row_num = current_matrix_row + 3 + r_offset
            
            for col_i, f_mult in enumerate(freight_surges, start=2):
                col_letter = get_column_letter(col_i)
                axis_y_coordinate = f"$A{grid_row_num}"
                axis_x_coordinate = f"{col_letter}${current_matrix_row + 2}"
                
                sim_formula = (
                    f"=IFERROR((('4. B2B Client Pricing Tiers'!$C${target_data_row} * (1-'1. Global Assumptions'!$B$10)) - "
                    f"(('3. Product Costing Matrix'!$B${target_data_row} * {axis_y_coordinate}) + "
                    f"'3. Product Costing Matrix'!$C${target_data_row} + "
                    f"'3. Product Costing Matrix'!$D${target_data_row} + "
                    f"('3. Product Costing Matrix'!$E${target_data_row} * {axis_x_coordinate}))) / "
                    f"('4. B2B Client Pricing Tiers'!$C${target_data_row} * (1-'1. Global Assumptions'!$B$10)), 0)"
                )

                m_cell = ws_matrix.cell(row=grid_row_num, column=col_i, value=sim_formula)
                m_cell.number_format = '0.0%'
                m_cell.font = reg_font
                m_cell.border = grid_border

        # Advance down the sheet to leave ample space for the next SKU matrix block
        # (Header/Labels = 3 rows) + (Data grid = 4 rows) + (Buffer spacing = 3 rows) = 10 row step down
        current_matrix_row += 10

    # Auto-adjust column layouts cleanly across all matrix dimensions
    ws_matrix.column_dimensions['A'].width = 30
    for c in range(2, 6):
        ws_matrix.column_dimensions[get_column_letter(c)].width = 24

    # ---------------------------------------------------------
    # TAB 6: 6. Subscription Risk Matrix
    # ---------------------------------------------------------
    ws_sub_matrix = wb.create_sheet(title="6. Subscription Risk Matrix")
    ws_sub_matrix.views.sheetView[0].showGridLines = True
    
    ws_sub_matrix.cell(row=2, column=1, value="2D Portfolio Gross Profit Disruption Modeling (Subscriptions)").font = title_font
    
    current_sub_matrix_row = 5

    for seq_idx, (df_idx, row) in enumerate(df_pricing.iterrows()):
        sku_name = row['SKU']
        target_data_row = 5 + seq_idx
        
        ws_sub_matrix.cell(row=current_sub_matrix_row, column=1, value=f"SKU Matrix: {sku_name} (12oz)").font = bold_font
        ws_sub_matrix.cell(row=current_sub_matrix_row + 1, column=2, value="Logistics Cost Multiplier ->").font = bold_font
        ws_sub_matrix.cell(row=current_sub_matrix_row + 2, column=1, value="Material Inflation Multiplier").font = bold_font

        for col_i, f_mult in enumerate(freight_surges, start=2):
            cell = ws_sub_matrix.cell(row=current_sub_matrix_row + 2, column=col_i, value=f_mult)
            cell.font = bold_font
            cell.number_format = '0.0x'
            cell.fill = zebra_fill
            cell.border = grid_border

        for r_offset, y_mult in enumerate(yield_shifts):
            cell = ws_sub_matrix.cell(row=current_sub_matrix_row + 3 + r_offset, column=1, value=y_mult)
            cell.font = bold_font
            cell.number_format = '0.0x'
            cell.fill = zebra_fill
            cell.border = grid_border

        for r_offset, y_mult in enumerate(yield_shifts):
            grid_row_num = current_sub_matrix_row + 3 + r_offset
            
            for col_i, f_mult in enumerate(freight_surges, start=2):
                col_letter = get_column_letter(col_i)
                axis_y_coordinate = f"$A{grid_row_num}"
                axis_x_coordinate = f"{col_letter}${current_matrix_row + 2}"
                
                # ENHANCED: Targets Column F (Subscription Price) and handles division safety
                sim_formula = (
                    f"=IFERROR((('4. B2B Client Pricing Tiers'!$F${target_data_row} * (1-'1. Global Assumptions'!$B$10)) - "
                    f"(('3. Product Costing Matrix'!$B${target_data_row} * {axis_y_coordinate}) + "
                    f"'3. Product Costing Matrix'!$C${target_data_row} + "
                    f"'3. Product Costing Matrix'!$D${target_data_row} + "
                    f"('3. Product Costing Matrix'!$E${target_data_row} * {axis_x_coordinate}))) / "
                    f"('4. B2B Client Pricing Tiers'!$F${target_data_row} * (1-'1. Global Assumptions'!$B$10)), 0)"
                )

                m_cell = ws_sub_matrix.cell(row=grid_row_num, column=col_i, value=sim_formula)
                m_cell.number_format = '0.0%'
                m_cell.font = reg_font
                m_cell.border = grid_border

        current_sub_matrix_row += 10

    ws_sub_matrix.column_dimensions['A'].width = 30
    for c in range(2, 6):
        ws_sub_matrix.column_dimensions[get_column_letter(c)].width = 24


    # ---------------------------------------------------------
    # TAB 7: Executive Summary
    # ---------------------------------------------------------
    ws_dash = wb.create_sheet(title="Summary")
    ws_dash.views.sheetView[0].showGridLines = True
    
    ws_dash.cell(row=2, column=1, value="Salsa Pricing Engine: Executive Summary").font = title_font
    
    # Calculate dataset metrics locally to support the summary tab
    num_skus = len(df_pricing)
    tot_row = 5 + num_skus
    
    # Re-initialize summary specific fills
    kpi_fill = PatternFill(fill_type="solid", start_color="F2F4F7", end_color="F2F4F7")
    thick_side = Side(style='medium', color='0F1E36')
    
    # KPI 1 Block: Portfolio Average Landed Cost (Direct Calculation)
    ws_dash.merge_cells("B4:C4")
    ws_dash.merge_cells("B5:C5")
    ws_dash.cell(row=4, column=2, value="AVG LANDED COST").font = kpi_title_font
    c1 = ws_dash.cell(row=5, column=2, value=f"=AVERAGE('3. Product Costing Matrix'!F5:F{4 + num_skus})")
    c1.font = kpi_value_font
    c1.number_format = '$#,##0.00'
    
    # KPI 2 Block: Portfolio Average Wholesale Profit (Direct Calculation)
    ws_dash.merge_cells("E4:F4")
    ws_dash.merge_cells("E5:F5")
    ws_dash.cell(row=4, column=5, value="AVG GROCERY PROFIT").font = kpi_title_font
    c2 = ws_dash.cell(row=5, column=5, value=f"=AVERAGE('4. B2B Client Pricing Tiers'!E5:E{4 + num_skus})")
    c2.font = kpi_value_font
    c2.number_format = '$#,##0.00'
    
    # KPI 3 Block: Portfolio Average Subscription Profit (Direct Calculation)
    ws_dash.merge_cells("H4:I4")
    ws_dash.merge_cells("H5:I5")
    ws_dash.cell(row=4, column=8, value="AVG SUBSCRIPTION PROFIT").font = kpi_title_font
    c3 = ws_dash.cell(row=5, column=8, value=f"=AVERAGE('4. B2B Client Pricing Tiers'!G5:G{4 + num_skus})")
    c3.font = kpi_value_font
    c3.number_format = '$#,##0.00'
    
    # Apply KPI card aesthetic fills & alignment
    for r in range(4, 6):
        for col_idx in (2, 3, 5, 6, 8, 9):
            cell = ws_dash.cell(row=r, column=col_idx)
            cell.fill = kpi_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
    # Apply KPI borders cleanly
    for col_start in (2, 5, 8):
        ws_dash.cell(row=4, column=col_start).border = Border(left=thin_side, right=thin_side, top=thick_side)
        ws_dash.cell(row=4, column=col_start+1).border = Border(left=thin_side, right=thin_side, top=thick_side)
        ws_dash.cell(row=5, column=col_start).border = Border(left=thin_side, right=thin_side, bottom=thin_side)
        ws_dash.cell(row=5, column=col_start+1).border = Border(left=thin_side, right=thin_side, bottom=thin_side)

    # Inject Cost Structure Stacked Chart Layer
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Landed Cost Structural Breakdown"
    chart.height = 14
    chart.width = 18
    
    # Grab cost buckets data series from Tab 3 (Cols B through E)
    data = Reference(wb["3. Product Costing Matrix"], min_col=2, min_row=4, max_col=5, max_row=4 + num_skus)
    # Grab SKU Names for Y-Axis labels (Col A)
    cats = Reference(wb["3. Product Costing Matrix"], min_col=1, min_row=5, max_row=4 + num_skus)
    
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws_dash.add_chart(chart, "B8")

    ws_dash.column_dimensions['A'].width = 4
    ws_dash.column_dimensions['B'].width = 16
    ws_dash.column_dimensions['C'].width = 16

    try:
        wb.save(output_path)
    finally:
        wb.close()
