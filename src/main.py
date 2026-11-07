import os
import sys
from data_pipeline import run_reconciliation_pipeline
from workbook_generator import create_corporate_workbook

def main():
    """
    Main orchestration entry point for the Salsa Costing & Pricing Engine.
    Executes data collection, pipeline reconciliation, and corporate workbook generation.
    """
    print("=" * 65)
    print("STARTING: Automated B2B Fresh Foods Costing & Pricing Engine")
    print("=" * 65)
    
    # ---------------------------------------------------------
    # 1. Path Resolutions & Directory Safety Checks
    # ---------------------------------------------------------
    # Resolve the base repository root path
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    spec_path = os.path.join(base_dir, "data", "product_specifications.json")
    output_dir = os.path.join(base_dir, "output")
    output_workbook_path = os.path.join(output_dir, "sams_salsa_pricing_matrix.xlsx")
    
    # Ensure directory architectures exist in the environment (e.g., GitHub Codespace)
    if not os.path.exists(output_dir):
        print(f"Creating missing directory: {output_dir}")
        os.makedirs(output_dir)
        
    if not os.path.exists(spec_path):
        print(f"CRITICAL ERROR: Baseline specification file not found at {spec_path}")
        print("Please build your data/product_specifications.json file first.")
        sys.exit(1)

    # ---------------------------------------------------------
    # 2. Run Data Ingestion & Yield Processing Pipeline
    # ---------------------------------------------------------
    print("\n[Phase 1/3] Triggering Data Ingestion Pipelines...")
    print(" -> Scraping public agricultural text summaries...")
    print(" -> Querying secondary keyless Open Food Facts API benchmarks...")
    
    try:
        # Run reconciliation mechanism via data_pipeline module
        df_pricing_matrix, market_prices = run_reconciliation_pipeline(spec_path)
        
        print("\n[Phase 2/3] Data Pipeline Executed Successfully.")
        print("-" * 50)
        print(df_pricing_matrix.to_string(index=False))
        print("-" * 50)
        
    except Exception as e:
        print(f"\nCRITICAL PIPELINE EXCEPTION UNHANDLED: {e}")
        sys.exit(1)

    # ---------------------------------------------------------
    # 3. Trigger openpyxl Corporate Spreadsheet Generator
    # ---------------------------------------------------------
    print("\n[Phase 3/3] Compiling Corporate Spreadsheet Infrastructure...")
    print(" -> Formatting workspace elements with Classic Navy theme rules...")
    print(" -> Injecting dynamic, capitalized cross-tab formula schemas...")
    print(" -> Executing defensive text truncation safety adjustments...")
    
    try:
        # Pass clean dataframe along with specification path path down to layout engine
        create_corporate_workbook((df_pricing_matrix, market_prices), output_workbook_path, spec_path)
        
        print("\n" + "=" * 65)
        print("SUCCESS: End-to-End Workbook Compilation Completed!")
        print(f"Location: {output_workbook_path}")
        print("=" * 65)
        
    except Exception as e:
        print(f"\nCRITICAL WORKBOOK EXCEPTION UNHANDLED: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
