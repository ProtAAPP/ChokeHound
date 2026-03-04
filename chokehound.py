"""
ChokeHound - Tier 0 Choke Points Analysis Tool

A specialized tool to identify and analyze Tier 0 choke points in Active Directory
and Azure/Entra ID environments from data collected by BloodHound.

Usage:
    python chokehound.py [--output OUTPUT_FILENAME] [--log-risk] [--skip-risk-calculation]
    
Examples:
    python chokehound.py --output my_security_report.xlsx
    python chokehound.py --skip-risk-calculation  # Faster execution without risk scoring
"""

import argparse

# Import core modules
from chokehound.core.database import DatabaseConnection
from chokehound.core.query_executor import QueryExecutor
from chokehound.queries.registry import get_registry
from chokehound.reporting.excel_report import ExcelReportGenerator

# Import query modules to register them - only choke points
from chokehound.queries import choke_points  # noqa: F401

import chokehound.config.settings as config


def banner():
    """Print ChokeHound banner."""
    font = r"""
   ____ _           _        _   _                       _ 
  / ___| |__   ___ | | _____| | | | ___  _   _ _ __   __| |
 | |   | '_ \ / _ \| |/ / _ \ |_| |/ _ \| | | | '_ \ / _` |
 | |___| | | | (_) |   <  __/  _  | (_) | |_| | | | | (_| |
  \____|_| |_|\___/|_|\_\___|_| |_|\___/ \__,_|_| |_|\__,_|
                                      
    Tier 0 Choke Points Analysis Tool
    v1.1.0
    """
    print(font)


def main():
    """Main function to execute queries and generate Excel report."""
    parser = argparse.ArgumentParser(
        description="Analyze Tier 0 choke points in Active Directory and Azure from BloodHound data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python chokehound.py
  python chokehound.py --output my_report.xlsx
  python chokehound.py --output my_report.xlsx --log-risk
  python chokehound.py --skip-risk-calculation
  python chokehound.py -o report.xlsx --skip-risk-calculation
        """
    )
    parser.add_argument(
        '--output', '-o',
        type=str,
        default=config.DEFAULT_OUTPUT_FILENAME,
        help=f'Output Excel filename (default: {config.DEFAULT_OUTPUT_FILENAME})'
    )
    parser.add_argument(
        '--log-risk',
        action='store_true',
        default=False,
        help='Generate a detailed log file explaining risk calculations for choke points (includes raw risk scores)'
    )
    parser.add_argument(
        '--skip-risk-calculation',
        action='store_true',
        default=False,
        help='Skip risk score calculation for choke points (faster execution, results sorted by affected attack paths)'
    )
    args = parser.parse_args()
    
    output_filename = args.output
    enable_logging = args.log_risk
    skip_risk_calculation = args.skip_risk_calculation
    
    # Validate conflicting arguments
    if enable_logging and skip_risk_calculation:
        print("[ERROR] Cannot use --log-risk and --skip-risk-calculation together.")
        print("        --log-risk requires risk calculation to generate the log.")
        return
    
    # Ensure .xlsx extension
    if not output_filename.endswith('.xlsx'):
        output_filename += '.xlsx'
    
    # Connect to Neo4j
    print(f"Connecting to Neo4j at {config.NEO4J_URI}...")
    try:
        db = DatabaseConnection()
        driver = db.connect()
        print("[OK] Connected to Neo4j")
    except Exception as e:
        print(f"[ERROR] Error connecting to Neo4j: {e}")
        return
    
    try:
        # Check for Active Directory data
        print("Checking for Active Directory data...")
        has_ad = db.has_ad_data()
        domains_detailed = []
        if has_ad:
            domains_detailed = db.get_domains_detailed()
            print(f"  [OK] Found {len(domains_detailed)} Active Directory domain(s):")
            for domain in domains_detailed:
                print(f"       - {domain['name']} (ID: {domain['objectid']})")
        else:
            print("  [INFO] No Active Directory data found in database")
        
        # Check for Azure data
        print("\nChecking for Azure data...")
        has_azure = db.has_azure_data()
        tenants = []
        if has_azure:
            tenants = db.get_tenants()
            print(f"  [OK] Found {len(tenants)} Azure tenant(s):")
            for tenant in tenants:
                print(f"       - {tenant['name']} (ID: {tenant['objectid']})")
        else:
            print("  [INFO] No Azure data found in database")
        
        # Check if we have any data at all
        if not has_ad and not has_azure:
            print("\n[ERROR] No Active Directory or Azure data found in database.")
            print("        Please ensure BloodHound data is imported before running ChokeHound.")
            return
        
        # Get registered queries
        registry = get_registry()
        all_queries = registry.get_queries_dict()
        
        if not all_queries:
            print("[ERROR] No security queries registered. Please check query modules.")
            return
        
        # Filter queries based on available data
        queries_dict = {}
        for query_name, cypher_query in all_queries.items():
            # Include AD queries only if AD data exists
            if "AD " in query_name or "Active Directory" in query_name:
                if has_ad:
                    queries_dict[query_name] = cypher_query
                else:
                    print(f"[INFO] Skipping '{query_name}' - No AD data in database")
            # Include Azure queries only if Azure data exists
            elif "Azure" in query_name or "AZ " in query_name:
                if has_azure:
                    queries_dict[query_name] = cypher_query
                else:
                    print(f"[INFO] Skipping '{query_name}' - No Azure data in database")
            # Include other queries by default
            else:
                queries_dict[query_name] = cypher_query
        
        if not queries_dict:
            print("\n[ERROR] No applicable security queries to execute.")
            return
        
        print(f"\nExecuting {len(queries_dict)} security check(s)")
        
        # Show execution mode
        if skip_risk_calculation:
            print("[INFO] Running in fast mode: Risk calculation disabled")
            print("       Results will be sorted by affected attack paths\n")
        elif enable_logging:
            print("[INFO] Running with detailed risk logging enabled\n")
        else:
            print("[INFO] Running with risk calculation enabled\n")
        
        # Execute queries
        executor = QueryExecutor(driver)
        dataframes = {}
        risk_breakdowns = []
        query_descriptions = {}
        query_objects = {}
        
        for query_name, cypher_query in queries_dict.items():
            print(f"Running query: {query_name}")
            df = executor.execute_query(cypher_query, query_name)
            
            # Get query object for post-processing
            query_obj = registry.get_query(query_name)
            if query_obj:
                query_descriptions[query_name] = query_obj.description
                query_objects[query_name] = query_obj
                
                # Handle special case for choke points risk query
                if query_name == "AD Tier 0 Choke Points Risk":
                    if skip_risk_calculation:
                        # Skip risk calculation, use simple post-processing
                        from chokehound.queries.choke_points import post_process_choke_points_no_risk
                        df = post_process_choke_points_no_risk(df)
                    elif enable_logging:
                        # Calculate risk with detailed logging
                        from chokehound.queries.choke_points import post_process_choke_points_risk
                        df, breakdowns = post_process_choke_points_risk(df, enable_logging=True)
                        risk_breakdowns = breakdowns
                    else:
                        # Normal risk calculation
                        df = query_obj.process_results(df)
                else:
                    df = query_obj.process_results(df)
            else:
                query_descriptions[query_name] = "Security check query"
            
            dataframes[query_name] = df
            
            # Check if we have actual results or just placeholder messages
            if df.empty or 'Info' in df.columns or 'Error' in df.columns:
                print(f"  [INFO] No results found")
            else:
                print(f"  [OK] {len(df)} rows returned")
        
        # Generate Excel report
        report_generator = ExcelReportGenerator(output_filename, domains_detailed, tenants)
        report_generator.generate(
            dataframes=dataframes,
            query_descriptions=query_descriptions,
            query_objects=query_objects,
            risk_breakdowns=risk_breakdowns if enable_logging else None,
            enable_logging=enable_logging
        )
    
    finally:
        # Always close the driver connection
        print("\nClosing database connection...")
        db.close()
        print("[OK] Connection closed")


if __name__ == "__main__":
    banner()
    main()

