"""
Query executor module.

Handles execution of security check queries against Neo4j.
"""

from neo4j import Driver
import pandas as pd
from typing import Dict, Any, Optional
from chokehound.utils.label_processor import simplify_labels, process_dataframe_labels
import chokehound.config.settings as config


class QueryExecutor:
    """Executes security check queries against Neo4j."""
    
    def __init__(self, driver: Driver):
        """
        Initialize query executor.
        
        Args:
            driver: Neo4j Driver object
        """
        self.driver = driver
    
    def execute_query(self, query: str, query_name: str, timeout: Optional[int] = None) -> pd.DataFrame:
        """
        Execute a Cypher query and return results as DataFrame.
        
        Args:
            query: Cypher query string
            query_name: Name of the query (for error reporting)
            timeout: Query timeout in seconds (defaults to config.QUERY_TIMEOUT_SECONDS)
            
        Returns:
            pandas DataFrame with query results
        """
        if timeout is None:
            # Use Azure timeout for Azure queries, default timeout for others
            if "Azure" in query_name:
                timeout = config.AZURE_QUERY_TIMEOUT_SECONDS
            else:
                timeout = config.QUERY_TIMEOUT_SECONDS
        
        try:
            with self.driver.session() as session:
                # Run query with timeout
                result = session.run(query, timeout=timeout)
                
                # Stream results into list (efficient memory usage)
                records = []
                for record in result:
                    # Convert Record to dict
                    records.append(record.data())
                
                # Convert to DataFrame
                if not records:
                    df = pd.DataFrame([{"Info": "No results found"}])
                else:
                    df = pd.DataFrame(records)
                    # Process labels to simplify them
                    df = process_dataframe_labels(df)
                
                return df
                
        except Exception as e:
            print(f"  [WARNING] Error running '{query_name}': {e}")
            return pd.DataFrame([{"Error": str(e)}])
    
    def execute_queries(self, queries: Dict[str, str]) -> Dict[str, pd.DataFrame]:
        """
        Execute multiple queries and return results.
        
        Args:
            queries: Dictionary mapping query names to Cypher queries
            
        Returns:
            Dictionary mapping query names to DataFrames
        """
        results = {}
        for query_name, cypher_query in queries.items():
            print(f"Running query: {query_name}")
            df = self.execute_query(cypher_query, query_name)
            results[query_name] = df
            print(f"  [OK] {len(df)} rows returned")
        return results



