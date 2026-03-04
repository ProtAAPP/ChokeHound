"""
Query registry system for managing security check queries.

This module provides a plugin-like system for registering and executing
security check queries. New queries can be easily added by creating
a query module and registering it.
"""

from typing import Dict, Callable, Optional
import pandas as pd


class SecurityQuery:
    """
    Base class for security check queries.
    
    Each security check query should inherit from this class and implement
    the required methods.
    """
    
    def __init__(self, name: str, description: str, cypher_query: str,
                 post_process: Optional[Callable[[pd.DataFrame], pd.DataFrame]] = None,
                 query_formatter: Optional[Callable[[], str]] = None):
        """
        Initialize a security query.
        
        Args:
            name: Display name for the query (used as sheet name in Excel)
            description: Description of what the query checks
            cypher_query: Cypher query string to execute (can be a template)
            post_process: Optional function to post-process the DataFrame
            query_formatter: Optional function that returns formatted query string
        """
        self.name = name
        self.description = description
        self.cypher_query = cypher_query
        self.post_process = post_process
        self.query_formatter = query_formatter
    
    def get_query(self) -> str:
        """Get the Cypher query string (formatted dynamically if formatter provided)."""
        if self.query_formatter:
            return self.query_formatter()
        return self.cypher_query
    
    def process_results(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Process query results.
        
        Args:
            df: Raw DataFrame from query execution
            
        Returns:
            Processed DataFrame
        """
        if self.post_process:
            return self.post_process(df)
        return df


class QueryRegistry:
    """
    Registry for managing security check queries.
    
    This class maintains a registry of all available security check queries
    and provides methods to register, retrieve, and execute them.
    """
    
    def __init__(self):
        """Initialize the query registry."""
        self._queries: Dict[str, SecurityQuery] = {}
    
    def register(self, query: SecurityQuery):
        """
        Register a security query.
        
        Args:
            query: SecurityQuery instance to register
        """
        if query.name in self._queries:
            print(f"[WARNING] Query '{query.name}' is already registered. Overwriting.")
        self._queries[query.name] = query
    
    def get_query(self, name: str) -> Optional[SecurityQuery]:
        """
        Get a registered query by name.
        
        Args:
            name: Query name
            
        Returns:
            SecurityQuery instance or None if not found
        """
        return self._queries.get(name)
    
    def get_all_queries(self) -> Dict[str, SecurityQuery]:
        """
        Get all registered queries.
        
        Returns:
            Dictionary mapping query names to SecurityQuery instances
        """
        return self._queries.copy()
    
    def get_query_names(self) -> list:
        """
        Get list of all registered query names.
        
        Returns:
            List of query names
        """
        return list(self._queries.keys())
    
    def get_queries_dict(self) -> Dict[str, str]:
        """
        Get dictionary mapping query names to Cypher queries.
        
        Returns:
            Dictionary mapping query names to query strings
        """
        return {name: query.get_query() for name, query in self._queries.items()}


# Global registry instance
_registry = QueryRegistry()


def get_registry() -> QueryRegistry:
    """
    Get the global query registry instance.
    
    Returns:
        QueryRegistry instance
    """
    return _registry


def register_query(query: SecurityQuery):
    """
    Register a query with the global registry.
    
    Args:
        query: SecurityQuery instance to register
    """
    _registry.register(query)



