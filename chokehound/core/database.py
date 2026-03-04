"""
Database connection module for Neo4j.

Handles connection to BloodHound Neo4j database.
"""

from neo4j import GraphDatabase, Driver
from typing import Optional, List, Dict
import chokehound.config.settings as config


class DatabaseConnection:
    """Manages Neo4j database connection."""
    
    def __init__(self, uri: Optional[str] = None, user: Optional[str] = None, 
                 password: Optional[str] = None):
        """
        Initialize database connection.
        
        Args:
            uri: Neo4j URI (defaults to config.NEO4J_URI)
            user: Neo4j username (defaults to config.NEO4J_USER)
            password: Neo4j password (defaults to config.NEO4J_PASSWORD)
        """
        self.uri = uri or config.NEO4J_URI
        self.user = user or config.NEO4J_USER
        self.password = password or config.NEO4J_PASSWORD
        self.driver: Optional[Driver] = None
    
    def connect(self) -> Driver:
        """
        Connect to Neo4j database using official driver.
        
        Returns:
            Neo4j Driver object
            
        Raises:
            Exception: If connection fails
        """
        try:
            self.driver = GraphDatabase.driver(
                self.uri, 
                auth=(self.user, self.password),
                max_connection_lifetime=3600,  # 1 hour
                max_connection_pool_size=50,
                connection_timeout=30.0,  # 30 seconds
                keep_alive=True
            )
            # Verify connectivity
            self.driver.verify_connectivity()
            return self.driver
        except Exception as e:
            raise Exception(f"Error connecting to Neo4j: {e}")
    
    def close(self):
        """Close the driver connection."""
        if self.driver:
            self.driver.close()
            self.driver = None
    
    def get_domains(self) -> List[str]:
        """
        Query Active Directory domains from Neo4j.
        
        Returns:
            List of domain names
        """
        if not self.driver:
            self.connect()
        
        try:
            with self.driver.session() as session:
                domain_query = "MATCH (d:Domain) RETURN d.name"
                result = session.run(domain_query)
                domains = sorted([record['d.name'] for record in result 
                                if record.get('d.name')])
                return domains
        except Exception as e:
            print(f"[WARNING] Error querying domains: {e}")
            return []
    
    def get_domains_detailed(self) -> List[Dict[str, str]]:
        """
        Query Active Directory domains with details from Neo4j.
        
        Returns:
            List of dictionaries with domain information (name, objectid)
        """
        if not self.driver:
            self.connect()
        
        try:
            with self.driver.session() as session:
                domain_query = "MATCH (d:Domain) RETURN d.name AS name, d.objectid AS objectid"
                result = session.run(domain_query)
                domains = [{'name': record['name'], 'objectid': record.get('objectid', 'N/A')} 
                          for record in result if record.get('name')]
                return sorted(domains, key=lambda x: x['name'])
        except Exception as e:
            print(f"[WARNING] Error querying domains: {e}")
            return []
    
    def get_tenants(self) -> List[Dict[str, str]]:
        """
        Query Azure tenants from Neo4j.
        
        Returns:
            List of dictionaries with tenant information (name, objectid)
        """
        if not self.driver:
            self.connect()
        
        try:
            with self.driver.session() as session:
                tenant_query = "MATCH (t:AZTenant) RETURN t.name AS name, t.objectid AS objectid"
                result = session.run(tenant_query)
                tenants = [{'name': record['name'], 'objectid': record.get('objectid', 'N/A')} 
                          for record in result if record.get('name')]
                return sorted(tenants, key=lambda x: x['name'])
        except Exception as e:
            print(f"[WARNING] Error querying tenants: {e}")
            return []
    
    def has_ad_data(self) -> bool:
        """
        Check if Active Directory data exists in the database.
        
        Returns:
            True if at least one Domain node exists, False otherwise
        """
        if not self.driver:
            self.connect()
        
        try:
            with self.driver.session() as session:
                check_query = "MATCH (d:Domain) RETURN count(d) > 0 AS exists LIMIT 1"
                result = session.run(check_query)
                record = result.single()
                return record['exists'] if record else False
        except Exception as e:
            print(f"[WARNING] Error checking AD data: {e}")
            return False
    
    def has_azure_data(self) -> bool:
        """
        Check if Azure data exists in the database.
        
        Returns:
            True if at least one AZTenant node exists, False otherwise
        """
        if not self.driver:
            self.connect()
        
        try:
            with self.driver.session() as session:
                check_query = "MATCH (t:AZTenant) RETURN count(t) > 0 AS exists LIMIT 1"
                result = session.run(check_query)
                record = result.single()
                return record['exists'] if record else False
        except Exception as e:
            print(f"[WARNING] Error checking Azure data: {e}")
            return False



