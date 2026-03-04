"""
Configuration file for ChokeHound settings.
Adjust these values to match your BloodHound instance and analysis requirements.
"""

# ============================================================================
# NEO4J CONNECTION SETTINGS
# ============================================================================

# Neo4j connection settings
NEO4J_URI = "bolt://localhost:7687"
NEO4J_USER = "neo4j"
NEO4J_PASSWORD = "bloodhoundcommunityedition"


# ============================================================================
# OUTPUT SETTINGS
# ============================================================================

# Default output filename (will be used if not specified via command line)
DEFAULT_OUTPUT_FILENAME = "ChokeHound_report.xlsx"


# ============================================================================
# CHOKE POINTS QUERY SETTINGS
# ============================================================================

# Hop limit for AD Tier 0 Choke Points query (maximum number of hops to traverse when finding origins)
# 
# WARNING: Query complexity grows EXPONENTIALLY with each additional hop!
# 
# Performance guidelines based on environment size:
#   Small (<1,000 nodes):     hop_limit 3-5 usually works fine
#   Medium (1,000-10,000):    hop_limit 2-3 recommended
#   Large (10,000-50,000):    hop_limit 1-2 recommended
#   Very Large (>50,000):     hop_limit 0-1 recommended (direct relationships only)
# 
# If queries take too long or timeout, reduce this value by 1 and try again.
AD_CHOKE_POINTS_HOP_LIMIT = 3

# Result limit for AD choke points query
# Controls the maximum number of choke points returned in the report.
# Lower values = faster queries, but may miss some findings.
AD_CHOKE_POINTS_LIMIT = 200

# Hop limit for Azure choke points query (maximum number of hops to traverse when finding origins)
# 
# WARNING: Azure queries are MORE EXPENSIVE than AD queries due to:
#   - 39 Azure relationship types (vs ~54 AD types, but broader Azure patterns)
#   - More interconnected graph structure in Azure/Entra ID environments
#   - Each hop creates exponentially more path combinations
# 
# Performance guidelines based on environment size:
#   Small Azure tenant (<500 objects):     hop_limit 1-2 may work
#   Medium tenant (500-5,000 objects):     hop_limit 0-1 recommended
#   Large tenant (>5,000 objects):         hop_limit 0 recommended (direct only)
# 
# If the Azure query hangs or takes >60 seconds:
#   1. Reduce AZURE_CHOKE_POINTS_HOP_LIMIT to 0
#   2. Reduce AZURE_CHOKE_POINTS_LIMIT to 20-30
# 
# Calculation: With hop_limit=3, each choke candidate can spawn ~60,000 path evaluations
#              (1 + 39 + 39² + 39³). This can easily create millions of paths to evaluate.
AZURE_CHOKE_POINTS_HOP_LIMIT = 2

# Result limit for Azure choke points query
# 
# This limit is applied EARLY in the query (after finding choke candidates),
# which dramatically reduces the number of paths to compute.
# 
# Performance impact:
#   - limit=200 with hop_limit=1: ~8,000 path evaluations
#   - limit=100 with hop_limit=1: ~4,000 path evaluations
#   - limit=50 with hop_limit=1:  ~2,000 path evaluations
#   - limit=20 with hop_limit=1:  ~800 path evaluations
# 
# For large environments, start with limit=20-50 and increase if needed.
AZURE_CHOKE_POINTS_LIMIT = 60

# ============================================================================
# QUERY EXECUTION SETTINGS
# ============================================================================

# Query timeout in seconds (default: 300 = 5 minutes)
# Increase for very large environments or complex queries
QUERY_TIMEOUT_SECONDS = 300

# Azure query timeout (may need longer than AD queries in some environments)
AZURE_QUERY_TIMEOUT_SECONDS = 600  # 10 minutes

