"""
Choke Points security check queries.

This module contains queries for identifying choke points - critical privilege
edges that connect non-Tier-0 objects to Tier-0 objects.
"""

import pandas as pd
from chokehound.queries.registry import SecurityQuery, register_query
from chokehound.config.risk_config import (
    calculate_risk_score,
    normalize_risk_score
)
from chokehound.config import settings


# ============================================================================
# RELATIONSHIP TYPES AND QUERY HELPERS
# ============================================================================

# List of relationship types used in queries
RELATIONSHIP_TYPES = [
    "Owns", "GenericAll", "GenericWrite", "WriteOwner", "WriteDacl", "MemberOf",
    "ForceChangePassword", "AllExtendedRights", "AddMember", "HasSession", "GPLink",
    "AllowedToDelegate", "CoerceToTGT", "AllowedToAct", "AdminTo", "CanPSRemote",
    "CanRDP", "ExecuteDCOM", "HasSIDHistory", "AddSelf", "DCSync", "ReadLAPSPassword",
    "ReadGMSAPassword", "DumpSMSAPassword", "SQLAdmin", "AddAllowedToAct", "WriteSPN",
    "AddKeyCredentialLink", "SyncLAPSPassword", "WriteAccountRestrictions", "WriteGPLink",
    "GoldenCert", "ADCSESC1", "ADCSESC3", "ADCSESC4", "ADCSESC6a", "ADCSESC6b",
    "ADCSESC9a", "ADCSESC9b", "ADCSESC10a", "ADCSESC10b", "ADCSESC13",
    "CoerceAndRelayNTLMToSMB", "CoerceAndRelayNTLMToADCS", "WriteOwnerLimitedRights",
    "OwnsLimitedRights", "ClaimSpecialIdentity", "CoerceAndRelayNTLMToLDAP",
    "CoerceAndRelayNTLMToLDAPS", "ContainsIdentity", "PropagatesACEsTo", "GPOAppliesTo",
    "CanApplyGPO", "HasTrustKeys", "ManageCA", "ManageCertificates", "Contains", "DCFor",
    "SameForestTrust", "SpoofSIDHistory", "AbuseTGTDelegation", "SyncedToADUser"
]


def get_relationship_pattern():
    """Returns the relationship pattern string for use in Cypher queries."""
    return "|".join(RELATIONSHIP_TYPES)


# Azure relationship types used in choke points queries
AZURE_RELATIONSHIP_TYPES = [
    "AZAvereContributor", "AZContributor", "AZGetCertificates", "AZGetKeys", 
    "AZGetSecrets", "AZHasRole", "AZMemberOf", "AZOwner", "AZRunsAs", 
    "AZVMContributor", "AZAutomationContributor", "AZKeyVaultContributor", 
    "AZVMAdminLogin", "AZAddMembers", "AZAddSecret", "AZExecuteCommand", 
    "AZGlobalAdmin", "AZPrivilegedAuthAdmin", "AZGrant", "AZGrantSelf", 
    "AZPrivilegedRoleAdmin", "AZResetPassword", "AZUserAccessAdministrator", 
    "AZOwns", "AZCloudAppAdmin", "AZAppAdmin", "AZAddOwner", "AZManagedIdentity", 
    "AZAKSContributor", "AZNodeResourceGroup", "AZWebsiteContributor", 
    "AZLogicAppContributor", "AZMGAddMember", "AZMGAddOwner", "AZMGAddSecret", 
    "AZMGGrantAppRoles", "AZMGGrantRole", "AZRoleEligible", 
    "AZRoleApprover", "AZContains", "SyncedToEntraUser"
]


def get_azure_relationship_pattern():
    """Returns the Azure relationship pattern string for use in Cypher queries."""
    return "|".join(AZURE_RELATIONSHIP_TYPES)


# ============================================================================
# POST-PROCESSING FUNCTIONS
# ============================================================================

def post_process_choke_points_risk(df: pd.DataFrame, enable_logging: bool = False):
    """
    Post-process choke points risk query results.
    Adds risk scores and sorts by risk.
    
    Args:
        df: DataFrame with choke points data
        enable_logging: If True, returns risk breakdowns for logging
        
    Returns:
        Processed DataFrame (and optionally breakdowns)
    """
    if df.empty:
        return df if not enable_logging else (df, [])
    
    # Check if required columns exist
    required_cols = ['SourceType', 'RelationshipType', 'TargetType', 'AffectedAttackPaths']
    if not all(col in df.columns for col in required_cols):
        return df if not enable_logging else (df, [])
    
    df = df.copy()
    risk_breakdowns = []
    
    if enable_logging:
        # Calculate risk with breakdowns
        results_list = []
        for idx in df.index:
            row = df.loc[idx]
            risk_score, breakdown = calculate_risk_score(row, return_breakdown=True)
            results_list.append((risk_score, breakdown))
        
        df['RiskScore'] = [normalize_risk_score(r[0]) for r in results_list]
        risk_breakdowns = [r[1] for r in results_list]
        df['_breakdown_idx'] = range(len(df))
    else:
        risk_scores = df.apply(calculate_risk_score, axis=1)
        df['RiskScore'] = risk_scores.apply(normalize_risk_score)
    
    # Sort by RiskScore descending, then by AffectedAttackPaths descending
    df = df.sort_values(['RiskScore', 'AffectedAttackPaths'], ascending=[False, False])
    
    # Reorder breakdowns to match sorted dataframe
    if enable_logging and risk_breakdowns:
        sorted_indices = df['_breakdown_idx'].tolist()
        risk_breakdowns = [risk_breakdowns[int(idx)] for idx in sorted_indices]
        df = df.drop(columns=['_breakdown_idx'])
    
    # Reset index after sorting
    df = df.reset_index(drop=True)
    
    # Add unique ID column (AD1, AD2, AD3, etc.)
    df.insert(0, 'ID', [f'AD{i}' for i in range(1, len(df) + 1)])
    
    if enable_logging:
        return df, risk_breakdowns
    return df


def _post_process_wrapper(df: pd.DataFrame):
    """Wrapper for post-processing without logging."""
    return post_process_choke_points_risk(df, enable_logging=False)


def _post_process_azure_no_risk(df: pd.DataFrame):
    """Wrapper for Azure post-processing without risk calculation (uses AZ prefix)."""
    return post_process_choke_points_no_risk(df, id_prefix="AZ")


def post_process_choke_points_no_risk(df: pd.DataFrame, id_prefix: str = "AD"):
    """
    Post-process choke points WITHOUT risk calculation.
    Sorts by affected attack paths and adds ID column.
    
    This is a faster alternative when risk scoring is not needed.
    
    Args:
        df: DataFrame with choke points data
        id_prefix: Prefix for ID column (default: "AD")
        
    Returns:
        Processed DataFrame (sorted by AffectedAttackPaths, no risk scores)
    """
    if df.empty:
        return df
    
    df = df.copy()
    
    # Sort by AffectedAttackPaths descending (most critical first)
    # Then by SourceName for consistency
    if 'AffectedAttackPaths' in df.columns:
        df = df.sort_values(['AffectedAttackPaths', 'SourceName'], 
                           ascending=[False, True])
    
    # Reset index after sorting
    df = df.reset_index(drop=True)
    
    # Add unique ID column (e.g., AD1, AD2, AD3 or AZ1, AZ2, AZ3, etc.)
    df.insert(0, 'ID', [f'{id_prefix}{i}' for i in range(1, len(df) + 1)])
    
    return df


# ============================================================================
# QUERY DEFINITIONS
# ============================================================================

# Query template for AD Tier 0 Choke Points
_AD_CHOKE_POINTS_QUERY_TEMPLATE = """
    // Step 1: identify all direct edges into Tier-0 (choke candidates)
    MATCH (src)-[r:{rel_pattern}]->(t:Tag_Tier_Zero)
    WHERE NOT src:Tag_Tier_Zero
    WITH DISTINCT src, t, type(r) AS RelationshipType

    // Step 2: for each choke edge, find all non–Tier-0 origins that can reach src (up to {hop_limit} hops)
    MATCH p = (o)-[:{rel_pattern}*0..{hop_limit}]->(src)
    WHERE NOT o:Tag_Tier_Zero
    // make sure no Tier-0 nodes appear in the origin→src path
    AND ALL(n IN nodes(p) WHERE NOT n:Tag_Tier_Zero)

    // Step 3: aggregate & rank choke edges by how many unique origins can reach them
    WITH src, t, RelationshipType, count(DISTINCT o) AS ReachableOrigins
    RETURN
        src.name AS SourceName,
        labels(src) AS SourceType,
        src.objectid AS SourceObjectID,
        src.distinguishedname AS SourceDN,
        RelationshipType,
        t.name AS TargetName,
        labels(t) AS TargetType,
        t.objectid AS TargetObjectID,
        t.distinguishedname AS TargetDN,
        ReachableOrigins AS AffectedAttackPaths
    
    LIMIT {limit}
"""

# Create query instances
_choke_points_risk_query = SecurityQuery(
    name="AD Tier 0 Choke Points Risk",
    description="Prioritized Active Directory Tier 0 choke points that connect non-Tier-0 objects with Tier-0 objects. Results are ranked by risk score (highest risk first) and number of affected attack paths.",
    cypher_query="",  # Placeholder, will use query_formatter
    post_process=_post_process_wrapper,
    query_formatter=lambda: _AD_CHOKE_POINTS_QUERY_TEMPLATE.format(
        rel_pattern=get_relationship_pattern(),
        hop_limit=settings.AD_CHOKE_POINTS_HOP_LIMIT,
        limit=settings.AD_CHOKE_POINTS_LIMIT
    )
)

_direct_relationships_query = SecurityQuery(
    name="ADirect relationships into Tier0",
    description="Ranking of Active Directory relationship types (edges) that directly connect to Tier-0 objects. This identifies which types of privileges are most commonly used to access Tier-0, helping identify edge-type choke points.",
    cypher_query="""
        MATCH (src)-[r]->(dst:Tag_Tier_Zero)
        WHERE NOT src:Tag_Tier_Zero
        RETURN type(r) AS RelationshipType,
            count(DISTINCT src.name + '|' + dst.name) AS DistinctSourceTargetPairs,
            count(*) AS TotalEdges
        ORDER BY DistinctSourceTargetPairs DESC, TotalEdges DESC
    """
)

_source_nodes_query = SecurityQuery(
    name="Source nodes into T0 Rank",
    description="Top Active Directory non-Tier-0 nodes ranked by the number of distinct Tier-0 targets they directly connect to. These are immediate choke points - objects that have direct relationships to many Tier-0 objects.",
    cypher_query="""
        MATCH (src)-[r]->(dst:Tag_Tier_Zero)
        WHERE NOT src:Tag_Tier_Zero
        RETURN src.name AS SourceNode,
            labels(src) AS SourceType,
            src.objectid AS SourceObjectID,
            src.distinguishedname AS SourceDN,
            count(DISTINCT dst.name) AS DistinctTier0Targets,
            collect(DISTINCT type(r)) AS RelationshipTypes
        ORDER BY DistinctTier0Targets DESC, SourceNode
        LIMIT 100
    """
)

# Query template for Azure Tier 0 Choke Points
_AZURE_CHOKE_POINTS_QUERY_TEMPLATE = """
    // Step 1: identify all direct Azure edges into Tier-0 (choke candidates)
    MATCH (src)-[r:{azure_rel_pattern}]->(t:Tag_Tier_Zero)
    WHERE NOT src:Tag_Tier_Zero
    WITH DISTINCT src, t, type(r) AS RelationshipType
    LIMIT {limit}  // CRITICAL: Limit choke candidates BEFORE expensive path traversal

    // Step 2: for each choke edge, find all non–Tier-0 origins that can reach src (up to {hop_limit} hops)
    MATCH p = (o)-[:{azure_rel_pattern}*0..{hop_limit}]->(src)
    WHERE NOT o:Tag_Tier_Zero
    // make sure no Tier-0 nodes appear in the origin→src path
    AND ALL(n IN nodes(p) WHERE NOT n:Tag_Tier_Zero)

    // Step 3: aggregate & rank choke edges by how many unique origins can reach them
    WITH src, t, RelationshipType, count(DISTINCT o) AS ReachableOrigins
    RETURN
        src.name AS SourceName,
        labels(src) AS SourceType,
        src.objectid AS SourceObjectID,
        RelationshipType,
        t.name AS TargetName,
        labels(t) AS TargetType,
        t.objectid AS TargetObjectID,
        ReachableOrigins AS AffectedAttackPaths
    
    ORDER BY ReachableOrigins DESC
"""

def _format_azure_query():
    formatted_query = _AZURE_CHOKE_POINTS_QUERY_TEMPLATE.format(
        azure_rel_pattern=get_azure_relationship_pattern(),
        hop_limit=settings.AZURE_CHOKE_POINTS_HOP_LIMIT,
        limit=settings.AZURE_CHOKE_POINTS_LIMIT
    )
    return formatted_query

_azure_choke_points_query = SecurityQuery(
    name="Azure Tier 0 Choke Points",
    description="Azure choke points that connect non-Tier-0 objects with Tier-0 objects. Results are sorted by number of affected attack paths (no risk scoring yet).",
    cypher_query="",  # Placeholder, will use query_formatter
    post_process=_post_process_azure_no_risk,
    query_formatter=_format_azure_query
)

# Register queries
register_query(_choke_points_risk_query)
register_query(_azure_choke_points_query)
register_query(_direct_relationships_query)
register_query(_source_nodes_query)
