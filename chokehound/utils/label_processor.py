"""
Label processing utilities.

Simplifies Neo4j node labels for better readability.
"""

import pandas as pd


def simplify_labels(label_array):
    """
    Simplify label arrays by removing 'Base', 'AZBase', and 'Tag_Tier_Zero', 
    keeping the most relevant label.
    
    Examples:
        ['Base', 'Container'] -> 'Container'
        ['Base', 'Group'] -> 'Group'
        ['Base', 'ADLocalGroup', 'Group'] -> 'Group'
        ['Base', 'Computer', 'Tag_Tier_Zero'] -> 'Computer'
        ['AZServicePrincipal', 'AZBase'] -> 'AZServicePrincipal'
        ['AZUser', 'AZBase'] -> 'AZUser'
        ['Base'] -> 'UNKNOWN'
        '[Base]' -> 'UNKNOWN'
    
    Args:
        label_array: List of labels or string representation of list
        
    Returns:
        Simplified label string, or 'UNKNOWN' if only Base remains
    """
    if not label_array:
        return ""
    
    # Handle string representation of list (from Neo4j)
    if isinstance(label_array, str):
        # Check if it's the string "[Base]" or "Base" or "[AZBase]"
        if label_array.strip() in ['[Base]', 'Base', '["Base"]', '[AZBase]', 'AZBase', '["AZBase"]']:
            return "UNKNOWN"
        # Try to evaluate if it's a string representation of a list
        try:
            import ast
            label_array = ast.literal_eval(label_array)
        except (ValueError, SyntaxError):
            return label_array
    
    # Ensure it's a list
    if not isinstance(label_array, list):
        return str(label_array)
    
    # Filter out 'Base', 'AZBase', and 'Tag_Tier_Zero'
    filtered_labels = [label for label in label_array 
                      if label not in ['Base', 'AZBase', 'Tag_Tier_Zero']]
    
    # If nothing left, return UNKNOWN (only Base/AZBase/Tag_Tier_Zero were present)
    if not filtered_labels:
        return "UNKNOWN"
    
    # Return the last (most specific) label
    return filtered_labels[-1]


def process_dataframe_labels(df):
    """
    Process columns containing node labels to simplify them.
    Processes columns ending with 'Type' (SourceType, TargetType, PrincipalType, etc.)
    Converts "[Base]", "Base", "[AZBase]", or "AZBase" to "UNKNOWN" when no other labels are present.
    
    Args:
        df: pandas DataFrame
        
    Returns:
        Processed DataFrame
    """
    df = df.copy()
    
    # Process all columns that end with 'Type' (typically containing Neo4j labels)
    for col in df.columns:
        if col.endswith('Type'):
            df[col] = df[col].apply(simplify_labels)
            # Handle case where result is "[Base]" or "[AZBase]" string (shouldn't happen after simplify_labels, but just in case)
            df[col] = df[col].replace(['[Base]', 'Base', '[AZBase]', 'AZBase'], 'UNKNOWN')
    
    return df



