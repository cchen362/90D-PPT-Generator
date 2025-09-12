"""
Utility functions for the 90D-PPT Generator
"""

import pandas as pd
import streamlit as st
from typing import Dict, List, Tuple, Optional, Any
import os
import tempfile
from pathlib import Path


def get_file_size_mb(file_bytes: bytes) -> float:
    """Get file size in megabytes"""
    return len(file_bytes) / (1024 * 1024)


def validate_file_size(file_bytes: bytes, max_size_mb: int = 200) -> bool:
    """Validate if file size is within acceptable limits"""
    size_mb = get_file_size_mb(file_bytes)
    return size_mb <= max_size_mb


def get_supported_file_extensions() -> List[str]:
    """Get list of supported file extensions"""
    return ['.xlsx', '.xls', '.csv']


def is_supported_file(filename: str) -> bool:
    """Check if file extension is supported"""
    file_ext = Path(filename).suffix.lower()
    return file_ext in get_supported_file_extensions()


def normalize_column_name(col_name: str) -> str:
    """Normalize column names for consistent mapping"""
    if pd.isna(col_name) or col_name is None:
        return ""
    return str(col_name).strip().lower()


def detect_header_row(df: pd.DataFrame, sheet_name: Optional[str] = None) -> int:
    """
    Enhanced header row detection with multiple strategies with defensive checks
    Returns the row index (0-based)
    """
    # Defensive check for None or invalid dataframe
    if df is None:
        return 0
    
    try:
        if len(df) == 0:
            return 0
    except (TypeError, AttributeError):
        # Handle case where df is not a valid DataFrame
        return 0
    
    best_row = 0
    best_score = 0
    
    # Check first 10 rows for header patterns
    for i in range(min(10, len(df))):
        row_data = df.iloc[i].fillna('')
        
        # Skip completely empty rows
        if all(str(cell).strip() == '' for cell in row_data):
            continue
            
        # Convert to strings and check for patterns
        str_row = [str(cell).lower().strip() for cell in row_data if str(cell).strip()]
        
        if len(str_row) == 0:
            continue
        
        score = 0
        
        # 1. Check for header-like keywords
        header_indicators = [
            'key', 'jira', 'ticket', 'id', 'issue',
            'status', 'state', 'condition', 'progress',
            'description', 'summary', 'title', 'desc', 'name',
            'component', 'team', 'area', 'group',
            'ranking', 'priority', 'category', 'type',
            'date', 'created', 'updated', 'modified',
            'risk', 'risks', 'notes', 'comments'
        ]
        
        keyword_matches = sum(1 for cell in str_row if any(indicator in cell for indicator in header_indicators))
        score += keyword_matches * 10
        
        # 2. Check for non-numeric content (headers are typically text)
        non_numeric_cells = sum(1 for cell in str_row if not _is_numeric(cell))
        score += non_numeric_cells * 2
        
        # 3. Check for reasonable length (not too short, not too long)
        reasonable_length = sum(1 for cell in str_row if 3 <= len(cell) <= 50)
        score += reasonable_length * 1
        
        # 4. Penalize rows that look like data
        data_like_patterns = ['todo', 'in progress', 'done', 'closed', 'completed']
        data_penalties = sum(1 for cell in str_row if any(pattern in cell for pattern in data_like_patterns))
        score -= data_penalties * 5
        
        # 5. Bonus for having standard column count (4-8 columns is typical)
        if 4 <= len(str_row) <= 8:
            score += 5
        
        if score > best_score:
            best_score = score
            best_row = i
    
    return best_row


def _is_numeric(value: str) -> bool:
    """Check if a string represents a number"""
    try:
        float(value)
        return True
    except (ValueError, TypeError):
        return False


def categorize_status(status) -> str:
    """
    Categorize status into 'closed' or 'non-closed' for counting
    
    Handles any data type: str, int, float, None, etc. with robust type checking
    """
    # Handle None, NaN, or empty values
    if status is None or pd.isna(status):
        return 'non-closed'
    
    try:
        # Safely convert any type to string
        if isinstance(status, (int, float)):
            # Handle numeric values that might represent status codes
            if pd.isna(status):
                return 'non-closed'
            status_str = str(status).strip()
        elif isinstance(status, str):
            status_str = status.strip()
        else:
            # Handle other types (datetime, etc.)
            status_str = str(status).strip()
        
        # Handle empty string after conversion
        if not status_str or status_str in ['nan', 'None', 'null', '']:
            return 'non-closed'
            
        status_lower = status_str.lower()
        
        closed_statuses = {
            'closed', 'done', 'resolved', 'completed', 'finished', 
            'complete', 'fixed', 'verified', 'delivered'
        }
        
        return 'closed' if status_lower in closed_statuses else 'non-closed'
        
    except Exception as e:
        # Log the error for debugging but don't break the flow
        print(f"Warning: Error processing status value {status}: {e}")
        return 'non-closed'


def get_ranking_variations() -> Dict[str, List[str]]:
    """
    Get variations of ranking values to handle case sensitivity and typos
    """
    return {
        'accepted': ['accepted', 'accept', 'approved', 'confirmed'],
        'up_next': ['up next', 'upnext', 'next', 'upcoming', 'up-next'],
        'maybe': ['maybe', 'possible', 'perhaps', 'potentially'],
        'likely_no': ['likely no', 'likelyno', 'probably not', 'unlikely', 'likely-no']
    }


def normalize_ranking(ranking) -> Optional[str]:
    """
    Normalize ranking values to standard categories with robust type handling
    Returns None if ranking should be ignored
    
    Handles any data type: str, int, float, None, etc.
    """
    # Handle None, NaN, or empty values
    if ranking is None or pd.isna(ranking):
        return None
    
    try:
        # Safely convert any type to string with explicit type checking
        if isinstance(ranking, (int, float)):
            # Handle numeric values that might represent rankings
            if pd.isna(ranking):
                return None
            ranking_str = str(ranking).strip()
        elif isinstance(ranking, str):
            ranking_str = ranking.strip()
        else:
            # Handle other types (datetime, bool, etc.)
            ranking_str = str(ranking).strip()
        
        # Handle empty string after conversion
        if not ranking_str or ranking_str in ['nan', 'None', 'null', '']:
            return None
            
        ranking_lower = ranking_str.lower()
        ranking_variations = get_ranking_variations()
        
        for standard_ranking, variations in ranking_variations.items():
            if any(var in ranking_lower for var in variations):
                return standard_ranking.replace('_', ' ').title()
        
        return None  # Ignore unrecognized rankings
        
    except Exception as e:
        # Log the error for debugging but don't break the flow
        print(f"Warning: Error processing ranking value {ranking}: {e}")
        return None


def create_temp_file(content: bytes, suffix: str = '.pptx') -> str:
    """
    Create a temporary file with the given content
    Returns the file path
    """
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    temp_file.write(content)
    temp_file.close()
    return temp_file.name


def show_success_message(message: str, additional_info: str = ""):
    """Display success message with consistent styling"""
    full_message = f"✅ {message}"
    if additional_info:
        full_message += f"\n\n{additional_info}"
    st.success(full_message)


def show_error_message(message: str, suggestion: str = "", debug_info: str = ""):
    """Display error message with actionable guidance"""
    full_message = f"❌ {message}"
    
    if suggestion:
        full_message += f"\n\n**💡 What you can do:**\n{suggestion}"
    
    if debug_info:
        full_message += f"\n\n**🔍 Technical details:**\n{debug_info}"
    
    st.error(full_message)


def show_warning_message(message: str, recommendation: str = ""):
    """Display warning message with recommendations"""
    full_message = f"⚠️ {message}"
    
    if recommendation:
        full_message += f"\n\n**💡 Recommendation:**\n{recommendation}"
    
    st.warning(full_message)


def show_info_message(message: str, details: str = ""):
    """Display info message with optional details"""
    full_message = f"ℹ️ {message}"
    
    if details:
        full_message += f"\n\n{details}"
    
    st.info(full_message)


def format_file_size(size_bytes: int) -> str:
    """Format file size in human readable format"""
    if size_bytes == 0:
        return "0 B"
    size_names = ["B", "KB", "MB", "GB"]
    import math
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return f"{s} {size_names[i]}"


def get_column_suggestions(df_columns: List[str]) -> Dict[str, str]:
    """
    Suggest column mappings based on common patterns
    Returns dict with target_column: suggested_source_column
    """
    suggestions = {}
    
    # Convert column names to lowercase for matching - handle both string and non-string column names
    col_mapping = {str(col).lower(): col for col in df_columns}
    
    # Mapping patterns (target -> possible source patterns)
    patterns = {
        'jira': ['key', 'jira', 'ticket', 'issue', 'id'],
        'description': ['description', 'summary', 'title', 'desc'],
        'status': ['status', 'state', 'condition'],
        'component': ['component', 'components', 'component/s', 'team', 'area'],
        'ranking': ['ranking', 'priority', 'category', 'type']
    }
    
    for target_col, source_patterns in patterns.items():
        for pattern in source_patterns:
            # Look for exact matches or partial matches
            matches = [col for col in col_mapping.keys() if pattern in col or col in pattern]
            if matches:
                # Pick the best match (shortest name, likely most specific)
                best_match = min(matches, key=len)
                suggestions[target_col] = col_mapping[best_match]
                break
    
    return suggestions


def validate_mapped_data(df: pd.DataFrame, column_mapping: Dict[str, str]) -> List[str]:
    """
    Validate that mapped columns contain usable data
    Returns list of validation warnings
    """
    warnings = []
    
    # Check if required columns are mapped and contain data
    required_columns = ['jira', 'description', 'status', 'ranking']
    
    for req_col in required_columns:
        if req_col not in column_mapping or not column_mapping[req_col]:
            warnings.append(f"Required column '{req_col}' is not mapped")
            continue
            
        mapped_col = column_mapping[req_col]
        if mapped_col not in df.columns:
            warnings.append(f"Mapped column '{mapped_col}' not found in data")
            continue
            
        # Check for empty data
        non_empty_count = df[mapped_col].notna().sum()
        if non_empty_count == 0:
            warnings.append(f"Column '{mapped_col}' contains no data")
        elif non_empty_count < len(df) * 0.5:  # Less than 50% filled
            warnings.append(f"Column '{mapped_col}' is mostly empty ({non_empty_count}/{len(df)} rows filled)")
    
    return warnings