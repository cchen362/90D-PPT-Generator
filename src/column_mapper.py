"""
Column mapping module for 90D-PPT Generator
Handles intelligent column mapping and user interface for column selection
"""

import pandas as pd
import streamlit as st
from typing import Dict, List, Optional, Tuple
from utils import get_column_suggestions, normalize_column_name


class ColumnMapper:
    """Handles column mapping between source data and PowerPoint requirements"""
    
    def __init__(self):
        # OrderedDict to maintain the exact order: Jira Key > Target Complete Date > Description > Status > Risks/Issues/Watch Items > Component
        from collections import OrderedDict
        self.columns = OrderedDict([
            ('jira', {
                'display_name': 'Jira Key',
                'description': 'Unique identifier for the JIRA ticket',
                'patterns': ['key', 'jira', 'ticket', 'issue', 'id']
            }),
            ('target_date', {
                'display_name': 'Target Complete Date',
                'description': 'Target completion date',
                'patterns': ['target', 'date', 'due', 'completion', 'deadline', 'target_date', 'complete_date', 'finish']
            }),
            ('description', {
                'display_name': 'Description',
                'description': 'Description or summary of the task',
                'patterns': ['description', 'summary', 'title', 'desc', 'name']
            }),
            ('status', {
                'display_name': 'Status',
                'description': 'Current status of the task',
                'patterns': ['status', 'state', 'condition', 'progress']
            }),
            ('risks', {
                'display_name': 'Risks/Issues/Watch Items',
                'description': 'Risk items or issues',
                'patterns': ['risk', 'risks', 'issues', 'notes', 'comments', 'concerns']
            }),
            ('component', {
                'display_name': 'Component',
                'description': 'Component or team responsible',
                'patterns': ['component', 'components', 'component/s', 'team', 'area', 'group']
            }),
            ('ranking', {
                'display_name': 'Calculator Ranking',
                'description': 'Priority ranking (Accepted, Up Next, Maybe, Likely No)',
                'patterns': ['ranking', 'priority', 'category', 'type', 'classification']
            })
        ])
    
    def get_auto_suggestions(self, df_columns: List[str]) -> Dict[str, str]:
        """Generate automatic column mapping suggestions"""
        suggestions = {}
        
        # Convert to lowercase for matching - handle both string and non-string column names
        col_lower_map = {str(col).lower(): col for col in df_columns}
        
        for target_col, config in self.columns.items():
            best_match = None
            best_score = 0
            
            for pattern in config['patterns']:
                for col_lower, col_original in col_lower_map.items():
                    # Calculate match score
                    score = self._calculate_match_score(pattern, col_lower)
                    
                    if score > best_score and score > 0.5:  # Minimum threshold
                        best_score = score
                        best_match = col_original
            
            if best_match:
                suggestions[target_col] = best_match
        
        return suggestions
    
    def _calculate_match_score(self, pattern: str, column_name: str) -> float:
        """Calculate matching score between pattern and column name"""
        # Exact match gets highest score
        if pattern == column_name:
            return 1.0
        
        # Pattern is contained in column name
        if pattern in column_name:
            return 0.8
        
        # Column name is contained in pattern (less likely but possible)
        if column_name in pattern:
            return 0.6
        
        # Check for partial matches (e.g., "desc" matches "description")
        if len(pattern) >= 3 and pattern in column_name:
            return 0.7
        
        # Fuzzy matching for common variations
        fuzzy_matches = {
            'key': ['jira', 'ticket', 'issue'],
            'jira': ['key', 'ticket', 'issue'],
            'desc': ['description', 'summary'],
            'description': ['desc', 'summary'],
            'status': ['state', 'condition'],
            'component': ['team', 'area', 'group'],
            'ranking': ['priority', 'type']
        }
        
        if pattern in fuzzy_matches:
            for fuzzy in fuzzy_matches[pattern]:
                if fuzzy in column_name:
                    return 0.6
        
        return 0.0
    
    def render_mapping_interface(self, df: pd.DataFrame) -> Dict[str, str]:
        """Render Streamlit interface for column mapping"""
        st.markdown("### 🔗 Column Mapping")
        st.markdown("Map your data columns to the required PowerPoint fields:")
        
        # Get automatic suggestions
        suggestions = self.get_auto_suggestions(list(df.columns))
        
        # Initialize mapping from session state or suggestions
        if 'column_mapping' not in st.session_state:
            st.session_state.column_mapping = suggestions
        
        mapping = {}
        
        # Create mapping interface - adjusted spacing
        col1, col2 = st.columns([0.8, 1.2])
        
        with col1:
            st.markdown("#### Required Fields")
            
        with col2:
            st.markdown("#### Your Data Columns")
        
        # Clean column names and create options
        clean_columns = []
        for col in df.columns:
            clean_col = str(col).strip() if pd.notna(col) else f"Column_{len(clean_columns)}"
            if not clean_col or clean_col == "":
                clean_col = f"Column_{len(clean_columns)}"
            clean_columns.append(clean_col)
        
        # Update dataframe with clean column names
        df.columns = clean_columns
        column_options = ["-- Select Column --"] + clean_columns
        
        for target_col, config in self.columns.items():
            col1, col2 = st.columns([0.8, 1.2])
            
            with col1:
                # Show field info - all fields are now optional
                st.markdown(f"📋 **{config['display_name']}**")
                st.markdown(f"*{config['description']}*")
            
            with col2:
                # Get current selection
                current_selection = st.session_state.column_mapping.get(target_col, "-- Select Column --")
                if current_selection not in column_options:
                    current_selection = "-- Select Column --"
                
                # Create selectbox with clean column names
                selected = st.selectbox(
                    f"Map to:",
                    column_options,
                    index=column_options.index(current_selection),
                    key=f"mapping_{target_col}",
                    help=f"Select the column from your data that contains {config['display_name'].lower()}"
                )
                
                if selected != "-- Select Column --":
                    mapping[target_col] = selected
                    st.session_state.column_mapping[target_col] = selected
                else:
                    st.session_state.column_mapping.pop(target_col, None)
        
        # Show mapping summary
        st.markdown("---")
        st.markdown("### 📋 Mapping Summary")
        
        mapped_count = 0
        total_columns = len(self.columns)
        
        for target_col, config in self.columns.items():
            mapped_col = mapping.get(target_col)
            
            if mapped_col:
                status_icon = "✅"
                mapped_count += 1
            else:
                status_icon = "⚪"
            
            mapped_display = mapped_col if mapped_col else "Not mapped (will be blank)"
            st.markdown(f"{status_icon} **{config['display_name']}** → {mapped_display}")
        
        # Show status - no validation required, all mappings are optional
        if mapped_count > 0:
            st.success(f"✅ {mapped_count}/{total_columns} columns mapped")
        else:
            st.info("ℹ️ No columns mapped - all PowerPoint columns will be blank")
        
        return mapping
    
    def validate_mapping(self, df: pd.DataFrame, mapping: Dict[str, str]) -> Tuple[bool, List[str]]:
        """Validate the column mapping"""
        errors = []
        warnings = []
        
        # Only check that mapped columns actually exist in the dataframe
        for target_col, mapped_col in mapping.items():
            if mapped_col and mapped_col not in df.columns:
                errors.append(f"Mapped column '{mapped_col}' not found in data")
                continue
            
            # Warn about empty columns but don't treat as error
            if mapped_col and mapped_col in df.columns:
                non_null_count = df[mapped_col].notna().sum()
                if non_null_count == 0:
                    warnings.append(f"Column '{mapped_col}' contains no data - PowerPoint column will be blank")
                elif non_null_count < len(df) * 0.3:  # Less than 30% filled
                    warnings.append(f"Column '{mapped_col}' is mostly empty ({non_null_count}/{len(df)} rows)")
        
        # Special validation for ranking column - now just a warning
        if 'ranking' in mapping and mapping['ranking']:
            ranking_col = mapping['ranking']
            if ranking_col in df.columns:
                # Check if ranking values are recognizable
                from utils import normalize_ranking
                valid_rankings = df[ranking_col].apply(
                    lambda x: normalize_ranking(x) is not None if pd.notna(x) else False
                ).sum()
                
                if valid_rankings == 0:
                    warnings.append(f"No valid ranking values found in column '{ranking_col}' - slides may be empty")
                elif valid_rankings < len(df) * 0.5:
                    # No longer show this warning since rankings are optional
                    pass
        
        # Show warnings if any
        if warnings:
            for warning in warnings:
                st.warning(f"⚠️ {warning}")
        
        return len(errors) == 0, errors
    
    def preview_mapped_data(self, df: pd.DataFrame, mapping: Dict[str, str], max_rows: int = 10):
        """Preview how the mapped data will look"""
        if not mapping:
            return
        
        st.markdown("### 👀 Data Preview")
        st.markdown("Here's how your mapped data will appear in the PowerPoint:")
        
        # Create preview dataframe
        preview_data = {}
        ppt_columns = [
            ('JIRA', 'jira'),
            ('Target Complete Date', 'target_date'),
            ('Description', 'description'),
            ('Status', 'status'),
            ('Risks/Issues', 'risks'),
            ('Component', 'component')
        ]
        
        for ppt_col_name, target_col in ppt_columns:
            if target_col is None:
                # Empty column
                preview_data[ppt_col_name] = [""] * min(max_rows, len(df))
            elif target_col in mapping and mapping[target_col] in df.columns:
                # Mapped column
                preview_data[ppt_col_name] = df[mapping[target_col]].head(max_rows).fillna("").astype(str)
            else:
                # Not mapped
                preview_data[ppt_col_name] = ["(Not mapped)"] * min(max_rows, len(df))
        
        preview_df = pd.DataFrame(preview_data)
        
        # Display the preview
        st.dataframe(
            preview_df,
            use_container_width=True,
            height=min(400, (len(preview_df) + 1) * 35)
        )
        
        # Show statistics
        if 'ranking' in mapping and mapping['ranking'] in df.columns:
            st.markdown("#### 📊 Data Statistics")
            
            # Ranking distribution
            from utils import normalize_ranking
            ranking_col = mapping['ranking']
            ranking_counts = {}
            
            for ranking in ['Accepted', 'Up Next', 'Maybe', 'Likely No']:
                count = df[ranking_col].apply(
                    lambda x: normalize_ranking(x) == ranking if pd.notna(x) else False
                ).sum()
                if count > 0:
                    ranking_counts[ranking] = count
            
            if ranking_counts:
                st.markdown("**Ranking Distribution:**")
                for ranking, count in ranking_counts.items():
                    st.markdown(f"- {ranking}: {count} items")
            else:
                st.warning("⚠️ No valid ranking data found")
            
            # Status distribution
            if 'status' in mapping and mapping['status'] in df.columns:
                status_col = mapping['status']
                from utils import categorize_status
                
                closed_count = df[status_col].apply(
                    lambda x: categorize_status(x) == 'closed'
                ).sum()
                non_closed_count = len(df) - closed_count
                
                st.markdown("**Status Distribution:**")
                st.markdown(f"- Closed: {closed_count} items")
                st.markdown(f"- Non-closed: {non_closed_count} items")
    
    def get_mapping_from_session(self) -> Dict[str, str]:
        """Get current mapping from session state"""
        return st.session_state.get('column_mapping', {})