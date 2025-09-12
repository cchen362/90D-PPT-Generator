"""
90D-PPT Generator - Main Streamlit Application
Converts Excel/CSV data into branded PowerPoint presentations for 90-day planning cycles
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import sys

# Add src directory to Python path
sys.path.append(str(Path(__file__).parent / "src"))
sys.path.append(str(Path(__file__).parent / "assets"))

# Import modules
from src.utils import (
    show_success_message, show_error_message, show_warning_message, 
    show_info_message, is_supported_file, validate_file_size
)
from src.file_processor import FileProcessor
from src.column_mapper import ColumnMapper
from src.ppt_generator import PowerPointGenerator
from src.pdf_exporter import PDFExporter
from src.debug_logger import debug_logger, debug_function
from assets.brand_colors import NAVY, LIGHT_BLUE, GREY

# Page configuration
st.set_page_config(
    page_title="90D-PPT Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load custom CSS
def load_css():
    """Load custom CSS styling"""
    css_file = Path(__file__).parent / "assets" / "styles.css"
    if css_file.exists():
        with open(css_file) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css()

# Initialize session state
def initialize_session_state():
    """Initialize session state variables"""
    if 'current_step' not in st.session_state:
        st.session_state.current_step = 1
    if 'uploaded_file' not in st.session_state:
        st.session_state.uploaded_file = None
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'original_df' not in st.session_state:
        st.session_state.original_df = None
    if 'sheet_names' not in st.session_state:
        st.session_state.sheet_names = []
    if 'selected_sheet' not in st.session_state:
        st.session_state.selected_sheet = None
    if 'column_mapping' not in st.session_state:
        st.session_state.column_mapping = {}
    if 'selected_rankings' not in st.session_state:
        st.session_state.selected_rankings = []
    if 'rows_per_slide' not in st.session_state:
        st.session_state.rows_per_slide = 10
    if 'generated_ppt_path' not in st.session_state:
        st.session_state.generated_ppt_path = None
    if '_file_just_uploaded' not in st.session_state:
        st.session_state._file_just_uploaded = False
    if 'selected_header_row' not in st.session_state:
        st.session_state.selected_header_row = None
    
    # Initialize debug info for debug logger
    if 'debug_info' not in st.session_state:
        from datetime import datetime
        st.session_state.debug_info = {
            'session_start': datetime.now().isoformat(),
            'steps_completed': [],
            'errors': [],
            'warnings': [],
            'file_info': {},
            'performance_metrics': {}
        }


def validate_session_state(step: int) -> tuple[bool, list[str]]:
    """
    Validate session state consistency for a given step
    Returns (is_valid, list_of_errors)
    """
    errors = []
    
    # Common validations
    if step < 1 or step > 6:
        errors.append(f"Invalid step number: {step}")
    
    # Step-specific validations
    if step >= 2:  # Steps 2+ require file upload
        if st.session_state.uploaded_file is None:
            errors.append("No file uploaded - cannot proceed beyond Step 1")
    
    if step >= 3:  # Steps 3+ require data processing to be completed
        if not hasattr(st.session_state, 'df') or st.session_state.df is None or st.session_state.df.empty:
            errors.append("No data available - please complete Step 2 data processing")
    
    if step >= 4:  # Steps 4+ require column mapping (even if empty)
        if not hasattr(st.session_state, 'column_mapping'):
            errors.append("Column mapping not initialized - please complete Step 3")
    
    if step >= 6:  # Step 6 requires configuration (Step 5 is skipped)
        if not hasattr(st.session_state, 'selected_rankings'):
            errors.append("Rankings not selected - please complete Step 4")
        if not st.session_state.selected_rankings:
            errors.append("At least one ranking must be selected in Step 4")
    
    return len(errors) == 0, errors


def reset_to_safe_state(target_step: int = 1):
    """Reset session state to a safe, known good state"""
    debug_logger.logger.info(f"Resetting to safe state - target step: {target_step}")
    
    # Clear step-dependent state
    if target_step <= 1:
        for key in ['uploaded_file', 'df', 'original_df', 'sheet_names', 'selected_sheet']:
            if key in st.session_state:
                del st.session_state[key]
    
    if target_step <= 2:
        for key in ['column_mapping', 'selected_header_row']:
            if key in st.session_state:
                del st.session_state[key]
    
    if target_step <= 3:
        for key in ['selected_rankings', 'filtered_data']:
            if key in st.session_state:
                del st.session_state[key]
    
    if target_step <= 4:
        for key in ['generated_ppt_path']:
            if key in st.session_state:
                del st.session_state[key]
    
    # Reset to target step
    st.session_state.current_step = target_step
    st.session_state._file_just_uploaded = False
    
    # Reinitialize required state
    initialize_session_state()
    
    show_success_message(f"Reset to Step {target_step} successfully!")


def with_error_boundary(step_func):
    """
    Decorator to add error boundary to step functions
    Catches exceptions and provides recovery options
    """
    def wrapper(*args, **kwargs):
        try:
            # Validate session state before executing step
            # Skip validation for Step 2 since that's where we process data
            current_step = st.session_state.current_step
            
            if current_step != 2:  # Don't validate Step 2 since it processes the data
                is_valid, errors = validate_session_state(current_step)
                
                if not is_valid:
                    st.error("🚨 **Session State Error Detected**")
                    for error in errors:
                        st.error(f"• {error}")
                    
                    st.markdown("---")
                    st.markdown("### 🔧 Recovery Options")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        if st.button("🔄 Reset to Step 1", type="primary"):
                            reset_to_safe_state(1)
                            st.rerun()
                    
                    with col2:
                        if current_step > 1:
                            if st.button(f"⬅️ Go Back to Step {current_step - 1}"):
                                reset_to_safe_state(current_step - 1)
                                st.rerun()
                    
                    with col3:
                        if st.button("🧹 Clear All Data"):
                            for key in list(st.session_state.keys()):
                                del st.session_state[key]
                            initialize_session_state()
                            st.rerun()
                    
                    return  # Don't execute the step function
            
            # Execute the step function
            return step_func(*args, **kwargs)
            
        except Exception as e:
            debug_logger.log_error(e, step_func.__name__)
            
            st.error(f"🚨 **Critical Error in {step_func.__name__}**")
            st.error(f"**Error:** {str(e)}")
            st.error(f"**Type:** {type(e).__name__}")
            
            # Show error details in expander
            with st.expander("🔍 Technical Details", expanded=False):
                import traceback
                st.code(traceback.format_exc())
            
            st.markdown("---")
            st.markdown("### 🔧 Recovery Options")
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Reset to Safe State", type="primary"):
                    reset_to_safe_state(max(1, st.session_state.current_step - 1))
                    st.rerun()
            
            with col2:
                if st.button("🏠 Return to Start"):
                    reset_to_safe_state(1)
                    st.rerun()
    
    return wrapper

initialize_session_state()

# Initialize processors
@st.cache_resource
def get_processors():
    return {
        'file_processor': FileProcessor(),
        'column_mapper': ColumnMapper(), 
        'ppt_generator': PowerPointGenerator(),
        'pdf_exporter': PDFExporter()
    }

processors = get_processors()

# Main header
def render_header():
    """Render main application header"""
    debug_logger.logger.debug("Rendering main header")
    st.markdown("""
        <div class="main-header">
            <h1>📊 90D-PPT Generator</h1>
            <p>Convert Excel/CSV data into branded PowerPoint presentations for 90-day planning cycles</p>
        </div>
    """, unsafe_allow_html=True)

# Progress indicator
def render_progress_indicator():
    """Render step progress indicator"""
    steps = [
        ("📁", "Upload File"),
        ("📋", "Select Data"), 
        ("🔗", "Map Columns"),
        ("⚙️", "Configure"),
        ("🎯", "Generate"),
        ("📥", "Download")
    ]
    
    current = st.session_state.current_step
    
    st.markdown("### Progress")
    
    # Create columns for progress steps
    cols = st.columns(len(steps))
    
    for i, (icon, step_name) in enumerate(steps, 1):
        with cols[i-1]:
            if i < current:
                # Completed step - green
                st.markdown(f"""
                    <div style='text-align: center; color: #28A745;'>
                        <div style='font-size: 1.5rem;'>✅</div>
                        <div style='font-size: 0.8rem; margin-top: 0.25rem;'>{icon} {step_name}</div>
                    </div>
                """, unsafe_allow_html=True)
            elif i == current:
                # Current step - blue
                st.markdown(f"""
                    <div style='text-align: center; color: #66A9F2;'>
                        <div style='font-size: 1.5rem;'>🔄</div>
                        <div style='font-size: 0.8rem; margin-top: 0.25rem; font-weight: bold;'>{icon} {step_name}</div>
                    </div>
                """, unsafe_allow_html=True)
            else:
                # Pending step - grey
                st.markdown(f"""
                    <div style='text-align: center; color: #A7A8AA;'>
                        <div style='font-size: 1.5rem;'>⏳</div>
                        <div style='font-size: 0.8rem; margin-top: 0.25rem;'>{icon} {step_name}</div>
                    </div>
                """, unsafe_allow_html=True)

# Step 1: File Upload
@with_error_boundary
def render_step_upload():
    """Render file upload step with error handling"""
    try:
        debug_logger.log_step_start(1, "File Upload")
        st.markdown("## Step 1: Upload Your Excel/CSV File")
    except Exception as e:
        debug_logger.log_error(e, "render_step_upload - initialization")
        show_error_message("Error initializing upload step", "Please refresh the page and try again.")
    
    st.markdown("""
        <div class="upload-area">
            <h3>📁 Drag and drop your file here or click to browse</h3>
            <p>Supported formats: Excel (.xlsx, .xls) and CSV (.csv)</p>
            <p>Maximum file size: 200MB</p>
        </div>
    """, unsafe_allow_html=True)
    
    try:
        uploaded_file = st.file_uploader(
            "Choose file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload your Excel or CSV file containing Jira data"
        )
        
        if uploaded_file is not None:
            # Validate file with comprehensive error handling
            try:
                if not is_supported_file(uploaded_file.name):
                    show_error_message(
                        f"Unsupported file type: {Path(uploaded_file.name).suffix}",
                        f"Please upload one of these file types: {', '.join(['.xlsx', '.xls', '.csv'])}",
                        f"Current file: {uploaded_file.name}"
                    )
                    return
                
                if not validate_file_size(uploaded_file.getvalue()):
                    file_size_mb = len(uploaded_file.getvalue()) / (1024 * 1024)
                    show_error_message(
                        f"File size too large: {file_size_mb:.1f}MB",
                        "Try these solutions:\n• Split your data into smaller files\n• Remove unnecessary columns\n• Export only the data you need",
                        f"Current limit: 200MB\nYour file: {file_size_mb:.1f}MB"
                    )
                    return
                
                # Store file in session state
                st.session_state.uploaded_file = uploaded_file
                
                # Log file info for debugging
                file_size = len(uploaded_file.getvalue()) / (1024 * 1024)  # Size in MB
                debug_logger.log_file_info(uploaded_file.name, file_size, uploaded_file.type)
                
                show_success_message(f"File '{uploaded_file.name}' uploaded successfully!")
                
                # Show file info
                st.markdown(f"**File Info:**")
                st.markdown(f"- Name: {uploaded_file.name}")
                st.markdown(f"- Size: {file_size:.2f} MB")
                st.markdown(f"- Type: {uploaded_file.type}")
                
                # Set flag to prevent showing duplicate button below
                st.session_state._file_just_uploaded = True
                    
            except Exception as e:
                debug_logger.log_error(e, "file_validation_and_processing")
                show_error_message(
                    "Error processing uploaded file",
                    "Please try uploading the file again or choose a different file.",
                    f"Technical details: {str(e)}"
                )
                return
    
    except Exception as e:
        debug_logger.log_error(e, "file_uploader_widget")
        show_error_message(
            "Error with file upload interface",
            "Please refresh the page and try again.",
            f"Technical details: {str(e)}"
        )
    
    # Show current file status and navigation - only if not just uploaded
    if st.session_state.uploaded_file is not None and not st.session_state.get('_file_just_uploaded', False):
        file = st.session_state.uploaded_file
        st.markdown(f"**Current file:** {file.name}")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Continue to Data Selection →", type="primary", key="upload_continue"):
                st.session_state.current_step = 2
                st.rerun()
        with col2:
            if st.button("Upload Different File", key="upload_different"):
                st.session_state.uploaded_file = None
                st.session_state.current_step = 1
                st.rerun()
    
    # Show continue button for just uploaded files
    elif st.session_state.uploaded_file is not None and st.session_state.get('_file_just_uploaded', False):
        if st.button("Continue to Data Selection →", type="primary", key="upload_continue_new"):
            st.session_state.current_step = 2
            # Clear the flag
            st.session_state._file_just_uploaded = False
            st.rerun()

# Step 2: Data Selection 
@with_error_boundary
def render_step_data_selection():
    """Render data selection step"""
    debug_logger.log_step_start(2, "Data Selection")
    st.markdown("## Step 2: Select Your Data")
    
    if st.session_state.uploaded_file is None:
        show_error_message(
            "No file uploaded",
            "Please go back to Step 1 to upload your Excel or CSV file",
            "Required: .xlsx, .xls, or .csv file"
        )
        return
    
    # Process file
    file_processor = processors['file_processor']
    
    if 'df' not in st.session_state or st.session_state.df is None:
        with st.spinner("Processing file..."):
            df, sheet_names, error_msg = file_processor.read_file(st.session_state.uploaded_file)
            
            if error_msg:
                show_error_message(
                    "Failed to read your file",
                    "Try these solutions:\n• Check if the file is corrupted\n• Ensure the file is not password protected\n• Try re-saving the file in Excel",
                    f"Technical error: {error_msg}"
                )
                return
            
            st.session_state.df = df
            st.session_state.sheet_names = sheet_names
            # Store original dataframe immediately after processing
            if 'original_df' not in st.session_state:
                st.session_state.original_df = df.copy()
    
    # Sheet selection for Excel files
    if len(st.session_state.sheet_names) > 1:
        st.markdown("### 📄 Select Sheet")
        selected_sheet = st.selectbox(
            "Choose the sheet containing your data:",
            st.session_state.sheet_names,
            index=0 if st.session_state.selected_sheet is None else st.session_state.sheet_names.index(st.session_state.selected_sheet)
        )
        
        if selected_sheet != st.session_state.selected_sheet:
            st.session_state.selected_sheet = selected_sheet
            # Re-read the selected sheet
            with st.spinner(f"Loading sheet '{selected_sheet}'..."):
                new_df = file_processor.read_excel_sheet(st.session_state.uploaded_file, selected_sheet)
                if new_df is not None:
                    st.session_state.df = new_df
                    # Update original_df when sheet changes
                    st.session_state.original_df = new_df.copy()
                    show_success_message(f"Sheet '{selected_sheet}' loaded successfully!")
                    st.rerun()
    
    # Intelligent Header Detection and Auto-Application
    if st.session_state.df is not None and not st.session_state.df.empty:
        st.markdown("### 📝 Data Processing")
        
        # Ensure we have original_df available
        if 'original_df' not in st.session_state or st.session_state.original_df is None:
            st.session_state.original_df = st.session_state.df.copy()
        
        # Intelligent header detection and auto-application
        detected_header_row = file_processor.detect_headers(st.session_state.original_df)
        
        # Auto-apply headers intelligently without user intervention
        if st.session_state.df.columns[0] == 'Column_0' or str(st.session_state.df.columns[0]).startswith('Column_'):
            with st.spinner("🔍 Detecting and applying headers intelligently..."):
                processed_df = file_processor.apply_header_row(
                    st.session_state.original_df.copy(), 
                    detected_header_row
                )
                st.session_state.df = processed_df
                st.session_state.selected_header_row = detected_header_row + 1
                
                show_success_message(
                    f"Headers automatically detected and applied from row {detected_header_row + 1}",
                    f"Found {len(processed_df.columns)} columns with meaningful headers"
                )
        
        # Show current processed data with clean headers
        st.markdown("### 👀 Your Data with Headers Applied")
        preview_df = file_processor.get_data_preview(st.session_state.df, max_rows=8)
        st.dataframe(preview_df, use_container_width=True)
        
        # Optional: Show original data in expander for reference
        with st.expander("🔍 View Original Data Structure", expanded=False):
            st.markdown("*Original data before header processing:*")
            preview_for_header = st.session_state.original_df.head(5).reset_index(drop=True)
            preview_for_header.index = preview_for_header.index + 1
            st.dataframe(preview_for_header, use_container_width=True)
            
            # Advanced options for manual override (collapsed by default)
            st.markdown("**Manual Header Override** *(if needed)*")
            col1, col2 = st.columns([2, 1])
            with col1:
                manual_header_row = st.number_input(
                    "Override header row:",
                    min_value=1,
                    max_value=min(10, len(st.session_state.original_df)),
                    value=detected_header_row + 1,
                    step=1,
                    key="manual_header_override",
                    help="Only change if auto-detection is incorrect"
                )
            with col2:
                if st.button("Apply Override", type="secondary"):
                    with st.spinner(f"Applying header row {manual_header_row}..."):
                        processed_df = file_processor.apply_header_row(
                            st.session_state.original_df.copy(), 
                            manual_header_row - 1
                        )
                        st.session_state.df = processed_df
                        st.session_state.selected_header_row = manual_header_row
                        show_success_message(f"Header row {manual_header_row} applied successfully!")
                        st.rerun()
        
        # Show data summary
        summary = file_processor.get_data_summary(st.session_state.df)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Rows", summary['total_rows'])
        with col2:
            st.metric("Total Columns", summary['total_columns'])
        with col3:
            st.metric("Sheet", st.session_state.selected_sheet or st.session_state.sheet_names[0] if st.session_state.sheet_names else "N/A")
    
    # Navigation buttons
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back to Upload", key="step2_back"):
            st.session_state.current_step = 1
            st.rerun()
    with col2:
        if st.session_state.df is not None and not st.session_state.df.empty:
            if st.button("Continue to Column Mapping →", type="primary", key="step2_continue"):
                st.session_state.current_step = 3
                st.rerun()
        else:
            st.button("Continue to Column Mapping →", disabled=True, key="step2_continue_disabled")

# Step 3: Column Mapping
@with_error_boundary
def render_step_column_mapping():
    """Render column mapping step"""
    st.markdown("## Step 3: Map Your Columns")
    
    st.markdown("""
        <div class="info-panel" style="background-color: #e8f4fd; border-left: 4px solid #66A9F2; padding: 1rem; margin: 1rem 0; border-radius: 5px;">
            <h4 style="color: #00175A; margin-bottom: 0.5rem;">ℹ️ Column Mapping is Optional</h4>
            <p style="margin-bottom: 0;">Map your data columns to PowerPoint fields, or leave unmapped for blank columns. All mappings are optional - you can proceed even without mapping any columns!</p>
        </div>
    """, unsafe_allow_html=True)
    
    if st.session_state.df is None or st.session_state.df.empty:
        st.error("No data available. Please go back to Step 2.")
        return
    
    column_mapper = processors['column_mapper']
    
    # Render column mapping interface
    mapping = column_mapper.render_mapping_interface(st.session_state.df)
    
    if mapping:
        # Validate mapping
        is_valid, errors = column_mapper.validate_mapping(st.session_state.df, mapping)
        
        if errors:
            st.markdown("### ❌ Validation Errors")
            for error in errors:
                st.error(error)
        
        if is_valid:
            # Show preview of mapped data
            column_mapper.preview_mapped_data(st.session_state.df, mapping, max_rows=5)
            
            # Store mapping in session state
            st.session_state.column_mapping = mapping
    
    # Navigation buttons
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back to Data Selection", key="step3_back"):
            st.session_state.current_step = 2
            st.rerun()
    with col2:
        # Allow progression regardless of mapping count - all columns are now optional
        if st.button("Continue to Configuration →", type="primary", key="step3_continue"):
            st.session_state.current_step = 4
            st.rerun()

# Step 4: Configuration
@with_error_boundary
def render_step_configuration():
    """Render configuration step"""
    st.markdown("## Step 4: Configure Your Presentation")
    
    if not st.session_state.column_mapping:
        st.error("No column mapping available. Please go back to Step 3.")
        return
    
    # Clean configuration interface without redundant tips
    
    file_processor = processors['file_processor']
    
    # Top section: Key configuration in compact layout
    st.markdown("---")
    config_col1, config_col2 = st.columns([1, 1])
    
    with config_col1:
        st.markdown("### 🎯 Data to Include")
        # Get available rankings from data
        ranking_col = st.session_state.column_mapping.get('ranking')
        if ranking_col and ranking_col in st.session_state.df.columns:
            ranking_counts = file_processor.get_ranking_counts(st.session_state.df, ranking_col)
            
            if ranking_counts:
                st.markdown("**Select Rankings:**")
                selected_rankings = []
                for ranking, count in ranking_counts.items():
                    if st.checkbox(f"{ranking} ({count} items)", value=True, key=f"ranking_{ranking}"):
                        selected_rankings.append(ranking)
                
                st.session_state.selected_rankings = selected_rankings
                
                if not selected_rankings:
                    st.warning("⚠️ Select at least one ranking.")
            else:
                st.info("ℹ️ No valid ranking data found. Will create general presentation.")
                st.session_state.selected_rankings = ['General']
        else:
            st.info("ℹ️ No ranking column mapped. Will create general presentation with all data.")
            st.session_state.selected_rankings = ['General']
    
    with config_col2:
        st.markdown("### 📊 Layout Settings")
        
        rows_per_slide = st.number_input(
            "Rows per slide:",
            min_value=5,
            max_value=25,
            value=st.session_state.rows_per_slide,
            step=1,
            help="Number of data rows to display on each slide"
        )
        st.session_state.rows_per_slide = rows_per_slide
        
        st.markdown(f"""
            **Current:** {rows_per_slide} rows per slide
            
            *Controls data density per slide.*
        """)
    
    # Show generation summary
    if st.session_state.selected_rankings and ranking_counts:
        st.markdown("### 📋 Generation Summary")
        
        total_items = sum(ranking_counts[ranking] for ranking in st.session_state.selected_rankings if ranking in ranking_counts)
        total_slides = 0
        
        for ranking in st.session_state.selected_rankings:
            if ranking in ranking_counts:
                items = ranking_counts[ranking]
                slides = (items + rows_per_slide - 1) // rows_per_slide  # Ceiling division
                total_slides += slides
                st.markdown(f"- **{ranking}**: {items} items → {slides} slide(s)")
        
        st.markdown(f"**Total slides to generate: {total_slides}**")
        
        # Status distribution
        if 'status' in st.session_state.column_mapping:
            status_col = st.session_state.column_mapping['status']
            if status_col in st.session_state.df.columns:
                status_counts = file_processor.get_status_counts(st.session_state.df, status_col)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Closed Items", status_counts.get('closed', 0))
                with col2:
                    st.metric("Non-Closed Items", status_counts.get('non-closed', 0))
    
    # Navigation buttons
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back to Column Mapping", key="step4_back"):
            st.session_state.current_step = 3
            st.rerun()
    with col2:
        # Can always generate now since all mappings are optional
        can_generate = (
            st.session_state.selected_rankings and 
            len(st.session_state.selected_rankings) > 0
        )
        
        if can_generate:
            if st.button("Generate Presentation →", type="primary", key="step4_generate"):
                # Skip Step 5 (slide preview) and go directly to generation
                st.session_state.current_step = 6
                st.rerun()
        else:
            st.button("Generate Presentation →", disabled=True, help="Please select at least one ranking", key="step4_generate_disabled")

# Step 5: Generation and Download (Combined)
@with_error_boundary
def render_step_generation_and_download():
    """Streamlined generation and download - no redundant previews or clicks"""
    st.markdown("## Step 5: Generate & Download Presentation")
    
    if not all([st.session_state.df is not None, st.session_state.column_mapping, st.session_state.selected_rankings]):
        st.error("Missing required data. Please complete previous steps.")
        return
    
    file_processor = processors['file_processor']
    ppt_generator = processors['ppt_generator']
    
    # Auto-generate presentation without previews
    if 'generated_ppt_path' not in st.session_state or st.session_state.generated_ppt_path is None:
        st.markdown("### 🚀 Generating Your Presentation...")
        st.markdown("*This will take a moment - we're creating slides with your data.*")
        
        with st.spinner("🎨 Creating PowerPoint slides..."):
            data_by_ranking = file_processor.prepare_data_for_ppt(
                st.session_state.df, 
                st.session_state.column_mapping
            )
            
            # Filter by selected rankings
            filtered_data = {
                ranking: data_by_ranking[ranking] 
                for ranking in st.session_state.selected_rankings 
                if ranking in data_by_ranking and not data_by_ranking[ranking].empty
            }
            
            if not filtered_data:
                st.error("❌ No data found for selected rankings.")
                return
            
            # Direct generation without previews
            try:
                ppt_path = ppt_generator.create_presentation(
                    filtered_data,
                    st.session_state.rows_per_slide
                )
                
                # Verify the file was actually created
                if ppt_path and Path(ppt_path).exists():
                    st.session_state.generated_ppt_path = ppt_path
                    st.session_state.filtered_data = filtered_data
                    
                    # Show immediate success and download options
                    st.rerun()
                else:
                    show_error_message("Failed to generate presentation file. Please try again.")
                    return
                    
            except Exception as e:
                debug_logger.log_error(e, "powerpoint_generation")
                show_error_message(
                    "Failed to generate PowerPoint presentation",
                    "Try these solutions:\n• Check if your data has all required columns mapped\n• Try reducing the number of rows per slide\n• Ensure your file doesn't contain special characters",
                    f"Error: {str(e)}"
                )
                return
    else:
        # Presentation generated successfully - show download options immediately
        st.markdown("### 🎉 Presentation Generated Successfully!")
        show_success_message("Your 90-day planning PowerPoint is ready!")
        
        # Show file info
        if Path(st.session_state.generated_ppt_path).exists():
            file_size = Path(st.session_state.generated_ppt_path).stat().st_size / (1024 * 1024)
            filename = Path(st.session_state.generated_ppt_path).name
            st.markdown(f"**File:** {filename} ({file_size:.2f} MB)")
            
            # Presentation summary
            if st.session_state.filtered_data:
                total_items = sum(len(df) for df in st.session_state.filtered_data.values())
                rankings_included = len(st.session_state.filtered_data)
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Rankings", rankings_included)
                with col2:
                    st.metric("Total Items", total_items)
                with col3:
                    st.metric("File Size", f"{file_size:.1f} MB")
            
            # Download Options (consolidated under single header)
            st.markdown("### 📥 Download Options")
            
            # Read file for download
            with open(st.session_state.generated_ppt_path, 'rb') as file:
                ppt_bytes = file.read()
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📄 Download PowerPoint",
                    data=ppt_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary"
                )
            with col2:
                if st.button("📖 Generate PDF Version", help="Convert to PDF format"):
                    show_info_message("PDF conversion feature coming soon!")
        
        else:
            st.warning("⚠️ Generated file not found. Please try regenerating.")
            if st.button("🔄 Regenerate Presentation"):
                if 'generated_ppt_path' in st.session_state:
                    del st.session_state.generated_ppt_path
                st.rerun()
    
    # Navigation
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Back to Configuration"):
            st.session_state.current_step = 4
            st.rerun()

# Old functions - keeping for compatibility but they won't be called
@with_error_boundary  
def render_step_generation():
    """Legacy function - redirects to combined function"""
    render_step_generation_and_download()

# Step 6: Download
@with_error_boundary
def render_step_download():
    """Render download step"""
    st.markdown("## Step 6: Download Your Presentation")
    
    if 'generated_ppt_path' not in st.session_state:
        st.error("No presentation generated. Please go back to Step 5.")
        return
    
    # Success message
    st.markdown("""
        <div class="success-message">
            <h3>🎉 Presentation Generated Successfully!</h3>
            <p>Your branded 90-day planning PowerPoint presentation is ready for download.</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Presentation summary
    if 'filtered_data' in st.session_state:
        st.markdown("### 📊 Presentation Summary")
        
        # Calculate summary from filtered data
        total_items = sum(len(df) for df in st.session_state.filtered_data.values())
        rankings_count = len([k for k, v in st.session_state.filtered_data.items() if len(v) > 0])
        
        # Calculate estimated slides (using same logic as ppt_generator)
        import math
        rows_per_slide = st.session_state.get('rows_per_slide', 10)
        total_slides = sum(math.ceil(len(df) / rows_per_slide) for df in st.session_state.filtered_data.values() if len(df) > 0)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Slides", total_slides)
        with col2:
            st.metric("Rankings Included", rankings_count)
        with col3:
            st.metric("Total Items", total_items)
        
        # Rankings breakdown
        st.markdown("**Breakdown by Ranking:**")
        for ranking, df in st.session_state.filtered_data.items():
            if len(df) > 0:
                slides = math.ceil(len(df) / rows_per_slide)
                # Calculate status counts
                closed_count = 0
                non_closed_count = 0
                if '_status_category' in df.columns:
                    closed_count = (df['_status_category'] == 'closed').sum()
                    non_closed_count = (df['_status_category'] == 'non-closed').sum()
                st.markdown(f"- **{ranking}**: {slides} slide(s), {len(df)} items ({closed_count} closed, {non_closed_count} non-closed)")
    
    # Download options
    st.markdown("### 📥 Download Options")
    
    # Check if presentation was generated
    if 'generated_ppt_path' not in st.session_state or st.session_state.generated_ppt_path is None:
        show_error_message(
            "No presentation available for download",
            "Please go back to Step 5 to generate your PowerPoint presentation first",
            "The presentation generation step must complete successfully before downloading"
        )
        return
    
    # Verify file exists
    if not Path(st.session_state.generated_ppt_path).exists():
        show_error_message(
            "Presentation file not found",
            "The generated file may have been deleted. Please regenerate your presentation from Step 5",
            f"Expected file: {st.session_state.generated_ppt_path}"
        )
        return
    
    # Read the generated file
    try:
        with open(st.session_state.generated_ppt_path, 'rb') as f:
            ppt_data = f.read()
        
        # Generate filename
        upload_filename = st.session_state.uploaded_file.name if st.session_state.uploaded_file else "data"
        base_name = Path(upload_filename).stem
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        download_filename = f"90D_PPT_{base_name}_{timestamp}.pptx"
        
        # Download buttons in side-by-side layout
        col1, col2 = st.columns(2)
        
        with col1:
            # PowerPoint download button
            st.download_button(
                label="📊 Download PowerPoint (.pptx)",
                data=ppt_data,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary"
            )
            # File info
            file_size_mb = len(ppt_data) / (1024 * 1024)
            st.markdown(f"**PowerPoint file size:** {file_size_mb:.2f} MB")
        
        with col2:
            # PDF Export section
            pdf_exporter = processors['pdf_exporter']
            export_info = pdf_exporter.get_export_info()
            
            if export_info['available']:
                if st.button("📄 Generate PDF Version", type="secondary"):
                    with st.spinner("Converting to PDF... This may take a moment."):
                        try:
                            # Try to export using the generated PowerPoint file first
                            pdf_path = None
                            
                            if export_info['method'] == 'powerpoint':
                                pdf_path = pdf_exporter.export_presentation_to_pdf(
                                    ppt_path=st.session_state.generated_ppt_path
                                )
                            elif export_info['method'] == 'reportlab':
                                # Use the data to recreate as PDF
                                pdf_path = pdf_exporter.export_presentation_to_pdf(
                                    data_by_ranking=st.session_state.get('filtered_data', {}),
                                    rows_per_slide=st.session_state.rows_per_slide
                                )
                            
                            if pdf_path and Path(pdf_path).exists():
                                # Read PDF data for download
                                with open(pdf_path, 'rb') as f:
                                    pdf_data = f.read()
                                
                                # Generate PDF filename
                                pdf_filename = download_filename.replace('.pptx', '.pdf')
                                
                                st.download_button(
                                    label="📄 Download PDF",
                                    data=pdf_data,
                                    file_name=pdf_filename,
                                    mime="application/pdf",
                                    type="secondary"
                                )
                                
                                pdf_size_mb = len(pdf_data) / (1024 * 1024)
                                st.markdown(f"**PDF file size:** {pdf_size_mb:.2f} MB")
                                
                                show_success_message("PDF version generated successfully!")
                            else:
                                show_error_message(
                                    "Failed to generate PDF",
                                    f"PDF export method '{export_info['method']}' did not work. You can still download the PowerPoint version above.",
                                    f"Export method: {export_info['method']}\nPlatform: {export_info['platform']}"
                                )
                        
                        except Exception as e:
                            show_error_message(
                                "PDF conversion failed",
                                "You can still download the PowerPoint version above. PDF export may require additional software.",
                                f"Error: {str(e)}\nMethod: {export_info['method']}"
                            )
            else:
                st.markdown("### 📄 PDF Export")
                show_info_message(
                    "PDF export not available",
                    f"PDF export requires additional libraries. Method attempted: {export_info['method']}\nPlatform: {export_info['platform']}"
                )
        
    except FileNotFoundError:
        show_error_message("Presentation file not found. Please regenerate the presentation.")
    except PermissionError:
        show_error_message("Cannot access presentation file. Please try regenerating.")
    except Exception as e:
        show_error_message(f"Error preparing download: {str(e)}")
        st.markdown("**Debug info:**")
        st.code(f"Generated PPT path: {st.session_state.get('generated_ppt_path', 'None')}")
        st.code(f"File exists: {Path(st.session_state.get('generated_ppt_path', '')).exists() if st.session_state.get('generated_ppt_path') else 'N/A'}")
    
    # Additional spacing for better visual separation
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    # Additional actions
    st.markdown("---")
    st.markdown("### 🔄 What's Next?")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("🆕 Process New File", type="secondary"):
            # Clear session state safely
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            # Reinitialize session state
            initialize_session_state()
            st.rerun()
    
    with col2:
        if st.button("⚙️ Regenerate with Different Settings"):
            st.session_state.current_step = 4
            # Clear generated data but keep everything else
            if 'generated_ppt_path' in st.session_state:
                del st.session_state.generated_ppt_path
            if 'filtered_data' in st.session_state:
                del st.session_state.filtered_data
            st.rerun()
    
    with col3:
        if st.button("🔙 Back to Generation"):
            st.session_state.current_step = 5
            st.rerun()
    
    # Feedback section
    st.markdown("---")
    st.markdown("### 📝 How did we do?")
    st.markdown("We hope this tool saved you time! If you encounter any issues or have suggestions for improvement, please let us know.")
    
    # Show generation info
    if st.session_state.uploaded_file:
        st.markdown("**Session Information:**")
        st.markdown(f"- Original file: {st.session_state.uploaded_file.name}")
        st.markdown(f"- Sheet: {st.session_state.selected_sheet or 'Default'}")
        st.markdown(f"- Rows per slide: {st.session_state.rows_per_slide}")
        st.markdown(f"- Rankings: {', '.join(st.session_state.selected_rankings)}")

# Sidebar
def render_sidebar():
    """Render sidebar with additional information and controls"""
    st.sidebar.markdown("## 📊 90D-PPT Generator")
    st.sidebar.markdown("---")
    
    # Current step info
    st.sidebar.markdown(f"**Current Step:** {st.session_state.current_step}/6")
    
    # Render debug panel
    debug_logger.render_debug_panel()
    
    # File info if uploaded
    if st.session_state.uploaded_file:
        st.sidebar.markdown("### 📁 Current File")
        st.sidebar.markdown(f"**Name:** {st.session_state.uploaded_file.name}")
        file_size = len(st.session_state.uploaded_file.getvalue()) / (1024 * 1024)
        st.sidebar.markdown(f"**Size:** {file_size:.2f} MB")
    
    st.sidebar.markdown("---")
    
    # Help section
    st.sidebar.markdown("### ❓ Need Help?")
    st.sidebar.markdown("""
    **Supported Formats:**
    - Excel (.xlsx, .xls)
    - CSV (.csv)
    
    **Available Columns (All Optional):**
    - JIRA Key/ID
    - Description/Summary
    - Status
    - Ranking/Priority
    - Component/Team
    - Risks/Issues
    
    **Ranking Values:**
    - "Accepted"
    - "Up Next"
    - "Maybe"
    - "Likely No"
    
    *Note: All column mappings are optional. Unmapped columns will appear blank in the PowerPoint output.*
    """)
    
    if st.sidebar.button("🔄 Reset Application", key="sidebar_reset"):
        # Clear all session state safely
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        # Reinitialize session state
        initialize_session_state()
        st.rerun()

# Main application
def main():
    """Main application logic with comprehensive error handling"""
    try:
        render_header()
        render_progress_indicator()
        
        # Render current step with individual error handling
        try:
            if st.session_state.current_step == 1:
                render_step_upload()
            elif st.session_state.current_step == 2:
                render_step_data_selection()
            elif st.session_state.current_step == 3:
                render_step_column_mapping()
            elif st.session_state.current_step == 4:
                render_step_configuration()
            elif st.session_state.current_step == 5:
                # Step 5 is skipped - redirect to Step 6
                st.session_state.current_step = 6
                st.rerun()
            elif st.session_state.current_step == 6:
                render_step_generation_and_download()
            else:
                # Handle invalid step
                st.error("❌ Invalid step. Resetting to step 1.")
                st.session_state.current_step = 1
                st.rerun()
                
        except Exception as e:
            debug_logger.log_error(e, f"render_step_{st.session_state.current_step}")
            show_error_message(
                f"Error in step {st.session_state.current_step}",
                "Please try refreshing the page. If the problem persists, try resetting the application using the button in the sidebar.",
                f"Technical details: {str(e)}"
            )
        
        # Render sidebar with error handling
        try:
            render_sidebar()
        except Exception as e:
            debug_logger.log_error(e, "render_sidebar")
            st.sidebar.error("❌ Error loading sidebar. Please refresh the page.")
            
    except Exception as e:
        debug_logger.log_error(e, "main_application")
        st.error("❌ Critical application error. Please refresh the page.")
        st.exception(e)

if __name__ == "__main__":
    main()