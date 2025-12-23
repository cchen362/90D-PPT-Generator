# 90D-PPT Generator

A web-based application that converts Excel/CSV data into branded PowerPoint presentations for 90-day planning cycles. Built with Streamlit and designed for non-technical team members who need to create standardized PowerPoint slides from Jira/project data.

## Features

- **File Upload Support**: Excel (.xlsx, .xls) and CSV files up to 200MB
- **Smart Column Mapping**: Intelligent auto-detection and mapping of data columns
- **Multi-Sheet Excel Support**: Select specific sheets from Excel workbooks
- **Data Filtering**: Filter by ranking categories (Accepted, Up Next, Maybe, Likely No)
- **Brand Compliant**: Professional styling with enterprise brand colors
- **Configurable Layout**: Adjustable rows per slide (5-25 rows)
- **Slide Preview**: Preview slides before generation
- **Status Tracking**: Automatic status categorization and counting
- **PowerPoint Export**: Download as .pptx format

## Required Data Columns

Your Excel/CSV file should contain these columns (column names can vary):

### Required Columns
- **JIRA Key/ID**: Unique identifier (e.g., "Key", "JIRA", "Ticket")
- **Description**: Task description (e.g., "Description", "Summary")
- **Status**: Current status (e.g., "Status", "State")
- **Ranking**: Priority ranking (e.g., "Ranking", "Priority")

### Optional Columns
- **Region/POS**: Geographic region or point of sale (e.g., "Region", "POS", "Location")
- **Risks/Issues**: Risk items (e.g., "Risk", "Issues", "Notes")

## Supported Ranking Values

The application processes these ranking values (case-insensitive):
- **Accepted**: "Accepted", "Accept", "Approved", "Confirmed"
- **Up Next**: "Up Next", "Next", "Upcoming"
- **Maybe**: "Maybe", "Possible", "Perhaps", "Potentially"
- **Likely No**: "Likely No", "Probably Not", "Unlikely"

*All other ranking values are ignored.*

## Installation

1. **Clone or download** this repository
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the application**:
   ```bash
   streamlit run app.py
   ```

## Usage Workflow

1. **Upload File**: Drag and drop your Excel or CSV file
2. **Select Data**: Choose sheet (for Excel) and verify header row
3. **Map Columns**: Map your data columns to PowerPoint fields
4. **Configure**: Select rankings and set rows per slide
5. **Generate**: Create slide previews and generate PowerPoint
6. **Download**: Download your branded presentation

## Brand Styling

The application uses enterprise brand colors:
- **Navy**: #00175A (headers, accents)
- **Light Blue**: #66A9F2 (highlights, status indicators)  
- **Grey**: #A7A8AA (borders, secondary text)

PowerPoint slides feature:
- Professional table formatting with alternating row colors
- Status summaries in top-right corner
- Branded title slides with green circle icons
- Page numbering (P1, P2, etc.)
- Consistent typography and spacing

## PowerPoint Output Structure

Each slide contains:
- **Title**: Ranking name (e.g., "Accepted JIRA")
- **Subtitle**: "Will be worked this 90"
- **Status Summary**: Count of closed vs non-closed items
- **Data Table** with 6 columns:
  1. JIRA (from mapped data)
  2. Target Complete Date (blank for user input)
  3. Description (from mapped data)
  4. Status (from mapped data)
  5. Risks/Issues/Watch Items (from mapped data or blank)
  6. Region (from mapped data)

## File Structure

```
90d-ppt-generator/
├── app.py                 # Main Streamlit application
├── src/
│   ├── file_processor.py  # Excel/CSV processing logic
│   ├── ppt_generator.py   # PowerPoint creation engine
│   ├── column_mapper.py   # Column mapping interface
│   └── utils.py          # Helper functions and utilities
├── assets/
│   ├── styles.css        # Custom CSS styling
│   └── brand_colors.py   # Brand color definitions
├── requirements.txt      # Python dependencies
└── README.md            # This file
```

## Dependencies

- **streamlit**: Web application framework
- **pandas**: Data processing and analysis
- **openpyxl**: Excel file reading/writing
- **python-pptx**: PowerPoint file creation
- **pillow**: Image processing for previews

## Troubleshooting

### Common Issues

**File Upload Errors**:
- Ensure file is under 200MB
- Check file format (.xlsx, .xls, .csv)
- Try re-saving Excel file if corrupted

**Column Mapping Issues**:
- Verify all required columns are mapped
- Check that mapped columns contain data
- Ensure ranking column has valid values

**Generation Errors**:
- Confirm at least one ranking is selected
- Check that filtered data is not empty
- Verify column mapping is complete

### Performance Notes

- Large files (>50MB) may take longer to process
- Files with 1000+ rows are supported but may require patience
- Consider splitting very large datasets across multiple presentations

## Future Enhancements

Planned features for future releases:
- PDF export functionality
- Batch processing of multiple files
- Custom slide templates
- Integration with Jira API
- Save column mapping templates
- Advanced filtering options

## Support

For issues, bugs, or feature requests, please review the troubleshooting section above. The application includes built-in error handling and user-friendly error messages to help diagnose issues.

## Technical Notes

- Built with Streamlit for ease of deployment and use
- Supports Python 3.9+ (excluding 3.9.7)
- Memory-efficient processing for large files
- Professional enterprise-grade output formatting
- Cross-platform compatibility (Windows, macOS, Linux)