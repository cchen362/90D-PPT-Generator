# Product Requirements Document (PRD)
## Excel to PowerPoint Generator for 90-Day Planning

### 1. Executive Summary

**Product Name:** 90D-PPT Generator  
**Purpose:** A web-based application that converts Excel/CSV data into formatted PowerPoint presentations for 90-day planning cycles  
**Target Users:** Non-technical team members who need to create standardized PowerPoint slides from Jira/project data  
**Deployment:** Docker container on POP server  

### 2. Problem Statement

Teams regularly need to create PowerPoint presentations from Excel data containing Jira tickets, project status, and planning information. Currently, this process is manual, time-consuming, and inconsistent. Non-technical users need a simple interface to upload Excel files and generate standardized, professionally formatted PowerPoint slides.

### 3. Core Requirements

#### 3.1 Functional Requirements

**File Upload & Processing:**
- Support multiple Excel formats: .xlsx, .xls, .csv
- Handle files up to 100-200MB
- Auto-detect header rows (typically first row, but may vary)
- Support multi-sheet Excel files with sheet selection
- Intelligent column mapping with user confirmation

**Data Processing:**
- Filter data by ticket rankings: "Accepted", "Up Next", "Maybe", "Likely No"
- Ignore all other ranking values
- Generate separate slides for each ranking category
- Configurable rows per slide (default: 10, user adjustable)
- Handle cases with fewer than 10 rows gracefully

**PowerPoint Generation:**
- Create slides with ranking as title
- Generate tables with 6 columns:
  1. JIRA (mapped from "Key" column)
  2. Target Complete Date (blank - user fills later)
  3. Description (mapped from "Description" column)
  4. Status (mapped from "Status" column)
  5. Risks/Issues/Watch Items (blank or user-mapped)
  6. Component (mapped from "Component/s" column)
- Add status summary in top-right corner: "In Progress/On Hold/Pending Review: X, Closed: Y"
- Professional enterprise styling with brand colors and fonts

**Export Options:**
- Download as .pptx format
- PDF export capability
- Slide preview thumbnails in browser

#### 3.2 User Interface Requirements

**Workflow:**
1. Landing page with upload area
2. File upload with drag-and-drop support
3. Sheet selection (if multiple sheets)
4. Data preview with column mapping interface
5. Ranking selection (multi-select checkboxes)
6. Rows-per-slide configuration (slider)
7. Generation and preview
8. Download options

**Design Requirements:**
- Brand colors: Navy (#00175A), Light Blue (#66A9F2), Grey (#A7A8AA)
- Typography: Benton Sans (primary), Arial (fallback)
- Professional, clean enterprise styling
- Responsive design for desktop/tablet use
- Toast notifications for errors and success states

#### 3.3 Technical Requirements

**Architecture:**
- Streamlit-based web application
- Python backend with pandas, openpyxl, python-pptx
- Docker containerization ready
- File processing with progress indicators
- Error handling with user-friendly messages

**Performance:**
- File upload progress indicators
- Processing status updates
- Responsive UI during file operations
- Memory-efficient handling of large files

### 4. User Stories

**As a project manager, I want to:**
- Upload an Excel file and quickly generate standardized PowerPoint slides
- Preview slides before downloading to ensure accuracy
- Configure how many rows appear per slide based on presentation needs
- Map columns correctly when Excel headers vary between reports

**As a non-technical user, I want to:**
- Use a simple drag-and-drop interface without needing technical knowledge
- Get clear feedback if there are errors in my data
- Download both PowerPoint and PDF versions of my presentation
- See exactly what my slides will look like before generating

### 5. Technical Architecture

#### 5.1 Technology Stack
- **Frontend:** Streamlit
- **Backend:** Python 3.9+
- **Data Processing:** pandas, openpyxl, xlrd
- **PowerPoint Generation:** python-pptx
- **PDF Export:** python-pptx + additional PDF conversion
- **Deployment:** Docker container

#### 5.2 Key Libraries
```python
streamlit>=1.28.0
pandas>=2.0.0
openpyxl>=3.1.0
xlrd>=2.0.0
python-pptx>=0.6.21
pillow>=10.0.0  # for slide previews
```

#### 5.3 File Structure
```
90d-ppt-generator/
├── app.py                 # Main Streamlit app
├── src/
│   ├── file_processor.py  # Excel/CSV processing
│   ├── ppt_generator.py   # PowerPoint creation
│   ├── column_mapper.py   # Column mapping logic
│   └── utils.py          # Helper functions
├── assets/
│   ├── styles.css        # Custom styling
│   └── brand_colors.py   # Brand color definitions
├── requirements.txt
├── Dockerfile
└── README.md
```

### 6. Data Mapping Specifications

#### 6.1 Expected Excel Columns
| Required Data | Typical Column Names | Mapping Priority |
|---------------|---------------------|------------------|
| Jira Number | "Key", "Jira", "Ticket" | High |
| Status | "Status", "State" | High |
| Description | "Description", "Summary" | High |
| Component | "Component/s", "Components", "Team" | Medium |
| Ranking | "Ranking", "Priority", "Category" | High |
| Risk Items | "Risk", "Issues", "Notes" | Low (optional) |

#### 6.2 Ranking Values
- **Process:** "Accepted", "Up Next", "Maybe", "Likely No"
- **Ignore:** All other values
- **Case Sensitivity:** Handle variations in capitalization

#### 6.3 Status Categorization
- **Non-Closed:** "In Progress", "On Hold", "Pending Review", "New", "Open"
- **Closed:** "Closed", "Done", "Resolved", "Completed"

### 7. PowerPoint Specifications

#### 7.1 Slide Layout
- **Title:** Ranking name (e.g., "Accepted JIRA") with green circle icon
- **Subtitle:** "Will be worked this 90" (matching template style)
- **Table:** 6 columns as specified
- **Status Summary:** Top-right corner with counts
- **Footer:** Page numbering (P1, P2, etc.)

#### 7.2 Table Styling
- **Header Row:** Black background (#000000), white text
- **Data Rows:** Alternating light grey (#F5F5F5) and white
- **Borders:** Clean professional borders
- **Font:** Benton Sans 10pt for data, 12pt for headers
- **Status Indicators:** Blue circles for "In Progress" (matching template)

#### 7.3 Brand Compliance
- **Primary Navy:** #00175A (headers, accents)
- **Light Blue:** #66A9F2 (highlights, status indicators)
- **Grey:** #A7A8AA (borders, secondary text)
- **Professional typography and spacing**

### 8. Error Handling

#### 8.1 File Upload Errors
- Unsupported file format
- File size exceeds limit
- Corrupted file
- Empty file

#### 8.2 Data Processing Errors
- No valid data found
- Missing required columns
- No ranking data available
- Invalid data types

#### 8.3 User Experience
- Toast notifications for all errors
- Clear error messages with suggested solutions
- Graceful fallbacks where possible
- Progress indicators during processing

### 9. Success Metrics

#### 9.1 User Experience
- File upload success rate > 95%
- Average processing time < 30 seconds
- User completes full workflow > 90% of attempts

#### 9.2 Technical Performance
- Support files up to 200MB
- Handle 1000+ rows of data
- Memory usage remains stable
- No crashes during processing

### 10. Development Phases

#### Phase 1: Core Functionality (Week 1-2)
- Basic Streamlit app structure
- File upload and Excel processing
- Column mapping interface
- Basic PowerPoint generation

#### Phase 2: Advanced Features (Week 3)
- Multi-sheet support
- Slide preview functionality
- PDF export
- Error handling improvements

#### Phase 3: Polish & Deployment (Week 4)
- Brand styling implementation
- Performance optimization
- Docker containerization
- Testing and bug fixes

### 11. Future Enhancements (Post-MVP)

- Save column mapping templates
- Batch processing multiple files
- Custom slide templates
- Integration with Jira API
- User session management
- Advanced filtering options

### 12. Acceptance Criteria

✅ **File Upload**
- Supports .xlsx, .xls, .csv files up to 200MB
- Drag-and-drop interface works smoothly
- Progress indicator during upload

✅ **Data Processing**
- Auto-detects header row correctly
- Column mapping interface is intuitive
- Handles missing data gracefully

✅ **PowerPoint Generation**
- Creates slides matching template format exactly
- Correct status counts in top-right corner
- Professional styling with brand colors
- Configurable rows per slide

✅ **Export & Preview**
- Slide thumbnails display correctly
- PowerPoint download works
- PDF export functions properly

✅ **User Experience**
- Complete workflow takes < 2 minutes
- Error messages are clear and helpful
- Interface is intuitive for non-technical users