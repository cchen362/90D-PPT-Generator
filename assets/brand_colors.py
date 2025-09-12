# Brand colors and styling definitions for 90D-PPT Generator

# Primary brand colors
NAVY = "#00175A"
LIGHT_BLUE = "#66A9F2"
GREY = "#A7A8AA"

# Additional colors for UI elements
WHITE = "#FFFFFF"
BLACK = "#000000"
LIGHT_GREY = "#F5F5F5"
SUCCESS_GREEN = "#28A745"
ERROR_RED = "#DC3545"
WARNING_ORANGE = "#FFC107"

# Ranking circle colors
RANKING_COLORS = {
    "Accepted": "#28A745",    # Green
    "Up Next": "#FFC107",     # Yellow  
    "Maybe": "#A7A8AA",       # Grey
    "Likely No": "#6F42C1"    # Purple
}

# Status color mappings
STATUS_COLORS = {
    "in_progress": LIGHT_BLUE,
    "on_hold": WARNING_ORANGE,
    "pending_review": WARNING_ORANGE,
    "closed": SUCCESS_GREEN,
    "done": SUCCESS_GREEN,
    "resolved": SUCCESS_GREEN,
    "completed": SUCCESS_GREEN
}

# Status circle colors for PowerPoint display
STATUS_CIRCLE_COLORS = {
    "New Request": "#9E9E9E",      # Light Gray, Background 2, Darker 25%
    "Closed": "#28A745",           # Standard green
    # All other statuses will use standard blue
    "default": "#66A9F2"           # Standard blue (LIGHT_BLUE)
}

# Typography - Updated to Aptos as specified
PRIMARY_FONT = "Aptos"
TITLE_FONT = "Aptos Display"  # Specific font for slide titles
FALLBACK_FONT = "Arial"
FONT_STACK = f"{PRIMARY_FONT}, {FALLBACK_FONT}, sans-serif"

# Font sizes (in points for PowerPoint) - Updated per requirements
TITLE_FONT_SIZE = 36        # Slide title: Aptos Display 36
HEADER_FONT_SIZE = 14       # Headers: Aptos 14 Bold
DATA_FONT_SIZE = 11         # Rows: Aptos 11
STATUS_SUMMARY_FONT_SIZE = 14  # Status summary: Aptos 14

# Spacing and layout
TABLE_ROW_HEIGHT = 25
SLIDE_MARGIN = 50
STATUS_SUMMARY_TOP_MARGIN = 20
STATUS_SUMMARY_RIGHT_MARGIN = 50