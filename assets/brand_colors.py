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

# Typography
PRIMARY_FONT = "Benton Sans"
FALLBACK_FONT = "Arial"
FONT_STACK = f"{PRIMARY_FONT}, {FALLBACK_FONT}, sans-serif"

# Font sizes (in points for PowerPoint, pixels for web)
HEADER_FONT_SIZE = 12
DATA_FONT_SIZE = 10
TITLE_FONT_SIZE = 16

# Spacing and layout
TABLE_ROW_HEIGHT = 25
SLIDE_MARGIN = 50
STATUS_SUMMARY_TOP_MARGIN = 20
STATUS_SUMMARY_RIGHT_MARGIN = 50