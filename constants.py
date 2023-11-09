"""
    This file contains constants used throughout this project.
"""

# Labels used in excel file (e.g. plots labels)
JET_EXIT_VELOCITY_LABEL = "Jet Exit Velocity"
POINT_ANALYSIS_LABEL = "Point Analysis"

STAT_LABEL = "Stat"
AVERAGE_LABEL = "Average"
STD_LABEL = "STD"
MODE_LABEL = "Mode"
MEDIAN_LABEL = "Median"
PERCENTILE_25_LABEL = "25th Percentile"
PERCENTILE_75_LABEL = "75th Percentile"
NORMALIZED_LABEL = "Normalized"

# Filenames
EXCEL_FILE_EXTENSION = ".xlsx"
DAVIS_SET_FILE_EXTENSION = ".set"
PNG_FILE_EXTENSION = ".png"
EPS_FILE_EXTENSION = ".eps"
COLLATER_EXCEL_FILENAME = f"Cumulative{EXCEL_FILE_EXTENSION}"

# Excel sheet names
EXCEL_DEFAULT_SHEETNAME = "Sheet"

# Excel constants
NAN_FROM_EXCEL = "#N/A"
ALT_NAN_FROM_EXCEL = "=NA()"
ERROR_FROM_EXCEL = "#VALUE!"
REF_VALUE_FROM_EXCEL = "#REF!"
DIV0_EXCEL_ERROR = "#DIV/0!"
ERROR_VALUE_FOR_INPUT = "NA()"
EXCEL_ERROR_VALUES = [
    NAN_FROM_EXCEL,
    ALT_NAN_FROM_EXCEL,
    ERROR_FROM_EXCEL,
    REF_VALUE_FROM_EXCEL,
    ERROR_VALUE_FOR_INPUT,
    DIV0_EXCEL_ERROR
]

# plotter.py constants
HEADER_HEADER, X_HEADER, Y_HEADER = "header", "x", "y"
Y_ERROR_HEADER, Y_POS_ERROR_HEADER, Y_NEG_ERROR_HEADER = "yerror", "yerror+", "yerror-"
X_ERROR_HEADER, X_POS_ERROR_HEADER, X_NEG_ERROR_HEADER = "xerror", "xerror+", "xerror-"
INDEX_HEADER = 'index'
SECONDARY_SERIES_PREFIX = 'secondary_'
SECONDARY_2_SERIES_PREFIX = 'secondary2_'
SECONDARY_3_SERIES_PREFIX = 'secondary3_'
