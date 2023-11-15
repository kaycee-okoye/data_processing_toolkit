"""
Recommended file to import the directory as a regular package.
Used to expose methods and classes to library level.
"""
from data_processing_toolkit.collater import collate_excel_files
from data_processing_toolkit.plot_info import PlotInfo, PlotType
from plot_generator import generate_plots
