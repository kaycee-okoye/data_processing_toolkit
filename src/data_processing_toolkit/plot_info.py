"""
    Module conatining contains the PlotInfo data class which is
    used to store figure formatting information used in
    plotting matplotlib figures in plotter.py
"""
from enum import Enum
import data_processing_toolkit.constants as constants


class PlotType(Enum):
    """Enum describing the type of matplotlib plot used in representing data"""

    SCATTER = "scatter"
    LINE = "line"


class PlotInfo:
    '''
    Data class containing matplotlib figure formatting information.

    This information will be interpreted by plotter.py in order to plot figures
    based on x and y datasets where y = f(x).

    Parameters
    ----------
    plot_type : PlotType
        The type of plot to be used to represent the data.
    xlabel : str
        x-axis label.
    ylabel : str
        y-axis label.
    save_sub_directory : str
        Relative path (from folder containing the excel file) to directory where the
        figure will be saved.
    secondary_series_stagger : float
        Horizontal offset of secondary series-values from the original series. This
        offset can be used to enhance clarity.
    legends : list[str]
        Legend values for each data series.
    vline_at_x : float
        x-value through which a vertical line should be drawn for y = [min(y), max(y)].
    hline_at_y : float
        y-value through which a horizontal line should be drawn for x = [min(x), max(x)].
    ylim : tuple of float
        (min(y), max(y)) describing the extent of the y-axis.
    xlim : tuple of float
        (min(x), max(x)) describing the extent of the x-axis.
    yaxis_log_scale : bool
        Y-axis will use log-log scale if True.
    xaxis_log_scale : bool
        X-axis will use log-log scale if True.
    save_file_extension : str
        File extension used in saving the figure, e.g., ".png".
    reverse_legends : bool
        Provided legends will be displayed on the figure in reverse order if True.
    font_size : int
        Font size of text in the figure.
    legend_size : int
        Font size of legend text (defaults to font-size if not provided).
    figure_size_in_inches : tuple of float
        (width, height) describing figure dimensions.
    subplot_tuple : tuple of float
        (row no, col no, index) describing the design of a subplot figure.
    title : str
        Figure title.
    fit_lines : tuple
        (x start, x end, function: lambda) used to draw a fit curve on the plot.
        The function should be a lambda function with one parameter (y = function(x)).
    sort_data : bool
        Series will be plotted based on x-values if True.
    show_x_axis_numbers : bool
        Numbers on the x-axis will be visible if True.
    show_y_axis_numbers : bool
        Numbers on the y-axis will be visible if True.
    reverse_series : bool
        Series in the excel spreadsheet will be plotted in reverse of the order they are
        in the spreadsheet if True.
    legend_loc : str
        Location where legends should be placed on the figure (see matplotlib documentation).
    legend_n_cols : int
        Number of columns to display legends in.
    legend_bbox_anchor : str
        See matplotlib documentation.
    draw_line_at_maximum : bool
        A vertical line for y = [min(y), max(y)] that passes through max(y) will be drawn if True.
    color_map : list[str]
        Hexadecimal colors representing the colors for representing each series.
    marker_size : int
        Size of markers of a scatter plot.
    use_color_map_for_fit_color : bool
        Colors of fit-lines will match the series they are fitting if True.
        Otherwise, the default color for fit-lines will be used.
    dpi : int
        The pixel density of the figure.

    Notes
    -----
    For proper functionality, all excel sheets which are part of the same
    subplot figure must be side by side in the excel workbook. The order they occur
    in the excel sheet are irrelevant (as long as they are side-by-side) since their
    order is defined in the PlotInfo list.
    '''

    def __init__(
        self,
        plot_type,
        xlabel,
        ylabel,
        save_sub_directory,
        secondary_series_stagger=0,
        legends=[],
        vline_at_x=None,
        hline_at_y=None,
        ylim=(None, None),
        xlim=(None, None),
        yaxis_log_scale=False,
        xaxis_log_scale=False,
        save_file_extension=constants.EPS_FILE_EXTENSION,
        reverse_legends=False,
        font_size=None,
        legend_size=None,
        figure_size_in_inches=None,
        suplot_tuple=None,
        title=None,
        fit_lines=[],
        sort_data=False,
        show_x_axis_numbers=True,
        show_y_axis_numbers=True,
        reverse_series=False,
        legend_loc="best",
        legend_n_cols=1,
        legend_bbox_anchor=None,
        draw_line_at_maximum=False,
        color_map=None,
        marker_size=None,
        use_color_map_for_fit_color=False,
        dpi=None,
    ):
        self.plot_type = plot_type
        self.xlabel = xlabel
        self.ylabel = ylabel
        self.save_sub_directory = save_sub_directory
        self.secondary_series_stagger = secondary_series_stagger
        self.legends = legends
        self.vline_at_x = vline_at_x
        self.hline_at_y = hline_at_y
        self.ylim = ylim
        self.xlim = xlim
        self.yaxis_log_scale = yaxis_log_scale
        self.xaxis_log_scale = xaxis_log_scale
        self.save_file_extension = save_file_extension
        self.reverse_legends = reverse_legends
        self.font_size = font_size
        self.legend_size = legend_size if legend_size is not None else font_size
        self.figure_size_in_inches = figure_size_in_inches
        self.suplot_tuple = suplot_tuple
        self.title = title
        self.fit_lines = fit_lines
        self.sort_data = sort_data
        self.show_x_axis_numbers = show_x_axis_numbers
        self.show_y_axis_numbers = show_y_axis_numbers
        self.reverse_series = reverse_series
        self.legend_loc = legend_loc
        self.legend_n_cols = legend_n_cols
        self.legend_bbox_anchor = legend_bbox_anchor
        self.draw_line_at_maximum = draw_line_at_maximum
        self.color_map = color_map
        self.marker_size = marker_size
        self.use_color_map_for_fit_color = use_color_map_for_fit_color
        self.dpi = dpi
