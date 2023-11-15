"""
    This module contains functions used in generating common
    excel formulas as well as plotting figures.
"""

from os.path import split
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol, xl_col_to_name
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.chart.data_source import NumDataSource, NumRef
from openpyxl.chart import ScatterChart, Reference, Series, marker
import data_processing_toolkit.constants as constants


def get_excel_column_range(col_no):
    """
    Function to convert a column number to an excel range formula
    representing the entire column

    Parameters
    ----------
    col_no : int
        column number to be converted

    Returns
    -------
    str
        appropriate excel formula
    """

    return f"{xl_col_to_name(col_no)}:{xl_col_to_name(col_no)}"


def get_cah_equation(
    left_angle_address, right_angle_address, flow_is_left_to_right=True
):
    """
    Function to generate excel formula that calculates contact angle hysteresis

    Parameters
    ----------
    left_angle_address : str
        excel cell address holding the current left contact angle
    right_angle_address : str
        excel cell address holding the current right contact angle
    flow_is_left_to_right : bool, optional
        whether the flow is moving from left to right from the
        perspective of the side view camera(s), by default True

    Returns
    -------
    str
        appropriate excel formula
    """

    upstream_contact_angle = (
        left_angle_address if flow_is_left_to_right else right_angle_address
    )
    downstream_contact_angle = (
        right_angle_address if flow_is_left_to_right else left_angle_address
    )

    return error_guard_formula(
        f"=COS(RADIANS({upstream_contact_angle})) - COS(RADIANS({downstream_contact_angle}))"
    )


def get_scale_equation(value_address, scale_address):
    """
    Function to generate excel formula that scales a value to another using a scale

    Parameters
    ----------
    value_address : str
        excel cell address holding the value to be converted
    scale_address : str
        excel cell address holding the scale to be used for conversion

    Returns
    -------
    str
        appropriate excel formula
    """

    return error_guard_formula(f"={value_address}*{scale_address}")


def get_ratio_equation(
    numerator_address, denominator_address, always_positive=False, default=None
):
    """
    Function to generate excel formula that finds the ratio between two values

    Parameters
    ----------
    numerator_address : str
        excel cell address holding the numerator
    denominator_address : str
        excel cell address holding the denominator
    always_positive : bool, optional
        whether to return the absolute value of the ratio, by default False
    default : any, optional
        value to store if the division fails, by default None

    Returns
    -------
    str
        appropriate excel formula
    """

    formula = (
        f"=ABS({numerator_address}/{denominator_address})"
        if always_positive
        else f"={numerator_address}/{denominator_address}"
    )
    return error_guard_formula(formula, default=default)


def get_average_equation(value1_address, value2_address):
    """
    Function to generate excel formula that calculates the average of two values

    Parameters
    ----------
    value1_address : str
        excel cell address holding the first value
    value2_address : str
        excel cell address holding the second value

    Returns
    -------
    str
        appropriate excel formula
    """

    return error_guard_formula(f"=({value1_address} + {value2_address}) / 2")


def get_average_equation_for_row_range(
    row_no, start_col_no, end_col_no, positive_only=True
):
    """
    Function to generate excel formula that calculates the average of a row range

    Parameters
    ----------
    row_no : int
        row number of the range
    start_col_no : int
        column number of leftmost column in range
    end_col_no : int
        column number of rightmost column in range
    positive_only : bool, optional
        whether to only consider positive values, by default True

    Returns
    -------
    str
        appropriate excel formula
    """

    start_address = convert_to_excel_address(row_no, start_col_no)
    end_address = convert_to_excel_address(row_no, end_col_no)

    if positive_only:
        avg_equation = f'=AVERAGEIFS({start_address}:{end_address}, {start_address}:{end_address}, "<>#N/A", {start_address}:{end_address}, ">= 0")'
    else:
        avg_equation = f'=AVERAGEIFS({start_address}:{end_address}, {start_address}:{end_address}, "<>#N/A")'
    return error_guard_formula(avg_equation)


def get_std_equation_for_row_range(
    row_no, start_col_no, end_col_no, positive_only=True
):
    """
    Function to generate excel formula that calculates the standard deviation of a row range

    Parameters
    ----------
    row_no : int
        row number of the range
    start_col_no : int
        column number of leftmost column in range
    end_col_no : int
        column number of rightmost column in range
    positive_only : bool, optional
        whether to only consider positive values, by default True

    Returns
    -------
    str
        appropriate excel formula
    """

    start_address = convert_to_excel_address(row_no, start_col_no)
    end_address = convert_to_excel_address(row_no, end_col_no)
    range_formula = f"{start_address}:{end_address}"

    if positive_only:
        condition = f'IF(({range_formula} > 0), {range_formula}, "")'
        std_equation = f"=STDEV({condition}"
    else:
        std_equation = f"=STDEV({range_formula}"
    return error_guard_formula(std_equation)


def calculate_jet_exit_velocity(image_time_address, flow_accel):
    """
    Function to generate excel formula that calculates the jet exit velocity.

    Parameters
    ----------
    image_time_address : str
        excel cell address holding the current elapsed time
    flow_accel : float
        current flow acceleration

    Returns
    -------
    str
        appropriate excel formula
    """

    jet_velocity_fit_equation = {
        4.4: {"slope": 4.7303408085, "intercept": -1.6190716868},
        3.2: {"slope": 3.4542, "intercept": -1.3055},
        2.2: {"slope": 2.3498, "intercept": -1.0454},
        1.2: {"slope": 1.2631, "intercept": -0.7672},
    }
    jet_velocity_fit_equation = jet_velocity_fit_equation[flow_accel]

    return f'=({image_time_address} * {jet_velocity_fit_equation["slope"]}) + {jet_velocity_fit_equation["intercept"]}'


def calculate_local_flow_velocity(image_time_address, flow_accel):
    """
    Function to generate excel formula that calculates the flow velocity
    at the depinning location.

    Parameters
    ----------
    image_time_address : str
        excel cell address holding the current elapsed time
    flow_accel : float
        current flow acceleration

    Returns
    -------
    str
        appropriate excel formula
    """

    local_velocity_fit_equations = {
        4.4: {"slope": 4.5387, "intercept": -1.5401},
        3.2: {"slope": 3.2511, "intercept": -1.2167},
        2.2: {"slope": 2.2221, "intercept": -0.935},
        1.2: {"slope": 1.1916, "intercept": -0.6331},
    }
    local_velocity_fit_equation = local_velocity_fit_equations[flow_accel]
    return f'=({image_time_address} * {local_velocity_fit_equation["slope"]}) + {local_velocity_fit_equation["intercept"]}'


def error_guard_formula(formula, default=None):
    """
    Function to generate generate excel formula that prevents other formula from return errors

    Parameters
    ----------
    formula : str
        formula to error guard
    default : any, optional
        value to default to if there's an error, by default None

    Returns
    -------
    str
        appropriate excel formula
    """

    formula = formula.lstrip("=")
    if default is None:
        return f"=IFERROR({formula}, NA())"
    else:
        return f"=IFERROR({formula}, {default})"


def convert_to_excel_address(
    row_no, col_no, file_path="", sheet_name="", fixed_row=False, fixed_column=False
):
    """
    Function to generate excel address

    Parameters
    ----------
    row_no : int
        row number of cell
    col_no : int
        column number of cell
    file_path : str, optional
        file containing cell, by default ""
    sheet_name : str, optional
        sheet containing cell, by default ""
    fixed_row : bool, optional
        whether the row is Constant e.g. A$1, by default False
    fixed_column : bool, optional
        whether the column is Constant e.g. $A1, by default False

    Returns
    -------
    str
        appropriate excel formula
    e.g. ='C:\\Folder\\[Dimensionless Test Matrix.xlsx]Sheet1'!$A$1
    """

    cell_address = xl_rowcol_to_cell(
        row_no, col_no, row_abs=fixed_row, col_abs=fixed_column
    )
    if file_path:
        root_path = split(file_path)[0]
        file_name = split(file_path)[1]
    return (
        f"'{root_path}\\[{file_name}]{sheet_name}'!{cell_address}"
        if file_path
        else f"'{sheet_name}'!{cell_address}"
        if sheet_name
        else cell_address
    )


def get_excel_row_range(start_row_no, end_row_no, col_no, sheet_name=""):
    """
    Function to generate generate excel row range

    Parameters
    ----------
    start_row_no : int
        row number of the topmost row
    end_row_no : int
        row number of the bottommost row
    col_no : int
        column number of row
    sheet_name : str, optional
        name of sheet containing the range, by default ""

    Returns
    -------
    str
        appropriate excel formula
    """

    start_address = convert_to_excel_address(
        start_row_no, col_no, fixed_row=True, fixed_column=True
    )
    end_address = convert_to_excel_address(
        end_row_no, col_no, fixed_row=True, fixed_column=True
    )
    if sheet_name:
        return f"='{sheet_name}'!{start_address}:{end_address}"
    else:
        return f"={start_address}:{end_address}"


def convert_from_excel_cell_address(cell_address):
    """
    Function to convert excel cell notation to its corresponding parts

    Parameters
    ----------
    cell_address : str
        excel cell address to be converted to its parts

    Returns
    -------
    sheet_name : str
        excel sheet containing the address if available,
        empty string otherwise
    row_no : int
        row number of cell
    col_no : int
        col number of cell
    """

    if "!" in cell_address:
        sheet_name = cell_address.split("!")[0].strip("'")
        address = cell_address.split("!")[1]
    else:
        sheet_name = ""
        address = cell_address
    (row_no, col_no) = xl_cell_to_rowcol(address)
    return (sheet_name, row_no, col_no)


def is_bad_input(value):
    """
    Function to check if value being input to excel sheet will cause error

    Parameters
    ----------
    value
        value being evaluated

    Returns
    -------
    bool
        whether it is a common error value
    """
    return (value is None) or (np.isnan(value)) or (np.isinf(value))


def write_cell(worksheet, row_no, col_no, data):
    """
    Function to write a value to an excel cell

    Parameters
    ----------
    worksheet : xlsxwriter worksheet
        sheet containing cell
    row_no : int
        row number of cell
    col_no : int
        column number of cell
    data : any
        value to write to cell
    """

    try:
        if str(data).startswith("="):
            worksheet.write_formula(row_no, col_no, data)
        else:
            if (
                isinstance(data, np.float64) or isinstance(data, float)
            ) and is_bad_input(data):
                worksheet.write_formula(row_no, col_no, "")
            else:
                worksheet.write(row_no, col_no, data)
    except Exception:
        worksheet.write(row_no, col_no, data)


def get_cell_value(worksheet, row_no, col_no, is_number=True, default_value=None):
    """
    Function to read a value from an excel cell

    Parameters
    ----------
    worksheet : openpyxl worksheet
        sheet containing cell
    row_no : int
        row number of cell
    col_no : int
        column number of cell
    is_number : bool, optional
        whether the cell is expected to contain a number, by default True
    default_value : any, optional
        value to default to if there's an error, by default None

    Returns
    -------
    any
        content of cell
    """

    cell_data = worksheet[convert_to_excel_address(row_no, col_no)].value
    try:
        if is_number:
            return (
                float(cell_data)
                if cell_data not in constants.EXCEL_ERROR_VALUES
                else default_value
            )
        else:
            return (
                str(cell_data)
                if cell_data not in constants.EXCEL_ERROR_VALUES
                else default_value
            )
    except:
        return default_value


def write_row(worksheet, contents, row_no, start_col=0):
    """
    Function to write data into an excel row

    Parameters
    ----------
    worksheet : xlsxwriter worksheet
        sheet containing row
    contents : iterable
        values to write to row
    row_no : int
        row number of row
    start_col : int, optional
        column number of leftmost cell in the row, by default 0
    """

    for col_no, data in enumerate(contents):
        write_cell(worksheet, row_no, start_col + col_no, data)


def write_col(worksheet, contents, col_no, start_row=0):
    """
    Function to write data into an excel column

    Parameters
    ----------
    worksheet : xlsxwriter worksheet
        sheet containing column
    contents : iterable
        values to write to column
    col_no : int
        column number of column
    start_row : int, optional
        row number of topmost cell in the start_row, by default 0
    """

    for row_no, data in enumerate(contents):
        write_cell(worksheet, start_row + row_no, col_no, data)


def write_openpyxl_row(worsksheet, row_no, col_no, contents):
    """
    Function to write data into an excel row

    Parameters
    ----------
    worsksheet : openpyxl worksheet
        sheet containing row
    row_no : int
        row number of row
    col_no : int
        column number of leftmost cell in the row
    contents : iterable
        values to write to row
    """

    for idx, value in enumerate(contents):
        worsksheet[convert_to_excel_address(row_no=row_no, col_no=col_no + idx)] = (
            value if value is not None else ""
        )


def create_sheet_name(name):
    """
    Function to edit a proposed sheetname to ensure its valid in excel

    Parameters
    ----------
    name : str
        proposed sheet name

    Returns
    -------
    sheet_name : str
        appropriate sheetname
    """

    sheet_name = (
        name if len(name) <= 31 else name[:31]
    )  # excel sheet titles have a max lenght of 31 chars
    sheet_name = "".join(
        c for c in sheet_name if not c in "[]:*?/\\"
    )  # remove illegal characters from shet name
    return sheet_name


def add_sheet_statistics(stats_sheet, source_sheet_name, headers, n_rows):
    """
    Function to insert excel formulas in a worksheet that
    calculate the statistics (mean, mode, median, 25th & 75th percentiles) of
    all columns in another worksheet.

    Parameters
    ----------
    stats_sheet : xlsxwriter worksheet
        worksheet where formulas will be inserted
    source_sheet_name : str
        name of the worksheet whose statistics will be calculated
    headers : iterable[str]
        titles of each column whose statistics will be calculated.
        It is assumed that the index of a header in this iterable, is its column index
    n_rows : int
        the maximum number of rows accross columns whose statistics will be
        calculated
    """

    column_ranges = list(
        map(
            lambda x: f"'{source_sheet_name}'!{get_excel_row_range(1, n_rows, headers.index(x)).lstrip('=')}",
            headers,
        )
    )
    write_row(stats_sheet, [constants.STAT_LABEL] + headers, 0)
    formulas = {
        constants.AVERAGE_LABEL: (
            lambda x: f"=AVERAGE({column_ranges[headers.index(x)]})"
        ),
        constants.STD_LABEL: (lambda x: f"=STDEV({column_ranges[headers.index(x)]})"),
        constants.MODE_LABEL: (lambda x: f"=MODE({column_ranges[headers.index(x)]})"),
        constants.MEDIAN_LABEL: (
            lambda x: f"=MEDIAN({column_ranges[headers.index(x)]})"
        ),
        constants.PERCENTILE_25_LABEL: (
            lambda x: f"=PERCENTILE({column_ranges[headers.index(x)]}, 0.25)"
        ),
        constants.PERCENTILE_75_LABEL: (
            lambda x: f"=PERCENTILE({column_ranges[headers.index(x)]}, 0.75)"
        ),
    }
    for i, (stat, eq) in enumerate(formulas.items()):
        write_row(
            stats_sheet,
            [stat]
            + list(
                map(
                    eq,
                    headers,
                )
            ),
            i + 1,
        )

def insert_chart(
    workbook,
    sheet,
    chart_title,
    xlabel,
    ylabel,
    x_sets,
    y_sets,
    series_names=[],
    error_bars={},
):
    """
    Function that plots a scatter plot with lines in an excel file

    Parameters
    ----------
    workbook : xlsxwriter workbook
        where the chart will be inserted
    sheet : xlsxwriter worksheet
        where the chart will be inserted
    chart_title : str
        title of the chart
    xlabel : str
        x-axis title
    ylabel : str
        y-axis title
    x_sets : iterable[str]
        excel range formula of x-values of each series being plotted
    y_sets : iterable[str]
        excel range formula of y-values of each series being plotted
    series_names : iterable[str], optional
        name of each series of data being plotted, by default []
    error_bars : dict, optional
        error_bars[series name] = excel range formula of
        error bars of each series being plotted. Series name whose error bars are
        included must be in the series_names iterable to be plotted, by default {}
    """

    chart = workbook.add_chart({"type": "scatter", "subtype": "straight"})
    if not series_names:
        series_names = [str(i).zfill(5) for i in range(len(y_sets))]
    for i, series_name in enumerate(series_names):
        if series_name in error_bars.keys():
            chart.add_series(
                {
                    "name": series_name,
                    "categories": x_sets[i],
                    "values": y_sets[i],
                    "y_error_bars": {
                        "type": "custom",
                        "plus_values": error_bars[series_name],
                        "minus_values": error_bars[series_name],
                    },
                }
            )
        else:
            chart.add_series(
                {
                    "name": convert_to_excel_address(
                        row_no=0, col_no=i, sheet_name=sheet.get_name()
                    ),
                    "categories": x_sets[i],
                    "values": y_sets[i],
                }
            )
    chart.set_title({"name": chart_title})
    chart.set_x_axis({"name": xlabel})
    chart.set_y_axis({"name": ylabel})
    sheet.insert_chart("A1", chart)


def insert_openpyxl_chart(
    output_sheet,
    title,
    x_title,
    y_title,
    x_sheet,
    x_cols,
    x_min_rows,
    x_max_rows,
    y_sheet,
    y_cols,
    y_min_rows,
    y_max_rows,
    labels,
    location,
    draw_line=True,
    error_bars=None,
):
    """
    Function that plots a scatter plot with lines in an excel file

    Parameters
    ----------
    output_sheet : xlsxwriter worksheet
        where the chart will be inserted
    title : str
        title of the chart
    x_title : str
        x-axis title
    y_title : str
        y-axis title
    x_sheet : openpyxl worksheet
        where the x-values are stored
    x_cols : iterable[int]
        column number in x_sheet of x-values of each series being plotted
    x_min_rows : iterable[int]
        topmost row number in x_sheet of x-values of each series being
        plotted
    x_max_rows : iterable[int]
        bottommost row number in x_sheet of x-values of each series being
        plotted
    y_sheet : openpyxl worksheet
        where the y-values are stored
    y_cols : iterable[int]
        column number in y_sheet of y-values of each series being plotted
    y_min_rows : iterable[int]
        topmost row number in y_sheet of y-values of each series being
        plotted
    y_max_rows : iterable[int]
        bottommost row number in y_sheet of y-values of each series being
        plotted
    labels : iterable[str]
        legend name of each series of data being plotted
    location : str
        excel formula of location where the chart will be inserted in output_sheet
    draw_line : bool, optional
        scatter points are connected with lines if true, by default True
    error_bars : dict, optional
        error_bars[series name] = excel range formula of
        error bars of each series being plotted. Series name whose error bars are
        included must be in the series_names iterable to be plotted, by default None
    """

    chart = ScatterChart()
    n_series = len(x_cols)
    for i in range(n_series):
        xvalues = Reference(
            worksheet=x_sheet,
            min_col=x_cols[i],
            min_row=x_min_rows[i],
            max_row=x_max_rows[i],
        )
        yvalues = Reference(
            worksheet=y_sheet,
            min_col=y_cols[i],
            min_row=y_min_rows[i],
            max_row=y_max_rows[i],
        )
        size = Reference(worksheet=output_sheet, min_col=2, min_row=2, max_row=10)
        series = Series(values=yvalues, xvalues=xvalues, zvalues=size, title=labels[i])

        if error_bars is not None:
            error_num_val = NumRef(f=error_bars)
            error_val = NumDataSource(numRef=error_num_val)
            error_bar = ErrorBars(
                errDir="y", plus=error_val, minus=error_val, errValType="cust"
            )
            series.errBars = error_bar

        if draw_line is False:
            series.marker = marker.Marker("circle")
            series.graphicalProperties.line.noFill = True
        chart.series.append(series)

    chart.title = title
    chart.x_axis.title = x_title
    chart.y_axis.title = y_title
    chart.legend.position = "t"
    output_sheet.add_chart(chart, location)
