# Python Data Processing Toolbox

Welcome to the Python Data Processing Toolbox, a collection of simple functions commonly used in data processing. The purpose of this package is to save a programmer's time while working on more elaborate projects. The repository includes the following modules:

1. **collater.py**
   - Module containing a function for combining Excel files into a single Excel file.
   - The function assumes that the Excel files being combined have similar structures.
   - It looks through each sheet in each file for target column headers (e.g., 'Average') and then writes the content of each target column into the corresponding sheet in the combined Excel file.

2. **constants.py**
   - This file contains constants used throughout this project.

3. **excel_formulas.py**
   - This module contains functions used in generating common Excel formulas as well as plotting figures.

4. **file_handler.py**
   - Module containing functions to perform common file/folder handling operations.

5. **least_square_fit.py**
   - Module containing a function for applying different kinds of fits to data of the form `y = f(x)` using a least-square fit method.

6. **plot_info.py**
   - Module containing the `PlotInfo` data class, which is used to store figure formatting information used in plotting matplotlib figures in `plot_generator.py`.

7. **plot_generator.py**
   - Module containing a function for plotting data from a selected Excel file using matplotlib.
   - A list of `PlotInfo` objects corresponding to the desired sheet names is used to provide formatting information for the figures.

## Getting Started

To get started with this project, follow these steps:

1. Install the package to your local machine:

   ```bash
   pip install --upgrade pip
   pip install https://github.com/kaycee-okoye/data_processing_toolkit
   ```

To use the functions in this toolbox, follow the instructions provided in each module's documentation. You can find detailed information on how to use the functions in the respective files.

## Contribution Guidelines

If you would like to contribute to this project, please refer to the [CONTRIBUTING.md](CONTRIBUTING.md) file for guidelines on how to submit your contributions.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE.md) file for details.

## Acknowledgments

Thank you for using the Python Data Processing Toolbox. We hope that it helps streamline your data processing tasks. If you have any questions, issues, or suggestions, please don't hesitate to reach out. Happy coding!