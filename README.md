# Python Data Processing Toolbox

Welcome to the Python Data Processing Toolbox, a collection of simple functions and scripts commonly used in data processing. The purpose of this repository is to save a programmer's time while working on more elaborate projects. The repository includes the following modules and scripts:

1. **collater.py**
   - Module containing a script for combining Excel files into a single Excel file.
   - The script assumes that the Excel files being combined have similar structures.
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
   - Module containing the `PlotInfo` data class, which is used to store figure formatting information used in plotting matplotlib figures in `plotter.py`.

7. **plotter.py**
   - Module containing a script for plotting data from a selected Excel file.
   - A list of `PlotInfo` objects corresponding to the desired sheet names is used to provide formatting information for the figures.

## Getting Started

To use the functions and scripts in this toolbox, follow the instructions provided in each module's documentation. You can find detailed information on how to use the functions and scripts in the respective files.

## Dependencies

This Python Data Processing Toolbox relies on several external libraries and packages to perform various tasks. To ensure the toolbox functions correctly, please make sure you have the following dependencies installed:

1. **matplotlib**
   - Matplotlib is a comprehensive library for creating static, animated, and interactive visualizations in Python.

   To install, run:
   ```
   pip install matplotlib
   ```

2. **numpy**
   - NumPy is a fundamental package for scientific computing with Python. It contains, among other things, a powerful N-dimensional array object useful for numerical computations.

   To install, run:
   ```
   pip install numpy
   ```

3. **openpyxl**
   - Openpyxl is a Python library for reading and writing Excel (xlsx) files.

   To install, run:
   ```
   pip install openpyxl
   ```

4. **xlsxwriter**
   - XlsxWriter is a Python module for writing Excel files.

   To install, run:
   ```
   pip install XlsxWriter
   ```

5. **cv2 (OpenCV)**
   - OpenCV (Open Source Computer Vision Library) is an open-source computer vision and machine learning software library.

   To install, run:
   ```
   pip install opencv-python
   ```

6. **scipy**
   - SciPy is a library used for scientific and technical computing. It provides modules for optimization, integration, interpolation, eigenvalue problems, and more.

   To install, run:
   ```
   pip install scipy
   ```

Please make sure to install these dependencies before using the Python Data Processing Toolbox to ensure the proper functioning of the scripts and modules.

## Contribution Guidelines

If you would like to contribute to this project, please refer to the [CONTRIBUTING.md](CONTRIBUTING.md) file for guidelines on how to submit your contributions.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE.md) file for details.

## Acknowledgments

Thank you for using the Python Data Processing Toolbox. We hope that it helps streamline your data processing tasks. If you have any questions, issues, or suggestions, please don't hesitate to reach out. Happy coding!