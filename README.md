# Directory Structure Creator

This Python application allows users to create a directory structure based on the contents of an Excel file. It provides a graphical user interface (GUI) built with Tkinter for ease of use.

## Features

- Load an Excel file and select columns to be used for creating directories.
- Add or remove selected fields for directory creation.
- Create directories in a temporary location.
- Export the created directory structure as a ZIP file.

## Requirements

To run this application, you need Python installed on your system along with the following packages:
- pandas
- openpyxl
- Tkinter (usually included with Python)

## Usage

1. Run the script to open the GUI.
2. Click "Load Excel File" to load an Excel file and display its columns.
3. Select a column from the list and click "Add Field" to add it to the directory structure criteria.
4. Optionally, select a field from the selected fields list and click "Remove Field" to remove it.
5. Click "Create Directories" to create the directory structure in a temporary location.
6. Click "Export as ZIP" to save the directory structure as a ZIP file.

## Notes

- The application uses a temporary directory to store the created directory structure until the program is closed or the structure is exported.
- The temporary directory is cleaned up when the program exits to avoid leaving residual files on the system.

## License

This project is open-source and available under the MIT License.
