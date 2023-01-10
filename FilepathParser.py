import os
import re
import sys
from collections import Counter

from pkg_resources import DistributionNotFound, get_distribution

# Check if the required dependencies are installed
dependencies = ["pandas", "openpyxl", "PyQt5"]
missing_dependencies = []
for dependency in dependencies:
    try:
        version = get_distribution(dependency).version
    except DistributionNotFound:
        missing_dependencies.append(dependency)

if missing_dependencies:
    print(f"Missing dependencies: {missing_dependencies}.\nPlease install them using pip or conda.")
    sys.exit(1)

import pandas as pd
from fuzzywuzzy import fuzz
from PyQt5.QtWidgets import (QApplication, QFileDialog, QPushButton,
                             QVBoxLayout, QWidget)

# Create the GUI window
app = QApplication(sys.argv)
window = QWidget()
window.setWindowTitle("Excel File Uploader")

# Create a function to browse for and upload an Excel file
def browse_file():
    options = QFileDialog.Options()
    options |= QFileDialog.ReadOnly
    file_path, _ = QFileDialog.getOpenFileName(window, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)",
                                               options=options)
    if file_path:
        df = pd.read_excel(file_path)
        # Use a regular expression to search for file paths in the cells of the DataFrame
        results = regex_filepath_search(df)
        # Create a new DataFrame with the search results and the most used folder structures
        results_df = pd.DataFrame(results, columns=['Command Line', 'Filepath', 'Filename']).sort_values(by='Filename')
        folder_structures = get_folder_structures(results)
        folder_structures_df = pd.DataFrame(folder_structures, columns=['Folder Structure', 'Count'])
        # Allow the user to select the destination for the Excel file
        file_path, _ = QFileDialog.getSaveFileName(window, "Save Results", "", "Excel Files (*.xlsx);;All Files (*)",
                                                   options=options)
        if file_path:
            # Write the DataFrame to the Excel file
            with pd.ExcelWriter(file_path) as writer:
                results_df.to_excel(writer, sheet_name='Results', index=False)
                folder_structures_df.to_excel(writer, sheet_name='Folder Structures', index=False)

def regex_filepath_search(df):
    pattern = r'(?:[A-Za-z]\:|\\\\[\w\.]+\\[\w.$]+)\\(?:[\w]+\\)*\w([\w.])+'
    results = []
    for col in df.columns:
        for cell in df[col]:
            match = re.search(pattern, str(cell))
            if match:
                file_path = match.group()
                filename = os.path.basename(file_path)
                result = {'Command Line': cell, 'Filepath': file_path, 'Filename': filename}
                results.append(result)
    return results

def get_folder_structures(results):
    # Create a list of the filepaths without the filenames
    filepaths = [os.path.dirname(result['Filepath']) for result in results]
    # Create a list of the folder structures of the filepaths
    folder_structures = []
    for filepath in filepaths:
        # Split the filepath into a list of folders
        folders = filepath.split('\\')
        # Remove the first folder if it is a drive letter
        if len(folders[0]) == 2 and folders[0][1] == ':':
            folders.pop(0)
        # Remove the last folder if it is a filename
        if '.' in folders[-1]:
            folders.pop(-1)
        # Remove any folders that are unique to the filepath
        for folder in folders:
            if fuzz.ratio(folder, filepath) > 90:
                folders.remove(folder)
        # Join the folders back together into a string
        folder_structure = '\\'.join(folders)
        folder_structures.append(folder_structure)
    # Count the number of times each folder structure is used
    folder_structures_count = Counter(folder_structures)
    # Sort the folder structures by the number of times they are used
    folder_structures_count = sorted(folder_structures_count.items(), key=lambda x: x[1], reverse=True)
    return folder_structures_count

# Create a button to trigger the file browser
browse_button = QPushButton("Browse")
browse_button.clicked.connect(browse_file)

# Set up the layout for the GUI
layout = QVBoxLayout()
layout.addWidget(browse_button)
window.setLayout(layout)

# Run the GUI
window.show()
sys.exit(app.exec_())