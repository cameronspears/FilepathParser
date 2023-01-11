import os
import sys
import pandas as pd
import re
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QVBoxLayout, QPushButton
from pkg_resources import get_distribution, DistributionNotFound

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

# Create the GUI window
app = QApplication(sys.argv)
window = QWidget()
window.setWindowTitle("Excel File Uploader")

# Create a function to browse for and upload an Excel file
def browse_file():
    options = QFileDialog.Options()
    options |= QFileDialog.ReadOnly
    file_path, _ = QFileDialog.getOpenFileName(window, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
    if file_path:
        df = pd.read_excel(file_path)
        # Use a regular expression to search for file paths in the cells of the DataFrame
        pattern = r'(?:[A-Za-z]\:|\\\\[\w\.]+\\[\w.$]+)\\(?:[\w]+\\)*[\w\s\(\)\-\,\.]+'
        results = []
        for col in df.columns:
            for cell in df[col]:
                match = re.search(pattern, str(cell))
                if match:
                    file_path = match.group()
                    filename = os.path.basename(file_path)
                    # remove the filename from filepath
                    file_path = file_path.replace(filename, "")
                    # remove the filepath and filename from command line 
                    command_line = cell.replace(file_path, "").replace(filename, "")
                    result = {'Command Line': command_line, 'Filepath': file_path, 'Filename': filename}
                    results.append(result)
        # Create a new DataFrame with the search results
        results_df = pd.DataFrame(results)
        # Allow the user to select the destination for the Excel file
        file_path, _ = QFileDialog.getSaveFileName(window, "Save Results", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_path:
            # Write the DataFrame to the Excel file
            results_df.to_excel(file_path, index=False)

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
