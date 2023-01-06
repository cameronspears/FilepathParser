import os
import sys
import subprocess
import pandas as pd
import re

from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QFileDialog,
    QVBoxLayout,
    QPushButton,
    QDesktopWidget,
    QLabel,
    QHBoxLayout,
)

from PyQt5.QtCore import Qt

def check_dependencies():
    dependencies = ["pandas", "openpyxl", "PyQt5", "re", "os"]
    for dependency in dependencies:
        try:
            __import__(dependency)
        except ImportError:
            print(f"Missing required dependency: {dependency}")
            subprocess.run(["pip", "install", dependency])

def create_gui():
    # Create the GUI window
    app = QApplication(sys.argv)
    window = QWidget()
    window.setWindowTitle("Excel File Uploader")
    window.setGeometry(100, 100, 400, 300)

    # Get the screen dimensions
    screen = QDesktopWidget().screenGeometry()
    screen_width = screen.width()
    screen_height = screen.height()

    # Calculate the position of the top left corner of the window
    x = (screen_width - window.width()) // 2
    y = (screen_height - window.height()) // 2

    # Set the position of the window
    window.move(x, y)

    # Make the window always stay on top
    window.setWindowFlag(Qt.WindowStaysOnTopHint)

    # Create a function to browse for and upload an Excel file
    def browse_file():
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(
            window, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if file_path:
            df = pd.read_excel(file_path)
            # Use a regular expression to search for file paths in the cells of the DataFrame
            pattern = r'(?:[A-Za-z]\:|\\\\[\w\.]+\\[\w.$]+)\\(?:[\w]+\\)*\w([\w.])+'
            results = [
                {
                    "Command Line": cell,
                    "Filepath": match.group(),
                    "Filename": os.path.basename(match.group()),
                }
                for col in df.columns
                for cell in df[col]
                if (match := re.search(pattern, str(cell)))
            ]
            # Create a new DataFrame with the search results
            results_df = pd.DataFrame(results)
            # Allow the user to select the destination for the Excel file
            file_path, _ = QFileDialog.getSaveFileName(
                window, "Save Results", "", "Excel Files (*.xlsx);;All Files (*)", options=options
            )
            if file_path:
                # Write the DataFrame to the Excel file
                results_df.to_excel(file_path, index=False)
                # Create a horizontal layout
                h_layout = QHBoxLayout()
                # Create a label to display the success message
                label = QLabel("Success!")
                # Add the label to the layout
                h_layout.addWidget(label)
                # Set the alignment of the layout to center
                h_layout.setAlignment(Qt.AlignCenter)
                # Add the layout to the window
                layout.addLayout(h_layout)
                # Show the label
                label.show()

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

# Check dependencies and run the GUI
check_dependencies()
create_gui()
