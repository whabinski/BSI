from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFileDialog, QCheckBox, QComboBox
from PyQt5.QtGui import QPixmap, QPainter
from PyQt5.QtCore import Qt
import os
from tally import read_tally, increase_tally, decrease_tally


class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        
        self.file_path = None

        # Set up the window
        self.setWindowTitle('BSI File Processor')
        self.setGeometry(100, 100, 500, 300)

        # Create a vertical layout
        self.layout = QVBoxLayout()
        
        # Create tally section
        self.create_tally_section()

        # Label to display messages
        self.message_label = QLabel('Select a file', self)
        self.layout.addWidget(self.message_label)

        # Create a horizontal layout for the checkboxes
        dropdown_layout = QHBoxLayout()

        # First checkbox
        self.dropdown1 = QComboBox(self)
        self.dropdown1.addItems(["Canada", "United States"]) 
        dropdown_layout.addWidget(self.dropdown1)

        # Second Dropdown
        self.dropdown2 = QComboBox(self)
        self.dropdown2.addItems(["MA", "PS"]) 
        dropdown_layout.addWidget(self.dropdown2)

        # Add the horizontal layout to the main vertical layout
        self.layout.addLayout(dropdown_layout)

        # Button to load file
        self.load_button = QPushButton('Load COC File', self)
        self.load_button.clicked.connect(self.load_file)
        self.layout.addWidget(self.load_button)

        # Button to start processing
        self.process_button = QPushButton('Process Data', self)
        self.process_button.clicked.connect(self.process_data)
        self.layout.addWidget(self.process_button)

        # Clear button (optional)
        self.clear_button = QPushButton('Clear', self)
        self.clear_button.clicked.connect(self.clear_file)
        self.layout.addWidget(self.clear_button)

        # Set the layout
        self.setLayout(self.layout)

    def create_tally_section(self):
        # Create a horizontal layout for the "-" label "+"
        increment_layout = QHBoxLayout()

        # Create the "-" button
        self.decrement_button = QPushButton('-')
        self.decrement_button.setFixedSize(40, 30)
        self.decrement_button.clicked.connect(self.decrement_value)
        increment_layout.addWidget(self.decrement_button)

        # Create the label to show the current value
        self.value_label = QLabel(read_tally())
        self.value_label.setAlignment(Qt.AlignCenter)
        increment_layout.addWidget(self.value_label)

        # Create the "+" button
        self.increment_button = QPushButton('+')
        self.increment_button.setFixedSize(40, 30)
        self.increment_button.clicked.connect(self.increment_value)
        increment_layout.addWidget(self.increment_button)

        # Align the increment section to the top-right of the window
        self.layout.addLayout(increment_layout)
        self.layout.setAlignment(increment_layout, Qt.AlignTop | Qt.AlignRight)

    def increment_value(self):
        increase_tally()
        self.value_label.setText(read_tally())

    def decrement_value(self):
        current_value = int(self.value_label.text())
        if current_value > 0:
            decrease_tally()
        self.value_label.setText(read_tally())

    def load_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        self.file_path, _ = QFileDialog.getOpenFileName(self, "Select an Excel file", "", "Excel Files (*.xlsx *.xls)", options=options)
        if self.file_path:
            if self.file_path.endswith('.xlsx'):
                file_name = os.path.basename(self.file_path)
                self.message_label.setText(file_name)
            else:
                self.message_label.setText('Invalid file type. Please select an .xlsx file')
        else:
            self.message_label.setText('Select a file')

    def process_data(self):
        # TEST
        # self.file_path = '/Users/wyatthabinski/Documents/Work/BSI/Data/New COC Tester.xlsx'
        if self.file_path and self.file_path.endswith('.xlsx'):
            try:
                from data_processing import process_file  # Importing here to avoid circular imports
                countryDropdown = self.dropdown1.currentText()
                typeDropdown = self.dropdown2.currentText()
                process_file(self.file_path, countryDropdown, typeDropdown)  # Pass the checkbox state
                self.message_label.setText('Processing complete.')
                self.file_path = 'Select a file'
                self.value_label.setText(read_tally())
            except Exception as e:
                self.message_label.setText(f'Error: {str(e)}')
        else:
            self.message_label.setText('Select a file')

    def clear_file(self):
        self.file_path = None
        self.message_label.setText('Select a file')

if __name__ == '__main__':
    app = QApplication([])
    window = MyApp()
    window.show()
    app.exec_()