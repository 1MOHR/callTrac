# Import necessary libraries
import csv
import sys
from datetime import datetime

# Import PyQt6 components
from PyQt6.QtCore import QModelIndex, QSettings, Qt, QTimer
from PyQt6.QtGui import QColor, QKeySequence, QPalette
from PyQt6.QtWidgets import QApplication, QHBoxLayout, QHeaderView, QMainWindow, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget


# Define a custom QTableWidget class to handle keyboard events
class CustomTableWidget(QTableWidget):
    # Define the custom keyPressEvent method for handling keyboard events
    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Copy):
            self.copy_selection()
        elif event.matches(QKeySequence.StandardKey.Paste):
            self.paste_data()
        elif event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            self.delete_or_clear()
        else:
            super().keyPressEvent(event)

    # Copy the selected data from the table to the clipboard
    def copy_selection(self):
        selected_ranges = self.selectedRanges()
        if not selected_ranges:
            return

        first_row = selected_ranges[0].topRow()
        last_row = selected_ranges[0].bottomRow()
        first_col = selected_ranges[0].leftColumn()
        last_col = selected_ranges[0].rightColumn()

        table_data = []
        for row in range(first_row, last_row + 1):
            row_data = []
            for col in range(first_col, last_col + 1):
                item = self.item(row, col)
                row_data.append(item.text() if item else '')
            table_data.append('\t'.join(row_data))

        clipboard = QApplication.clipboard()
        clipboard.setText('\n'.join(table_data))

    # Paste data from the clipboard to the table
    def paste_data(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return

        rows = text.split('\n')
        row_count = len(rows)

        col_count = 0
        for row in rows:
            col_count = max(col_count, len(row.split('\t')))

        current_row = self.currentRow()
        current_col = self.currentColumn()

        if current_row + row_count > self.rowCount():
            self.setRowCount(current_row + row_count)

        if current_col + col_count > self.columnCount():
            self.setColumnCount(current_col + col_count)

        for r, row in enumerate(rows):
            cols = row.split('\t')
            for c, cell in enumerate(cols):
                self.setItem(current_row + r, current_col + c, QTableWidgetItem(cell))

    # Delete or clear the selected data in the table
    def delete_or_clear(self):
        selected_indexes = self.selectedIndexes()
        selected_rows = sorted(set(index.row() for index in selected_indexes))

        if len(selected_rows) == 1 and self.selectionModel().isRowSelected(selected_rows[0], QModelIndex()):
            self.removeRow(selected_rows[0])
        else:
            for index in selected_indexes:
                self.setItem(index.row(), index.column(), QTableWidgetItem(""))


# Define the main TimeTracker class that extends QMainWindow
class TimeTracker(QMainWindow):
    def __init__(self):
        super().__init__()

        self.autosave_timer = QTimer(self)
        self.end_call_button = QPushButton("End Call", self)
        self.new_call_button = QPushButton("New Call", self)
        self.new_activity_button = QPushButton("New Activity", self)  # New button for new activity
        self.table = CustomTableWidget(0, 6, self)
        self.activity_table = CustomTableWidget(0, 2, self)  # New table for activities
        self.call_counter = 1

        self.init_ui()

    # Initialize the user interface
    def init_ui(self):
        self.setWindowTitle("Time Tracker")
        self.setGeometry(100, 100, 800, 400)

        central_widget = QWidget(self)
        layout = QVBoxLayout(central_widget)

        self.table.setHorizontalHeaderLabels(["Call Number", "Time Call Taken", "Time on Call", "Ticket Number", "Closed/Client/Callback", "Notes"])
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        layout.addWidget(self.table)

        self.activity_table.setHorizontalHeaderLabels(["Activity", "Notes"])  # Set labels for the new table
        self.activity_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)  # Stretch the Notes column
        self.activity_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        self.activity_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        layout.addWidget(self.activity_table)

        buttons_layout = QHBoxLayout()  # New layout for buttons
        buttons_layout.addWidget(self.new_activity_button)  # Add new_activity_button to the layout
        buttons_layout.addWidget(self.new_call_button)
        layout.addLayout(buttons_layout)  # Add buttons_layout to the main layout

        self.new_activity_button.clicked.connect(self.new_activity)  # Connect new_activity_button to new_activity function
        self.new_call_button.clicked.connect(self.new_call)

        self.end_call_button.clicked.connect(self.end_call)
        layout.addWidget(self.end_call_button)

        self.setCentralWidget(central_widget)

        self.autosave_timer.timeout.connect(self.save_data)
        self.autosave_timer.start(60 * 1000)  # Save data every 60 seconds

        self.load_data()

    # New function to create a new activity entry in the activity table
    def new_activity(self):
        row_position = self.activity_table.rowCount()
        self.activity_table.insertRow(row_position)

        self.activity_table.setCurrentCell(row_position, 0)

    # Create a new call entry in the table
    def new_call(self):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)

        call_number = f"Call {row_position + 1}"  # Use row_position + 1 as the call number
        time_taken = datetime.now().strftime("%I:%M %p")

        self.table.setItem(row_position, 0, QTableWidgetItem(call_number))
        self.table.setItem(row_position, 1, QTableWidgetItem(time_taken))

        self.table.setCurrentCell(row_position, 0)

    # End the current call and calculate the duration
    def end_call(self):
        row_position = self.table.currentRow()

        if row_position >= 0:
            time_taken_item = self.table.item(row_position, 1)

            if time_taken_item:
                time_taken_str = time_taken_item.text()
                time_taken = datetime.strptime(time_taken_str, "%I:%M %p")

                current_time = datetime.now()
                duration = int(round((current_time - time_taken).seconds / 60))

                self.table.setItem(row_position, 2, QTableWidgetItem(f"{duration} minutes"))

    # Save the table data to a CSV file
    # Save the table data to CSV files
    def save_data(self):
        self.save_table_data(self.table, 'autosave_calls.csv')
        self.save_table_data(self.activity_table, 'autosave_activities.csv')

    @staticmethod
    def save_table_data(table, file_name):
        with open(file_name, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile)
            for row in range(table.rowCount()):
                row_data = []
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    row_data.append(item.text() if item else '')
                csv_writer.writerow(row_data)

    # Load the table data from CSV files
    def load_data(self):
        self.load_table_data(self.table, 'autosave_calls.csv')
        self.load_table_data(self.activity_table, 'autosave_activities.csv')

    @staticmethod
    def load_table_data(table, file_name):
        try:
            with open(file_name, 'r', newline='', encoding='utf-8') as csvfile:
                csv_reader = csv.reader(csvfile)
                for row_data in csv_reader:
                    row_position = table.rowCount()
                    table.insertRow(row_position)
                    for col, data in enumerate(row_data):
                        table.setItem(row_position, col, QTableWidgetItem(data))
        except FileNotFoundError:
            pass  # If the file doesn't exist, start with an empty table


# Check if dark mode is enabled in the system settings
def is_dark_mode_enabled():
    qsettings = QSettings("HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", QSettings.Format.NativeFormat)
    return bool(qsettings.value("AppsUseLightTheme") == 0)


# Apply dark mode styles to the application
def set_dark_mode(app_instance):
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    palette.setColor(QPalette.ColorRole.Highlight, QColor(64, 128, 224))  # Changed highlight color
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)  # Changed highlighted text color

    app_instance.setPalette(palette)

    style_sheet = '''
    QTableWidget {
        background-color: #353535;
        gridline-color: #1E1E1E;
    }
    QHeaderView::section {
        background-color: #353535;
        color: white;
    }
    QPushButton {
        background-color: #111111;
        color: white;
        border: 1px solid #454545;
    }
    QPushButton:hover {
        background-color: #191919;
    }
    QPushButton:pressed {
        background-color: #190016;
    }
    QWidget {
        background-color: #070707;
        color: white;
    }
    QScrollBar {
        background-color: #353535;
    }
    QTableCornerButton::section {
        background-color: #353535;
    }
    '''
    app_instance.setStyleSheet(style_sheet)


# Main entry point for the application
if __name__ == "__main__":
    app = QApplication(sys.argv)

    if is_dark_mode_enabled():
        set_dark_mode(app)

    time_tracker = TimeTracker()
    time_tracker.show()
    sys.exit(app.exec())
