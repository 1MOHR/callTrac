import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QPushButton, QWidget, QHeaderView
from PyQt6.QtGui import QKeySequence
from PyQt6.QtCore import Qt
from datetime import datetime


class CustomTableWidget(QTableWidget):
    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Copy):
            self.copy_selection()
        else:
            super().keyPressEvent(event)

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


class TimeTracker(QMainWindow):
    def __init__(self):
        super().__init__()

        self.call_counter = 1

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Time Tracker")
        self.setGeometry(100, 100, 800, 400)

        central_widget = QWidget(self)
        layout = QVBoxLayout(central_widget)

        self.table = CustomTableWidget(0, 6, self)
        self.table.setHorizontalHeaderLabels(["Call Number", "Time Call Taken", "Time on Call", "Ticket Number", "Closed/Client/Callback", "Notes"])
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        layout.addWidget(self.table)

        self.new_call_button = QPushButton("New Call", self)
        self.new_call_button.clicked.connect(self.new_call)
        layout.addWidget(self.new_call_button)

        self.end_call_button = QPushButton("End Call", self)
        self.end_call_button.clicked.connect(self.end_call)
        layout.addWidget(self.end_call_button)

        self.setCentralWidget(central_widget)

    def new_call(self):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)

        call_number = f"Call {self.call_counter}"
        time_taken = datetime.now().strftime("%I:%M %p")

        self.table.setItem(row_position, 0, QTableWidgetItem(call_number))
        self.table.setItem(row_position, 1, QTableWidgetItem(time_taken))

        self.call_counter += 1

        self.table.setCurrentCell(row_position, 0)

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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    time_tracker = TimeTracker()
    time_tracker.show()
    sys.exit(app.exec())
