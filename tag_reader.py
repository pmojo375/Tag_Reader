from PySide6.QtCore import QSettings, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QMessageBox,
    QHBoxLayout,
    QCheckBox,
    QFileDialog
)
from PySide6 import QtGui
import sys
from pycomm3 import LogixDriver
import csv
import pandas as pd
import re
import qdarktheme
import os
from collections import defaultdict

# Constants for better maintainability
WINDOW_SIZE = (1000, 800)
IP_INPUT_WIDTH = 400

basedir = os.path.dirname(__file__)

try:
    from ctypes import windll  # Only exists on Windows.
    tag_reader_tool_id = 'PM_Development.Tag_Reader_Tool.1.2'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(tag_reader_tool_id)
except ImportError:
    pass

def sanitize_filename(filename):
    """Remove or replace invalid characters from filename"""
    # Remove invalid characters
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    # Remove leading/trailing whitespace and dots
    filename = filename.strip().strip('.')
    # Remove .csv extension if present
    filename = re.sub(r'\.csv$', '', filename)
    return filename

def extract_index(tag):
    tag = re.sub(r'^Program:[^.]+\.', '', tag)
    match = re.findall(r'\[(\d+)\]', tag)
    if len(match) >= 2:
        return (int(match[0]) + int(match[1]))
    elif len(match) == 1:
        return int(match[0])
    else:
        return None

def should_format_as_array(data, top_level_is_list=False):
    """Pivot only for direct array reads or pure array data (no mixed struct fields)."""
    if top_level_is_list:
        return True

    has_unindexed_tags = any(extract_index(tag) is None for tag in data)
    if has_unindexed_tags:
        return False

    index_to_children = defaultdict(set)
    for tag in data:
        idx = extract_index(tag)
        if idx is not None:
            index_to_children[idx].add(extract_child_names(tag))

    if not index_to_children:
        return False

    return max(len(children) for children in index_to_children.values()) > 1

def extract_child_names(tag):
    match = re.search(r'\]\.(.+)', tag)
    if match:
        return match.group(1)

    match = re.search(r'^(.*?)(?=\[)', tag)
    if match:
        return match.group(1)

    return tag.split('.')[-1]

def format_csv(og_file, file, include_raw, is_array, save_location):
    try:
        # Use os.path.join for cross-platform compatibility
        df = pd.read_csv(os.path.join(save_location, f'{file}.csv'))
        df = df.fillna('')

        if is_array:
            df['index'] = df['tag'].apply(extract_index)
            df['child_name'] = df['tag'].apply(extract_child_names)

            df_pivot = df.pivot_table(index='index', columns='child_name', values='value', aggfunc='first')

            df_pivot.reset_index(inplace=True)

        else:
            df_pivot = df.set_index('tag').T

        rev_num = 1

        if os.path.exists(os.path.join(save_location, f'{og_file}.csv')):
            while os.path.exists(os.path.join(save_location, f'{og_file}_{rev_num}.csv')):
                rev_num += 1

            og_file = f'{og_file}_{rev_num}'

        df_pivot.to_csv(os.path.join(save_location, f'{og_file}.csv'), index=False)

        if not include_raw:
            os.remove(os.path.join(save_location, f'{file}.csv'))

        output_path = os.path.join(save_location, f'{og_file}.csv')
        return True, output_path

    except Exception as e:
        return False, f"Error formatting CSV: {e}"


def flatten_dict(d, parent_key='', sep='.'):
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        elif isinstance(v, list):
            for i, item in enumerate(v):
                if isinstance(item, dict):
                    items.extend(flatten_dict(
                        item, f"{new_key}[{i}]", sep=sep).items())
                else:
                    items.append((f"{new_key}[{i}]", item))
        else:
            items.append((new_key, v))

    return dict(items)


def write_to_csv(data, csv_file, include_raw, is_array, save_location):
    try:
        rev_num = 1
        og_file = csv_file

        if os.path.exists(os.path.join(save_location, f'{csv_file}_raw.csv')):
            while os.path.exists(os.path.join(save_location, f'{csv_file}_raw_{rev_num}.csv')):
                rev_num += 1
            csv_file = f'{csv_file}_raw_{rev_num}'
        else:
            csv_file = f'{csv_file}_raw'

        csv_path = os.path.join(save_location, f'{csv_file}.csv')
        with open(csv_path, 'w', newline='') as cf:
            writer = csv.DictWriter(cf, fieldnames=['tag', 'value'])
            writer.writeheader()
            for tag, value in data.items():
                writer.writerow({'tag': tag, 'value': value})

        return format_csv(og_file, csv_file, include_raw, is_array, save_location)

    except Exception as e:
        return False, f"Error writing CSV: {e}"


def read_tag(tag, ip, file_name_input, include_raw, save_location):
    try:
        with LogixDriver(ip) as plc:
            read_result = plc.read(tag)

        if read_result.error:
            return False, f"PLC read error: {read_result.error}"

        file_name_input = sanitize_filename(file_name_input)
        data = {read_result.tag: read_result.value}
        data = flatten_dict(data)
        is_array = should_format_as_array(
            data, top_level_is_list=type(read_result.value) is list)

        success, message = write_to_csv(
            data, file_name_input, include_raw, is_array, save_location)
        if success:
            return True, f"Tag read successfully!\n\nSaved to:\n{message}"
        return False, message

    except Exception as e:
        return False, f"Error reading tag: {e}"


class TagReadWorker(QThread):
    finished = Signal(bool, str)

    def __init__(self, tag, ip, file_name, include_raw, save_location):
        super().__init__()
        self.tag = tag
        self.ip = ip
        self.file_name = file_name
        self.include_raw = include_raw
        self.save_location = save_location

    def run(self):
        success, message = read_tag(
            self.tag, self.ip, self.file_name,
            self.include_raw, self.save_location)
        self.finished.emit(success, message)


class MainWindow(QMainWindow):
    
    def __init__(self):
        super(MainWindow, self).__init__()
        self.settings = QSettings("PM Development", "Tag Reader Tool")

        self.setWindowTitle("Tag Reader Tool")
        self.layout = QVBoxLayout()

        self.tag_input = QLineEdit()
        self.ip_input = QLineEdit()
        self.raw_file_checkbox = QCheckBox("Output Raw File")
        self.read_tag_button = QPushButton("Read Tag")
        self.file_name_input = QLineEdit()
        self.about_button = QPushButton("About")
        self.help_button = QPushButton("Help")

        # CSV Save Location Layout
        self.csv_save_path_layout = QHBoxLayout()
        self.csv_save_path_input = QLineEdit()
        self.csv_save_path_browse_button = QPushButton('Browse')
        self.csv_save_path_input.setPlaceholderText("CSV Save Location")
        self.csv_save_path_layout.addWidget(self.csv_save_path_input)
        self.csv_save_path_layout.addWidget(self.csv_save_path_browse_button)

        self.file_name_input.setPlaceholderText("Output File Name")
        self.ip_input.setPlaceholderText("Enter PLC IP")
        self.tag_input.setPlaceholderText("Enter Tag")

        # size ip input to be able to handle 40 characters
        self.ip_input.setFixedWidth(IP_INPUT_WIDTH)

        self.hor_layout = QHBoxLayout()

        self.layout.addWidget(self.ip_input)
        self.layout.addWidget(self.tag_input)
        self.layout.addWidget(self.raw_file_checkbox)
        self.layout.addWidget(self.read_tag_button)
        self.layout.addWidget(self.file_name_input)
        self.layout.addLayout(self.csv_save_path_layout)
        self.hor_layout.addWidget(self.about_button)
        self.hor_layout.addWidget(self.help_button)
        self.layout.addLayout(self.hor_layout)

        self._read_worker = None

        self.read_history()

        self.read_tag_button.clicked.connect(
            lambda: self.read_tag_clicked(self.tag_input.text(), self.ip_input.text(), self.csv_save_path_input.text()))
        
        self.about_button.clicked.connect(
            lambda: QMessageBox.about(self, "About", "This tool was written by Parker Mojsiejenko.\n\nIt uses the following libraries:\n - pycomm3\n - PySide6\n - pandas\n - qdarktheme"))

        self.help_button.clicked.connect(
            lambda: QMessageBox.about(self, "Help", "This tool requires tag names to be formatted in a specific way to read their data.\n\nIf the tag is an array, it will be in the following format: tag_name[start]{length}\n\nThe [start] can be omitted if you want to start at [0] and if the length is omitted, it will only read the [x] (or [0] if its omitted) member of the array.\n\nIf the tags are program scope tags, the tag name will need to start with Program:program_name.rest_of_tag_name.\n\nFor example, if you want to read a program scope array tag named my_array and start at the 5th member and read 50 members, the tag name would be my_array[4]{50} and if it was a program scope tag in the program my_program it would be Program:my_program.my_array[4]{50}\n\nIf the Output Raw File check box is checked, the non-formatted file will also be created. Both files will be saved in the specified folder you selected and if no folder was specified, the files will be saved where the EXE file resides.\n\nThe tool will not overwrite any previous CSV files. If one already exists with the same name, it will append a revision number that will increment each time a file is created and saved."))
        
        self.csv_save_path_browse_button.clicked.connect(
            lambda: self.csv_save_path_input.setText(QFileDialog.getExistingDirectory()))

        self.setFixedSize(self.layout.sizeHint())
        # Set central widget
        widget = QWidget()
        widget.setLayout(self.layout)
        self.setCentralWidget(widget)
        

    def read_tag_clicked(self, tag_input, ip_input, save_location):
        if self._read_worker and self._read_worker.isRunning():
            return

        try:
            if self.file_name_input.text() == '':
                file_name = tag_input
            else:
                file_name = self.file_name_input.text()

            self.validate_inputs(tag_input, ip_input, file_name)

            if save_location == '':
                save_location = '.'

            self._set_reading_state(True)
            self._read_worker = TagReadWorker(
                tag_input, ip_input, file_name,
                self.raw_file_checkbox.isChecked(), save_location)
            self._read_worker.finished.connect(self._on_read_finished)
            self._read_worker.start()

        except ValueError as e:
            QMessageBox.critical(self, "Input Error", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {e}")

    def _set_reading_state(self, reading):
        self.read_tag_button.setEnabled(not reading)
        self.read_tag_button.setText("Reading..." if reading else "Read Tag")

    def _on_read_finished(self, success, message):
        self._set_reading_state(False)
        self._read_worker = None

        if success:
            self.save_history()
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.warning(self, "Tag Read Failed", message)


    def read_history(self):
        self.ip_input.setText(self.settings.value('ip', ''))
        self.tag_input.setText(self.settings.value('tag', ''))
        self.file_name_input.setText(self.settings.value('file', ''))
        checked = self.settings.value('raw', "False")
        self.csv_save_path_input.setText(self.settings.value('save_path', ''))

        if checked == "True":
            checked = True
        else:
            checked = False

        self.raw_file_checkbox.setChecked(checked)


    def validate_inputs(self, tag, ip, file_name):
        """Validate user inputs before processing"""
        if not tag.strip():
            raise ValueError("Tag name cannot be empty")
        
        if not ip.strip():
            raise ValueError("IP address cannot be empty")
        
        if not file_name.strip():
            raise ValueError("File name cannot be empty")


    def save_history(self):
        self.settings.setValue('ip', self.ip_input.text())
        self.settings.setValue('tag', self.tag_input.text())
        self.settings.setValue('file', self.file_name_input.text())
        self.settings.setValue('raw', self.raw_file_checkbox.isChecked())
        self.settings.setValue('save_path', self.csv_save_path_input.text())

app = QApplication(sys.argv)
app.processEvents()
app.setWindowIcon(QtGui.QIcon(os.path.join(basedir, 'icon.ico')))
qdarktheme.setup_theme()
window = MainWindow()
window.resize(*WINDOW_SIZE)
window.show()

app.exec()
