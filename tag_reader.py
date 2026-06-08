from PySide6.QtCore import QSettings, QThread, Signal, QUrl
from PySide6.QtGui import QDesktopServices
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
    QFileDialog,
)
from PySide6 import QtGui
import sys
from pycomm3 import LogixDriver
import csv
import logging
import re
import qdarktheme
import os
from collections import defaultdict

# Constants for better maintainability
WINDOW_SIZE = (1000, 800)
IP_INPUT_WIDTH = 400

basedir = os.path.dirname(__file__)


def get_default_save_location():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return basedir


logger = logging.getLogger('tag_reader')
_log_handler = None


def set_debug_logging(enabled):
    global _log_handler
    pycomm3_logger = logging.getLogger('pycomm3')

    if enabled:
        if _log_handler is None:
            log_path = os.path.join(get_default_save_location(), 'tag_reader.log')
            _log_handler = logging.FileHandler(log_path, encoding='utf-8')
            _log_handler.setFormatter(logging.Formatter(
                '%(asctime)s - %(levelname)s - %(message)s',
                datefmt='%m/%d/%Y %I:%M:%S %p',
            ))
            logger.addHandler(_log_handler)
            pycomm3_logger.addHandler(_log_handler)
        logger.setLevel(logging.DEBUG)
        pycomm3_logger.setLevel(logging.DEBUG)
    elif _log_handler is not None:
        logger.removeHandler(_log_handler)
        pycomm3_logger.removeHandler(_log_handler)
        _log_handler.close()
        _log_handler = None


def resolve_save_location(save_location):
    if not save_location.strip():
        save_location = get_default_save_location()

    save_location = os.path.abspath(save_location)

    if not os.path.isdir(save_location):
        raise ValueError(f"Save location does not exist:\n{save_location}")

    if not os.access(save_location, os.W_OK):
        raise ValueError(f"Save location is not writable:\n{save_location}")

    return save_location


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
    match = re.search(r'\[(\d+)\]', tag)
    if match:
        return int(match.group(1))
    return None


def natural_sort_key(name):
    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r'(\d+)', name)
    ]


def parse_array_field(field):
    match = re.match(r'^(.+)\[(\d+)\]$', field)
    if match:
        return match.group(1), int(match.group(2))
    return field, None

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

def get_revisioned_filename(base_name, save_location, suffix=''):
    name = f'{base_name}{suffix}'
    if not os.path.exists(os.path.join(save_location, f'{name}.csv')):
        return name

    rev_num = 1
    while os.path.exists(os.path.join(save_location, f'{base_name}{suffix}_{rev_num}.csv')):
        rev_num += 1
    return f'{base_name}{suffix}_{rev_num}'


def pivot_array_data(data, compact=False):
    if compact:
        return _pivot_array_compact(data)

    rows = defaultdict(dict)
    columns = set()

    for tag, value in data.items():
        idx = extract_index(tag)
        child = extract_child_names(tag)
        columns.add(child)
        rows[idx][child] = value

    column_list = sorted(columns, key=natural_sort_key)
    header = ['index'] + column_list
    body = [
        [idx] + [rows[idx].get(col, '') for col in column_list]
        for idx in sorted(rows.keys())
    ]
    return header, body


def _pivot_array_compact(data):
    scalars = defaultdict(dict)
    arrays = defaultdict(lambda: defaultdict(dict))

    for tag, value in data.items():
        idx = extract_index(tag)
        field = extract_child_names(tag)
        base_name, element_idx = parse_array_field(field)

        if element_idx is not None:
            arrays[idx][base_name][element_idx] = value
        else:
            scalars[idx][field] = value

    columns = set()
    for idx_values in scalars.values():
        columns.update(idx_values.keys())
    for idx_values in arrays.values():
        columns.update(idx_values.keys())

    column_list = sorted(columns, key=natural_sort_key)
    all_indices = set(scalars.keys()) | set(arrays.keys())
    header = ['index'] + column_list
    body = []

    for idx in sorted(all_indices):
        row = [idx]
        for col in column_list:
            if col in scalars.get(idx, {}):
                val = scalars[idx][col]
                row.append(val if val is not None else '')
            elif col in arrays.get(idx, {}):
                elements = arrays[idx][col]
                ordered = [
                    elements[i] if elements[i] is not None else ''
                    for i in sorted(elements.keys())
                ]
                row.append(', '.join(str(v) for v in ordered))
            else:
                row.append('')
        body.append(row)

    return header, body


def transpose_tag_data(data):
    tags = list(data.keys())
    return tags, [data[tag] for tag in tags]


def write_raw_csv(data, path):
    with open(path, 'w', newline='') as cf:
        writer = csv.DictWriter(cf, fieldnames=['tag', 'value'])
        writer.writeheader()
        for tag, value in data.items():
            writer.writerow({'tag': tag, 'value': value if value is not None else ''})


def write_formatted_csv(data, is_array, output_path, compact=False):
    with open(output_path, 'w', newline='') as f:
        writer = csv.writer(f)
        if is_array:
            header, body = pivot_array_data(data, compact=compact)
            writer.writerow(header)
            writer.writerows(body)
        else:
            header, values = transpose_tag_data(data)
            writer.writerow(header)
            writer.writerow([v if v is not None else '' for v in values])


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


def write_to_csv(data, csv_file, include_raw, is_array, save_location, compact=False):
    try:
        if include_raw:
            raw_name = get_revisioned_filename(csv_file, save_location, '_raw')
            write_raw_csv(data, os.path.join(save_location, f'{raw_name}.csv'))

        formatted_name = get_revisioned_filename(csv_file, save_location)
        output_path = os.path.join(save_location, f'{formatted_name}.csv')
        write_formatted_csv(data, is_array, output_path, compact=compact)
        return True, output_path

    except Exception as e:
        return False, f"Error writing CSV: {e}"


def read_tag(tag, ip, file_name_input, include_raw, save_location,
             log_enabled=False, compact=False):
    try:
        if log_enabled:
            logger.info(
                "Tag read requested: Tag: %s, IP: %s, File Name: %s",
                tag, ip, file_name_input)

        with LogixDriver(ip) as plc:
            read_result = plc.read(tag)

        if read_result.error:
            if log_enabled:
                logger.error("PLC read error: %s", read_result.error)
            return False, f"PLC read error: {read_result.error}", ''

        file_name_input = sanitize_filename(file_name_input)
        data = {read_result.tag: read_result.value}
        data = flatten_dict(data)
        is_array = should_format_as_array(
            data, top_level_is_list=type(read_result.value) is list)

        success, output_path = write_to_csv(
            data, file_name_input, include_raw, is_array, save_location,
            compact=compact)
        if success:
            if log_enabled:
                logger.info("Tag read successful, saved to: %s", output_path)
            return True, f"Tag read successfully!\n\nSaved to:\n{output_path}", output_path

        if log_enabled:
            logger.error("CSV write failed: %s", output_path)
        return False, output_path, ''

    except Exception as e:
        if log_enabled:
            logger.exception("Error reading tag")
        return False, f"Error reading tag: {e}", ''


class TagReadWorker(QThread):
    finished = Signal(bool, str, str)

    def __init__(self, tag, ip, file_name, include_raw, save_location,
                 log_enabled, compact):
        super().__init__()
        self.tag = tag
        self.ip = ip
        self.file_name = file_name
        self.include_raw = include_raw
        self.save_location = save_location
        self.log_enabled = log_enabled
        self.compact = compact

    def run(self):
        success, message, output_path = read_tag(
            self.tag, self.ip, self.file_name,
            self.include_raw, self.save_location,
            log_enabled=self.log_enabled, compact=self.compact)
        self.finished.emit(success, message, output_path)


class MainWindow(QMainWindow):
    
    def __init__(self):
        super(MainWindow, self).__init__()
        self.settings = QSettings("PM Development", "Tag Reader Tool")

        self.setWindowTitle("Tag Reader Tool")
        self.layout = QVBoxLayout()

        self.tag_input = QLineEdit()
        self.ip_input = QLineEdit()
        self.raw_file_checkbox = QCheckBox("Output Raw File")
        self.compact_wide_checkbox = QCheckBox("Compact Wide Output")
        self.debug_log_checkbox = QCheckBox("Enable Debug Logging")
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
        self.layout.addWidget(self.compact_wide_checkbox)
        self.layout.addWidget(self.debug_log_checkbox)
        self.layout.addWidget(self.read_tag_button)
        self.layout.addWidget(self.file_name_input)
        self.layout.addLayout(self.csv_save_path_layout)
        self.hor_layout.addWidget(self.about_button)
        self.hor_layout.addWidget(self.help_button)
        self.layout.addLayout(self.hor_layout)

        self._read_worker = None

        self.read_history()

        self.read_tag_button.clicked.connect(self._trigger_read)

        for line_edit in (
            self.ip_input, self.tag_input,
            self.file_name_input, self.csv_save_path_input,
        ):
            line_edit.returnPressed.connect(self._trigger_read)

        self.about_button.clicked.connect(
            lambda: QMessageBox.about(self, "About", "This tool was written by Parker Mojsiejenko.\n\nIt uses the following libraries:\n - pycomm3\n - PySide6\n - qdarktheme"))

        self.help_button.clicked.connect(
            lambda: QMessageBox.about(self, "Help", "This tool requires tag names to be formatted in a specific way to read their data.\n\nIf the tag is an array, it will be in the following format: tag_name[start]{length}\n\nThe [start] can be omitted if you want to start at [0] and if the length is omitted, it will only read the [x] (or [0] if its omitted) member of the array.\n\nIf the tags are program scope tags, the tag name will need to start with Program:program_name.rest_of_tag_name.\n\nFor example, if you want to read a program scope array tag named my_array and start at the 5th member and read 50 members, the tag name would be my_array[4]{50} and if it was a program scope tag in the program my_program it would be Program:my_program.my_array[4]{50}\n\nIf the Output Raw File check box is checked, the non-formatted file will also be created. Both files will be saved in the specified folder you selected and if no folder was specified, the files will be saved where the EXE file resides.\n\nThe tool will not overwrite any previous CSV files. If one already exists with the same name, it will append a revision number that will increment each time a file is created and saved."))
        
        self.csv_save_path_browse_button.clicked.connect(
            lambda: self.csv_save_path_input.setText(QFileDialog.getExistingDirectory()))

        widget = QWidget()
        widget.setLayout(self.layout)
        self.setCentralWidget(widget)
        self.setFixedSize(self.layout.sizeHint())

    def _trigger_read(self):
        self.read_tag_clicked(
            self.tag_input.text(),
            self.ip_input.text(),
            self.csv_save_path_input.text())

    def read_tag_clicked(self, tag_input, ip_input, save_location):
        if self._read_worker and self._read_worker.isRunning():
            return

        try:
            if self.file_name_input.text() == '':
                file_name = tag_input
            else:
                file_name = self.file_name_input.text()

            self.validate_inputs(tag_input, ip_input, file_name)
            save_location = resolve_save_location(save_location)

            log_enabled = self.debug_log_checkbox.isChecked()
            set_debug_logging(log_enabled)

            self._set_reading_state(True)
            self._read_worker = TagReadWorker(
                tag_input, ip_input, file_name,
                self.raw_file_checkbox.isChecked(), save_location,
                log_enabled, self.compact_wide_checkbox.isChecked())
            self._read_worker.finished.connect(self._on_read_finished)
            self._read_worker.start()

        except ValueError as e:
            QMessageBox.critical(self, "Input Error", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {e}")

    def _set_reading_state(self, reading):
        self.read_tag_button.setEnabled(not reading)
        self.read_tag_button.setText("Reading..." if reading else "Read Tag")

    def _on_read_finished(self, success, message, output_path):
        self._set_reading_state(False)
        self._read_worker = None

        if success:
            self.save_history()
            self._show_success_dialog(message, output_path)
        else:
            QMessageBox.warning(self, "Tag Read Failed", message)

    def _show_success_dialog(self, message, output_path):
        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Information)
        dialog.setWindowTitle("Success")
        dialog.setText(message)
        open_folder_btn = dialog.addButton("Open Folder", QMessageBox.ActionRole)
        dialog.addButton(QMessageBox.Ok)
        dialog.exec()

        if dialog.clickedButton() == open_folder_btn:
            folder = os.path.dirname(output_path)
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))


    def read_history(self):
        self.ip_input.setText(self.settings.value('ip', ''))
        self.tag_input.setText(self.settings.value('tag', ''))
        self.file_name_input.setText(self.settings.value('file', ''))
        self.raw_file_checkbox.setChecked(
            self.settings.value('raw', False, type=bool))
        self.compact_wide_checkbox.setChecked(
            self.settings.value('compact_wide', False, type=bool))
        self.debug_log_checkbox.setChecked(
            self.settings.value('debug_log', False, type=bool))
        self.csv_save_path_input.setText(self.settings.value('save_path', ''))


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
        self.settings.setValue('compact_wide', self.compact_wide_checkbox.isChecked())
        self.settings.setValue('debug_log', self.debug_log_checkbox.isChecked())
        self.settings.setValue('save_path', self.csv_save_path_input.text())

app = QApplication(sys.argv)
app.processEvents()
app.setWindowIcon(QtGui.QIcon(os.path.join(basedir, 'icon.ico')))
qdarktheme.setup_theme()
window = MainWindow()
window.show()

app.exec()
