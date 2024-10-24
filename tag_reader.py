from PySide6.QtCore import QSettings
from PySide6.QtWidgets import (
    QApplication,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QMessageBox,
    QHBoxLayout,
    QCheckBox
)
from PySide6 import QtGui
import sys
from pycomm3 import LogixDriver
import csv
import pandas as pd
import re
import qdarktheme
import os
import logging
from logging.handlers import RotatingFileHandler

handler = RotatingFileHandler('tag_reader.log', maxBytes=100000, backupCount=5)
logging.basicConfig(handlers=[handler], level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')


def format_csv(og_file, file, include_raw, is_array):
    '''
    Formats the csv file to be more readable by pivoting the data and saving it to a new file.

    Parameters:
        og_file (str):The original file name.
        file (str):The name of the raw file to be formatted.
        include_raw (bool):Whether or not to keep the raw file or delete it.
        is_array (bool):Whether or not the tag is an array.

    Returns:
        None
    '''

    logging.info(f"Formatting CSV: Original File: {og_file}, File: {file}, Include Raw: {include_raw}, Is Array: {is_array}")

    def extract_index(tag):
        '''
        Extracts the index from the tag.

        Parameters:
            tag (str):The tag name to extract the index from.

        Returns:
            int: The index of the tag.
        '''
        match = re.search(r'\[(\d+)\]', tag)
        return int(match.group(1)) if match else None


    def extract_child_names(tag):
        '''
        Extracts the child name from the tag.

        Parameters:
            tag (str):The tag name to extract the child name from.
        
        Returns:
            str: The child name of the tag.
        '''

        match = re.search(r'\]\.(.+)', tag)
        
        if match:
            return match.group(1)
        else:
            match = re.search(r'^(.*?)(?=\[)', tag)
            return match.group(1)


    df = pd.read_csv(f'{file}.csv')
    df = df.fillna('')

    if is_array:
        df['index'] = df['tag'].apply(extract_index)
        df['child_name'] = df['tag'].apply(extract_child_names)

        df_pivot = df.pivot_table(index='index', columns='child_name', values='value', aggfunc='first')

        df_pivot.reset_index(inplace=True)

    else:
        df_pivot = df.set_index('tag').T

    rev_num = 1

    if os.path.exists(f'{og_file}.csv'):
        while os.path.exists(f'{og_file}_{rev_num}.csv'):
            rev_num += 1

        og_file = f'{og_file}_{rev_num}'

    df_pivot.to_csv(f'{og_file}.csv', index=False)

    if not include_raw:
        # remove raw file
        os.remove(f'{file}.csv')


def flatten_dict(d, parent_key='', sep='.'):
    '''
    Flattens a dictionary to be able to write it to a csv file.
    
    Parameters:
        d (dict): The dictionary to flatten.
        parent_key (str): The parent key of the dictionary.
        sep (str): The separator to use between keys.
    
    Returns:
        dict: The flattened dictionary
    '''

    logging.info(f"Flattening Dictionary: Dictionary: {d}, Parent Key: {parent_key}, Separator: {sep}")

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


def write_to_csv(data, csv_file, include_raw, is_array):
    '''
    Writes the data to a csv file.

    Parameters:
        data (dict): The data to write to the csv file.
        csv_file (str): The name of the csv file to write to.
        include_raw (bool): Whether or not to include the raw file.
        is_array (bool): Whether or not the tag is an array.

    Returns:
        None
    '''

    logging.info(f"Writing to CSV: Data: {data}, CSV File: {csv_file}, Include Raw: {include_raw}, Is Array: {is_array}")

    rev_num = 1
    og_file = csv_file

    if os.path.exists(f'{csv_file}_raw.csv'):
        while os.path.exists(f'{csv_file}_raw_{rev_num}.csv'):
            rev_num += 1

        csv_file = f'{csv_file}_raw_{rev_num}'
    else:
        csv_file = f'{csv_file}_raw'

    with open(f'{csv_file}.csv', 'w', newline='') as cf:
        writer = csv.DictWriter(cf, fieldnames=['tag', 'value'])
        writer.writeheader()
        for tag, value in data.items():
            writer.writerow({'tag': tag, 'value': value})

    format_csv(og_file, csv_file, include_raw, is_array)


def read_tag(tag, ip, file_name_input, include_raw):
    '''
    Reads the tag from the PLC and writes it to a csv file.

    Parameters:
        tag (str): The tag to read from the PLC.
        ip (str): The IP address of the PLC.
        file_name_input (str): The name of the file to write the data to.
        include_raw (bool): Whether or not to include the raw file.
    
    Returns:
        None
    '''
    logging.info(f"Tag Read Requested: Tag: {tag}, IP: {ip}, File Name: {file_name_input}")

    try:
        with LogixDriver(ip) as plc:
            read_result = plc.read(tag)
    except Exception as e:
        logging.error(f"Connection Error: {e}")
        return

    # check if the file_name contains illegal characters
    file_name_input = re.sub(r'[<>:"/\\|?*]', '', file_name_input)

    # remove any leading or trailing whitespace
    file_name_input = file_name_input.strip()

    # remove file name extension if it exists
    file_name_input = re.sub(r'\.csv$', '', file_name_input)
    
    if not read_result.error:
        data = {read_result.tag: read_result.value}
    
        if type(read_result.value) is list:
            is_array = True
        else:
            is_array = False

        data = flatten_dict(data)
        write_to_csv(data, file_name_input, include_raw, is_array)
    else:
        logging.error(f"Tag Read Error: {read_result.error}")


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

        self.file_name_input.setPlaceholderText("Output File Name")
        self.ip_input.setPlaceholderText("Enter PLC IP")
        self.tag_input.setPlaceholderText("Enter Tag")

        # size ip input to be able to handle 40 characters
        self.ip_input.setFixedWidth(250)

        self.hor_layout = QHBoxLayout()

        self.layout.addWidget(self.ip_input)
        self.layout.addWidget(self.tag_input)
        self.layout.addWidget(self.raw_file_checkbox)
        self.layout.addWidget(self.file_name_input)
        self.layout.addWidget(self.read_tag_button)
        self.hor_layout.addWidget(self.about_button)
        self.hor_layout.addWidget(self.help_button)
        self.layout.addLayout(self.hor_layout)

        self.read_history()
        
        self.read_tag_button.clicked.connect(
            lambda: self.read_tag_clicked(self.tag_input.text(), self.ip_input.text()))
        
        self.about_button.clicked.connect(
            lambda: QMessageBox.about(self, "About", "This tool was written by Parker Mojsiejenko.\n\nIt uses the following libraries:\n - pycomm3\n - PySide6\n - pandas\n - qdarktheme"))

        self.help_button.clicked.connect(
            lambda: QMessageBox.about(self, "Help", "This tool requires tag names to be formatted in a specific way to read their data.\n\nIf the tag is an array, it will be in the following format: tag_name[start]{length}\n\nThe [start] can be omitted if you want to start at [0] and if the length is omitted, it will only read the [x] (or [0] if its omitted) member of the array.\n\nIf the tags are program scope tags, the tag name will need to start with Program:program_name.rest_of_tag_name.\n\nFor example, if you want to read a program scope array tag named my_array and start at the 5th member and read 50 members, the tag name would be my_array[4]{50} and if it was a program scope tag in the program my_program it would be Program:my_program.my_array[4]{50}\n\nTwo files will be outputted: one with the raw data and one with the data formatted in a more readable way. If you use a file name that already exists, it will overwrite the existing file. And if the file name is not entered, it will output the file name as tag_data.csv. These files will be saved in the same directory as the tool."))
        
        self.setFixedSize(self.layout.sizeHint())
        # Set central widget
        widget = QWidget()
        widget.setLayout(self.layout)
        self.setCentralWidget(widget)
        

    def read_tag_clicked(self, tag_input, ip_input):
        if self.file_name_input.text() == '':
            file_name = tag_input
        else:
            file_name = self.file_name_input.text()
        read_tag(tag_input, ip_input, file_name, include_raw=self.raw_file_checkbox.isChecked())
        self.save_history()


    def read_history(self):
        self.ip_input.setText(self.settings.value('ip', ''))
        self.tag_input.setText(self.settings.value('tag', ''))
        self.file_name_input.setText(self.settings.value('file', ''))
        checked = self.settings.value('raw', "False")

        if checked == "True":
            checked = True
        else:
            checked = False

        self.raw_file_checkbox.setChecked(checked)


    def save_history(self):
        self.settings.setValue('ip', self.ip_input.text())
        self.settings.setValue('tag', self.tag_input.text())
        self.settings.setValue('file', self.file_name_input.text())
        self.settings.setValue('raw', self.raw_file_checkbox.isChecked())

app = QApplication(sys.argv)
app.processEvents()
app.setWindowIcon(QtGui.QIcon('icon.ico'))
qdarktheme.setup_theme()
window = MainWindow()
window.resize(1000, 800)
window.show()

app.exec()
