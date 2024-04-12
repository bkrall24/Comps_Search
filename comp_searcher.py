
import sys
import polars as pl
import os
from PyQt5.QtWidgets import QApplication, QLabel, QVBoxLayout, QHBoxLayout, QWidget, QStackedLayout, QComboBox, QLineEdit, QCompleter, QListWidget, QAbstractItemView, QSpinBox, QPushButton, QMessageBox, QScrollArea
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtCore import  Qt, QUrl

from search_funcs import *

def generate_link_html(file_path):
    fn = os.path.basename(file_path)
    return f'<a href="file://{file_path}">{fn}</a>'

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.doc_ref = pl.read_excel("all_filename_detail_050442.xlsx")
        self.methods_list = None
        self.methods_str = None
        self.cmpd_list = None
        self.client = None
        (self.minyear, self.maxyear) = get_possible_years(self.doc_ref)
        self.lower_year = self.minyear
        self.upper_year = self.maxyear
        self.found_docs = pl.read_excel("all_filename_detail_050442.xlsx")
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Comps Finder')
        self.stacked_layout = QStackedLayout()
        self._createSearchPage()
        self._createFileDisplay()
        
        main_layout = QVBoxLayout()
        main_layout.addLayout(self.stacked_layout)
        self.setLayout(main_layout)

    def _createSearchPage(self):
        self.searchPage = QWidget()
        search_layout = QVBoxLayout(self.searchPage)

        info = QLabel("Search for Proposals/Reports Matching Criteria")
        font = info.font()
        font.setPointSize(16)
        font.setBold(True)
        info.setFont(font)
        search_layout.addWidget(info)

        method_label = QLabel('Choose from standard assays')
        self.method_combo = QListWidget()
        all_methods = get_possible_methods()
        self.method_combo.addItems(all_methods)
        self.method_combo.itemSelectionChanged.connect(self.method_choice)
        self.method_combo.setSelectionMode(QAbstractItemView.MultiSelection)
        search_layout.addWidget(method_label)
        search_layout.addWidget(self.method_combo)

        head_label = QLabel('Search for method name')
        self.heading_edit = QLineEdit()
        self.heading_edit.textChanged.connect(self.heading_choice)
        search_layout.addWidget(head_label)
        search_layout.addWidget(self.heading_edit)

        cmpd_label = QLabel('Choose from common compounds')
        self.cmpds_combo = QListWidget()
        all_cmpds = get_possible_compounds()
        self.cmpds_combo.addItems(all_cmpds)
        self.cmpds_combo.itemSelectionChanged.connect(self.cmpds_choice)
        self.cmpds_combo.setSelectionMode(QAbstractItemView.MultiSelection)
        search_layout.addWidget(cmpd_label)
        search_layout.addWidget(self.cmpds_combo)

        client_label = QLabel('Choose client')
        self.client_combo = QComboBox()
        all_clients = get_possible_clients()
        all_clients.insert(0, "")
        self.client_combo.addItems(all_clients)
        self.client_combo.setEditable(True)
        completer = QCompleter(all_clients, parent=self.client_combo)
        self.client_combo.setCompleter(completer)
        self.client_combo.activated.connect(self.client_choice)
        search_layout.addWidget(client_label)
        search_layout.addWidget(self.client_combo)

        # (self.minyear, self.maxyear) = get_possible_years(self.doc_ref)
        hbox = QHBoxLayout()
        label = QLabel('Select Year Range')
        hbox.addWidget(label)

        self.lower = QSpinBox()
        self.lower.setRange(self.minyear, self.maxyear)
        self.lower.setValue(self.minyear)
        self.lower.valueChanged.connect(self.update_lower)
        hbox.addWidget(self.lower)

        inbetween = QLabel('to')
        hbox.addWidget(inbetween)

        self.upper = QSpinBox()
        self.upper.setRange(self.minyear, self.maxyear)
        self.upper.setValue(self.maxyear)
        self.upper.valueChanged.connect(self.update_upper)
        hbox.addWidget(self.upper)
        search_layout.addLayout(hbox)

        enter_button = QPushButton("Search")
        enter_button.clicked.connect(self.search_docs)
        search_layout.addWidget(enter_button, alignment=Qt.AlignRight)
        self.stacked_layout.addWidget(self.searchPage)

    def _createFileDisplay(self):
        
        self.fileStack = QWidget()
        self.fileStackLayout = QVBoxLayout(self.fileStack)
        back_button = QPushButton("Back")
        back_button.clicked.connect(self.last_page)
        self.fileStackLayout.addWidget(back_button, alignment=Qt.AlignRight)

        self.comp_count = QLabel(str(len(self.found_docs)) + " documents found")
        self.fileStackLayout.addWidget(self.comp_count)
        
        self.fileDisplay = QWidget()
        self.scroll = QScrollArea()
        self.display_layout = QVBoxLayout(self.fileDisplay)
        self.link_labels = []
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(self.fileDisplay)
        self.fileStackLayout.addWidget(self.scroll)
        
        
        self.stacked_layout.addWidget(self.fileStack)
        
    def add_link(self, filepath):

        self.link_labels.append(QLabel())
        fl = generate_link_html(filepath)
        self.link_labels[-1].setText(fl)
        self.link_labels[-1].setOpenExternalLinks(True)
        self.link_labels[-1].setAlignment(Qt.AlignLeft)
        self.display_layout.addWidget(self.link_labels[-1])
        self.link_labels[-1].linkActivated.connect(self.open_document)

    def open_document(self, url):
        QDesktopServices.openUrl(QUrl(url))

    def method_choice(self):
        self.methods_list = [item.text() for item in self.method_combo.selectedItems()]
        if len(self.methods_list) == 0:
            self.methods_list = None

    def heading_choice(self):
        self.methods_str = self.heading_edit.text()
        if self.methods_str == '':
            self.methods_str = None
        
    
    def cmpds_choice(self):
        self.cmpd_list = [item.text() for item in self.cmpds_combo.selectedItems()]
        if len(self.cmpd_list) == 0:
            self.cmpd_list = None
    
    def client_choice(self):
        self.client = self.client_combo.currentText()
        if self.client == "":
            self.client = None

    def update_lower(self):
        self.lower_year = self.lower.value()
        self.upper.setRange(self.lower.value(), self.maxyear)
    
    def update_upper(self):
        self.upper_year = self.upper.value()
        self.lower.setRange(self.minyear, self.upper.value())
    
    def search_docs(self):

        dr = pl.read_excel("all_filename_detail_050442.xlsx")
        print(self.methods_str)
        if self.methods_str is not None:
            sms = search_method_str(self.methods_str, dr)
            if sms is None:
                QMessageBox.information(self, 'Text not found', self.methods_str + " not found in any document methods")
                return 
            
        else:
            sms = dr
        
        print(f"{len(sms)} after method string search")
        
        if self.methods_list is not None:
            fmm = find_matching_methods(self.methods_list, sms)
            if fmm is None:
                QMessageBox.information(self, 'No match found', "No documents matched criteria")
                return
        else:
            fmm = sms
        
        # print(f"{len(fmm)} after method list search")

        if self.cmpd_list is not None:
            fmc = find_matching_compounds(self.cmpd_list, fmm)
            if fmc is None:
                QMessageBox.information(self, 'No match found', "No documents matched criteria")
                return
        else:
            fmc = fmm
        
        # print(f"{len(fmc)} after compound list search")
        
        if self.client is not None:
            fcm = find_client_match(self.client, fmc)
            if fcm is None:
                QMessageBox.information(self, 'No match found', "No documents matched criteria")
                return
        else:
            fcm = fmc
        
        # print(f"{len(fcm)} after client search")

        if self.lower_year is None and self.upper_year is None:
            fin = fcm
        elif self.lower_year is not None and self.upper_year is None:
            fin = find_date_range(self.lower_year, fcm)
        elif self.lower_year == self.upper_year:
            fin = find_date_range(self.lower_year, fcm)
        else:
            fin = find_date_range((self.lower_year, self.upper_year), fcm)
        

        if len(fin) == 0:
            QMessageBox.information(self, 'No match found', "No documents matched criteria")
            return
        else:
            self.found_docs = fin

            filepaths = self.found_docs['filepath'].to_list()
            

            for row in filepaths:
                self.add_link(row)
        
            self.comp_count.setText(str(len(self.found_docs)) + " documents found:")
            self.stacked_layout.setCurrentIndex(1)
            return 
        
    def last_page(self):
        self.clearLayout(self.display_layout)
        self.methods_list = None
        self.methods_str = None
        self.heading_edit.clear()
        self.cmpd_list = None
        self.client = None
        self.lower_year = self.minyear
        self.upper_year = self.maxyear
        # self.found_docs = pl.read_excel("all_filename_detail_050442.xlsx")
        self.upper.setValue(self.maxyear)
        self.lower.setValue(self.minyear)
        self.stacked_layout.setCurrentIndex(0)
        # self.doc_ref = pl.read_excel("all_filename_detail_050442.xlsx")
    
    
    def clearLayout(self, layout):
        if layout is not None:
            while layout.count():
                child = layout.takeAt(0)
                if child.widget() is not None:
                    child.widget().deleteLater()
                elif child.layout() is not None:
                    self.clearLayout(child.layout())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())