import sys
import excel

import fitz

from config import *
from document_models import KONTRAGENT, DOCUMENT
from pyqt_methods import *
from PyQt5 import QtCore, QtWidgets, QtGui,uic, Qt
from PyQt5.QtWidgets import QApplication

from excel import get_documents, get_kontragent, list_to_csv

from database import *
from database_models import Svazi

class QTreeView(QtWidgets.QTreeView):
    def __init__(self):
        super().__init__()
        self.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Space:
            self.test_method()

    def test_method(self):
        self.index = self.selectedIndexes()[0]
        print(self.windowFilePath())
        print('Space key pressed')

    def edit(self, index, trigger, event):
        if trigger == QtWidgets.QAbstractItemView.DoubleClicked:
            return False
        return QtWidgets.QTreeView.edit(self, index, trigger, event)

class WidgetModelsGUI():
    def __init__(self, window, dict_document_models):
        self.window = window
        self.model = QtWidgets.QFileSystemModel()
        self.model.setRootPath(DIR_HRANILIHE)
        self.model.setReadOnly(False)

        self.tree_view = QTreeView()
        self.window.tree_layout.addWidget(self.tree_view)
        self.tree_view.setModel(self.model)
        self.tree_view.setRootIndex(self.model.index(DIR_HRANILIHE))
        self.tree_view.doubleClicked.connect(self.test)
        self.tree_view.clicked.connect(self.save_selected_filepath)
        self.dict_document_models = dict_document_models

        self.window.openfile.clicked.connect(self.open_file)

        self.window.export_button.clicked.connect(self.export_documents_csv)

        self.window.svazi_button.clicked.connect(self.add_to_svazi)

        self.db = SessionLocal()
        Base.metadata.create_all(bind=engine)

        self.svazi_model = QSvaziModel(window=self.window, parent=self)

        self.content_indicators = QtWidgets.QVBoxLayout()
        self.content_indicators.setAlignment(QtCore.Qt.AlignTop)
        self.window.indicator_layout_widget.setLayout(self.content_indicators)


    def test(self, signal):
        deleteItemsOfLayout(self.window.metadata_layout)
        # deleteItemsOfLayout(self.window.indicator_layout)

        deleteItemsOfLayout(self.window.predprosmotr_layout)
        self.window.contragent_e.setStyleSheet("")
        self.window.contragent_b.setStyleSheet("")

        self.init_kontragent()
        self.file_path=self.model.filePath(signal)
        print('DIR_HRANILIHE: ', DIR_HRANILIHE)
        print(self.file_path)
        self.kontragent_name = self.file_path.replace(DIR_HRANILIHE, '').split('/')[0]
        print('self.kontragent_name: ', self.kontragent_name)

        if self.file_path.find('.') != -1:
            self.check_metadata(path_file=self.file_path)
            self.create_metadata_block()
            for self.a, self.b in self.metadata_form.items():
                self.b.setText( self.metadata_document[self.a] )

            if self.file_path.find('.pdf') != -1:
                doc = fitz.open(self.file_path)
                page = doc.loadPage(0)
                pix = page.getPixmap()
                pix.writePNG("outfile.png")

                label_image = QtWidgets.QLabel()
                # label_image.setMaximumSize(300, 300)
                label_image.setScaledContents(True)
                pred_image = QtGui.QPixmap('outfile.png')
                label_image.setPixmap(pred_image)
                self.window.predprosmotr_layout.addWidget(label_image)
            elif self.file_path.find('.png') != -1 or self.file_path.find('.jpg') != -1:
                label_image = QtWidgets.QLabel()
                # label_image.setMaximumSize(300, 300)
                label_image.setScaledContents(True)
                pred_image = QtGui.QPixmap(self.file_path)
                label_image.setPixmap(pred_image)
                self.window.predprosmotr_layout.addWidget(label_image)

        # self.create_indicator_block()
        self.create_check_and_paste_obligation_file()
        self.svazi_model.update_main_svaz(path_svaz=self.file_path)
        self.svazi_model.update_list()
        # self.window.NP.setCheckState(True)
        # print(self.file_path)
        # print(DIR_HRANILIHE)
        # print(self.file_path.replace(DIR_HRANILIHE, '').split('/')[0])

    def check_metadata(self, path_file):
        self.only_filename = path_file.split('/')[-1].split('.')[0]
        self.metadata_document = {}
        try:
            if len( self.only_filename.split('_') ) == 1:
                self.metadata_document = self.dict_document_models[self.only_filename].pass_metadata(self.only_filename)
            else:
                self.metadata_document = self.dict_document_models[self.only_filename.split('_')[1]].pass_metadata(self.only_filename)
                self.metadata_document['Бумажная версия'] = self.dict_document_models[self.only_filename.split('_')[1]].check_bum_v(self.only_filename)
        except Exception as err:
            print('Не найден прототип такого файла:', err)
            self.metadata_document['Название документа'] = self.only_filename

    def init_kontragent(self):
        pass

    def create_indicator_block(self):
        self.dict_indicators = {}
        self.content_indicators = QtWidgets.QVBoxLayout()
        self.content_indicators.setAlignment(QtCore.Qt.AlignTop)
        self.window.indicator_layout_widget.setLayout(self.content_indicators)

        for self.a, self.b in self.dict_document_models.items():
            self.label = QtWidgets.QLabel(self.a)
            self.label.setFixedHeight(20)
            self.indicator_e = QtWidgets.QTextEdit()
            self.indicator_e.setFixedSize(16,16)
            # self.indicator_e.setMinimumSize(16, 16)
            self.indicator_b = QtWidgets.QTextEdit()
            self.indicator_b.setFixedSize(16,16)
            # self.indicator_b.setMaximumSize(16, 16)
            # self.indicator_b.setMinimumSize(16, 16)

            self.hor_layout = QtWidgets.QHBoxLayout(self.window.indicator_layout_widget)
            self.hor_layout.addWidget(self.label)
            self.hor_layout.addWidget(self.indicator_e)
            self.hor_layout.addWidget(self.indicator_b)
            self.content_indicators.addLayout(self.hor_layout)

            self.dict_indicators[self.a+'_e'] = self.indicator_e
            self.dict_indicators[self.a + '_b'] = self.indicator_b

    def create_metadata_block(self):
        self.metadata_form = {}
        for self.a, self.b in self.metadata_document.items():
            if self.a == 'Бумажная версия':
                self.bum_v_checkbox = QtWidgets.QCheckBox('Бумажная версия')
                self.bum_v_checkbox.setChecked(self.b)
                self.window.metadata_layout.addWidget(self.bum_v_checkbox)
                continue

            self.label = QtWidgets.QLabel(self.a)
            self.textedit = QtWidgets.QTextEdit()
            self.textedit.setMaximumHeight(40)
            # self.textedit.setMinimumSize(280, 20)

            self.hor_layout = QtWidgets.QHBoxLayout()
            self.hor_layout.setAlignment(QtCore.Qt.AlignTop)
            self.hor_layout.addWidget(self.label)
            self.hor_layout.addWidget(self.textedit)
            self.window.metadata_layout.addLayout(self.hor_layout)

            self.metadata_form[self.a] = self.textedit

        self.button_save_metadata = QtWidgets.QPushButton('Сохранить')
        self.button_save_metadata.clicked.connect(self.save_metadata)
        self.hor_layout = QtWidgets.QHBoxLayout()
        self.hor_layout.setAlignment(QtCore.Qt.AlignTop)
        self.hor_layout.addWidget(self.button_save_metadata)
        self.window.metadata_layout.addLayout(self.hor_layout)

    def create_check_and_paste_obligation_file(self):
        try:
            deleteItemsOfLayout(self.content_indicators)


            self.kontragent = KONTRAGENT(document_models=self.dict_document_models, kontragent_name=self.kontragent_name)

            self.uniq_kontragent_obligation = get_kontragent()

            if self.uniq_kontragent_obligation.get(self.kontragent_name) != None:
                for self.kontragent_document_name, self.kontragent_document_model in self.kontragent.document_models.items():
                    if self.kontragent_document_name in self.uniq_kontragent_obligation[self.kontragent_name][0]:
                        self.kontragent_document_model.obligation = True
                    else:
                        self.kontragent_document_model.obligation = False

            if self.uniq_kontragent_obligation.get(self.kontragent_name) != None:
                for self.kontragent_document_name, self.kontragent_document_model in self.kontragent.document_models.items():
                    if self.kontragent_document_name in self.uniq_kontragent_obligation[self.kontragent_name][1]:
                        self.kontragent_document_model.obligation_b = True
                    else:
                        self.kontragent_document_model.obligation_b = False


            self.color_kontragent_e = GREEN
            self.color_kontragent_b = GREEN

            for self.a, self.b in self.kontragent.check().items():
                for self.b_ in self.b:
                    self.label = QtWidgets.QTextEdit(self.b_[2])
                    self.label.setFixedHeight(40)
                    self.label.setReadOnly(True)
                    self.indicator_e = QtWidgets.QTextEdit()
                    self.indicator_e.setFixedSize(16, 16)
                    self.indicator_b = QtWidgets.QTextEdit()
                    self.indicator_b.setFixedSize(16, 16)

                    self.hor_layout = QtWidgets.QHBoxLayout(self.window.indicator_layout_widget)
                    self.hor_layout.addWidget(self.label)
                    self.hor_layout.addWidget(self.indicator_e)
                    self.hor_layout.addWidget(self.indicator_b)
                    self.content_indicators.addLayout(self.hor_layout)

                    if self.dict_document_models[self.a].obligation == True and self.b_[0] == False:
                        self.indicator_e.setStyleSheet(RED)
                        self.color_kontragent_e = RED
                    elif self.dict_document_models[self.a].obligation == True and self.b_[0] == True:
                        self.indicator_e.setStyleSheet(GREEN)
                    elif self.dict_document_models[self.a].obligation == False and self.b_[0] == True:
                        self.indicator_e.setStyleSheet(GREEN)
                    elif self.dict_document_models[self.a].obligation == False and self.b_[0] == False:
                        self.indicator_e.setStyleSheet(YELLOW)
                        if self.color_kontragent_e != RED:
                            self.color_kontragent_e = YELLOW

                    if self.dict_document_models[self.a].obligation_b == True and self.b_[1] == False:
                        self.indicator_b.setStyleSheet(RED)
                        self.color_kontragent_b = RED
                    elif self.dict_document_models[self.a].obligation_b == True and self.b_[1] == True:
                        self.indicator_b.setStyleSheet(GREEN)
                    elif self.dict_document_models[self.a].obligation_b == False and self.b_[1] == True:
                        self.indicator_b.setStyleSheet(GREEN)
                    elif self.dict_document_models[self.a].obligation_b == False and self.b_[1] == False:
                        self.indicator_b.setStyleSheet(YELLOW)
                        if self.color_kontragent_b != RED:
                            self.color_kontragent_b = YELLOW

            self.window.contragent_e.setStyleSheet(self.color_kontragent_e)
            self.window.contragent_b.setStyleSheet(self.color_kontragent_b)

        except Exception as err:
            print('Error: ', err)

    def save_metadata(self):
        try:
            print("Сохраняем")
            self.rashir = self.file_path.split('.')[1]
            self.dir_file_save = self.file_path[:self.file_path.rfind('/')]
            print(self.dir_file_save)
            self.new_name_file = ''
            self.i = 0
            for self.a, self.b in self.metadata_form.items():
                if self.i != 0:
                    self.new_name_file += '_'
                    self.i += 1
                if self.a == 'Дата':
                    self.non_redact_date_split = self.b.toPlainText().split('.')
                    self.non_redact_date = str(self.non_redact_date_split[2])+\
                                           str(self.non_redact_date_split[1])+\
                                           str(self.non_redact_date_split[0])
                    self.new_name_file += self.non_redact_date
                    self.i += 1
                else:
                    self.new_name_file += self.b.toPlainText()
                    self.i += 1

            try:
                if self.bum_v_checkbox.isChecked() == True:
                    self.new_name_file += '_Б'
            except:
                pass

            self.new_name_file += '.' + self.rashir
            self.new_name_file_w_dir = self.dir_file_save + '/' + self.new_name_file
            print(self.new_name_file_w_dir)
            os.rename(src=self.file_path, dst=self.new_name_file_w_dir)
            self.file_path = self.new_name_file_w_dir
        except Exception as err:
            print(err)

    def open_file(self):
        try:
            filepath = self.file_path
            os.startfile(filepath)
        except:
            pass

    def export_documents_csv(self):
        self.dict_res_all_check = {}
        for self.kontragent_name in os.listdir(DIR_HRANILIHE):
            self.kontragent = KONTRAGENT(document_models=self.dict_document_models, kontragent_name=self.kontragent_name)
            self.res_check = self.kontragent.check_w_filename()
            self.dict_res_all_check[self.kontragent_name] = self.res_check
        print(self.dict_res_all_check)

        self.all_name_rows = []
        for self.name_f, self.model_document_f in self.dict_document_models.items():
            self.i = 0
            for self.metadata_row in self.model_document_f.form_metadata:
                if self.i == 1:
                    self.i += 1
                    continue
                if self.metadata_row not in self.all_name_rows:
                    self.all_name_rows.append(self.metadata_row)
                self.i += 1
        self.all_name_rows.append('Бумажная версия')
        self.all_name_rows.append('Контрагент')
        print(self.all_name_rows)

        self.list_export_data = []
        for self.name_kontragent, self.documents_kontragent in self.dict_res_all_check.items():
            for self.one_document_name, self.one_document_filename in self.documents_kontragent.items():
                for self.odf in self.one_document_filename:
                    self.dict_one_export_data = {}
                    self.one_check_metadata = self.dict_document_models[self.one_document_name].pass_metadata(self.odf)
                    for self.one_metadata in self.all_name_rows:
                        if self.one_metadata == 'Контрагент':
                            self.dict_one_export_data['Контрагент'] = self.name_kontragent
                        elif self.one_metadata == 'Бумажная версия':
                            self.ch_bum_v = self.dict_document_models[self.one_document_name].check_bum_v(self.odf)
                            if self.ch_bum_v == True:
                                self.dict_one_export_data['Бумажная версия'] = 'Да'
                            if self.ch_bum_v == False:
                                self.dict_one_export_data['Бумажная версия'] = 'Нет'
                        else:
                            try:
                                self.dict_one_export_data[self.one_metadata] = self.one_check_metadata[self.one_metadata]
                            except:
                                self.dict_one_export_data[self.one_metadata] = ''
                    self.list_export_data.append(self.dict_one_export_data)

        print(self.list_export_data)
        try:
            list_to_csv(list_export=self.list_export_data)
        except Exception as err:
            self.error_dialog = QtWidgets.QErrorMessage()
            self.error_dialog.showMessage('Ошибка: ' + str(err))

    def add_to_svazi(self):
        print(self.file_path, self.selected_path)
        self.active_svaz = self.db.query(Svazi).filter_by(main=self.file_path).all()
        print(self.active_svaz)

        self.new_svaz = Svazi(
            main=self.file_path,
            second=self.selected_path
        )
        self.db.add(self.new_svaz)
        self.db.commit()

        self.svazi_model.update_list()

    def save_selected_filepath(self, signal):
        self.selected_path = self.model.filePath(signal)

    def rec_from_qlist(self, path):
        print(path)
        deleteItemsOfLayout(self.window.metadata_layout)
        # deleteItemsOfLayout(self.window.indicator_layout)
        deleteItemsOfLayout(self.window.predprosmotr_layout)
        self.window.contragent_e.setStyleSheet("")
        self.window.contragent_b.setStyleSheet("")

        self.init_kontragent()
        self.file_path = path
        self.kontragent_name = self.file_path.replace(DIR_HRANILIHE, '').split('/')[0]

        if self.file_path.find('.') != -1:
            self.check_metadata(path_file=self.file_path)
            self.create_metadata_block()
            for self.a, self.b in self.metadata_form.items():
                self.b.setText(self.metadata_document[self.a])

            if self.file_path.find('.pdf') != -1:
                doc = fitz.open(self.file_path)
                page = doc.loadPage(0)
                pix = page.getPixmap()
                pix.writePNG("outfile.png")

                label_image = QtWidgets.QLabel()
                # label_image.setMaximumSize(300, 300)
                label_image.setScaledContents(True)
                pred_image = QtGui.QPixmap('outfile.png')
                label_image.setPixmap(pred_image)
                self.window.predprosmotr_layout.addWidget(label_image)
            elif self.file_path.find('.png') != -1 or self.file_path.find('.jpg') != -1:
                label_image = QtWidgets.QLabel()
                # label_image.setMaximumSize(300, 300)
                label_image.setScaledContents(True)
                pred_image = QtGui.QPixmap(self.file_path)
                label_image.setPixmap(pred_image)
                self.window.predprosmotr_layout.addWidget(label_image)

        self.create_check_and_paste_obligation_file()


class QSvaziModel():
    def __init__(self, window, parent):
        self.db = SessionLocal()
        self.parent = parent

        self.window = window
        self.hor_layout = QtWidgets.QHBoxLayout()
        self.obr_svazi_box = QtWidgets.QCheckBox('Обратные связи')
        self.update_all_documents_button = QtWidgets.QPushButton('Обновить')
        self.qlist_svazi = QtWidgets.QListWidget()
        self.label_main_svaz = QtWidgets.QLabel('Не выбрано')
        self.del_button = QtWidgets.QPushButton('Удалить')
        self.qlist_svazi.itemActivated.connect(self.send_to_main_widget)
        self.qlist_svazi.itemClicked.connect(self.update_selected_item)
        self.del_button.clicked.connect(self.del_item)
        self.update_all_documents_button.clicked.connect(self.update_os_files)

        self.hor_layout.addWidget(self.obr_svazi_box)
        self.hor_layout.addWidget(self.update_all_documents_button)
        self.window.svazi_layout.addLayout(self.hor_layout)
        self.window.svazi_layout.addWidget(self.label_main_svaz)
        self.window.svazi_layout.addWidget(self.qlist_svazi)
        self.window.svazi_layout.addWidget(self.del_button)

    def update_main_svaz(self, path_svaz):
        self.main_svaz_path = path_svaz
        self.label_main_svaz.setText(path_svaz.replace(DIR_HRANILIHE, '').split('/')[-1])

    def update_list(self):
        self.qlist_svazi.clear()
        self.list_documents_path = []
        self.dict_active_documents = {}

        if self.obr_svazi_box.isChecked():
            self.active_svaz_documents = self.db.query(Svazi).filter_by().all()
            for self.one_document in self.active_svaz_documents:
                if self.one_document.main == self.main_svaz_path:
                    self.only_file_name = self.one_document.second.replace(DIR_HRANILIHE, '').split('/')[-1]
                    self.dict_active_documents[self.only_file_name] = self.one_document.second
                    self.item1 = QtWidgets.QListWidgetItem(self.only_file_name)
                    self.qlist_svazi.addItem(self.item1)
                elif self.one_document.second == self.main_svaz_path:
                    self.only_file_name = self.one_document.main.replace(DIR_HRANILIHE, '').split('/')[-1]
                    self.dict_active_documents[self.only_file_name] = self.one_document.main
                    self.item1 = QtWidgets.QListWidgetItem(self.only_file_name)
                    self.qlist_svazi.addItem(self.item1)
        else:
            self.active_svaz_documents = self.db.query(Svazi).filter_by(main=self.main_svaz_path).all()
            for self.one_document in self.active_svaz_documents:
                self.only_file_name = self.one_document.second.replace(DIR_HRANILIHE, '').split('/')[-1]
                self.dict_active_documents[self.only_file_name] = self.one_document.second
                self.item1 = QtWidgets.QListWidgetItem(self.only_file_name)
                self.qlist_svazi.addItem(self.item1)

    def send_to_main_widget(self, item):
        self.parent.rec_from_qlist(self.dict_active_documents[item.text()])

    def del_item(self):
        self.active_svaz_documents = self.db.query(Svazi).filter_by(main=self.main_svaz_path).all()
        for self.one_document in self.active_svaz_documents:
            if self.one_document.second == self.dict_active_documents[self.select_item.text()]:
                self.db.delete(self.one_document)
        self.db.commit()
        self.qlist_svazi.takeItem(self.qlist_svazi.row(self.select_item))

        self.active_svaz_documents = self.db.query(Svazi).filter_by(second=self.main_svaz_path).all()
        for self.one_document in self.active_svaz_documents:
            if self.one_document.main == self.dict_active_documents[self.select_item.text()]:
                self.db.delete(self.one_document)
        self.db.commit()
        self.qlist_svazi.takeItem(self.qlist_svazi.row(self.select_item))

    def update_selected_item(self, item):
        self.select_item = item

    def update_os_files(self):
        self.list_del_svazi = []
        self.all_bd_svazi = self.db.query(Svazi).filter_by().all()
        for self.one_document in self.all_bd_svazi:
            if os.path.isfile(self.one_document.second) != True:
                self.db.delete(self.one_document)
                self.list_del_svazi.append(self.one_document.second)
            elif os.path.isfile(self.one_document.main) != True:
                self.db.delete(self.one_document)
                self.list_del_svazi.append(self.one_document.main)

        self.del_dialog = QtWidgets.QErrorMessage()
        print(self.list_del_svazi)
        self.del_dialog.showMessage('Были удалены: '+','.join(self.list_del_svazi))
        self.db.commit()

        self.update_list()
