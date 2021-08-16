from gui_models import WidgetModelsGUI, QSvaziModel
from document_models import DOCUMENT
from config import *
import sys
import os
from excel import *

from PyQt5 import QtCore, QtWidgets, QtGui,uic
from PyQt5.QtWidgets import QApplication

from qt_material import apply_stylesheet

app = QApplication(sys.argv)

apply_stylesheet(app, theme='light_blue.xml')

window = uic.loadUi("widget_gui.ui")

dict_document_models = {}
for a in get_documents():
    name = a['name']
    print(a['obligation'])
    if a['obligation'] == 'Да':
        obligation = True
    elif a['obligation'] == 'Нет':
        obligation = False
    if a['obligation_b'] == 'Да':
        obligation_b = True
    elif a['obligation_b'] == 'Нет':
        obligation_b = False
    if len(name.split('_')) != 1:
        name = name.split('_')[1]
    dict_document_models[name] = DOCUMENT(obligation=obligation,
                                          file_name=a['name'],
                                          folder_name=a['folder'],
                                          obligation_b=obligation_b
                                          )

# label_image = QtWidgets.QLabel()
# pred_image = QtGui.QPixmap('assets/outfile.png')
# label_image.setPixmap(pred_image)
# window.predprosmotr_layout.addWidget(label_image)

a = WidgetModelsGUI(window=window, dict_document_models=dict_document_models)
window.show()


sys.exit(app.exec_())