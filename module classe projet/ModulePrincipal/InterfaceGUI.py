# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'GraphicalLayout.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QDialog,QWidget, QMessageBox
from ClassesProjet import *
from FormulesCouts import *

class Ui_Form(QWidget):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(864, 613)
        self.tabWidget = QtWidgets.QTabWidget(Form)
        self.tabWidget.setGeometry(QtCore.QRect(390, 220, 461, 331))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.tabWidget.addTab(self.tab_2, "")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(694, 60, 141, 31))
        self.pushButton.setObjectName("pushButton")
        self.treeView = QtWidgets.QTreeView(Form)
        self.treeView.setGeometry(QtCore.QRect(20, 250, 256, 192))
        self.treeView.setObjectName("treeView")
        self.comboBox = QtWidgets.QComboBox(Form)
        self.comboBox.setGeometry(QtCore.QRect(310, 60, 351, 31))
        self.comboBox.setObjectName("comboBox")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(10, 60, 141, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(20, 230, 71, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(10, 120, 331, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setGeometry(QtCore.QRect(550, 120, 111, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(10, 150, 321, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setGeometry(QtCore.QRect(10, 180, 261, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.lineEdit_3 = QtWidgets.QLineEdit(Form)
        self.lineEdit_3.setGeometry(QtCore.QRect(550, 180, 111, 20))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.dateEdit = QtWidgets.QDateEdit(Form)
        self.dateEdit.setGeometry(QtCore.QRect(550, 150, 110, 22))
        self.dateEdit.setObjectName("dateEdit")
        self.progressBar = QtWidgets.QProgressBar(Form)
        self.progressBar.setGeometry(QtCore.QRect(30, 526, 261, 23))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(750, 570, 101, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.retranslateUi(Form)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(Form)
        self.initUI()
        
    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Antofagasta : Modélisation Standardisée des projets D\'investissement"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("Form", "Plan d\'investissement"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("Form", "Plan de financement"))
        self.pushButton.setText(_translate("Form", "Valider le choix"))
        self.label.setText(_translate("Form", "Selectionnez un projet"))
        self.label_2.setText(_translate("Form", "Activités"))
        self.label_3.setText(_translate("Form", "Quel horizon de qualification souhaitez-vous ?"))
        self.label_4.setText(_translate("Form", "Quelle est la date du début des investissements ?"))
        self.label_5.setText(_translate("Form", "Quel est la durée des investissements ?"))
        self.pushButton_2.setText(_translate("Form", "Suivant"))
    def initUI(self):
        self.Projet = []
        self.Activites = []
        #Le code D'initialisation des données et des liaisons
        self.ListeProjets = ['----------','Ecole Secondaire','Clinique','Restaurant','Projet Immobilier', 'Usine']
        self.pushButton.clicked.connect(self.on_confirm_click)
        self.comboBox.addItems(self.ListeProjets)
    def on_confirm_click(self):
        if self.comboBox.currentText() == '----------':
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Erreur choix de projet")
            msg.setInformativeText("Veuillez choisir un projet Valide")
            msg.setWindowTitle("Choix Invalide")
            msg.exec_()
        else: 
            print(self.comboBox.currentText())
            print(self.lineEdit.text()) #Horizon
            print(self.lineEdit_3.text()) #Durée des Investissement
            
app = QApplication(sys.argv)
window = QDialog()
ui = Ui_Form()
ui.setupUi(window)
window.show()
sys.exit(app.exec_())
