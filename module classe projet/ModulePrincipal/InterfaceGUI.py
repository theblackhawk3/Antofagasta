# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'GraphicalLayout.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import QApplication, QDialog,QWidget, QMessageBox,QTableWidget,QTableWidgetItem
from ClassesProjet import *
from FormulesCouts import *

class Ui_Form(QWidget):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1000, 700)
        self.tabWidget = QtWidgets.QTabWidget(Form)
        self.tabWidget.setGeometry(QtCore.QRect(390, 220, 600, 400))
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
        self.treeView = QtWidgets.QTreeWidget(Form)
        self.treeView.setGeometry(QtCore.QRect(20, 250, 256, 192))
        self.treeView.setObjectName("treeView")
        self.buttonConfirmActivity = QtWidgets.QPushButton(Form)
        self.buttonConfirmActivity.setGeometry(QtCore.QRect(290, 340, 80, 31))
        self.buttonConfirmActivity.setObjectName("pushButton_3")
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
        self.pushButton_2.setGeometry(QtCore.QRect(885, 630, 101, 31))
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
        self.buttonConfirmActivity.setText(_translate("Form", " >> "))
        
    def initUI(self):
        self.treeView.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.Projet = []
        self.Activites = []
        self.ActivitesSecondaires = []
        #Le code D'initialisation des données et des liaisons
        self.ListeProjets = ['----------','Ecole Secondaire','Clinique','Restaurant','Projet Immobilier', 'Usine']
        self.pushButton.clicked.connect(self.on_confirm_click)
        self.comboBox.addItems(self.ListeProjets)
        self.progressBar.setProperty("value", 0)
        self.buttonConfirmActivity.clicked.connect(self.on_confirm_click2)
        self.ProjetValide = False
        self.ActiviteValide = False
        self.createTable()
        
    def on_confirm_click(self):
        if self.ProjetValide == False:
            if self.comboBox.currentText() == '----------':
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Information)
                msg.setText("Erreur choix de projet")
                msg.setInformativeText("Veuillez choisir un projet Valide")
                msg.setWindowTitle("Choix Invalide")
                msg.exec_()
            else: 
                self.Projet = []
                self.Activites = []
                print(self.comboBox.currentText())
                print(self.lineEdit.text()) #Horizon
                print(self.lineEdit_3.text()) #Durée des Investissement
                print(self.dateEdit.text())
                self.Projet = Projet(self.comboBox.currentText())
                self.Projet.IdentifierActivitesPossibles()
                for Activite in self.Projet.getListeActivites():
                    ActiviteWidget = QtWidgets.QTreeWidgetItem(self.treeView)
                    ActiviteWidget.setText(0,Activite.getNom())
                    ActiviteWidget.setFlags(ActiviteWidget.flags() | Qt.ItemIsUserCheckable)
                    ActiviteWidget.setCheckState(0, Qt.Checked)
                    self.Activites.append(ActiviteWidget)
                for Activite in self.Projet.getListeActivitesSecondaires():
                    ActiviteWidget = QtWidgets.QTreeWidgetItem(self.treeView)
                    ActiviteWidget.setText(0,Activite.getNom())
                    ActiviteWidget.setFlags(ActiviteWidget.flags() | Qt.ItemIsUserCheckable)
                    ActiviteWidget.setCheckState(0, Qt.Unchecked)
                    self.Activites.append(ActiviteWidget)
            self.Projet.AjoutActivites(self.Projet.ListeActivitesSecondaires)
            self.ProjetValide = True
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Deja Validé")
            msg.setInformativeText("Vous avez deja validé le projet")
            msg.setWindowTitle("Projet déja validé")
            msg.exec_()
        
    def on_confirm_click2(self):
        if self.ActiviteValide == False:
            for i in self.Activites:
                if i.checkState(0) == 0:
                    for j in self.Projet.ListeActivites:
                        if j.getNom() == i.text(0):
                            self.Projet.ListeActivites.pop(self.Projet.ListeActivites.index(j))
            self.tableWidget.setColumnCount(int(self.lineEdit.text())+2)
            self.tableWidget_2.setColumnCount(int(self.lineEdit.text())+2)
            self.tableWidget.item(1,0).setBackground(QtGui.QColor("red"))
            self.tableWidget.item(1,1).setBackground(QtGui.QColor("red"))
            self.tableWidget.item(1,0).setFont(QFont("Arial",12,italic =  True))
            self.tableWidget.item(1,1).setFont(QFont("Arial",12,italic =  True))
            self.tableWidget_2.item(1,0).setBackground(QtGui.QColor("red"))
            self.tableWidget_2.item(1,1).setBackground(QtGui.QColor("red"))
            self.tableWidget_2.item(1,0).setFont(QFont("Arial",12,italic =  True))
            self.tableWidget_2.item(1,1).setFont(QFont("Arial",12,italic =  True))
            for i in range(2,int(self.lineEdit.text())+2):
                self.tableWidget.setItem(1,i,QTableWidgetItem("A"+str(i-1)))
                self.tableWidget_2.setItem(1,i,QTableWidgetItem("A"+str(i-1)))
                self.tableWidget.item(1,i).setBackground(QtGui.QColor("red"))
                self.tableWidget_2.item(1,i).setBackground(QtGui.QColor("red"))                
                self.tableWidget.item(1,i).setFont(QFont("Arial",12,italic =  True))
                self.tableWidget_2.item(1,i).setFont(QFont("Arial",12,italic =  True))
            for i in range(2,5):
                for j in range(2,int(self.lineEdit.text())+2):
                    self.tableWidget_2.setItem(i,j,QTableWidgetItem("a"))
                    self.tableWidget_2.item(i,j).setFlags( Qt.ItemIsSelectable | Qt.ItemIsEnabled )
            self.tableWidget.show()
            self.tableWidget_2.show()
            self.confirm_invest.show()
            self.tableWidget.resizeColumnsToContents()
            self.tableWidget_2.resizeColumnsToContents()
            self.ActiviteValide = True
            
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Deja Validé")
            msg.setInformativeText("Vous avez deja validé les Activités")
            msg.setWindowTitle("Activités déja validés")
            msg.exec_() 
    def createTable(self):
        self.createTableInvest()
        self.createTableFinanc()
                         
    def createTableInvest(self):
        self.tableWidget = QTableWidget(self.tab)
        self.tableWidget.setGeometry(QtCore.QRect(480, 220, 591, 310))
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.tableWidget.setRowCount(3)
        self.tableWidget.setColumnCount(5)
        #self.tableWidget.setItem(0,0, QTableWidgetItem("Cell (1,1)"))
        combo_box_options = ['Achat de terrain','Achat de biens immobiliers construits','Cout de construction de biens immobilier','Frais de licences et autorisations','Frais d’acquisition de machines et outils de travail','Frais d’acquisition d’ameublements']
        self.combo_invest = QtWidgets.QComboBox()
        button_invest = QtWidgets.QPushButton()
        button_invest.setText("+ Ajouter")
        button_invest.clicked.connect(self.clicked_add_invest)
        self.confirm_invest = QtWidgets.QPushButton(self.tab)
        self.confirm_invest.setText("Confirmer")
        self.confirm_invest.clicked.connect(self.confirmed_invest)
        self.confirm_invest.setGeometry(QtCore.QRect(490, 318, 101, 31))
        self.confirm_invest.hide()
        for t in combo_box_options:
            self.combo_invest.addItem(t)
        self.tableWidget.setItem(0,0, QTableWidgetItem("Quels investissements allez-vous effectuer ?"))
        self.tableWidget.item(0,0).setFont(QFont("Times New Roman",10,weight = 1))
        self.tableWidget.setCellWidget(2,0,self.combo_invest)
        self.tableWidget.setCellWidget(0,1, button_invest)
        self.tableWidget.setItem(1,0,QTableWidgetItem("Investissement"))
        self.tableWidget.setItem(1,1,QTableWidgetItem("Total"))
        self.tableWidget.move(0,0)
        self.tableWidget.hide()
        
    def createTableFinanc(self):
        self.tableWidget_2 = QTableWidget(self.tab_2)
        self.tableWidget_2.setGeometry(QtCore.QRect(480, 220, 591, 310))
        self.tableWidget_2.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.tableWidget_2.setRowCount(5)
        self.tableWidget_2.setColumnCount(5)
        button_financ = QtWidgets.QPushButton()
        button_financ.setText("Répartir les financements")
        button_financ.clicked.connect(self.clicked_spread_financ)
        self.tableWidget_2.setCellWidget(0,1, button_financ)
        self.tableWidget_2.setItem(0,0, QTableWidgetItem("Quelles sont vos ressources de financement ?"))
        self.tableWidget_2.item(0,0).setFont(QFont("Times New Roman",10,weight = 1))
        self.tableWidget_2.setItem(1,0,QTableWidgetItem("Souces de financements"))
        self.tableWidget_2.setItem(1,1,QTableWidgetItem("Total"))
        self.tableWidget_2.setItem(2,0,QTableWidgetItem("Fonds Propres"))
        self.tableWidget_2.setItem(3,0,QTableWidgetItem("CCA"))
        self.tableWidget_2.setItem(4,0,QTableWidgetItem("Dette"))
        self.tableWidget_2.move(0,0)
        self.tableWidget_2.hide()
            
    def clicked_add_invest(self):
        rowCount = self.tableWidget.rowCount()
        text = self.combo_invest.currentText()
        self.tableWidget.insertRow(rowCount)
        self.tableWidget.setCellWidget(rowCount,0,self.combo_invest)
        self.tableWidget.setItem(rowCount-1,0, QTableWidgetItem(text))
    def clicked_spread_financ(self):
        pass

    def confirmed_invest(self):
        pass
        
app = QApplication(sys.argv)
window = QDialog()
ui = Ui_Form()
ui.setupUi(window)
window.show()
sys.exit(app.exec_())
