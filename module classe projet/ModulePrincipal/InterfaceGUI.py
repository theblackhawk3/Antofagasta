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

##########################1ère Page ###################
class Ui_Form(QWidget):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1000, 700)
        self.connected = False
        self.connected_2 = False
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
        self.ListeProjets = ['----------','Ecole Secondaire','Clinique','Restaurant','Projet Immobilier']
        self.pushButton.clicked.connect(self.on_confirm_click)
        self.comboBox.addItems(self.ListeProjets)
        self.progressBar.setProperty("value", 0)
        self.buttonConfirmActivity.clicked.connect(self.on_confirm_click2)
        self.ProjetValide = False
        self.ActiviteValide = False
        self.pushButton_2.clicked.connect(go_to_page2)
        self.createTable()
        self.progressBar.hide()
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
                self.lineEdit,self.lineEdit_3 = self.lineEdit_3,self.lineEdit
                print(self.lineEdit_3.text()) #Horizon
                print(self.lineEdit.text()) #Durée des Investissement
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
                for j in range(1,int(self.lineEdit.text())+2):
                    self.tableWidget_2.setItem(i,j,QTableWidgetItem("0"))
                    #self.tableWidget_2.item(i,j).setFlags( Qt.ItemIsSelectable | Qt.ItemIsEnabled )
            if self.connected_2 == False:
                self.tableWidget_2.cellChanged.connect(self.changed_cell_2)
                self.connected_2 = True        
            self.tableWidget.show()
            self.tableWidget_2.show()
            #self.confirm_invest.show()
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
        self.tableWidget.setGeometry(QtCore.QRect(390, 220, 595, 376))
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
        self.tableWidget.setItem(1,0,QTableWidgetItem("Investissements"))
        self.tableWidget.setItem(1,1,QTableWidgetItem("Total"))
        self.tableWidget.move(0,0)
        self.tableWidget.hide()
        
    def createTableFinanc(self):
        self.tableWidget_2 = QTableWidget(self.tab_2)
        self.tableWidget_2.setGeometry(QtCore.QRect(390, 220, 595, 376))
        self.tableWidget_2.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.tableWidget_2.setRowCount(5)
        self.tableWidget_2.setColumnCount(5)
        button_financ = QtWidgets.QPushButton()
        button_financ.setText("Répartir les financements")
        button_financ.clicked.connect(self.clicked_spread_financ)
        #self.tableWidget_2.setCellWidget(0,1, button_financ)
        self.tableWidget_2.setItem(0,0, QTableWidgetItem("Quelles sont vos ressources de financement ?"))
        self.tableWidget_2.item(0,0).setFont(QFont("Times New Roman",10,weight = 1))
        self.tableWidget_2.setItem(1,0,QTableWidgetItem("Sources de financement"))
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
        self.tableWidget.setItem(rowCount-1,1, QTableWidgetItem("0"))
        for i in range(2,int(self.lineEdit.text())+2):
            self.tableWidget.setItem(rowCount-1,i, QTableWidgetItem("0"))
        if self.connected == False:
            self.tableWidget.cellChanged.connect(self.changed_cell)
            self.connected = True
    def clicked_spread_financ(self):
        FP = float(self.tableWidget_2.item(2,1).text())
        FP_Reste = 0
        CCA = float(self.tableWidget_2.item(3,1).text())
        CCA_Reste
        Dette = float(self.tableWidget_2.item(4,1).text())
        Dette_Reste = 0
        Liste_Resources = [FP,CCA,Dette]
        Sum_Invest = 0
        for i in range(2,self.tableWidget.rowCount()-1):
            for j in range(2,int(self.lineEdit.text())+2):
               Sum_Invest += float(self.tableWidget.item(i,j).text()) 
               
        if sum(Liste_Resources) >= Sum_Invest:
            for i in range(2,self.tableWidget.rowCount()-1):
                cout = 0
                for j in range(2,int(self.lineEdit.text())+2):
                    cout += self.tableWidget.item(i,j).text()
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Ressources Insuffisantes")
            msg.setInformativeText("Le montant de vos ressources n'est pas assez suffisant pour couvrir vos invesstissements, Veuillez modifier vos montants")
            msg.setWindowTitle("Pas suffisament de ressources")
            msg.exec_()
    
    def extractTable(self):
        return self.tableWidget
        
        #Sommer tout les Dettes
        

    def confirmed_invest(self):
        for i in range(2,self.tableWidget.rowCount()-1):
            for j in range(2,int(self.lineEdit.text())+2):
                #self.tableWidget.setItem(i,j,QTableWidgetItem(str(float(ui.tableWidget.item(i,j).text())*float(ui.tableWidget.item(i,1).text()))))
                self.tableWidget.setItem(i,j,QTableWidgetItem("0"))
                
    def changed_cell(self,row,column):
        if column == 1 or column == 0 or self.tableWidget.item(row,column).text() == "0":
            pass
        else:
            self.tableWidget.setItem(row,1,QTableWidgetItem(str(Sum_cells(self.tableWidget,row,2,row,int(self.lineEdit.text())+1))))
            
    def changed_cell_2(self,row,column):
        if column == 1 or column == 0 or self.tableWidget_2.item(row,column).text() == "0":
            pass
        else:
            self.tableWidget_2.setItem(row,1,QTableWidgetItem(str(Sum_cells(self.tableWidget_2,row,2,row,int(self.lineEdit.text())+1))))

##########################2ème Page##################                
class Ui_Form_2(QWidget):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(700, 700)
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(30, 620, 121, 41))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(550, 620, 121, 41))
        self.pushButton_2.setObjectName("pushButton_2")
        self.tableWidget = QtWidgets.QTableWidget(Form)
        self.tableWidget.setGeometry(QtCore.QRect(15, 91, 673, 511))
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setObjectName("tableWidget")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(210, 60, 291, 21))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)
        self.initUI()

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.pushButton.setText(_translate("Form", "< Retour"))
        self.pushButton_2.setText(_translate("Form", "Valider >"))
        self.label.setText(_translate("Form", "Récapitulatif phase Investissements"))
        self.pushButton.clicked.connect(go_to_page1)
        
    def initUI(self):
        self.pushButton_2.clicked.connect(go_to_page3)


class Ui_Form_3(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1000, 700)
        font = QtGui.QFont()
        font.setPointSize(10)
        Form.setFont(font)
        self.line = QtWidgets.QFrame(Form)
        self.line.setGeometry(QtCore.QRect(20, 60, 951, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(20, 40, 161, 16))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(20, 90, 331, 16))
        self.label_2.setObjectName("label_2")
        self.line_2 = QtWidgets.QFrame(Form)
        self.line_2.setGeometry(QtCore.QRect(20, 150, 951, 16))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(20, 130, 161, 16))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setGeometry(QtCore.QRect(20, 200, 461, 381))
        self.groupBox.setObjectName("groupBox")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.groupBox)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(30, 40, 411, 201))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.listWidget = QtWidgets.QListWidget(self.verticalLayoutWidget)
        self.listWidget.setObjectName("listWidget")
        self.verticalLayout.addWidget(self.listWidget)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.pushButton_ap = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.pushButton_ap.setObjectName("pushButton_ap")
        self.horizontalLayout_2.addWidget(self.pushButton_ap)
        self.pushButton_sp = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.pushButton_sp.setObjectName("pushButton_sp")
        self.horizontalLayout_2.addWidget(self.pushButton_sp)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.comboBox_p = QtWidgets.QComboBox(self.verticalLayoutWidget)
        self.comboBox_p.setObjectName("comboBox_p")
        self.verticalLayout.addWidget(self.comboBox_p)
        self.label_4 = QtWidgets.QLabel(self.groupBox)
        self.label_4.setGeometry(QtCore.QRect(30, 20, 111, 16))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.groupBox)
        self.label_5.setGeometry(QtCore.QRect(30, 250, 211, 16))
        self.label_5.setObjectName("label_5")
        self.comboBox_tfp = QtWidgets.QComboBox(self.groupBox)
        self.comboBox_tfp.setGeometry(QtCore.QRect(30, 270, 411, 22))
        self.comboBox_tfp.setObjectName("comboBox_tfp")
        self.groupBox_2 = QtWidgets.QGroupBox(Form)
        self.groupBox_2.setGeometry(QtCore.QRect(520, 200, 451, 381))
        self.groupBox_2.setObjectName("groupBox_2")
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(self.groupBox_2)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(20, 40, 411, 201))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.listWidget_2 = QtWidgets.QListWidget(self.verticalLayoutWidget_2)
        self.listWidget_2.setObjectName("listWidget_2")
        self.verticalLayout_2.addWidget(self.listWidget_2)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.pushButton_ac = QtWidgets.QPushButton(self.verticalLayoutWidget_2)
        self.pushButton_ac.setObjectName("pushButton_ac")
        self.horizontalLayout_3.addWidget(self.pushButton_ac)
        self.pushButton_sc = QtWidgets.QPushButton(self.verticalLayoutWidget_2)
        self.pushButton_sc.setObjectName("pushButton_sc")
        self.horizontalLayout_3.addWidget(self.pushButton_sc)
        self.verticalLayout_2.addLayout(self.horizontalLayout_3)
        self.comboBox_c = QtWidgets.QComboBox(self.verticalLayoutWidget_2)
        self.comboBox_c.setObjectName("comboBox_c")
        self.verticalLayout_2.addWidget(self.comboBox_c)
        self.comboBox_tfc = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox_tfc.setGeometry(QtCore.QRect(20, 270, 411, 22))
        self.comboBox_tfc.setObjectName("comboBox_tfc")
        self.label_6 = QtWidgets.QLabel(self.groupBox_2)
        self.label_6.setGeometry(QtCore.QRect(20, 20, 111, 16))
        self.label_6.setObjectName("label_6")
        self.label_7 = QtWidgets.QLabel(self.groupBox_2)
        self.label_7.setGeometry(QtCore.QRect(20, 250, 211, 16))
        self.label_7.setObjectName("label_7")
        self.dateEdit = QtWidgets.QDateEdit(Form)
        self.dateEdit.setGeometry(QtCore.QRect(365, 87, 110, 22))
        self.dateEdit.setObjectName("dateEdit")
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)
        self.initUI()
        
    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "II - L\'exploitation"))
        self.label_2.setText(_translate("Form", "Quelle est la date de démarrage de votre exploitation ?"))
        self.label_3.setText(_translate("Form", "Business Model"))
        self.groupBox.setTitle(_translate("Form", "Qu\'est ce que vous vendez ?"))
        self.pushButton_ap.setText(_translate("Form", "Ajouter Produit"))
        self.pushButton_sp.setText(_translate("Form", "Supprimer"))
        self.label_4.setText(_translate("Form", "Liste des Produits"))
        self.label_5.setText(_translate("Form", "Type de facturation"))
        self.groupBox_2.setTitle(_translate("Form", "Qu\'est ce que vous achetez ?"))
        self.pushButton_ac.setText(_translate("Form", "Ajouter Charge"))
        self.pushButton_sc.setText(_translate("Form", "Supprimer"))
        self.label_6.setText(_translate("Form", "Liste des Charges"))
        self.label_7.setText(_translate("Form", "Type de facturation"))

    def initUI(self):
        #Chargement de la liste selon les méthodes IdentifierCouts / Identifier Revenus
        TypeFacturationProduits = ['100% à la Vente','Réparti sur périodes']
        TypeFacturationCouts = ['100% à l\'achat','Réparti sur périodes']
        for i in ListeProduits:
            self.comboBox_p.addItem(i)
        for i in ListeCharges:
            self.comboBox_c.addItem(i)
        for i in TypeFacturationProduits:
            self.comboBox_tfp.addItem(i)
        for i in TypeFacturationCouts:
            self.comboBox_tfc.addItem(i)
        self.comboBox_tfp.activated.connect(self.combo_tfp_activated)
        self.comboBox_tfc.activated.connect(self.combo_tfc_activated)
        self.pushButton_ac.clicked.connect(self.clicked_ac)
        self.pushButton_ap.clicked.connect(self.clicked_ap)
        self.pushButton_sc.clicked.connect(self.clicked_sc)
        self.pushButton_sp.clicked.connect(self.clicked_sr)
    
    def clicked_ac(self):
        self.listWidget_2.addItem(self.comboBox_c.currentText())
    def clicked_ap(self):
        self.listWidget.addItem(self.comboBox_p.currentText())
    def clicked_sc(self):
        for i in self.listWidget_2.selectedItems():
            self.listWidget_2.takeItem(self.listWidget_2.row(i))
    def clicked_sr(self):
        for i in self.listWidget.selectedItems():
            self.listWidget.takeItem(self.listWidget.row(i))
    def combo_tfp_activated(self,text):
        print(text)
    def combo_tfc_activated(self,text):
        pass    
    #     self.label_1 = QtWidgets.QLabel(Form)
    #     self.label_1.setText("Nombre d'échéances")
    #     self.NEcheances = QtWidget.QLineEdit(Form)
    #     self.label_2 = QtWidgets.QLabel(Form)class Ui_Form_3(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1000, 700)
        font = QtGui.QFont()
        font.setPointSize(10)
        Form.setFont(font)
        self.line = QtWidgets.QFrame(Form)
        self.line.setGeometry(QtCore.QRect(20, 60, 951, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(20, 40, 161, 16))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(20, 90, 331, 16))
        self.label_2.setObjectName("label_2")
        self.line_2 = QtWidgets.QFrame(Form)
        self.line_2.setGeometry(QtCore.QRect(20, 150, 951, 16))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(20, 130, 161, 16))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setGeometry(QtCore.QRect(20, 200, 461, 381))
        self.groupBox.setObjectName("groupBox")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.groupBox)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(30, 40, 411, 201))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.listWidget = QtWidgets.QListWidget(self.verticalLayoutWidget)
        self.listWidget.setObjectName("listWidget")
        self.verticalLayout.addWidget(self.listWidget)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.pushButton_ap = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.pushButton_ap.setObjectName("pushButton_ap")
        self.horizontalLayout_2.addWidget(self.pushButton_ap)
        self.pushButton_sp = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.pushButton_sp.setObjectName("pushButton_sp")
        self.horizontalLayout_2.addWidget(self.pushButton_sp)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.comboBox_p = QtWidgets.QComboBox(self.verticalLayoutWidget)
        self.comboBox_p.setObjectName("comboBox_p")
        self.verticalLayout.addWidget(self.comboBox_p)
        self.label_4 = QtWidgets.QLabel(self.groupBox)
        self.label_4.setGeometry(QtCore.QRect(30, 20, 111, 16))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.groupBox)
        self.label_5.setGeometry(QtCore.QRect(30, 250, 211, 16))
        self.label_5.setObjectName("label_5")
        self.comboBox_tfp = QtWidgets.QComboBox(self.groupBox)
        self.comboBox_tfp.setGeometry(QtCore.QRect(30, 270, 411, 22))
        self.comboBox_tfp.setObjectName("comboBox_tfp")
        self.groupBox_2 = QtWidgets.QGroupBox(Form)
        self.groupBox_2.setGeometry(QtCore.QRect(520, 200, 451, 381))
        self.groupBox_2.setObjectName("groupBox_2")
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(self.groupBox_2)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(20, 40, 411, 201))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.listWidget_2 = QtWidgets.QTreeWidget(self.verticalLayoutWidget_2)
        self.listWidget_2.setObjectName("listWidget_2")
        self.verticalLayout_2.addWidget(self.listWidget_2)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.pushButton_ac = QtWidgets.QPushButton(self.verticalLayoutWidget_2)
        self.pushButton_ac.setObjectName("pushButton_ac")
        self.horizontalLayout_3.addWidget(self.pushButton_ac)
        self.pushButton_sc = QtWidgets.QPushButton(self.verticalLayoutWidget_2)
        self.pushButton_sc.setObjectName("pushButton_sc")
        self.horizontalLayout_3.addWidget(self.pushButton_sc)
        self.verticalLayout_2.addLayout(self.horizontalLayout_3)
        self.comboBox_c = QtWidgets.QComboBox(self.verticalLayoutWidget_2)
        self.comboBox_c.setObjectName("comboBox_c")
        self.verticalLayout_2.addWidget(self.comboBox_c)
        self.comboBox_tfc = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox_tfc.setGeometry(QtCore.QRect(20, 270, 411, 22))
        self.comboBox_tfc.setObjectName("comboBox_tfc")
        self.label_6 = QtWidgets.QLabel(self.groupBox_2)
        self.label_6.setGeometry(QtCore.QRect(20, 20, 111, 16))
        self.label_6.setObjectName("label_6")
        self.label_7 = QtWidgets.QLabel(self.groupBox_2)
        self.label_7.setGeometry(QtCore.QRect(20, 250, 211, 16))
        self.label_7.setObjectName("label_7")
        self.dateEdit = QtWidgets.QDateEdit(Form)
        self.dateEdit.setGeometry(QtCore.QRect(365, 87, 110, 22))
        self.dateEdit.setObjectName("dateEdit")
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)
        self.initUI()
        
    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "II - L\'exploitation"))
        self.label_2.setText(_translate("Form", "Quelle est la date de démarrage de votre exploitation ?"))
        self.label_3.setText(_translate("Form", "Business Model"))
        self.groupBox.setTitle(_translate("Form", "Qu\'est ce que vous vendez ?"))
        self.pushButton_ap.setText(_translate("Form", "Ajouter Produit"))
        self.pushButton_sp.setText(_translate("Form", "Supprimer"))
        self.label_4.setText(_translate("Form", "Liste des Produits"))
        self.label_5.setText(_translate("Form", "Type de facturation"))
        self.groupBox_2.setTitle(_translate("Form", "Qu\'est ce que vous achetez ?"))
        self.pushButton_ac.setText(_translate("Form", "Ajouter Charge"))
        self.pushButton_sc.setText(_translate("Form", "Supprimer"))
        self.label_6.setText(_translate("Form", "Liste des Charges"))
        self.label_7.setText(_translate("Form", "Type de facturation"))

    def initUI(self):
        #Chargement de la liste selon les méthodes IdentifierCouts / Identifier Revenus
        ListeProduits = ['Produits Industrialisés','Paiement mensuel des prestation de cours','Frais annuels de prestation de cours','Vente de Biens Immobiliers','Commercialisation de plats','Vente de soins médicaux']
        TypeFacturationProduits = ['100% à la Vente','Réparti sur périodes']
        TypeFacturationCouts = ['100% à l\'achat','Réparti sur périodes']
        for i in ListeProduits:
            self.comboBox_p.addItem(i)
        for i in TypeFacturationProduits:
            self.comboBox_tfp.addItem(i)
        for i in TypeFacturationCouts:
            self.comboBox_tfc.addItem(i)
        self.comboBox_tfp.activated.connect(self.combo_tfp_activated)
        self.comboBox_tfc.activated.connect(self.combo_tfc_activated)
        self.pushButton_ac.clicked.connect(self.clicked_ac)
        self.pushButton_ap.clicked.connect(self.clicked_ap)
        self.pushButton_sc.clicked.connect(self.clicked_sc)
        self.pushButton_sp.clicked.connect(self.clicked_sr)
    
    def clicked_ac(self):
        self.listWidget_2.addItem(self.comboBox_c.currentText())
    def clicked_ap(self):
        self.listWidget.addItem(self.comboBox_p.currentText())
    def clicked_sc(self):
        for i in self.listWidget_2.selectedItems():
            self.listWidget_2.takeItem(self.listWidget_2.row(i))
    def clicked_sr(self):
        for i in self.listWidget.selectedItems():
            self.listWidget.takeItem(self.listWidget.row(i))
    def combo_tfp_activated(self,text):
        print(text)
    def combo_tfc_activated(self,text):
        pass    
    #     self.label_1 = QtWidgets.QLabel(Form)
    #     self.label_1.setText("Nombre d'échéances")
    #     self.NEcheances = QtWidget.QLineEdit(Form)
    #     self.label_2 = QtWidgets.QLabel(Form)
    
    def load_combo_pc(self):
        ListeCharges = []
        for activite in ui.Projet.ListeActivites:
            ListeCharges.append("------ "+activite.getNom()+" ------")
            activite.IdentifierCouts()
            WidgetActivity = QtWidgets.QTreeWidgetItem(self.listWidget_2)
            WidgetActivity.setText(0,activite.getNom())
            for cout in activite.getlistCout():
                ListeCharges.append(cout.getNom())
                costWidget = QtWidgets.QTreeWidgetItem(WidgetActivity)
                costWidget.setText(0,cout.getNom())
                costWidget.setFlags(costWidget.flags() | Qt.ItemIsUserCheckable)
                costWidget.setCheckState(0, Qt.Unchecked)
        for i in ListeCharges:
            self.comboBox_c.addItem(i)

def SendTables():
    ui2.tableWidget.setColumnCount(ui.tableWidget.columnCount())
    ui2.tableWidget.setRowCount(ui.tableWidget.rowCount()+ui.tableWidget_2.rowCount()+4)
    for i in range(ui2.tableWidget.rowCount()):
        for j in range(ui2.tableWidget.columnCount()):
            ui2.tableWidget.setItem(i,j,QTableWidgetItem(""))
    
    for i in range(1,ui.tableWidget.rowCount()-1):
        ui2.tableWidget.setItem(i,0,QTableWidgetItem(ui.tableWidget.item(i,0).text()))
        #ui2.tableWidget.item(i,0).setBackground(QColor("grey"))
        #ui2.tableWidget.item(i,0).setForeground(QColor("white"))
        
    for i in range(1,ui.tableWidget.rowCount()-1):
        for j in range(2,int(ui.lineEdit.text())+2):
            ui2.tableWidget.setItem(i,j-1,QTableWidgetItem(ui.tableWidget.item(i,j).text()))

    
    ui2.tableWidget.setItem(ui.tableWidget.rowCount()-1,0,QTableWidgetItem("Total"))
    
    for i in range(1,int(ui.lineEdit.text())+1):
        ui2.tableWidget.setItem(ui.tableWidget.rowCount()-1,i,QTableWidgetItem(str(Sum_cells(ui2.tableWidget,2,i,ui.tableWidget.rowCount()-2,i))))
        
    #Refaire de même pour le tableau de financement
    
    #ui2.tableWidget.setItem(ui.tableWidget.rowCount(),0,QTableWidgetItem("Ressources"))
    for i in range(1,ui.tableWidget_2.rowCount()):
        ui2.tableWidget.setItem(i+ui.tableWidget.rowCount(),0,QTableWidgetItem(ui.tableWidget_2.item(i,0).text()))
    
    for i in range(1,ui.tableWidget_2.rowCount()):
        for j in range(2,int(ui.lineEdit.text())+2):
            ui2.tableWidget.setItem(i+ui.tableWidget.rowCount(),j-1,QTableWidgetItem(ui.tableWidget_2.item(i,j).text()))
    ui2.tableWidget.setItem(ui.tableWidget_2.rowCount()+ui.tableWidget.rowCount(),0,QTableWidgetItem("Total"))
    
    for i in range(1,int(ui.lineEdit.text())+1):
        ui2.tableWidget.setItem(ui.tableWidget_2.rowCount()+ui.tableWidget.rowCount(),i,QTableWidgetItem(str(Sum_cells(ui2.tableWidget,2+ui.tableWidget.rowCount(),i,ui.tableWidget.rowCount()+ui.tableWidget_2.rowCount() -1,i))))
        
    current_line = ui.tableWidget.rowCount()+ui.tableWidget_2.rowCount()

    ui2.tableWidget.setItem(current_line+2,0,QTableWidgetItem("Ecarts"))
    for i in range(1,int(ui.lineEdit.text())+1):
        ressource = int(ui2.tableWidget.item(ui.tableWidget_2.rowCount()+ui.tableWidget.rowCount(),i).text())
        emploi = int(ui2.tableWidget.item(ui.tableWidget.rowCount()-1,i).text())
        ui2.tableWidget.setItem(current_line+2,i,QTableWidgetItem(str(ressource - emploi)))
        
        
    # Mise en forme
    font = QFont()
    font.setBold(True)
    font.setItalic(True)
    for j in range(0,int(ui.lineEdit.text())+1):
    #Ligne emplois
        ui2.tableWidget.item(1,j).setBackground(QColor("grey"))
        ui2.tableWidget.item(1,j).setForeground(QColor("white"))
        ui2.tableWidget.item(1,j).setFont(font)
    #Ligne Total emplois 
        ui2.tableWidget.item(ui.tableWidget.rowCount()-1,j).setBackground(QColor("grey"))
        ui2.tableWidget.item(ui.tableWidget.rowCount()-1,j).setForeground(QColor("white"))
        ui2.tableWidget.item(ui.tableWidget.rowCount()-1,j).setFont(font)
    #Ligne Ressources
        ui2.tableWidget.item(1+ui.tableWidget.rowCount(),j).setBackground(QColor("grey"))
        ui2.tableWidget.item(1+ui.tableWidget.rowCount(),j).setForeground(QColor("white"))
        ui2.tableWidget.item(1+ui.tableWidget.rowCount(),j).setFont(font)
    #Ligne Total Ressources
        ui2.tableWidget.item(ui.tableWidget_2.rowCount()+ui.tableWidget.rowCount(),j).setBackground(QColor("grey"))
        ui2.tableWidget.item(ui.tableWidget_2.rowCount()+ui.tableWidget.rowCount(),j).setForeground(QColor("white"))
        ui2.tableWidget.item(ui.tableWidget_2.rowCount()+ui.tableWidget.rowCount(),j).setFont(font)
    #Ligne Ecarts
        ui2.tableWidget.item(current_line+2,j).setBackground(QColor("grey"))
        ui2.tableWidget.item(current_line+2,j).setForeground(QColor("white"))
        ui2.tableWidget.item(current_line+2,j).setFont(font)
        
def go_to_page2():
    SendTables()
    window.hide()
    ui2.tableWidget.resizeColumnsToContents()
    window2.show()

def go_to_page1():
    window2.hide()
    window.show()
    
def go_to_page3():
    ui3.load_combo_pc()
    window2.hide()
    window3.show()

def Sum_cells(table,firstRow,firstCol,secondRow,secondCol):
    sum = 0
    for r in range(firstRow, secondRow + 1):
        for c in range(firstCol, secondCol + 1):
            tableItem = table.item(r, c)
            try:
                sum += int(tableItem.text())
            except (ValueError,AttributeError) as e:
                pass
    return sum

#########################3ème Page####################       
app = QApplication(sys.argv)
window = QDialog()
window3 = QDialog()
ui = Ui_Form()
ui.setupUi(window)
window.show()
window2 = QDialog()
ui2 = Ui_Form_2()
ui2.setupUi(window2)
ui3 = Ui_Form_3()
ui3.setupUi(window3)
window3.hide()
window2.hide()
sys.exit(app.exec_())
