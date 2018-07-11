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
from PyQt5.QtWidgets import *
from ClassesProjet import *
from FormulesCouts import *

def bind(objectName, propertyName, type):
    def getter(self):
        return type(self.findChild(QObject, objectName).property(propertyName).toPyObject())
        
    def setter(self, value):
        self.findChild(QObject, objectName).setProperty(propertyName, QVariant(value))
        
    return property(getter, setter)

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
                self.Projet = Projet(self.comboBox.currentText(),Horizon = int(self.lineEdit_3.text()))
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
        Form.resize(1000, 800)
        self.Form = Form
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
        self.groupBox.setGeometry(QtCore.QRect(20, 200, 461, 491))
        self.groupBox.setObjectName("groupBox")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.groupBox)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(30, 40, 411, 201))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.listWidget = QtWidgets.QTreeWidget(self.verticalLayoutWidget)
        self.listWidget.setObjectName("listWidget")
        self.verticalLayout.addWidget(self.listWidget)
        self.groupBox_2 = QtWidgets.QGroupBox(Form)
        self.groupBox_2.setGeometry(QtCore.QRect(520, 200, 451, 280))
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
        self.buttonDetails = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.buttonDetails.setObjectName("buttonDetails")
        self.buttonDetails.setText("Détails")
        self.horizontalLayout_3.addWidget(self.buttonDetails)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.dateEdit = QtWidgets.QDateEdit(Form)
        self.dateEdit.setGeometry(QtCore.QRect(365, 87, 110, 22))
        self.dateEdit.setObjectName("dateEdit")
        self.detailsSectionP = QtWidgets.QGroupBox(self.groupBox)
        self.detailsSectionP.setGeometry(QtCore.QRect(20,250,431,230))
        self.detailsSectionP.setObjectName("detailsSectionP")   
        self.detailsSectionP.setTitle("Informations détaillés sur la facturation")
        self.scrollarea = QtWidgets.QScrollArea(self.detailsSectionP)
        self.scrollarea.setFixedWidth(431)
        self.scrollarea.setFixedHeight(230)
        self.scrollarea.setWidgetResizable(True)
        widget = QWidget()
        self.scrollarea.setWidget(widget)
        self.layout_SArea = QtWidgets.QVBoxLayout(widget)
        # self.test_Label = QtWidgets.QLabel()
        # self.test_Label.setText("Question 1")
        # self.layout_SArea.addWidget(self.test_Label)
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)
        self.buttonNext = QPushButton(Form)
        self.buttonNext.setGeometry(QtCore.QRect(360, 720, 121, 41))
        self.buttonNext.setText("Suivant")
        self.buttonNext.clicked.connect(go_to_page4)
        self.initUI()
        
    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "II - L\'exploitation"))
        self.label_2.setText(_translate("Form", "Quelle est la date de démarrage de votre exploitation ?"))
        self.label_3.setText(_translate("Form", "Business Model"))
        self.groupBox.setTitle(_translate("Form", "Qu\'est ce que vous vendez ?"))
        self.groupBox_2.setTitle(_translate("Form", "Qu\'est ce que vous achetez ?"))

    def initUI(self):
        self.buttonDetails.clicked.connect(self.details_clicked)
    def load_combo_pc(self):
        ListeProduits = []
        ListeCharges = []
        revenu_widget = []
        for activite in ui.Projet.ListeActivites:
            activite.IdentifierCouts()
            activite.IdentifierRevenus()
            WidgetActivity_C = QtWidgets.QTreeWidgetItem(self.listWidget_2)
            WidgetActivity_R = QtWidgets.QTreeWidgetItem(self.listWidget)
            WidgetActivity_C.setText(0,activite.getNom())
            WidgetActivity_R.setText(0,activite.getNom())
            for cout in activite.getlistCout():
                costWidget = QtWidgets.QTreeWidgetItem(WidgetActivity_C)
                costWidget.setText(0,cout.getNom())
                costWidget.setFlags(costWidget.flags() | Qt.ItemIsUserCheckable)
                costWidget.setCheckState(0, Qt.Unchecked)
            for revenu in activite.getlistRev():
                revenuWidget = QtWidgets.QTreeWidgetItem(WidgetActivity_R)
                revenuWidget.setText(0,revenu.getNom())
                revenuWidget.setFlags(revenuWidget.flags() | Qt.ItemIsUserCheckable)
                revenuWidget.setCheckState(0, Qt.Unchecked)
        
        for i in ListeCharges:
            self.comboBox_c.addItem(i)
    
    
    def details_clicked(self):
        RevenuWidget = self.listWidget.selectedItems()[0] 
        ##### Revenu associé au widget Revenu ####"
        for activite in ui.Projet.ListeActivites:
            for revenu in activite.getlistRev():
                if revenu.getNom() == RevenuWidget.text(0) and RevenuWidget.checkState(0) == 2:
                    if activite.getNom() == RevenuWidget.parent().text(0):
                        Revenu = revenu
                        self.detailed_produit(Revenu)
    def detailed_produit(self,Revenu):
        print("Working Button")
        
        ####### Partie 1 Facturation   ######
        List_TypeF = ['100% à la Vente','Abbonement','Echeances Personnalisés']
        List_Periodes = [('Jours',1),('Semaines',7),('mois',30),('Trimestre',90),('Semestre',180),('Année',360)]
        groupBox = QtWidgets.QGroupBox()
        groupBox.setTitle("Type et pas de facturation")
        vertic_gb = QtWidgets.QVBoxLayout(groupBox)
        Label1 = QtWidgets.QLabel()
        Label1.setText("Type de Facturation")
        combo_Facturation = QtWidgets.QComboBox()
        button = QtWidgets.QPushButton()
        button.setText("Valider")
        
        for i in List_TypeF:
            combo_Facturation.addItem(i)
        
        HorizLayout = QtWidgets.QHBoxLayout()
        HorizLayout.addWidget(Label1)
        HorizLayout.addWidget(combo_Facturation)
        vertic_gb.addLayout(HorizLayout)
        vertic_gb.addStretch(1)
        Label3 = QtWidgets.QLabel()
        Label3.setText("Pas de Facturation       ")
        combo_periodes_2 = QtWidgets.QComboBox()
        for i in List_Periodes:
            combo_periodes_2.addItem(i[0])
        HorizLayout = QtWidgets.QHBoxLayout()
        HorizLayout.addWidget(Label3)
        HorizLayout.addWidget(combo_periodes_2)
        vertic_gb.addLayout(HorizLayout)
        vertic_gb.addStretch(1)
        vertic_gb.addWidget(button)
        self.layout_SArea.addWidget(groupBox)
        self.layout_SArea.addStretch(1)
        button.clicked.connect(lambda: self.valider_facturation(combo_Facturation.currentText(),combo_periodes_2.currentText(),vertic_gb,Revenu))
        # Label2 = QtWidgets.QLabel()
        # Label2.setText("Date de Facturation")
        # Duree = QtWidgets.QLineEdit()
        # combo_periodes = QtWidgets.QComboBox()
        # for i in List_Periodes:
        #     combo_periodes.addItem(i[0])
        # HorizLayout = QtWidgets.QHBoxLayout()
        # HorizLayout.addWidget(Label2)
        # HorizLayout.addWidget(Duree)
        # HorizLayout.addWidget(combo_periodes)
        # self.layout_SArea.addLayout(HorizLayout)
        
    def valider_facturation(self,facturation,pas,vertic_gb,Revenu):
        print(facturation)
        groupBox = QGroupBox()
        groupBox.setTitle("Informations sur les produits")
        groupVW = QVBoxLayout(groupBox)
        if facturation == '100% à la Vente':
            Revenu.isOnce = True
            Revenu.step = pas
            HorizentalBox =  QHBoxLayout()
            label = QLabel()
            label.setText("Prix de l'unité")
            Prix = QLineEdit()
            HorizentalBox.addWidget(label)
            HorizentalBox.addWidget(Prix)
            button = QPushButton()
            button.setText("Valider")
            groupVW.addLayout(HorizentalBox)
            groupVW.addStretch(1)
            groupVW.addWidget(button)
            groupVW.addStretch(1)
            button.clicked.connect(lambda: self.assign_price(Prix,Revenu))
            vertic_gb.addWidget(groupBox)
        elif facturation == 'Abbonement':
            Revenu.step = pas
            List_Periodes_2 = [('Jours',1),('Semaines',7),('mois',30),('Trimestre',90),('Semestre',180),('Année',360)]
            List_Periodes = [('Quotidien',1),('Hebdomadaire',7),('Mensuel',30),('Trimestriel',90),('Semestriel',180),('Annuel',360)]
            HorizentalBox = QHBoxLayout()
            label = QLabel()
            label.setText("Type d'abbonnement")
            combo = QComboBox()
            for i in List_Periodes:
                combo.addItem(i[0])
            HorizentalBox.addWidget(label)
            HorizentalBox.addWidget(combo)
            groupVW.addLayout(HorizentalBox)
            groupVW.addStretch(1)
            HorizentalBox = QHBoxLayout()
            label_2 = QLabel()
            label_2.setText("Durée de l'Abbonement")
            lineEdit = QLineEdit()
            combo_2 = QComboBox()
            for i in List_Periodes_2:
                combo_2.addItem(i[0])
            HorizentalBox.addWidget(label_2)
            HorizentalBox.addWidget(lineEdit)
            HorizentalBox.addWidget(combo_2)
            groupVW.addLayout(HorizentalBox)
            groupVW.addStretch(1)
            HorizentalBox = QHBoxLayout()
            label_3 = QLabel()
            label_3.setText("Prix de l'Echéance")
            Prix= QLineEdit()
            HorizentalBox.addWidget(label_3)
            HorizentalBox.addWidget(Prix)
            groupVW.addLayout(HorizentalBox)
            groupVW.addStretch(1)
            button = QPushButton()
            button.setText("Valider")
            groupVW.addWidget(button)
            groupVW.addStretch(1)
            vertic_gb.addWidget(groupBox)
            button.clicked.connect(lambda: self.confirmed_Abbonement(lineEdit.text(),combo.currentText(),combo_2.currentText(),Prix.text(),Revenu))
        elif facturation == 'Echeances Personnalisés':
            HorizentalBox =  QHBoxLayout()
            label = QLabel()
            label.setText("Nombre D'Echeances  :")
            NEcheances = QLineEdit()
            HorizentalBox.addWidget(label)
            HorizentalBox.addWidget(NEcheances)
            groupVW.addLayout(HorizentalBox)
            groupVW.addStretch(1)
            HorizentalBox =  QHBoxLayout()
            label_2 = QLabel()
            label_2.setText("Prix Total")
            Prix = QLineEdit()
            HorizentalBox.addWidget(label_2)
            HorizentalBox.addWidget(Prix)
            groupVW.addLayout(HorizentalBox)
            groupVW.addStretch(1)
            HorizentalBox =  QHBoxLayout()
            label_3 = QLabel()
            label_3.setText("Duree")
            Duree = QLineEdit()
            HorizentalBox.addWidget(label_3)
            HorizentalBox.addWidget(Duree)
            groupVW.addLayout(HorizentalBox)
            groupVW.addStretch(1)
            button = QPushButton()
            button.setText("Valider")
            groupVW.addWidget(button)
            vertic_gb.addWidget(groupBox)
            button.clicked.connect(lambda: self.confirmed_echeances(NEcheances.text(),Prix.text(),Duree.text(),vertic_gb))
        self.layout_SArea.addStretch(1)
      
    def assign_price(self,Prix,Revenu):
        Revenu.Prix = int(Prix.text())
    
    def confirmed_Abbonement(self,Duree,TypeAbbonement,UniteDuree,Prix,Revenu):
        Revenu.Prix = Prix
        Revenu.VEchantillon = [0 for i in range(int(Duree)*int(getPas(UniteDuree)/getPas(Revenu.step)))]
        for i in range(0,len(Revenu.VEchantillon),int(getPas(TypeAbbonement)/getPas(Revenu.step))):
            Revenu.VEchantillon[i] = 1
        print(Duree)
        print(TypeAbbonement)
        print(UniteDuree)
        
    def confirmed_echeances(self,NEcheances,Prix,Duree,vertic_gb):
        groupBox = QGroupBox()
        groupBox.setTitle("Details de l'Echéancier")
        NEcheances = int(NEcheances)
        Prix = int(Prix)
        Duree = int(Duree)
        VerticalLayout = QVBoxLayout(groupBox)
        table = QTableWidget()
        table.setColumnCount(3)
        table.setRowCount(NEcheances)
        table.setHorizontalHeaderLabels(['Echéance','Période','Pourcentage'])
        table.verticalHeader().setDefaultSectionSize(30)
        table.horizontalHeader().setDefaultSectionSize(100)
        table.setMinimumHeight(30*NEcheances)
        for i in range(NEcheances):
            table.setItem(i,0,QTableWidgetItem(str(i+1)))
        VerticalLayout.addWidget(table)
        VerticalLayout.addStretch(1)
        vertic_gb.addWidget(groupBox)
        vertic_gb.addStretch(1)

class Ui_Form_4(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(878, 525)
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setGeometry(QtCore.QRect(20, 50, 821, 241))
        self.groupBox.setTitle("Informations sur les couts")
        self.scrollArea = QtWidgets.QScrollArea(self.groupBox)
        self.scrollArea.setGeometry(QtCore.QRect(0, 0, 821, 241))
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")
        self.widget = QWidget()
        self.scrollArea.setWidget(self.widget)
        self.verticalLayout = QtWidgets.QVBoxLayout(self.widget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(20, 20, 241, 16))
        font = QtGui.QFont()
        font.setPointSize(12)
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
        self.label.setText(_translate("Form", "Informations sur les revenus"))
    
    def initUI(self):
        for activite in ui.Projet.ListeActivites:
            label = QtWidgets.QLabel()
            label.setText(activite.getNom())
            self.verticalLayout.addWidget(label)
            self.verticalLayout.addStretch(1)
            for revenu in activite.getlistRev():
                listeQuestions = revenu.SaisieIntrinseque_1()
                for question in listeQuestions:
                    titreTableau = QtWidgets.QLabel()
                    titreTableau.setText("          "+str(question[0]))
                    self.verticalLayout.addWidget(titreTableau)
                    for i in range(1,len(question)):
                        HorizentalLayout = QHBoxLayout()
                        label = QLabel()
                        label.setText(question[i][1])
                        lineEdit = QLineEdit()
                        lineEdit.setObjectName(question[i][0]+str(i))
                        #revenu.DicoFormes[titreTableau][question[i][0]] = bind(question[i][0]+str(i),"text",str)
                        HorizentalLayout.addWidget(label)
                        HorizentalLayout.addWidget(lineEdit)
                        self.verticalLayout.addLayout(HorizentalLayout)
                        self.verticalLayout.addStretch(1)
                    
                listeQuestions = revenu.SaisieMarche_1()
                for question in listeQuestions:
                    titreTableau = QtWidgets.QLabel()
                    titreTableau.setText("          "+str(question[0]))
                    self.verticalLayout.addWidget(titreTableau)
                    for i in range(1,len(question)):
                        HorizentalLayout = QHBoxLayout()
                        label = QLabel()
                        label.setText(question[i][1])
                        lineEdit = QLineEdit()
                        print(revenu.DicoFormesMarche[question[0]][question[i][0]])
                        lineEdit.textChanged.connect(lambda arg1 = lineEdit.text(),arg2=question[0],arg3=question[i][0]: self.assign_field_to_value(arg1,arg2,arg3))
                        #,revenu.DicoFormesMarche[titreTableau][question[i][0]]))
                        lineEdit.setObjectName(question[i][0]+str(i))
                        #revenu.DicoFormes[titreTableau][question[i][0]] = bind(question[i][0]+str(i),"text",str)
                        HorizentalLayout.addWidget(label)
                        HorizentalLayout.addWidget(lineEdit)
                        self.verticalLayout.addLayout(HorizentalLayout)
                        self.verticalLayout.addStretch(1)    
                    
                # for i in revenu.SaisieMarche_1():
                #     titreTableau = QtWidgets.QLabel()
                #     titreTableau.setText("          "+str(i[0]))
                #     self.verticalLayout.addWidget(titreTableau)
                #     HorizentalLayout = QHBoxLayout()
                #     label = QLabel()
                #     label.setText(str(i[1][0]))
                #     lineEdit = QLineEdit()
                #     HorizentalLayout.addWidget(label)
                #     HorizentalLayout.addWidget(lineEdit)
                #     self.verticalLayout.addLayout(HorizentalLayout)
                #     self.verticalLayout.addStretch(1)
                #     HorizentalLayout = QHBoxLayout()
                #     label = QLabel()
                #     label.setText(i[1][1])
                #     lineEdit = QLineEdit()
                #     HorizentalLayout.addWidget(label)
                #     HorizentalLayout.addWidget(lineEdit)
                #     self.verticalLayout.addLayout(HorizentalLayout)
                #     self.verticalLayout.addStretch(1)
            button = QPushButton()
            self.verticalLayout.addWidget(button)
            button.clicked.connect(self.clicked_params_rev)
        
    def clicked_params_rev(self):
        pass
    def assign_field_to_value(self,line,cle1,cle2):
        print(line)
        print(cle1)
        print(cle2)
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
        
    
def getPas(pas):
    List_Periodes_2 = [('Jours',1),('Semaines',7),('mois',30),('Trimestre',90),('Semestre',180),('Année',360)]
    List_Periodes = [('Quotidien',1),('Hebdomadaire',7),('Mensuel',30),('Trimestriel',90),('Semestriel',180),('Annuel',360)]
    list = List_Periodes + List_Periodes_2
    for i in list:
        if i[0] == pas:
            return i[1]

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

def go_to_page4():
    window3.hide()
    global ui4
    ui4 = Ui_Form_4()
    ui4.setupUi(window4)
    window4.show()

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
window4 = QDialog()
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
