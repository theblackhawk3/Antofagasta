from openpyxl import *
from openpyxl.styles import Font, Color
import pandas as pd
import numpy as np
import os



def export_dfs(df_list, sheet, file_name, spaces,colstart=0):
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')
    row = 4
    for dataframe in df_list:
        dataframe.to_excel(writer,sheet_name=sheet,startrow=row , startcol=colstart)   
        row = row + len(dataframe.index) + spaces + 1
    writer.save()

def excel_dfs(Dict, file_name, spaces):
    fontsDict = {'Activite':Font(name='Calibri',size=16,bold = True),
                 'Cout':Font(name='Calibri',size=11,bold=True,italic=True),
                 'Titre':Font(name='Calibri',size=11,italic=True)}
    indexColAct = 0
    WordRowCol = {}
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')
    sheets = ['Int couts','Mar couts']
    #Début ecriture données
    for sheet in sheets:
        indexAct = 0
        WordRowCol[sheet] = []
        for activite in Dict[sheet].keys():
            indexAct+=(max([a[1] for a in A.getDictParams()[sheet][activite].values()])//2) + 2
            row = 5
            WordRowCol[sheet].append((activite,row-4,indexAct+2,fontsDict['Activite']))
            for cout in Dict[sheet][activite].keys():
                if Dict[sheet][activite][cout][0] != []:
                    WordRowCol[sheet].append((cout,row-2,indexAct,fontsDict['Cout']))
                    for dataframe in Dict[sheet][activite][cout][0]:
                        WordRowCol[sheet].append((dataframe[1],row+1,indexAct+1,fontsDict['Titre']))
                        dataframe[0].to_excel(writer,sheet_name=sheet,startrow=row , startcol=indexAct)   
                        row = row + len(dataframe[0].index) + spaces + 1
    writer.save()
    #Fin ecritre données
    wb = load_workbook('Parametres.xlsx')
    for sheet in WordRowCol.keys():
        ws = wb[sheet]
        for i in WordRowCol[sheet]:
            ws.cell(i[1],i[2]).value = i[0]
            ws.cell(i[1],i[2]).font = i[3]
        for col in ws.columns:
            max_length = 0
            column = col[0].column # Get the column name
            for cell in col:
                try: # Necessary to avoid error on empty cells
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width
    wb.save('Parametres.xlsx')
            

    

class Projet:
    
    def __init__(self,Nom="",Localisation="",ListeActivites=[],Horizon=0,CPC=[]):
        self.Nom = Nom
        self.Localisation = Localisation
        self.ListeActivites = ListeActivites
        self.Horizon = Horizon
        self.CPC = CPC
        self.DictParams = {"Int couts":{},
                           "Mar couts":{},
                           "Int revenus":{},
                           "Mar revenus":{}}
                           
    def PrepareExcelInput(self):
        for activite in self.ListeActivites:
            for key in self.DictParams.keys():
                self.DictParams[key][activite.getNom()] = {}
            for cout in activite.getlistCout():
                MaxCols = max([i.getTableauAffichage().shape[1] for i in cout.getListTableaux()])
                # Attention au changement du titre
                self.DictParams['Int couts'][activite.getNom()][cout.getNom()] = ([(T.getTableauAffichage(),T.getTitre()) for T in cout.getListTableaux()],MaxCols) 
                self.DictParams['Mar couts'][activite.getNom()][cout.getNom()] = ([(T.getTableauAffichage(),T.getTitre()) for T in cout.getListTableauxMarche()],MaxCols)
        excel_dfs(self.DictParams,"Parametres.xlsx",4)
    def getDictParams(self):
        return self.DictParams
    
    def setDictParams(self,Dict):
        self.DictParams = Dict
    
    #Méthode pour enregistrer tout les tableaux
        
    def IdentifierActivitesPossibles(self):
        print("Identification des Activités")
        wb = load_workbook('References.xlsx')
        ws = wb['Ref activite']
        for col in ws.iter_cols(min_row=2,min_col=3, max_col=78, max_row=2):
            for cell in col:
                if cell.value == self.getNom():
                    col_to_search = cell.col_idx
        for row in ws.iter_rows(min_row=3,max_row=87,min_col=col_to_search,max_col=col_to_search):
            for cell in row:
                if cell.value == 1:
                    A = Activite()
                    A.setNom(ws['C'+str(cell.row)].value)
                    self.ListeActivites.append(A)
        
    def AjoutActivites(self,NewListeActiv):
        self.ListeActivites += NewListeActiv 
    
    def CommencerSaisie(self):
        pass
    
    def GenerateCPC(self):
        #Programme de generation du cpc et retour Du CPC Final
        pass
        
    #Setters 
    def setNom(self,Nom):
        self.Nom = Nom
    def setLocalisation(self,Localisation):
        self.Localisation = Localisation
    def setListeActivite(self,ListeActivites):
        self.ListeActivites = ListeActivites
    def setHorizon(self,Horizon):
        self.Horizon = Horizon

    #Getters
    def getNom(self):
        return self.Nom
    def getLocalisation(self):
        return self.Localisation
    def getListeActivites(self):
        return self.ListeActivites
    def showListeActivites(self):
        for i in self.ListeActivites:
            print(i.getNom())
    def getHorizon(self):
        return self.Horizon
    
class TableauSaisie:
    def __init__(self,Titre="",IntituleLigne="",IntituleColonne="",Taille = [1,1]):
        self.Titre = Titre
        self.IntituleLigne = IntituleLigne
        self.IntituleColonne = IntituleColonne
        self.Taille = Taille # [Indice Ligne , Indice Colonne]
        self.TableauAffichage = pd.DataFrame(np.zeros((Taille[0],Taille[1])))
        self.TableauDonnees = np.zeros((Taille[0],Taille[1]))
        self.TableauAffichage.columns = [self.IntituleColonne+'='+str(i) for i in range(Taille[1])]
        self.TableauAffichage.index = [self.IntituleLigne+'='+str(i) for i in range(Taille[0])]
    
    #Setters
    def setTitre(self,Titre):
        self.Titre = Titre
    def setIntituleLignes(self,IntituleLigne):
        self.IntituleLigne = IntituleLigne
    def setIntituleColonne(self,IntituleColonne):
        self.IntituleColonne = IntituleColonne
    def setTaille(self,Taille):
        self.Taille = Taille
    #def setTableauAffichage(self,TableauAffichage):
        #self.TableauAffichage = TableauAffichage
    def setTableauDonnees(self,TableauDonnees):
        self.TableauDonnees = TableauDonnees
        self.TableauAffichage = pd.DataFrame(self.TableauDonnees)
        self.TableauAffichage.columns = [self.IntituleColonne+'='+str(i) for i in range(Taille[1])]
        self.TableauAffichage.index = [self.IntituleLigne+'='+str(i) for i in range(Taille[0])]
    
    #Getters
    def getTitre(self):
        return self.Titre    
    def getIntituleLigne(self):
        return self.IntituleLigne  
    def getIntituleColonne(self):
        return self.IntituleColonne  
    def getTaille(self):
        return self.Taille  
    def getTableauAffichage(self):
        return self.TableauAffichage  
    def getTableauDonnees(self):
        return self.TableauDonnees  
    
class Activite:
    def __init__(self,nom = "",listRev=[],listCout=[]):
        self.nom      = nom
        self.listRev  = listRev
        self.listCout = listCout
    
    #Setters
    def setNom(self,a) :
        self.nom = a
    def setlistRev(self,b) :
        self.listRev = b
    def setlistCout(self,c) :
        self.listCout = c
     
    #Getters
    def getNom(self) :
        return self.nom
    def getlistRev(self) :
        return self.listRev
    def getlistCout(self) :
        return self.listCout
    def showlistCout(self):
        for i in self.listCout:
            print(i.getNom())
    
    def IdentifierCouts(self):
        print("Identification des Couts")
        wb = load_workbook('References.xlsx')
        ws = wb['Ref couts']
        for col in ws.iter_cols(min_row=2,min_col=4, max_col=90, max_row=2):
            for cell in col:
                if (self.nom == cell.value):
                    coltosearch = cell.col_idx
        for row in ws.iter_rows(min_row=3,min_col=coltosearch, max_col=coltosearch, max_row=52):
            for cell in row:
                if (cell.value == "x"):
                    C = Cout(ws['C'+str(cell.row)].value)
                    self.listCout.append(C)
                    
    
    def IdentifierRevenus(self):
        print("Identification des Revenus")
        wb = load_workbook('References.xlsx')
        ws = wb['Ref revenus']
        for col in ws.iter_cols(min_row=2,min_col=3, max_col=89, max_row=2):
            for cell in col:
                if (self.nom == cell.value):
                    coltosearch = cell.col_idx
        for row in ws.iter_rows(min_row=3,min_col=coltosearch, max_col=coltosearch, max_row=71):
            for cell in row:
                if (cell.value == "x"):
                    R = Revenu(ws['A'+str(cell.row)].value)
                    self.listRev.append(R)
        
class Cpc:
    
    def __init__(self,ListeActivites = [],Dotation = [],ServiceDette = []) :
        self.ListeActivites    = ListeActivites
        self.Dotation          = Dotation
        self.ServiceDette      = ServiceDette
        
        
    #Setters
    def setListeActivites(self,a) :
        self.ListeActivites = a
    def setDotation(self,b) :
        self.Dotation = b
    def setServiceDette(self,c) :
        self.ServiceDette = c
     
    #Getters
    def getListeActivites(self) :
        return self.ListeActivites
    def getDotation(self) :
        return self.Dotation
    def getServiceDette(self) :
        return self.ServiceDette
    
#Definition des classes couts et activtés
class Cout:
    def __init__(self,Nom = "", Horizon=0,SaisieStartCol = 1,CoutCPC = []):
        self.Nom     = Nom
        self.Horizon = Horizon
        self.CoutCPC = CoutCPC
        self.listeTableaux = []
        self.listeTableauxMarche = []
        self.SaisieStartCol = SaisieStartCol
    #Setters
    def setNom(self, Nom):
        self.Nom = Nom
    def setHorizon(self,Horizon):
        self.Horizon = Horizon
    def setCoutCPC(self,CoutCPC):
        self.CoutCPC = CoutCPC
    def setSaisieStartCol(self,SaisieStartCol):
        self.SaisieStartCol = SaisieStartCol
    def setListTableauxMarche(self,listeTableauxMarche):
        self.listeTableauxMarche = listeTableauxMarche
    
    #Getters
    def getNom(self):
        return self.Nom
    def getHorizon(self):
        return self.Horizon
    def getCoutCPC(self):
        return self.CoutCPC
    def getListTableaux(self):
        return self.listeTableaux
    def getSaisieStartCol(self):
        return self.SaisieStartCol
    def getListTableauxMarche(self):
        return self.listeTableauxMarche
    
    def SaisieIntrinseque(self):
        #Au préalable necessite le nom du cout, l'horizon, SaisieStartCol
        wb = load_workbook("References.xlsx")
        ws=wb['Cout x Tableau']
        for col in ws.iter_cols(min_row=2,min_col=3, max_col=34, max_row=2):
            for cell in col:
                if (cell.value == self.Nom):
                    coltosearch = cell.col_idx
        #Liste temporaire occupant les titres des Tableaux de saisie
        Titres = []
        for row in ws.iter_rows(min_row=2,min_col=coltosearch, max_col=coltosearch, max_row=23):
            for cell in row:
                if cell.value == 'x':
                    if ws['B'+str(cell.row)].value == 'int':
                        Titres.append(ws['A'+str(cell.row)].value)
        wsTableau = wb['Tableau x Features']
        for j in Titres:
            for col in wsTableau.iter_cols(min_row=2,min_col=2, max_col=22, max_row=2):
                for cell in col:
                    if (cell.value == j):
                        Taille = [0,0]
                        IntituleLigne = wsTableau[cell.column+'3'].value
                        TypeCol = wsTableau[cell.column+'4'].value
                        IntituleColonne = wsTableau[cell.column+'5'].value
                        #Un Test sur le type d'indices ligne pour le tableau
                        if IntituleLigne == 'Horizon':
                            Taille[0] = self.Horizon
                        else:
                            Taille[0] = int(input("Veuillez saisir votre "+IntituleLigne))
                        if TypeCol == 1:
                            Taille[1] = 1
                        else:
                            Taille[1] = int(input("Veuillez saisir votre "+str(TypeCol)))
                        self.listeTableaux.append(TableauSaisie(j,IntituleLigne,IntituleColonne,Taille))
        
        #Ajouter les tableaux à Excel sans titres
        '''export_dfs([T.getTableauAffichage() for T in self.getListTableaux()],'Int couts','Parametres.xlsx',3,self.SaisieStartCol-1)
        wb=load_workbook('Parametres.xlsx')
        ws=wb['Int couts']
        ws.cell(1,self.SaisieStartCol+2).value = self.Nom
        self.TaillesTableaux = [len(T.getTableauAffichage()) for T in self.getListTableaux()]
        row = 4
        for i,ecartrow in enumerate(self.TaillesTableaux):
            ws.cell(row,self.SaisieStartCol).value = self.getListTableaux()[i].getTitre()
            row += ecartrow+4
        wb.save('Parametres.xlsx')'''
    def SaisieMarche(self):
        wb = load_workbook("References.xlsx")
        ws=wb['Cout x Tableau']
        for col in ws.iter_cols(min_row=2,min_col=3, max_col=34, max_row=2):
            for cell in col:
                if (cell.value == self.Nom):
                    coltosearch = cell.col_idx
        #Liste temporaire occupant les titres des Tableaux de saisie
        Titres = []
        for row in ws.iter_rows(min_row=2,min_col=coltosearch, max_col=coltosearch, max_row=23):
            for cell in row:
                if cell.value == 'x':
                    if ws['B'+str(cell.row)].value == 'ext':
                        Titres.append(ws['A'+str(cell.row)].value)
        wsTableau = wb['Tableau x Features']
        for j in Titres:
            for col in wsTableau.iter_cols(min_row=2,min_col=2, max_col=22, max_row=2):
                for cell in col:
                    if (cell.value == j):
                        Taille = [0,0]
                        IntituleLigne = wsTableau[cell.column+'3'].value
                        TypeCol = wsTableau[cell.column+'4'].value
                        IntituleColonne = wsTableau[cell.column+'5'].value
                        #Un Test sur le type d'indices ligne pour le tableau
                        if IntituleLigne == 'Horizon':
                            Taille[0] = self.Horizon
                        else:
                            Taille[0] = int(input("Veuillez saisir votre "+IntituleLigne))
                        if TypeCol == 1:
                            Taille[1] = 1
                        else:
                            Taille[1] = int(input("Veuillez saisir votre "+str(TypeCol)))
                        self.listeTableauxMarche.append(TableauSaisie(j,IntituleLigne,IntituleColonne,Taille))
    def CalculCout(self):
        pass

class Revenu:
    def __init__(self, Nom="",Horizon=0,RevenuCPC=[]):
        self.Nom = Nom
        self.Horizon = Horizon
        self.RevenuCPC = RevenuCPC
        
    #Setters
    def setNom(self, Nom):
        self.Nom = Nom
    def setHorizon(self,Horizon):
        self.Horizon = Horizon
    def setRevenuCPC(self,RevenuCPC):
        self.RevenuCPC = RevenuCPC
    
    #Getters
    def getNom(self):
        return self.Nom
    def getHorizon(self):
        return self.Horizon
    def getRevenuCPC(self):
        return self.RevenuCPC
    
    def SaisieIntrinseque():
        pass
    def SaisieMarche():
        pass
    def CalculRevenu():
        pass