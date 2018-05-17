#### Main Code ###
from FormulesCouts import *
import numpy as np
A = Projet(input("Veuillez Saisir le nom de votre projet : "))
A.IdentifierActivitesPossibles()
A.showListeActivites()
A.setHorizon(5)
compteur = 1
for activite in A.getListeActivites():
    activite.IdentifierCouts()
    for cout in activite.getlistCout():
        cout.setHorizon(5)
        cout.setSaisieStartCol(compteur)
        cout.SaisieIntrinseque()
        cout.SaisieMarche()
A.PrepareExcelInput()  

#Remplir les données dans le fichier Excel et puis exécuter :

A.GetExcelInput()

#Executer le fichier Formulescouts avant de calculer les couts
for activite in A.ListeActivites:
    for cout in activite.getlistCout():
            cout.CalculCout()

#Generer le CPC Final
A.GenerateCPC()