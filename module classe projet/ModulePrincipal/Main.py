#### Main Code ###
from FormulesCouts import *
import numpy as np
A = Projet(input("Veuillez Saisir le nom de votre projet : "))
A.IdentifierActivitesPossibles()
A.showListeActivites()
compteur = 1
for activite in A.getListeActivites():
    activite.IdentifierCouts()
    for cout in activite.getlistCout():
        cout.setHorizon(5)
        cout.setSaisieStartCol(compteur)
        cout.SaisieIntrinseque()
        cout.SaisieMarche()
A.PrepareExcelInput()  

for activite in A.getListeActivites():
    for cout in activite.getlistCout():
        L=cout.getListTableauxMarche()
        print(cout.getNom())
        for i in L:
            print(i.getTitre())
            print(i.getTableauAffichage())
Activity = A.getListeActivites()[0]

for activite in A.ListeActivites:
    for cout in activite.getlistCout():
        if cout.getNom() == "Cout de location, espace":
            CalculerCout(cout)
            print(cout.resultat)
