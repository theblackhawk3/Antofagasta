#### Main Code ###
## Modification ##
##Adam Pc Modification##
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
A.PrepareExcelInput()  
