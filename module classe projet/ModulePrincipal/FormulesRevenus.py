import numpy as np

def CalculerRevenu(revenu):
    DictTableaux = revenu.getDicoTableaux()
    if revenu.getNom() == 'Inscriptions enregistrées dans un cours':
        Capacite = DictTableaux['Capacité (places élèves - classes)']
        TauxRemp = DictTableaux['Taux de remplissage (classes)']
        Frais    = DictTableaux['Frais (annuels par élève - classes)']
        print(Capacite)
        print(TauxRemp)
        print(Frais)
        revenu.resultat=list(np.sum(np.multiply(np.multiply(Capacite,TauxRemp),Frais),axis=1))
        