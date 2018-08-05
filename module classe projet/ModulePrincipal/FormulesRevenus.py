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
        
    elif revenu.getNom() == 'Nuitées vendues':
        Capacite = DictTableaux['Capacité par type de chambre']
        TauxRemp = DictTableaux['Taux de remplissage chambres']
        Frais    = DictTableaux['Frais annuels par chambre']
        print(Capacite)
        print(TauxRemp)
        print(Frais)
        revenu.resultat=list(np.sum(np.multiply(np.multiply(Capacite,TauxRemp),Frais),axis=1))
        
    elif revenu.getNom() == 'Inscriptions au service de transport enregitrées':
        Capacite = DictTableaux['Capacité (places élèves - transport)']
        TauxRemp = DictTableaux['Taux de remplissage (véhicules transport)']
        Frais    = DictTableaux['Frais (annuels par élève - transport)']
        revenu.resultat = list(np.transpose(np.multiply(np.multiply(Capacite,TauxRemp),Frais))[0])