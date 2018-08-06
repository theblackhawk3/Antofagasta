def CalculerCout(cout):
    import numpy as np
    DictTableaux = cout.getDicoTableaux()
    
    if cout.getNom() == "Couts d'assurance":
        cout.resultat = [list(i)[0] for i in DictTableaux['Autres assurances (MAD)']]
        
    elif cout.getNom() == "Cout de location, espace":
        cout.resultat = list(np.sum(np.multiply(DictTableaux['Nombre espaces loués'],DictTableaux['Prix de Location'],DictTableaux["Nombre Locataires / Espace"]),axis=1))
        
    elif cout.getNom() == "Cout de contribution à la gestion d'une copropriété":
        pass 
    elif cout.getNom() == "Couts de nettoyage et gardiennage":
        pass
        
    elif cout.getNom() == "Cout MP obsolète":
        pass
    elif cout.getNom() == "Cout de marchandise achetée ":
        pass
        
    elif cout.getNom() == "Cout d'energie, gaz":
        pass
        
    elif cout.getNom() == "Cout d'energie, électricité":
        a = np.multiply(DictTableaux['Consommation Electrique par poste ( Kwh )'],DictTableaux['Prix du Kwh'])
        cout.resultat = list(np.sum(np.multiply(a, DictTableaux['Nombre de postes de consommation électriques']),axis=1))
        
    elif cout.getNom() == "Cout d'énérgie, carburant":
        pass
        
    elif cout.getNom() == "Cout de consommation en eau":
        a = np.multiply(DictTableaux['Nombre de postes de consommation Eau'],DictTableaux['Consommation d\'Eau par poste ( MAD )'])
        cout.resultat = list(np.sum(a,axis = 1))
        
    elif cout.getNom() == "Cout d'entretien":
        cout.resultat = [list(i)[0] for i in DictTableaux['Maintenance (MAD)']]
        
    elif cout.getNom() == "Cout RH":
        Nrh = DictTableaux['Nombre Ressources Humaines']
        SalairesRH = DictTableaux['Salaire Ressources Humaines']
        cout.resultat = list(np.sum((Nrh*SalairesRH),axis=1))
        
    elif cout.getNom() == "Couts de transport":
        Nrh = DictTableaux['Nombre Ressources Humaines']
        Nvehs = DictTableaux['Nombre de véhicules']
        ConsoCarb = DictTableaux['Consommation en carburant / véhicule']
        SalairesRH = DictTableaux['Salaire Ressources Humaines']
        CoutAssurance = DictTableaux["Cout d'assurance / Type de véhicule"]
        PrixLitreCarb = DictTableaux['Prix moyen du litre carburant']
        cout.resultat = list(((np.sum((Nrh*SalairesRH),axis=1)+
                               np.sum(Nvehs*ConsoCarb,axis=1)*np.transpose(PrixLitreCarb)                                   +np.sum((Nvehs*CoutAssurance),axis=1))[0]))
    elif cout.getNom() == "Cout des consommables de bureau":
        pass
    elif cout.getNom() == "Cout des consommables offerts au client":
        pass
    elif cout.getNom() == "Cout des consommables de santé":
        PrixConso = DictTableaux['Nombre Consommable santé']
        NbrConso = DictTableaux['Prix consommable santé']
        cout.resultat = list(np.sum((PrixConso*NbrConso),axis=1))
        
    elif cout.getNom() == "Cout MP Restauration" or cout.getNom() == "Cout MP utilisée ":
        QteMP = DictTableaux['Quantité achetée par type de MP']
        PrixMP = DictTableaux['Prix / Type MP']
        cout.resultat = list(np.sum(np.multiply(QteMP,PrixMP),axis = 1))