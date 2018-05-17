def CalculerCout(cout):
    import numpy as np
    DictTableaux = cout.getDicoTableaux()
    
    if cout.getNom() == "Couts d'assurance":
        cout.resultat = [list(i)[0] for i in DictTableaux['Autres assurances (MAD)']]
        
    elif cout.getNom() == "Cout de location, espace":
        cout.resultat = list(np.sum(np.multiply(DictTableaux['Nombre espaces loués'],DictTableaux['Prix de Location']),axis=1))
        
    elif cout.getNom() == "Cout de contribution à la gestion d'une copropriété":
        pass 
    elif cout.getNom() == "Couts de nettoyage et gardiennage":
        pass
        
    elif cout.getNom() == "Cout MP utilisée":
        PM = np.dot(DictTableaux['Prix / Type MP'],DictTableaux['Nomenclature Produits'])
        cout.resultat =  list(np.sum(np.multiply(PM, DictTableaux['Fabrication par type de produit ']),axis = 1))
        
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
        a = np.multiply(DictTableaux['Consommation d\'Eau par poste ( m3 )'],DictTableaux['Consommation d\'Eau par poste ( MAD )'])
        cout.resultat = np.sum(a,axis = 1)
        
    elif cout.getNom() == "Cout d'entretien":
        cout.resultat = [list(i)[0] for i in DictTableaux['Maintenance (MAD)']]
        
    elif cout.getNom() == "Cout RH":
        pass
    elif cout.getNom() == "Couts de transport":
        pass
    elif cout.getNom() == "Cout des consommables de bureau":
        pass
    elif cout.getNom() == "Cout des consommables offerts au client":
        pass
    elif cout.getNom() == "Cout des consommables de santé":
        pass
    elif cout.getNom() == "Cout d'achat de contenu artistique":
        pass
    elif cout.getNom() == "Cout des consommables liés à l'activité sportive":
        pass
    elif cout.getNom() == "Couts des consomables liés à l'élevage de bétail":
        pass
    elif cout.getNom() == "Cout des consommables liés à la culture de plantes":
        pass
    elif cout.getNom() == "Cout des consommables liés à l'elevage de poisson":
        pass
    elif cout.getNom() == "Couts des repas offerts  au client":
        pass
    elif cout.getNom() == "Couts des consomables liés à la culture de plantes aquatiques":
        pass
    elif cout.getNom() == "Couts des consomables liés à la pêche":
        pass
    elif cout.getNom() == "Couts des consomables de bien être":
        pass
    elif cout.getNom() == "Couts des consommables de nettoyage":
        pass
    elif cout.getNom() == "Consommables d'entretiens de bateaux ":
        pass
    elif cout.getNom() == "Couts liés à l'activité sur internet":
        pass
    elif cout.getNom() == "Couts des consommables de destruction de batiments":
        pass
