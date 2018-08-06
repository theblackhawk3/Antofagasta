import numpy as np

def affluence_conversion(horizon,pas):
    t=0
    if pas == "M":  t=1
    if pas == "T":  t=3
    if pas == "S":  t=6
    if pas == "A":  t=12
    return t
    
def roundN(p):
    if p != int(p):
        return int(p) +1
    else:
        return p

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
    elif revenu.getNom() == 'Plats vendus':
        Affluence = DictTableaux['Affluence de clients']
        PartsPaniers = DictTableaux['Parts paniers moyens']
        PrixPaniers = DictTableaux['Prix paniers moyens']
        SP = np.sum(np.multiply(PartsPaniers,PrixPaniers)[0],axis=0)
        Affluence = SP*Affluence
        Affluence.resize((1,Affluence.size))
        Affluence = Affluence[0]
        revenu.resultat = [0]*revenu.Horizon
        pas = affluence_conversion(revenu.Horizon,revenu.pasVisualisation)
        for i in range(len(revenu.resultat)):
            revenu.resultat[i] = sum(Affluence[pas*i:min(pas*(i+1),len(Affluence))])
    elif revenu.getNom() == "Boites de médicaments vendues":
        VentesMedic = DictTableaux['Ventes par type de médicaments']
        PrixMedic = DictTableaux['Prix par type médicaments']
        revenu.resultat = list(np.sum(np.multiply(VentesMedic,PrixMedic),axis=1))
    elif revenu.getNom() == "Opérations effectuées":
        PrixPrestation = DictTableaux['Prix par type de prestation de soin médical']
        PartsPrestation = DictTableaux['Parts par type de prestation de soin médical']
        TempsJourn = DictTableaux['Temps d\'ouverture journalier ']
        DureeMoyPrest = DictTableaux['Durée moyenne de la prestation de soin médical']
        SP = np.sum(np.multiply(PartsPrestation,PrixPrestation)[0],axis=0)
        SP = SP*TempsJourn[0][0]/DureeMoyPrest[0][0]
        revenu.resultat = [SP*22*affluence_conversion(revenu.Horizon,revenu.pasVisualisation)]*revenu.Horizon
    elif revenu.getNom() == "Diagnostics médicaux efféctués ":
        PrixPrestation = DictTableaux['Prix par type de prestation de diagnostic médical']
        PartsPrestation = DictTableaux['Parts par type de prestation de diagnostic médical']
        TempsJourn = DictTableaux['Temps d\'ouverture journalier ']
        DureeMoyPrest = DictTableaux['Durée moyenne de la prestation de diagnostic médical']
        SP = np.sum(np.multiply(PartsPrestation,PrixPrestation)[0],axis=0)
        SP = SP*TempsJourn[0][0]/DureeMoyPrest[0][0]
        revenu.resultat = [SP*22*affluence_conversion(revenu.Horizon,revenu.pasVisualisation)]*revenu.Horizon
    elif revenu.getNom() == "Biens immobiliers vendus":
        PrixM2 = DictTableaux['Prix du m2 par type de bien']
        NombreM2 = DictTableaux['Nombre de m2 par type de bien']
        EclatementTitres = DictTableaux['Date éclatement des titres'][0][0]
        PrcAvances = DictTableaux['Pourcentage avance'][0][0]
        VentesPrev = DictTableaux['Ventes prévisionnelles par type de bien']
        #Calcul 
        PrixBien =(PrixM2*NombreM2)[0]
        MatriceIntermediaire = np.array([[0]*VentesPrev.shape[1]]*VentesPrev.shape[0])
        for i in range(MatriceIntermediaire.shape[0]):
            for j in range(MatriceIntermediaire.shape[1]):
                if VentesPrev[i][j] != 0:
                    if i < EclatementTitres - 1:
                        MatriceIntermediaire[i][j] += PrixBien[j]*PrcAvances
                        MatriceIntermediaire[EclatementTitres - 1][j] += PrixBien[j]*(1-PrcAvances)
                    else:
                        MatriceIntermediaire[i][j] += PrixBien[j]*PrcAvances
        
        revenu.resultat = np.sum(MatriceIntermediaire,axis=1)
        
    elif revenu.getNom() == "Produits fabriqués vendus":
        Ventes = DictTableaux['Ventes en unités de produit']
        PrixVentes = DictTableaux['Prix de vente par produit']
        revenu.resultat = list(np.sum(PrixVentes*Ventes,axis=1))
        