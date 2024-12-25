def nettoyer_tableau(tableau_recap_global, dossier_courant):
    """
    Nettoyer le tableau récapitulatif global.
    """
    tableau_nettoye = []

    for ligne in tableau_recap_global:
        if dossier_courant == "part_2":
            ligne[1], ligne[2], ligne[3] = ligne[2], ligne[3], ligne[1]
        tableau_nettoye.append(ligne)

    return tableau_nettoye


