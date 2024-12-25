import os
from modules.file_loader import charger_fichier_excel
from modules.sheet_processor import traiter_feuilles
from modules.data_cleaner import nettoyer_tableau
from modules.exporter import exporter_donnees


def creer_sous_dossier_recap(dossier):
    """
    Crée un sous-dossier 'recap' dans le dossier spécifié s'il n'existe pas.
    """
    sous_dossier_recap = os.path.join(dossier, "recap")
    if not os.path.exists(sous_dossier_recap):
        os.makedirs(sous_dossier_recap)
    return sous_dossier_recap


def traiter_dossiers(dossiers):
    """
    Parcourt les dossiers spécifiés et traite tous les fichiers Excel dans chaque dossier.
    """
    for dossier in dossiers:
        print(f"Traitement du dossier : {dossier}")
        sous_dossier_recap = creer_sous_dossier_recap(dossier)

        fichiers_excel = [f for f in os.listdir(dossier) if f.endswith('.xlsx')]
        if not fichiers_excel:
            print(f"Aucun fichier Excel trouvé dans {dossier}.")
            continue

        for fichier in fichiers_excel:
            chemin_fichier = os.path.join(dossier, fichier)
            nom_fichier_sortie = f"recapitulatif_global_{os.path.splitext(fichier)[0]}.xlsx"
            chemin_sortie = os.path.join(sous_dossier_recap, nom_fichier_sortie)

            try:
                # Charger le fichier Excel
                xls, feuilles = charger_fichier_excel(chemin_fichier)

                # Traiter les feuilles et générer le tableau récapitulatif
                tableau_recap_global = traiter_feuilles(xls, feuilles)

                # Nettoyer le tableau
                tableau_nettoye = nettoyer_tableau(tableau_recap_global, dossier)

                # Exporter les données
                exporter_donnees(tableau_nettoye, chemin_sortie)

                print(f"Fichier traité et exporté : {chemin_sortie}")
            except Exception as e:
                print(f"Erreur lors du traitement de {fichier} : {e}")


def main():
    dossiers = ["part_1", "part_2", "part_3", "part_4"]
    traiter_dossiers(dossiers)


if __name__ == "__main__":
    main()
