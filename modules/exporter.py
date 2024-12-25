import pandas as pd

def exporter_donnees(tableau, chemin_sortie):
    """
    Exporter les données dans un fichier Excel.
    """
    colonnes_recap = ['NOM ET PRENOMS', 'MONT.PAYE', 'MONT.DÜ', 'MONT.TOTAL', 'Filiere', 'feuille']

    df_recap_global = pd.DataFrame(tableau, columns=colonnes_recap)

    #print(df_recap_global)
    df_recap_global.to_excel(chemin_sortie, index=False)
    print(f"Tableau récapitulatif global enregistré dans {chemin_sortie}.")
