import pandas as pd

def charger_fichier_excel(chemin_fichier):
    """
    Charger un fichier Excel et récupérer les noms des feuilles.
    """
    try:
        xls = pd.ExcelFile(chemin_fichier)
        feuilles = xls.sheet_names 
        if not feuilles:
            raise ValueError("Aucune feuille trouvée dans le fichier Excel.")
        return xls, feuilles
    except Exception as e:
        raise RuntimeError(f"Erreur lors du chargement du fichier : {e}")
