import pandas as pd
import numpy as np

FEUILLES_A_IGNORER = {"A_IGNORER_1", "A_IGNORER_2", "A_IGNORER_3", "A_IGNORER_4", "A_IGNORER_5"}

ELEMENTS_DE_JEU = ['N°', 'NOM  ET  PRENOMS', 'MONT.TOTAL', 'MONT.PAYE', 'MONT.DÜ', 'Mont. Versé', 'Restant dû', 'Total', 'NOM', 'PRENOMS', 'Montant recouvré ', 'Solde', 'Nom & Prénoms', 'Nom et Prenoms', 'NOM ET PRENOMS', 'MONTANT RECOUVRE', 'SOLDE A PAYER', 'TOTAL A PAYER', 'NOM ET PRENOMS', 'MONTANT VERSE', 'MONTANT DÛ', 'TOTAL' ]

def traiter_feuilles(xls, feuilles):
    """
    Traiter toutes les feuilles du fichier Excel et générer un tableau récapitulatif global.
    """
    tableau_recap_global = []

    for feuille in feuilles:
        if feuille in FEUILLES_A_IGNORER:
            print(f"Feuille ignorée : {feuille}")
            continue

        print(f"Traitement de la feuille : {feuille}")
        try:
            df = pd.read_excel(xls, sheet_name=feuille)
            df = nettoyer_feuille(df)
            tableau = extraire_donnees_recap(df, feuille)

            if feuille == "GC2 ok":
                for ligne in tableau:
                    print(ligne)
            
            if tableau:
                # Ajoute la première ligne avec la chaîne 'feuille' à la dernière colonne
                tableau_recap_global.append(tableau[0] + ['feuille'])
                # Ajoute les autres lignes avec le vrai nom de la feuille
                tableau_recap_global.extend([ligne + [feuille] for ligne in tableau[1:]])
        except Exception as e:
            print(f"Erreur lors du traitement de la feuille '{feuille}' : {e}")

    ligne_a_supprimer = ['NOM  ET  PRENOMS', 'MONT.TOTAL', 'MONT.PAYE', 'MONT.DÜ', '', 'Filiere', 'feuille']

    # Conserver uniquement la première occurrence de `ligne_a_supprimer` et les autres lignes différentes
    nouveau_tableau_recap = []
    premiere_occurrence = False

    for ligne in tableau_recap_global:
        if ligne == ligne_a_supprimer:
            if not premiere_occurrence:
                nouveau_tableau_recap.append(ligne)
                premiere_occurrence = True
        else:
            if len(ligne) == 8:
                del ligne[3]

            if len(ligne) == 6:
                ligne.insert(5, '')
            nouveau_tableau_recap.append(ligne)

    for ligne in tableau_recap_global:
        del ligne[4]
        for i in range(len(ligne)):
            if ligne[i] == '':
                ligne[i] = 0

    # Afficher le tableau modifié
    tableau_recap_global = nouveau_tableau_recap

    return tableau_recap_global


def nettoyer_feuille(df):
    """
    Nettoyer une feuille Excel :
    - Supprimer les lignes et colonnes vides
    - Enlever les colonnes contenant 'Unnamed'
    """
    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(how='all').dropna(how='all', axis=1)  
    return df.reset_index(drop=True)


def extraire_donnees_recap(df, feuille):
    """
    Extraire les lignes utiles d'une feuille après nettoyage.
    """
    tableau = df.values.tolist()

    if feuille == "GC2 ok":
        for ligne in tableau:
            print(ligne)
    

    tableau_sans_nan = [ligne for ligne in tableau if not all(pd.isna(val) for val in ligne)]
    tableau_sans_total_general = [ligne for ligne in tableau_sans_nan if not any("TOTAL GENERAL" in str(val) for val in ligne)]
    tableau_sans_licence = [ligne for ligne in tableau_sans_total_general if not any("LICENCE" in str(val) for val in ligne)]
    tableau_sans_filiere = [ligne for ligne in tableau_sans_licence if not any("Filière" in str(val) for val in ligne)]
    tableau_sans_annee = [ligne for ligne in tableau_sans_filiere if not any("Année" in str(val) for val in ligne)]
    tableau_sans_total = [ligne for ligne in tableau_sans_annee if not any("BALANCE" in str(val) for val in ligne)]
    tableau_sans_totaux = [ligne for ligne in tableau_sans_total if not any("TOTAUX" in str(val) for val in ligne)]
    
    # Appeler la fonction pour splitter le tableau en fonction de la ligne de split

    tableau_split = split_table_with_header(pd.DataFrame(tableau_sans_totaux), ELEMENTS_DE_JEU)
    return tableau_split


def split_table_with_header(df, elements_de_jeu):
    """
    Splitte le tableau lorsque des lignes contenant au moins 4 des éléments de jeu sont trouvées.
    Ajoute à chaque ligne d'un sous-tableau la valeur héritée (dernier élément non-NaN du sous-tableau précédent),
    puis supprime la dernière ligne du sous-tableau précédent après l'extraction de la valeur.
    Fusionne ensuite tous les sous-tableaux non vides en un seul tableau avec un seul titre.
    """
    def contains_enough_elements(row, elements_de_jeu):
        """Vérifie si une ligne contient au moins 4 des éléments spécifiés."""
        count = sum(elem in str(row) for elem in elements_de_jeu)
        return count >= 4

    def clean_row(row):
        """Convertit tous les éléments d'une ligne en chaînes, remplace les NaN par une chaîne vide."""
        return [str(val) if not pd.isna(val) else "" for val in row]

    sublists = [] 
    current_list = [] 
    last_non_nan_value = None  

    for row in df.values:
        if contains_enough_elements(row, elements_de_jeu):  
            if current_list:  # Si un sous-tableau est en cours, on le termine
                # Récupérer la dernière ligne et extraire le dernier élément non-NaN
                last_row = current_list.pop(-1)  # Supprime la dernière ligne après récupération
                last_non_nan_value = next((val for val in last_row if not pd.isna(val)), None)

                # Ajouter le sous-tableau actuel (sans la dernière ligne)
                sublists.append((current_list, last_non_nan_value))

            # Démarrer un nouveau sous-tableau
            current_list = [row]
        else:
            current_list.append(row)

    # Ajouter le dernier sous-tableau (s'il en reste)
    if current_list:
        # Récupérer la dernière ligne et extraire le dernier élément non-NaN
        last_row = current_list.pop(-1)  # Supprime la dernière ligne après récupération
        last_non_nan_value = next((val for val in last_row if not pd.isna(val)), None)

        # Ajouter le dernier sous-tableau (sans la dernière ligne)
        sublists.append((current_list, last_non_nan_value))

    # Fusionner les sous-tableaux non vides en un seul tableau
    merged_table = []  # Tableau final
    for i, (sublist, inherited_value) in enumerate(sublists):
        if sublist:  # Si le sous-tableau n'est pas vide
            # Ajouter la valeur héritée à chaque ligne du sous-tableau
            inherited_value_str = str(inherited_value) if inherited_value is not None else ""
            updated_sublist = [clean_row(ligne) + [inherited_value_str] for ligne in sublist]

            # Fusionner en un seul tableau, en supprimant les titres
            if merged_table:  # Si merged_table n'est pas vide, on ne garde pas la première ligne
                updated_sublist = updated_sublist[1:]  # Supprimer la première ligne (le titre)
            
            merged_table.extend(updated_sublist)

    merged_table = [ligne for ligne in merged_table if ligne[0] != '']


    merged_table = [ligne[1:] for ligne in merged_table]
    merged_table = [ligne for ligne in merged_table if ligne[0] != '0']
    merged_table = [ligne for ligne in merged_table if ligne[0] != '']

    while merged_table and (merged_table[0][0].startswith('NOM') or merged_table[0][0].startswith('NOM  ET  PRENOMS') or merged_table[0][0].startswith('Nom et Prénoms')):
        merged_table.pop(0)

    # Affichage du tableau fusionné
    #print("\n=== Tableau fusionné ===")
    #for ligne in merged_table:
        #print(ligne)

    return merged_table