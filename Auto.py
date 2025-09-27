import pandas as pd

def check_automated(file_path, sheet_name):
    # Lire le fichier Excel
    data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Rechercher le mot "Historique" dans la colonne A
    historique_indices = data[data[0].astype(str).str.contains("Historique", case=False, na=False)].index
    
    if not historique_indices.empty:
        start_index = historique_indices[0] + 3
        for i in range(start_index, len(data)):
            current_row = data.iloc[i]
            if isinstance(current_row[2], str):
                if 'automated' in current_row[2].lower():
                    previous_row = data.iloc[i-1]
                    if 'automated' in str(previous_row[2]).lower():
                        print(f"Automated existe dans la ligne precedente ({i}).")
                    else:
                        print(f"Automated supprimé dans la ligne {i}. Veuillez l'ajouter si nécessaire.")

check_automated(
    "C:/Users/ssuiquik/Desktop/Python/Tests_20_71_01272_19_01464_FSE_ACC_V27_VSM.xlsm",
    "VSM20_GC_20_71_0059",
)