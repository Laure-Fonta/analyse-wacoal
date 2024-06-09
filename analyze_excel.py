import pandas as pd

# Charger le fichier Excel
file_path = 'REWORK STOCK.xlsx'  # Assurez-vous que le chemin est correct
df = pd.read_excel(file_path)

# Filtrer les produits contenant 'non trouvé' ou 'création de déclinaison' dans la colonne Import
non_trouves = df[df['Import'].str.contains("non trouvé|création déclinaison", case=False, na=False)]

# Supprimer les lignes avec une quantité inférieure à 100
non_trouves = non_trouves[non_trouves['Quantité'] >= 100]

# Extraire la couleur et la collection
def extraire_collection_couleur(row):
    mots = row['Nom'].split()
    collection = ' '.join(mots[:3])
    couleur = row['Couleur']
    return pd.Series([collection, couleur])

non_trouves[['collection', 'couleur']] = non_trouves.apply(extraire_collection_couleur, axis=1)

# Supprimer les doublons
non_trouves = non_trouves.drop_duplicates()

# Calculer la somme des quantités pour chaque combinaison de couleur et collection
synthese = non_trouves.groupby(['collection', 'couleur'])['Quantité'].sum().reset_index()

# Classer par ordre décroissant des quantités
synthese = synthese.sort_values(by='Quantité', ascending=False)

# Afficher la synthèse
print(synthese)

# Exporter les résultats vers un fichier Excel
output_path = 'synthese_non_trouves_avec_declinaison.xlsx'  # Chemin relatif ou absolu
synthese.to_excel(output_path, index=False)

print(f"Le fichier de synthèse a été sauvegardé à l'emplacement : {output_path}")
