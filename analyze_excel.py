import pandas as pd
import re
from itertools import product

# Charger le fichier Excel
file_path = 'REWORK STOCK.xlsx'  # Assurez-vous que le fichier est dans le même répertoire que ce script
df = pd.read_excel(file_path)

# Garder uniquement les colonnes nécessaires (Réf, Descriptif, Coloris, Stock, Statut)
df = df[['Ref', 'Nom', 'Couleur', 'Quantité', 'Import']]

# Renommer les colonnes pour plus de clarté
df.columns = ['ref_article', 'descriptif', 'coloris', 'quantite', 'statut']

# Ne conserver que les deux premiers mots de la colonne 'descriptif'
df['descriptif'] = df['descriptif'].apply(lambda x: ' '.join(x.split()[:2]))

# Renommer les statuts
df['statut'] = df['statut'].replace({
    'Produit trouvé, création déclinaison': 'Création déclinaison',
    'Produit trouvé, couleur non trouvée': 'Couleur non trouvée',
    'Produit non trouvé': 'Produit absent',
    'Produit trouvé, taille non trouvée': 'Taille non trouvée'
})

# Remplacer les statuts avec "Produit trouvé, ancien stock : #{integer}"
df['statut'] = df['statut'].apply(lambda x: 'Produit trouvé' if re.match(r'Produit trouvé, ancien stock : \d+', x) else x)

# Filtrer les lignes où le stock est supérieur à 100
df = df[df['quantite'] > 100]

# Regrouper les lignes où ref_article, descriptif, et coloris sont les mêmes et cumuler les quantités
synthese = df.groupby(['ref_article', 'descriptif', 'coloris', 'statut'], as_index=False)['quantite'].sum()

# Ajouter une colonne pour les deux premières lettres de ref_article pour le tri
synthese['ref_prefix'] = synthese['ref_article'].apply(lambda x: x[:2])

# Séparer les tableaux en fonction des statuts
statuts_negatifs = ['Produit absent', 'Taille non trouvée', 'Couleur non trouvée']
statuts_positifs = ['Produit trouvé', 'Création déclinaison']

tableau_negatif = synthese[synthese['statut'].isin(statuts_negatifs)]
tableau_positif = synthese[synthese['statut'].isin(statuts_positifs)]

# Trier les tableaux : ordre alphabétique sur les deux premières lettres de ref_article, puis ordre décroissant sur quantite
tableau_negatif = tableau_negatif.sort_values(by=['ref_prefix', 'quantite'], ascending=[True, False])
tableau_positif = tableau_positif.sort_values(by=['ref_prefix', 'quantite'], ascending=[True, False])

# Supprimer la colonne temporaire 'ref_prefix'
tableau_negatif = tableau_negatif.drop(columns=['ref_prefix'])
tableau_positif = tableau_positif.drop(columns=['ref_prefix'])

# Regrouper les lignes ayant des descriptifs + coloris identiques et cumuler les quantités dans tableau_negatif
tableau_negatif_grouped = tableau_negatif.groupby(['descriptif', 'coloris'], as_index=False)['quantite'].sum()

# Trouver les combinaisons descriptif + coloris manquantes pour chaque ref_prefix dans tableau_negatif
ref_prefixes = tableau_negatif['ref_article'].apply(lambda x: x[:2]).unique()
combinations = []

for ref_prefix in ref_prefixes:
    prefix_df = tableau_negatif[tableau_negatif['ref_article'].str.startswith(ref_prefix)]
    all_descriptif_coloris = list(product(prefix_df['descriptif'].unique(), prefix_df['coloris'].unique()))
    existing_combinations = prefix_df[['descriptif', 'coloris']].drop_duplicates().apply(tuple, axis=1).tolist()
    missing_combinations = set(all_descriptif_coloris) - set(existing_combinations)
    for combination in missing_combinations:
        combinations.append({
            'ref_prefix': ref_prefix,
            'descriptif': combination[0],
            'coloris': combination[1]
        })

# Créer un DataFrame pour les combinaisons manquantes
missing_combinations_df = pd.DataFrame(combinations)

# Exporter les résultats vers des fichiers Excel séparés
output_path_negatif = 'synthese_stocks_negatifs.xlsx'
output_path_positif = 'synthese_stocks_positifs.xlsx'
output_path_missing = 'synthese_stocks_missing_combinations.xlsx'
output_path_negatif_grouped = 'synthese_stocks_negatifs_grouped.xlsx'

tableau_negatif.to_excel(output_path_negatif, index=False)
tableau_positif.to_excel(output_path_positif, index=False)
missing_combinations_df.to_excel(output_path_missing, index=False)
tableau_negatif_grouped.to_excel(output_path_negatif_grouped, index=False)

print(f"Le fichier de synthèse des statuts négatifs a été sauvegardé à l'emplacement : {output_path_negatif}")
print(f"Le fichier de synthèse des statuts positifs a été sauvegardé à l'emplacement : {output_path_positif}")
print(f"Le fichier des combinaisons manquantes a été sauvegardé à l'emplacement : {output_path_missing}")
print(f"Le fichier regroupé des statuts négatifs a été sauvegardé à l'emplacement : {output_path_negatif_grouped}")
