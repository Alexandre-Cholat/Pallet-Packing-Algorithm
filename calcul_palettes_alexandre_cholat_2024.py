# ----------------------------------------
# Programme : Estimation de Palettes avec py3dbp
# Auteur : Alexandre Cholat
# Date : 2024
# ----------------------------------------
# Ce programme utilise la bibliothèque py3dbp pour creer des palettes et visualiser
# la modelisation en 3D. Il lit les données des articles à partir d'un fichier Excel et génère une visualisation
# proportionnelle des palettes emballées. La fonction calculate_pallets prend en parametere un string, correspondant au code de une commande,
# et renvoie le nombre total, un entier, de palettes nécessaires pour cette commande. Noubliez pas de fermer les fenetres GUI pour poursuivre lexecution de la fonction.
#
# La bibliothèque py3dbp est utilisée pour résoudre le problème de 'bin packing', c'est-à-dire configurer des articles dans des conteneurs ('bins', ou dans notre cas, des palettes) 
# de manière optimale. L'algorithme est base sur cette publication : https://github.com/enzoruiz/3dbinpacking/blob/master/erick_dube_507-034.pdf
#
# Packer, Bin, et Item sont les classes de la bibliothèque py3dbp. Une fois un packer cree, nous ajoutons des bins et des items, 
# puis nous appelons la methode pack() pour effectuer la configuration.
#
# Installation des bibliothèques nécessaires :
# - pip install py3dbp
# - pip install pandas
# - pip install matplotlib
#
# Documentation :
# - py3dbp : https://pypi.org/project/py3dbp/
# - pandas : https://pandas.pydata.org/pandas-docs/stable/index.html
# - matplotlib : https://matplotlib.org/stable/contents.html
# ----------------------------------------

import pandas as pd
from py3dbp import Packer, Bin, Item
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D


def fetch_items_from_excel(command_code, excel_file=r'C:\Users\alexa\OneDrive\Desktop\Neolait_2024\Composition palettes.xlsx'): #!! remplacez la valeur excel_file par le chemin de votre fichier Excel

# Fonction pour lire un fichier Excel et créer une liste d'articles basée sur le code Commande


    # Lire le fichier Excel
    df = pd.read_excel(excel_file)

    # Débogage : Afficher les noms des colonnes
    print("Noms des colonnes :", df.columns)
    #print("Types de données :", df.dtypes)

    # Supprimer les espaces et convertir en chaîne de caractères pour la comparaison
    df['Commande'] = df['Commande'].astype(str).str.strip()

    # Débogage : Afficher les valeurs uniques de Commande
    print("Valeurs uniques de Commande :", df['Commande'].unique())

    # Filtrer les lignes en fonction du code Commande
    filtered_df = df[df['Commande'] == str(command_code)]

    # Débogage : Afficher les lignes filtrées
    print(f"Lignes filtrées pour Commande {command_code} :", filtered_df)

    items = []
    # Boucler à travers les lignes filtrées et créer des articles
    for _, row in filtered_df.iterrows():
        name = row['Nom produit']
        weight = row['Poids unitaire']
        quantity = row['Quantité']
        height = row['Hauteur']
        width = row['Largeur']
        depth = row['Longueur']

        # Créer un article pour chaque quantité commandée
        for _ in range(int(quantity)):
            items.append(Item(name, width, height, depth, weight))
    
    return items

def calculate_pallets(command_code):

    # Définir les dimensions de la palette et la capacité de poids
    bin_width = 100  # en cm
    bin_height = 120  # en cm
    bin_depth = 180  # en cm
    bin_max_weight = 1100  # en kg

    # Récupérer les articles à partir du fichier Excel en fonction du code Commande
    items = fetch_items_from_excel(command_code)

    if not items:
        print(f"Aucun article trouvé pour le code Commande : {command_code}")
        return None

    remaining_items = items[:]
    num_pallets = 0

    while remaining_items:
        num_pallets += 1
        packer = Packer()  # Réinitialiser le Packer à chaque itération
        current_bin = Bin(f'Bin{num_pallets}', bin_width, bin_height, bin_depth, bin_max_weight)
        packer.add_bin(current_bin)

        # Ajouter les articles restants au Packer
        for item in remaining_items:
            packer.add_item(item)

        # Effectuer l'emballage
        packer.pack()

        # Visualiser les palettes emballées - L'interface graphique doit être supprimée avant que le programme continue
        visualize_pallets(packer)

        # Afficher les articles emballés pour la palette actuelle
        print(f"\nArticles emballés dans la palette {num_pallets} :")
        packed_items = []
        for bin in packer.bins:
            if bin.items:
                for item in bin.items:
                    print(f" Article : {item.name} à la position {item.position} avec les dimensions {item.get_dimension()}")
                    packed_items.append(item)

        # Déterminer les articles non emballés pour la prochaine palette
        remaining_items = [item for item in remaining_items if item not in packed_items]
        print(f"\nArticles non emballés après la palette {num_pallets} : {len(remaining_items)}")
        for item in remaining_items:
            print(f"  {item.string()}")

    print(f"\nNombre total de palettes nécessaires : {num_pallets}")
    
    return num_pallets

def get_random_color():
    # Fonction pour générer une couleur aléatoire

    return np.random.rand(3,)

def add_box(ax, item, color):
# Fonction pour ajouter une boîte 3D (représentant un article)


    # Extraire la position et les dimensions
    pos = np.array(item.position, dtype=float)
    dim = np.array([item.width, item.height, item.depth], dtype=float)

    # Créer un prisme rectangulaire (boîte)
    xx, yy = np.meshgrid([pos[0], pos[0] + dim[0]], [pos[1], pos[1] + dim[1]])
    ax.plot_surface(xx, yy, np.full_like(xx, pos[2]), color=color, alpha=0.5)
    ax.plot_surface(xx, yy, np.full_like(xx, pos[2] + dim[2]), color=color, alpha=0.5)

    yy, zz = np.meshgrid([pos[1], pos[1] + dim[1]], [pos[2], pos[2] + dim[2]])
    ax.plot_surface(np.full_like(yy, pos[0]), yy, zz, color=color, alpha=0.5)
    ax.plot_surface(np.full_like(yy, pos[0] + dim[0]), yy, zz, color=color, alpha=0.5)

    xx, zz = np.meshgrid([pos[0], pos[0] + dim[0]], [pos[2], pos[2] + dim[2]])
    ax.plot_surface(xx, np.full_like(xx, pos[1]), zz, color=color, alpha=0.5)
    ax.plot_surface(xx, np.full_like(xx, pos[1] + dim[1]), zz, color=color, alpha=0.5)

def visualize_pallets(packer):
# Fonction pour visualiser les palettes


    palette_count = 0
    for bin_container in packer.bins:
        if bin_container.items:
            palette_count += 1
            fig = plt.figure(figsize=(10, 8))
            ax = fig.add_subplot(111, projection='3d')

            color_mapping = {}

            # Ajouter chaque article dans la palette au graphique
            for item in bin_container.items:
                color = get_random_color()
                add_box(ax, item, color)
                color_mapping[item.name] = color  # Stocker l'association des couleurs

            # Créer une légende pour les articles
            legend_labels = [plt.Line2D([0], [0], color=color, lw=4, label=name) for name, color in color_mapping.items()]
            plt.legend(handles=legend_labels, loc='upper center', bbox_to_anchor=(0.5, -0.05), fancybox=True, shadow=True, ncol=5)

            # Définir les limites pour correspondre à la taille de la palette
            ax.set_xlim([0, 100])
            ax.set_ylim([0, 120])
            ax.set_zlim([0, 180])

            # Définir un rapport d'aspect égal pour les axes
            ax.set_box_aspect([100, 120, 180])  # Rapport d'aspect : x:y:z

            # Étiquettes et titre
            ax.set_xlabel('Largeur (cm)')
            ax.set_ylabel('Profondeur (cm)')
            ax.set_zlabel('Hauteur (cm)')
            plt.title(f'Visualisation 3D de {bin_container.name}')

            plt.show()

# Exécution du programme
calculate_pallets('1002128')
