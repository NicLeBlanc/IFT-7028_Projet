import pandas as pd
import numpy as np
import random

# Permet de lire tout les feuilles d'un fichier Excel et de les mettre dans un dictionnaire
file = pd.ExcelFile(r'DataSet.xlsx')
dict_project = {sheet_name: file.parse(sheet_name) for sheet_name in file.sheet_names}

df_palettisation = dict_project['C-3312']
print(df_palettisation)

liste_mur = df_palettisation[['Type']].to_numpy()

print(liste_mur)

liste_palette = []

compteur_palette = 0
compteur_palette_actuelle_mur = 0
compteur_palette_actuelle_plafond = 0
compteur_palette_mur = 0
compteur_palette_plafond = 0

for panneau in liste_mur:

    if panneau == 'Mur' or panneau == 'Coin' or panneau == 'Retour':

        if compteur_palette_mur == 0:
            compteur_palette_actuelle_mur = compteur_palette + 1
            compteur_palette += 1

        if compteur_palette_mur < 7:
            liste_palette.append(compteur_palette_actuelle_mur)
            compteur_palette_mur += 1

        else:
            compteur_palette_mur = 1
            compteur_palette_actuelle_mur = compteur_palette + 1
            liste_palette.append(compteur_palette_actuelle_mur)
            compteur_palette += 1

    if panneau == 'Plafond':

        if compteur_palette_plafond == 0:
            compteur_palette_actuelle_plafond = compteur_palette + 1
            compteur_palette += 1

        if compteur_palette_plafond < 7:
            liste_palette.append(compteur_palette_actuelle_plafond)
            compteur_palette_plafond += 1

        else:
            compteur_palette_plafond = 1
            compteur_palette_actuelle_plafond = compteur_palette + 1
            liste_palette.append(compteur_palette_actuelle_plafond)
            compteur_palette += 1

print(liste_palette)
