import pandas as pd
import numpy as np
import random

# Permet de lire tout les feuilles d'un fichier Excel et de les mettre dans un dictionnaire
file = pd.ExcelFile(r'DataSet.xlsx')
dict_project = {sheet_name: file.parse(sheet_name) for sheet_name in file.sheet_names}


# Pour tout les classements de sous-type : mur, retour, coin, plafond
# En ordre (panneau 1, 2, 3, ... ,n-1, n)
def heuristic_ordre(df):
    df_ordered = df.sort_values('Id')
    return df_ordered


# Par type (plafonds, murs, murs retours ...)
def heuristic_type(df):
    df['Type'] = pd.Categorical(df['Type'], ['Mur', 'Retour', 'Coin', 'Plafond'])
    df_ordered = df.sort_values(['Type', 'Secteur'])
    return df_ordered


# Par secteur aléatoire
def heuristic_sector_random(df):
    list_int = list(range(1, df.shape[0] + 1))
    random.shuffle(list_int)
    df['random'] = list_int
    df_ordered = df.sort_values(['Secteur', 'random'])
    return df_ordered


# Par secteur en ordre
def heuristic_sector(df):
    df_ordered = df.sort_values(['Secteur', 'Id'])
    return df_ordered


# Par secteur par type (classé par type au sein d'un même secteur)
def heuristic_sector_type(df):
    df['Type'] = pd.Categorical(df['Type'], ['Mur', 'Retour', 'Coin', 'Plafond'])
    df_ordered = df.sort_values(['Secteur', 'Type'])
    return df_ordered


# Par secteur par temps (ordre décroissant par 7 panneaux, plus long au plus court)
def heuristic_sector_time(df):
    df_1 = heuristic_sector_random(df)
    df_2 = df_1.reset_index()
    df_2['index_reset'] = df_2.index
    df_2['group'] = df_2['index_reset'].apply(lambda x: int(x/7) + 1)
    df_ordered = df_2.sort_values(['Secteur', 'group', 'Soudure'], ascending=[True, True, False])
    # df_ordered2 = df_ordered.drop(['group', 'index_reset'], axis=1)
    return df_ordered

# Par secteur par temps (minimiser moyenne mobile sur 7 panneaux)
# En attente, requiert un modèle d'optimisation


for project, df in dict_project.items():
    heuristic_ordre(df).to_csv(r'results\heuristic_ordre_{}'.format(project))
    heuristic_type(df).to_csv(r'results\heuristic_type_{}'.format(project))
    heuristic_sector_random(df).to_csv(r'results\heuristic_sector_random_{}'.format(project))
    heuristic_sector(df).to_csv(r'results\heuristic_sector_{}'.format(project))
    heuristic_sector_type(df).to_csv(r'results\heuristic_sector_type_{}'.format(project))
    heuristic_sector_time(df).to_csv(r'results\heuristic_sector_time_{}'.format(project))
