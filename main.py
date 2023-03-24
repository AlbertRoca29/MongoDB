import pandas as pd
from pymongo import MongoClient
import json

dades_cp = []
dades_p = []
dades_a = []

# Lectura de les dades a l'Excel
dades_cp = pd.read_excel('Dades.xlsx', sheet_name='Colleccions-Publicacions')
dades_p = pd.read_excel('Dades.xlsx', sheet_name='Personatges')
dades_a = pd.read_excel('Dades.xlsx', sheet_name='Artistes')

dades_editorials = dades_cp[['NomEditorial','resposable','adreca','pais']].drop_duplicates().reset_index(drop=True)
dades_artistes = dades_a.drop_duplicates().reset_index(drop=True)
dades_colleccions = dades_cp[['NomColleccio','genere','idioma','any_inici',
                              'any_fi','tancada']].drop_duplicates().reset_index(drop=True)
dades_publicacions = dades_cp[['ISBN',	'titol','stock','autor','preu',	'num_pagines','guionistes','dibuixants']]

df_extra = pd.merge(dades_publicacions, dades_p, left_on='ISBN', right_on='isbn')
df_extra['personatges'] = df_extra[dades_p.columns[:2]].apply(lambda x: json.dumps(dict(x)), axis=1)
list_p = df_extra.groupby('ISBN')['personatges'].apply(list)
dades_publicacions = dades_publicacions.merge(list_p, on='ISBN')
dades_publicacions.head()
