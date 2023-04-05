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
                              'any_fi','tancada', 'NomEditorial']].drop_duplicates().reset_index(drop=True)

dades_publicacions = dades_cp[['ISBN','titol','stock','autor','preu', 'num_pagines', 'guionistes', 'dibuixants', 'NomColleccio']]

df_extra = pd.merge(dades_publicacions, dades_p, left_on='ISBN', right_on='isbn')
df_extra['personatges'] = df_extra[dades_p.columns[:2]].apply(lambda x: json.dumps(dict(x)), axis=1)
list_p = df_extra.groupby('ISBN')['personatges'].apply(list)
dades_publicacions = dades_publicacions.merge(list_p, on='ISBN')

# Passem els Pandas Dataframe a JSON
"""
dades_editorials = dades_editorials.to_json(force_ascii=False, orient='records')
dades_artistes = dades_artistes.to_json(force_ascii=False, orient='records')
dades_colleccions = dades_colleccions.to_json(force_ascii=False, orient='records')
dades_publicacions = dades_publicacions.to_json(force_ascii=False, orient='records')
"""
#dades_editorials = json.loads(dades_editorials).values()
#dades_artistes = json.loads(dades_artistes).values()
#dades_colleccions = json.loads(dades_colleccions).values()
#dades_publicacions = json.loads(dades_publicacions).values()

dades_editorials = dades_editorials.to_dict('records')
dades_artistes = dades_artistes.to_dict('records')
dades_colleccions = dades_colleccions.to_dict('records')
dades_publicacions = dades_publicacions.to_dict('records')

for i,publicacio in enumerate(dades_publicacions):
    new_publicacio = []
    #print(publicacio['personatges'])
    for posicio in publicacio['personatges']:
        posicio = eval(posicio)
        new_publicacio.append(posicio)
    #print(new_publicacio)
    
    dades_publicacions[i]['personatges'] = new_publicacio
    print(dades_publicacions[i]['personatges'])



# En execucio remota
Host = 'localhost' 
Port = 27017

# Connexio
DSN = "mongodb://{}:{}".format(Host,Port)
conn = MongoClient(DSN)

# Seleccio de la base de dades
db = conn['llibreria']

# Eliminem la base de dades anterior
conn.drop_database(db)

# Insercio de les dades a MongoDB
db['editorials'].insert_many(dades_editorials)
db['artistes'].insert_many(dades_artistes)
db['colleccions'].insert_many(dades_colleccions)
db['publicacions'].insert_many(dades_publicacions)

conn.close()
