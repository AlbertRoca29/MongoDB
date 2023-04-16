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

dades_publicacions = dades_cp[['ISBN','titol','stock','autor','preu', 'num_pagines', 'guionistes', 'dibuixants', 'NomColleccio', 'NomEditorial']]


# what.fillna('unknown', inplace=True)
df_extra = pd.merge(dades_publicacions, dades_p, left_on='ISBN', right_on='isbn',how='outer')
print(df_extra)
df_extra['personatges'] = df_extra[dades_p.columns[:2]].apply(lambda x: json.dumps(dict(x)), axis=1)
list_p = df_extra.groupby('ISBN')['personatges'].apply(list)
dades_publicacions = dades_publicacions.merge(list_p, on='ISBN')
# print(dades_publicacions)

dades_editorials = dades_editorials.to_dict('records')
dades_artistes = dades_artistes.to_dict('records')
dades_colleccions = dades_colleccions.to_dict('records')
dades_publicacions = dades_publicacions.to_dict('records')

# Diccionaris
for i,publicacio in enumerate(dades_publicacions):
    new_publicacio = []
    #print(publicacio['personatges'])
    for posicio in publicacio['personatges']:
        if(posicio[8:11]!='NaN'):
            posicio = eval(posicio)
            new_publicacio.append(posicio)

    if(len(new_publicacio)>0):
        dades_publicacions[i]['personatges'] = new_publicacio
    else: 
        del dades_publicacions[i]['personatges']
    
    
# Llistes
for doc in dades_publicacions:
    
    for key, value in doc.items():
        if isinstance(value, str) and '[' in value and ']' in value:
            doc[key] = value.strip('][').split(', ')

for doc in dades_colleccions:
    for key, value in doc.items():
        if isinstance(value, str) and '[' in value and ']' in value:
            doc[key] = value.strip('][').split(', ')

# filtrar any_fi
for i,doc in enumerate(dades_colleccions):
    if doc['tancada'] == False:
        del dades_colleccions[i]['any_fi']
    else:
        dades_colleccions[i]['any_fi'] = int(dades_colleccions[i]['any_fi'])


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