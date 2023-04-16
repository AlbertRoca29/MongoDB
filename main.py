import pandas as pd
from pymongo import MongoClient
import json

# Lectura de les dades a l'Excel
dades_cp = pd.read_excel('Dades.xlsx', sheet_name='Colleccions-Publicacions')
dades_p = pd.read_excel('Dades.xlsx', sheet_name='Personatges')
dades_a = pd.read_excel('Dades.xlsx', sheet_name='Artistes')

# 4 coleccions
dades_editorials = dades_cp[['NomEditorial','resposable','adreca','pais']].drop_duplicates().reset_index(drop=True).to_dict('records')

dades_artistes = dades_a.to_dict('records')

dades_colleccions = dades_cp[['NomColleccio','genere','idioma','any_inici',
                              'any_fi','tancada']].drop_duplicates().reset_index(drop=True).to_dict('records')

dades_publicacions = dades_cp[['ISBN','titol','stock','autor','preu', 'num_pagines', 'guionistes', 'dibuixants', 'NomColleccio', 'NomEditorial']].to_dict('records')



# Afegir personatges a publicacions

personatges = dades_p.to_dict('records')

for publicacio in dades_publicacions:
    isbn = publicacio['ISBN']
    personatges_publicacio = [{k: v for k, v in char.items() if k != 'isbn'} for char in personatges if char['isbn'] == isbn]
    if (len(personatges_publicacio)) :
        publicacio['personatges'] = personatges_publicacio
        
        
        
# Llistes
for dades in [dades_publicacions,dades_colleccions]:
    for doc in dades:
        for key, value in doc.items():
            if isinstance(value, str) and '[' in value and ']' in value:
                doc[key] = value.strip('][').split(', ')

# filtrar any_fi
for i,doc in enumerate(dades_colleccions):
    if not doc['tancada']:
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
