import pandas as pd
from pymongo import MongoClient

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

# TODO
ll = []
for row in dades_colleccions.itertuples():
    ll.append(row)
print(ll)

"""
# En execucio remota
Host = 'localhost' 
Port = 27017

# Connexio
DSN = "mongodb://{}:{}".format(Host,Port)
conn = MongoClient(DSN)

# Seleccio de la base de dades
db = conn['llibreria']

# Insercio de les dades a MongoDB
db['editorials'].insert_many(dades_editorials)
db['artistes'].insert_many(artistes_data)
db['colleccions'].insert_many(colleccions_data)
db['publicacions'].insert_many(publicacions_data)

conn.close()
"""