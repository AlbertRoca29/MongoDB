import pandas as pd
from pymongo import MongoClient

dades_cp = []
dades_p = []
dades_a = []

# Lectura de les dades a l'Excel
dades_cp = pd.read_excel('Dades.xlsx', sheet_name='Colleccions-Publicacions')
dades_p = pd.read_excel('Dades.xlsx', sheet_name='Personatges')
dades_a = pd.read_excel('Dades.xlsx', sheet_name='Artistes')

# TODO


# En execucio remota
Host = 'localhost' 
Port = 27017

# Connexio
DSN = "mongodb://{}:{}".format(Host,Port)
conn = MongoClient(DSN)

# Seleccio de la base de dades
db = conn['llibreria']

# Insercio de les dades a MongoDB
db['editorials'].insert_many(editorials_data)
db['artistes'].insert_many(artistes_data)
db['colleccions'].insert_many(colleccions_data)
db['publicacions'].insert_many(publicacions_data)
db['personatges'].insert_many(personatges_data)

conn.close()