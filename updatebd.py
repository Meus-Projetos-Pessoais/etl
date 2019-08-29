#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import pyodbc
df =  pd.read_excel('baseDeDados.xlsx')
df.head()
print(pyodbc.version)
#https://stackoverflow.com/questions/36180703/pyodbc-error-python-to-ms-access?noredirect=1&lq=1
conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\\Users\\Andre\\Documents\\bancoteste.accdb;'
cnxn =  pyodbc.connect(conn_str)
crsr = cnxn.cursor()
table_name =  "cadastroUsuario"
conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\\Users\\Andre\\Documents\\bancoteste.accdb;')
cursor = conn.cursor()
#cursor.execute("INSERT INTO cadastroUsuario (Nome,email,address,phonenumber,coments) VALUES ('Andre Emidio','andre.emidio@outlook.com' ,'Teste' ,'(12)98862-4725' ,'Entre em contato !')")
conn.commit()
cursor.execute('select * from cadastroUsuario')
for row in cursor.fetchall():
    print (row)
for index, row in df.iterrows():
    #print(row)
    print(index)
    with conn.cursor() as crsr:
       crsr.execute("INSERT INTO cadastroUsuario(Nome,email,address,phonenumber,coments) VALUES(?,?,?,?,?)",row["Name"], row["email"], row['address'], row['PhoneNumber'],row['Comments'])
conn.commit()
cursor.close()
conn.close()



