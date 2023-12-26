
import pyodbc
import pandas as pd 
import pymssql as sql
import warnings
import data as dt

server = '172.16.22.212' 
database = 'CIGAM' 
username = '' 
password = ''  
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

planilha = pd.read_excel('data1.xlsx')
for item, linha in planilha.iterrows():
    cursor.execute("insert into TABELA_TESTE(nome, idade, cpf, rg, data_nasc, sexo, email, cep, endereco, numero, bairro, cidade, estado, celular) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?); ", 
                   str(linha.nome), int(linha.idade), str(linha.cpf), str(linha.rg), str(linha.data_nasc), str(linha.sexo), str(linha.email), str(linha.cep), str(linha.endereco), int(linha.numero), str(linha.bairro), str(linha.cidade), str(linha.estado), str(linha.celular))
cnxn.commit()
cursor.close()


'''query = "SELECT * from TABELA_TESTE;"
df = pd.read_sql(query, cnxn)
print(df)'''


