import requests
import psycopg2
import pandas as pd
import os
import time
import ctypes

if os.path.exists( # valida se o Arquivo Existe  na mesma pasta que o que o fonte
        "CNPJ.xlsx"):

    arquivo_cnpj = pd.read_excel("CNPJ.xlsx", header=0)  # Le o arquivo excel que contém os cnpj

else:
    ctypes.windll.user32.MessageBoxW(0, "Arquivo excel com Cnpjs Não encontrato!"
                                     , "Erro!", 0)
    exit()
cont = 0  # Contador

for i, row in enumerate(arquivo_cnpj.values): # estrutura de repetição para percorrer todos os cnpjs listados

    cnpj = str((arquivo_cnpj["CNPJ"][cont])) # cacputura o valor da linha e coluna CNPJ
    cnpj = str(cnpj.replace('.','').replace('-','').replace('/','')) # Realiza tratamento caso Usuário tenha usado caracteries
    print(cnpj)
    r = requests.get('https://www.receitaws.com.br/v1/cnpj/{}'.format(cnpj)) # realiza requisição para API
    r.text # Capitura texto retorna pela API

    while r.text == 'Too many requests, please try again later.': # Valida casos que sobre carga na API onde requer consultar mais tarde
        time.sleep(10) # espera 10 seguntos para realizar nova consulta
        r.close() # fecha solicitação antiga
        r = requests.get('https://www.receitaws.com.br/v1/cnpj/{}'.format(cnpj)) # Realiza Nova consulta ate que retorne arquivo JSON
        r.text # capitura o retono da API

    arquivo = r.json() # VAriavel que armazena resulta em Json da API

    if 'message' not in arquivo: # Valida casos onde a API retona uma mensagem de erro
        nome =format(arquivo['nome'])
        Cnpj = format(arquivo['cnpj']).replace('.','').replace('-','').replace('/','')
        municipio = format(arquivo['municipio'])
        UF = format(arquivo['uf'])

        con = psycopg2.connect(host='localhost', database='alfa_transportes',
        user='postgres', password='31082') # conexão com banco

        cursor = con.cursor() # cria Cursor

        sql = (
            "select * from CAD_FILIAL where cnpj = '" + Cnpj + "'"
        )
        cursor.execute(sql) # consulta se o Cnpj existe na tabela
        result = cursor.fetchone()

        if result == None : # caso não exista realiza o Insert
            insert = (
                "INSERT INTO CAD_FILIAL ("
                    "cnpj"
                    ", nome"
                    ", cidade"
                    ", estado"
                        ") " 
                    "VALUES ('"
                       + Cnpj +
                        "' , '" + nome +
                        "' , '" + municipio +
                        "' , '" + UF +
                            "')"
                      )

            cursor.execute(insert)
            con.commit()
            con.close()
        else: # caso exista realiza o Update das informações

            update = (
                "UPDATE CAD_FILIAL "
                "SET NOME = '" + nome +
                    "' , cidade = '" + municipio +
                    "' , estado = '" + UF +
                    "' WHERE CODIGO_FIL = '" + str(result[0]) + "'"
            )
            cursor.execute(update)
            con.commit()
            con.close()

    else: # caso de Retorno Com erro Aprenseta na tela o erro e qual foi  Cnpj que apresentou o erro
        retorno = format(arquivo['message'])
        ctypes.windll.user32.MessageBoxW(0, " O Retorno da consulta para o CNPJ " +
                                         cnpj + " Foi; --> " + retorno , "Erro!", 0)
    cont = cont + 1 # alimenta contador
    r.close() # Fecha requisão para API

ctypes.windll.user32.MessageBoxW(0, "Registos Inseridos Com Sucesso!"
                                    , "OK", 0)