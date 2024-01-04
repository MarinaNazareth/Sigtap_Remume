
#Importações geral (tem mais bibliotecas do que serão utilizadas por reaproveitamento do código)
import pandas as pd
import math
import numpy as np
from datetime import datetime, timedelta, date
from IPython.display import HTML
import mysql.connector
import requests
import hashlib
import os
from unidecode import unidecode
import re as regex
from selenium import webdriver
from selenium.webdriver.common.by import By
import requests 
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import xlrd

df_key1 = pd.read_excel(r"C:\Users\Downloads\Remume-automatizado\LISTA.xlsx", sheet_name ='BA')
df_key2 = pd.read_excel(r"C:\Users\Downloads\Remume-automatizado\LISTA-902.xlsx", sheet_name ='BH')
df_regex = pd.read_excel(r"C:\Users\Downloads\Remume-automatizado\regex.xlsx")


df_key1['concat'] = df_key2['concat'].str.upper()
df_key2['concat'] = df_key2['concat'].str.upper()

df_key1['concat'] = df_key1['concat'].str.replace('/',' ', regex=False)
df_key2['concat'] = df_key2['concat'].str.replace('/',' ', regex=False)
df_key1['concat'] = df_key1['concat'].str.replace(',','.', regex=False)
df_key2['concat'] = df_key2['concat'].str.replace(',','.', regex=False)
df_regex['concat'] = df_regex['concat'].str.upper()
df_regex['concat'] = df_regex['concat'].str.replace('  ',' ', regex=False)
df_regex['concat'] = df_regex['concat'].str.replace(',','.', regex=False)

#função para retirar acentos
def remover_acentos(texto):
    return unidecode(texto)

df_regex['concat'] = df_regex['concat'].apply(remover_acentos)
df_key1['concat'] = df_key1['concat'].apply(remover_acentos)
df_key2['concat'] = df_key2['concat'].apply(remover_acentos)

#definicao dos caracteres especiais que devem ser escapados do regex

def escape_special_chars(string):
    special_chars = ['\\', '^', '$', '|', '?', '*', '(', ')', '[', ']', '{', '}', ':']

    for char in special_chars:
        string = string.replace(char, '\\' + char)
    
    return string

df_regex['concat'] = df_regex['concat'].apply(escape_special_chars)
df_key1['concat'] = df_key1['concat'].apply(escape_special_chars)
df_key2['concat'] = df_key2['concat'].apply(escape_special_chars)


#tentar correspondencia
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def encontrar_correspondencia_similaridade(texto, opcoes):
    melhor_correspondencia = process.extractOne(texto, opcoes)
    return melhor_correspondencia

def id_correspondente(row):
    texto = row['concat']
    opcoes = df_regex['concat'].tolist()
    
    melhor_correspondencia = encontrar_correspondencia_similaridade(texto, opcoes)
    
    if melhor_correspondencia[1] >= 90:  # similaridade aceitável de 90%
        ids = df_regex.loc[df_regex['concat'] == melhor_correspondencia[0], 'id'].tolist()
        return ids
    return None

def remover_duplicatas(lista):
    if lista is None:
        return []
    return list(set(lista))


df_key1['id_correspondente'] = df_key1.apply(id_correspondente, axis=1)
df_key1['id_correspondente'] = df_key1['id_correspondente'].apply(remover_duplicatas)

df_key2['id_correspondente'] = df_key2.apply(id_correspondente, axis=1)
df_key2['id_correspondente'] = df_key2['id_correspondente'].apply(remover_duplicatas)

writer = pd.ExcelWriter(r"C:\Users\Downloads\Remume-automatizado\remume_v1.xlsx", engine='xlsxwriter')
df_key1.to_excel(writer, sheet_name='BA')
df_key3.to_excel(writer, sheet_name='BH')