import pandas as pd
from datetime import datetime, date, time
import os
import math
import copy

#Função para carregar a planilha pré construída com dados de cargas
def carregar_dados(file,folder): 
    df_exante = pd.read_excel(f'{folder}/{file}',sheet_name=0)
    df_expost = pd.read_excel(f'{folder}/{file}',sheet_name=1)
    os.remove(f'{os.getcwd()}/{folder}/{file}')
    dados_dict_ante = {}
    dados_dict_post = {}
    for col in df_exante.columns:
        dados_dict_ante[col] = df_exante[col].to_list()
    for col in df_expost.columns:
        dados_dict_post[col] = df_expost[col].to_list()
    dados_dict = [dados_dict_ante,dados_dict_post]
    return dados_dict
#--------------------------------------------------------------------------------------------------------

def calculos(dados_dict):
    consumo_ante = calc_consumo(dados_dict=dados_dict[0])
    consumo_post = calc_consumo(dados_dict=dados_dict[1])
    results_dict = {
        'Data': dados_dict[0]['Data'],
        'Hora': dados_dict[0]['Hora'],
        'Consumo': consumo_ante,
        'Consumo - Ex Post': consumo_post
    }
    return results_dict


def calc_consumo(dados_dict): #Preciso testar se tá registrando corretamente os valores após o horário de ponta
    i=0
    pot=[]
    while i<len(dados_dict['V1']):
        pot.append((dados_dict['V1'][i]*dados_dict['I1'][i]*dados_dict['FP1'][i])+(dados_dict['V2'][i]*dados_dict['I2'][i]*dados_dict['FP2'][i])+(dados_dict['V3'][i]*dados_dict['I3'][i]*dados_dict['FP3'][i]))
        i+=1
    
    delta_t = str(datetime.combine(date.today(),dados_dict['Hora'][1]) - datetime.combine(date.today(),dados_dict['Hora'][0]))
    delta_t = float(delta_t.split(':')[0])*60+float(delta_t.split(':')[1])+float(delta_t.split(':')[2])/60
    t0 = dados_dict['Hora'][0]

    energia_t = []
    energia_fp = []
    energia_p = []
    energia_d = []
    reset = True
    fp=0
    i=0
    for p in pot:
        if i == 0: 
            energia_t.append(p/60)
            energia_fp.append(p/60)
            energia_p.append(p/60)
            energia_d.append(p/60)
            i+=1
            continue
        else: 
            energia_t.append((energia_t[i-1] + p/60000))
            energia_d.append((energia_d[i-1] + p/60000))

        if dados_dict['Hora'][i].hour >= 18 and dados_dict['Hora'][i].hour < 18+3 and dados_dict['Data'][i].weekday() not in [6,7]:   #Colocar a opção do horário de ponta ser input do usuário
            if fp == 0:
                fp=i-1
            reset = False
            energia_p.append((energia_p[i-1]+p/60000))
            energia_fp.append(0)
        else:
            energia_p.append(0)
            if dados_dict['Hora'][i].hour < 18 and dados_dict['Hora'][i] != t0 and reset:
                energia_fp.append((energia_fp[i-1]+p/60000))
            elif dados_dict['Hora'][i] == t0:
                energia_fp.append(0)
                energia_d[i] = 0
            else:
                energia_fp.append((energia_fp[fp]+p/60000))
                fp = 0
                reset = True
        i+=1
    dias = list(dict.fromkeys(dados_dict['Data']))
    valores_dict = {
        'Data': dados_dict['Data'],
        'Hora': dados_dict['Hora'],
        'Dias': dias,
        'Potência': pot,
        'Energia Total': energia_t,
        'Energia - Dia': energia_d,
        'Energia - FP': energia_fp,
        'Energia - P': energia_p
    }
    return valores_dict
    