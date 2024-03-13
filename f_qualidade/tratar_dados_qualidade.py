import pandas as pd
import os
import math
import copy

#Função para carregar a planilha pré construída com dados de cargas
def carregar_dados(file,folder): 
    df = pd.read_excel(f'{folder}/{file}',sheet_name=0)
    os.remove(f'{os.getcwd()}/{folder}/{file}')
    dados_dict = {}
    for col in df.columns:
        dados_dict[col] = df[col].to_list()
    return dados_dict
#--------------------------------------------------------------------------------------------------------

def calculos(dados_dict):
    tensoes = calc_tensoes(dados_dict=dados_dict)
    fp = calc_fp(dados_dict=dados_dict)
    correntes = calc_correntes(dados_dict=dados_dict)
    consumo = calc_consumo(dados_dict=dados_dict)
    fd = calc_fd(dados_dict=dados_dict)
    thd = calc_thd(dados_dict=dados_dict)
    
    results_dict = {
        'Tensões': tensoes,
        'FP': fp,
        'Correntes': correntes,
        'Consumo': consumo,
        'FD': fd,
        'THD': thd
    }
    return results_dict

def calc_tensoes(dados_dict):
    tr = 220
    t_min = [min(dados_dict['V1']),min(dados_dict['V2']),min(dados_dict['V3'])]
    desv_inf = [(1-(t_min[0]/tr))*100,(1-(t_min[1]/tr))*100,(1-(t_min[2]/tr))*100]
    t_max = [max(dados_dict['V1']),max(dados_dict['V2']),max(dados_dict['V3'])]
    desv_sup = [((t_max[0]/tr)-1)*100,((t_max[1]/tr)-1)*100,((t_max[2]/tr)-1)*100]

    valores_dict = {
        'Data': dados_dict['Data'],
        'Hora': dados_dict['Hora'],
        'Ref': tr,
        'Valor inf': t_min,
        'Desvio % inf': desv_inf,
        'Valor sup': t_max,
        'Desvio % sup': desv_sup
    }
    return valores_dict
    

def calc_fp(dados_dict):            #Talvez seja interessante adicionar uma opção para definir se quer ou não as linhas com 0
    new_dados = copy.deepcopy(dados_dict)
    fp1_list = new_dados['FP1']
    fp2_list = new_dados['FP2']
    fp3_list = new_dados['FP3']
    fpt_list = new_dados['FP Total']
    hora_fp = new_dados['Hora']
    data_fp = new_dados['Data']

    while True:
        if 0 in fpt_list:
            idx = fpt_list.index(0)
        elif 0 in fp1_list:
            idx = fp1_list.index(0)
        elif 0 in fp2_list:
            idx = fp2_list.index(0)
        elif 0 in fp3_list:
            idx = fp3_list.index(0)
        else:
            break
        hora_fp.pop(idx)
        data_fp.pop(idx)
        fpt_list.pop(idx)
        fp1_list.pop(idx)
        fp2_list.pop(idx)
        fp3_list.pop(idx)
    
    valores_dict = {
        'Data': data_fp,
        'Hora': hora_fp,
        'Valores': ['Médio','Mínimo','Máximo'],
        'FP1': fp1_list,
        'FP2': fp2_list,
        'FP3': fp3_list,
        'Fase 1': [sum(fp1_list)/len(fp1_list),min(fp1_list),max(fp1_list)],
        'Fase 2': [sum(fp2_list)/len(fp2_list),min(fp2_list),max(fp2_list)],
        'Fase 3': [sum(fp3_list)/len(fp3_list),min(fp3_list),max(fp3_list)],
        'Total': [sum(fpt_list)/len(fpt_list),min(fpt_list),max(fpt_list)]
    }
    return valores_dict

def calc_correntes(dados_dict):
    i1_list = dados_dict['I1']
    i2_list = dados_dict['I2']
    i3_list = dados_dict['I3']
    in_list = dados_dict['In']

    i1_list2 = i1_list.copy()
    i2_list2 = i2_list.copy()
    i3_list2 = i3_list.copy()
    in_list2 = in_list.copy()
    while True:
        if 0 in i1_list2:
            i1_list2.remove(0)
        elif 0 in i2_list2:
            i2_list2.remove(0)
        elif 0 in i3_list2:
            i3_list2.remove(0)
        elif 0 in in_list2:
            in_list2.remove(0)
        else:
            break

    valores_dict = {
        'Data': dados_dict['Data'],
        'Hora': dados_dict['Hora'],
        'Valores': ['Médio','Máximo','Mínimo'],
        'Fase 1': [sum(i1_list2)/len(i1_list2),max(i1_list),min(i1_list)],
        'Fase 2': [sum(i2_list2)/len(i2_list2),max(i2_list),min(i2_list)],
        'Fase 3': [sum(i3_list2)/len(i3_list2),max(i3_list),min(i3_list)],
        'Neutro': [sum(in_list2)/len(in_list2),max(in_list),min(in_list)],
        'I1': i1_list,
        'I2': i2_list,
        'I3': i3_list,
        'In': in_list
    }
    return valores_dict


def calc_consumo(dados_dict): #DEPOIS TENHO QUE COLOCAR A O PREENCHIMENTO DO HORÁRIO DE PONTA E FORA DE PONTA
    dias = [0]
    consumo = []
    pot = dados_dict['PT Calc.']
    energia = dados_dict['Energia Calc.']

    d = dados_dict['Data'][0]
    i=0
    j=0
    for dia in dados_dict['Data']:
        if dia != d:
            d = dia
            dias.append(0)
            consumo.append(dados_dict['E. dia Total'][j-1])
            i+=1
        else:
            dias[i] = d
        j+=1
    consumo.append(dados_dict['E. dia Total'][j-1])

    valores_dict = {
        'Data': dados_dict['Data'],
        'Hora': dados_dict['Hora'],
        'Dias': dias,
        'Consumo': consumo,
        'Potência': pot,
        'Energia': energia
    }
    return valores_dict
    
def calc_fd(dados_dict):        
    vab = dados_dict['U1']
    vbc = dados_dict['U2']
    vca = dados_dict['U3']
    beta = []
    fd = []
    lim = []

    i=0
    while i <len(dados_dict['U1']):
        beta.append((math.pow(vab[i],4)+math.pow(vbc[i],4)+math.pow(vca[i],4))/math.pow((math.pow(vab[i],2)+math.pow(vbc[i],2)+math.pow(vca[i],2)),2))
        e = pow(10,-13)
        if (3 - 6*beta[i]) > 1 and (3 - 6*beta[i]-e) <= 1:
            beta[i] = 1/3
        fd.append(100*math.sqrt((1-math.sqrt(3-6*beta[i]))/(1+math.sqrt(3-6*beta[i]))))
        lim.append(2)
        i+=1
    
    valores_dict = {
        'Data': dados_dict['Data'],
        'Hora': dados_dict['Hora'],
        'Valores': ['Médio','Mínimo','Máximo'],
        'FD [%]': [sum(fd)/len(fd),min(fd),max(fd)],
        'FD': fd,
        'Limite': lim
    }
    return valores_dict

def calc_thd(dados_dict):
    thd_v1 = dados_dict['thd_V1 [%]']
    thd_v2 = dados_dict['thd_V2 [%]']
    thd_v3 = dados_dict['thd_V3 [%]']

    thd_i1 = dados_dict['thd_I1 [%]']
    thd_i2 = dados_dict['thd_I2 [%]']
    thd_i3 = dados_dict['thd_I3 [%]']

    valores_dict = {
        'Data': dados_dict['Data'],
        'Hora': dados_dict['Hora'],
        'THD V1': thd_v1,
        'THD V2': thd_v2,
        'THD V3': thd_v3,
        'THD V': [sum([sum(thd_v1)/len(thd_v1),sum(thd_v2)/len(thd_v2),sum(thd_v3)/len(thd_v3)])/3,max([max(thd_v1),max(thd_v2),max(thd_v3)])],
        'THD I1': thd_i1,
        'THD I2': thd_i2,
        'THD I3': thd_i3,
        'THD I': [sum([sum(thd_i1)/len(thd_i1),sum(thd_i2)/len(thd_i2),sum(thd_i3)/len(thd_i3)])/3,max([max(thd_i1),max(thd_i2),max(thd_i3)])]
    }
    return valores_dict

#Função para verificar se os dados foram inseridos corretamente
def verificar_save(dados_dict):
    if 'V1' not in dados_dict:
        return 'Carregue um arquivo de dados'
    
    return 'Arquivo salvo com sucesso'
#--------------------------------------------------------------------------------------------------------