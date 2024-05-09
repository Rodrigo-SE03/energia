import pandas as pd
from datetime import datetime, date, time
import inspect

class FlukeData():
    intervalo=0         #O intervalo é dado em segundos
    data_dict = {
       'data':[],
       'hora':[],
       'tensao_1n':[],
       'tensao_2n':[],
       'tensao_3n':[],
       'corrente_1':[],
       'corrente_2':[],
       'corrente_3':[],
       'corrente_n':[],
       'freq':[],
       'fp_1':[],
       'fp_2':[],
       'fp_3':[],
       'fp_t':[],
       'pot_1':[],
       'pot_2':[],
       'pot_3':[],
       'pot_t':[],
       'energia_t':[],
       'energia_dia':[],
       'energia_p':[],
    }

    def __init__(self,file,folder):
        self.ler_dados(file=file,folder=folder)
        self.corrigir_hora()
        self.preencher()
        self.calc_energia()
    
    def ler_dados(self,file,folder):
        df = pd.read_excel(f'{folder}/{file}')
        df_dict = {}
        for col in df.columns:
            df_dict[col] = df[col].to_list()
        
        self.data_dict['data']=df_dict['Data']
        self.data_dict['hora']=df_dict['Hora']
        self.data_dict['tensao_1n']=df_dict['Voltagem L1N Méd.']
        self.data_dict['tensao_2n']=df_dict['Voltagem L2N Méd.']
        self.data_dict['tensao_3n']=df_dict['Voltagem L3N Méd.']
        self.data_dict['corrente_1']=df_dict['Corrente L1 Méd.']
        self.data_dict['corrente_2']=df_dict['Corrente L2 Méd.']
        self.data_dict['corrente_3']=df_dict['Corrente L3 Méd.']
        self.data_dict['corrente_n']=df_dict['Corrente N Méd.']
        self.data_dict['freq']=df_dict['Freqüência Méd.']
        self.data_dict['fp_1']=df_dict['PF clássico L1N Méd.']
        self.data_dict['fp_2']=df_dict['PF clássico L2N Méd.']
        self.data_dict['fp_3']=df_dict['PF clássico L3N Méd.']
        self.data_dict['fp_t']=df_dict['PF clássico Total Méd.']
        self.data_dict['pot_1']=df_dict['Potência Ativa L1N Méd.']
        self.data_dict['pot_2']=df_dict['Potência Ativa L2N Méd.']
        self.data_dict['pot_3']=df_dict['Potência Ativa L3N Méd.']
        self.data_dict['pot_t']=df_dict['Potência Ativa Total Méd.']
        self.data_dict['energia_t']=df_dict['Energia Ativa Total Méd.']
        self.data_dict['energia_dia'] = [0]*len(df_dict['Hora'])
        self.data_dict['energia_p'] = [0]*len(df_dict['Hora'])

        self.intervalo = int((datetime.strptime(self.data_dict['hora'][1], r'%H:%M:%S.%f') - datetime.strptime(self.data_dict['hora'][0], r'%H:%M:%S.%f')).total_seconds())
    
    def corrigir_hora(self):
        i=0
        for hora in self.data_dict['hora']:
            new_hora = hora.split(':')
            new_hora.pop()
            new_hora = ':'.join(new_hora)
            self.data_dict['hora'][i] = new_hora
            i+=1
    
    def preencher(self):
        if self.data_dict['hora'][0] == '00:00':
            return 0
        h0 = self.data_dict['hora'][0]
        d_semana = self.data_dict['data'][0].weekday()
        found_flag = False
        id=0
        for dia in self.data_dict['data']:
            if dia.weekday() == d_semana and dia != self.data_dict['data'][0]:
                found_flag=True
                break
            id+=1

        if found_flag == False:
            id=0
            for hora in self.data_dict['hora']:
                if hora == '00:00':
                    break
                id+=1

        i=0
        while True:
            for key in self.data_dict.keys():
                self.data_dict[key].insert(i,self.data_dict[key][id+i])
            if self.data_dict['hora'][id+i+2] == h0:
                break
            i+=1
            id+=1
        
    def calc_energia(self):
        ponta = []
        for h in range(18,21):
            for m in range(60):
                min = f'0{m}' if m<10 else m
                ponta.append(f'{h}:{min}')
        i=0
        for hora in self.data_dict['hora']:
            if hora != '00:00':
                self.data_dict['energia_dia'][i] = float(self.data_dict['pot_t'][i])*float(self.intervalo/3600) + self.data_dict['energia_dia'][i-1]
            else:
                self.data_dict['energia_dia'][i] = 0
            if hora == ponta[0]:
                self.data_dict['energia_p'][i] = 0
            elif hora in ponta:
                self.data_dict['energia_p'][i] = float(self.data_dict['pot_t'][i])*float(self.intervalo/3600) + self.data_dict['energia_p'][i-1]
            else:
                self.data_dict['energia_p'][i] = 0
            i+=1
        
        


            
        