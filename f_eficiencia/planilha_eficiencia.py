import pandas as pd
from datetime import datetime
import os

merge_style ={
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#dbdbdb"
    }

empty_style = {
        "bold": 1,
        "border": 1,
        "fg_color": "#dbdbdb",
    }


#Função geral de criação de planilhas
def criar_planilha(results_dict,folder,nome):
    writer = pd.ExcelWriter(f'{folder}/{nome}',engine="xlsxwriter")
    tab_consumo(consumo_dict=results_dict['Consumo'],writer=writer)
    writer.close()
#--------------------------------------------------------------------------------------------------------

#Criação da aba de consumo
def tab_consumo(consumo_dict,writer):
    global merge_style

    dados = {
        'Data': consumo_dict['Data'],
        'Hora': consumo_dict['Hora'],
        'Potência': consumo_dict['Potência'],
        'Energia': consumo_dict['Energia Total']
    }

    dados_df = pd.DataFrame(dados)
    start = 19
    dados_df.to_excel(writer,sheet_name='Consumo',startrow=start+1,header=False,index=False)

    workbook = writer.book
    worksheet = writer.sheets[f"Consumo"]
    (max_row, max_col) = dados_df.shape
    column_settings = [{"header": column} for column in dados_df.columns]
    worksheet.add_table(start, 0, max_row+start, max_col - 1, {"columns": column_settings})

    chart = workbook.add_chart({'type':'line'})
    chart.add_series({'categories':f"='Consumo'!$B${start+2}:$B${max_row+1+start}",'name': "PT Calc.",'values':f"='Consumo'!$C${start+2}:$C${max_row+1+start}",'line':{'color':'#4472C4','width':1}})
    chart.add_series({'y2_axis': True,'name': "Energia Calc.",'values':f"='Consumo'!$D${start+2}:$D${max_row+1+start}",'line':{'color':'red','width':1.5}})

    v_max = 2*int(max(dados['Potência'])/2)
    while v_max < max(dados['Potência']):
        v_max+=2
    
    chart.set_size({'width': 1071.496063, 'height': 348.8503937})
    chart.set_y_axis({'min': 0,'max': v_max,'name': 'Potência [kW]'})
    chart.set_y2_axis({'name': 'Energia [kWh]'})
    chart.set_legend({'position': 'bottom'})
    chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('F2', chart)

    merge_format = workbook.add_format(merge_style)
    border_format = workbook.add_format({'border':1,'align':'center','num_format':'0.00'})
    date_format = workbook.add_format({'border':1,'align':'center','num_format':'dd/mmm'})
    semana_format = workbook.add_format({'border':1,'align':'center','num_format':'ddd'})
    fds_format = workbook.add_format({'border':1,'align':'center','num_format':'ddd','color':'#FF0000'})

    worksheet.merge_range('W4:X4','Período',merge_format)
    worksheet.write('Y4','Consumo',merge_format)
    worksheet.merge_range('W5:X5','7 Dias',merge_format)  #PERMITIR PERSONALIZAÇÃO
    worksheet.write('Y5','[kWh]',merge_format)


    worksheet.write_column(5,22,consumo_dict['Dias'],date_format)
    i=0
    for dia in consumo_dict['Dias']:
        dia_s = str(dia).replace(' ','-').split('-')
        dia_s.pop()
        dia_s = datetime(int(dia_s[0]),int(dia_s[1]),int(dia_s[2])).weekday()
        if dia_s == 5 or dia_s == 6:
            worksheet.write(5+i,23,dia,fds_format)
        else:
            worksheet.write(5+i,23,dia,semana_format)
        i+=1
        dia_s = ''
    worksheet.write_column(5,24,consumo_dict['Consumo'],border_format)

    
    worksheet.set_column('W:Y',12)
    worksheet.set_column('A:A',18)
    worksheet.set_column(1,max_col-1,0.1)
#--------------------------------------------------------------------------------------------------------