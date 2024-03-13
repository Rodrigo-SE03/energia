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
def criar_planilha(results_dict,dados_dict,folder,nome):
    writer = pd.ExcelWriter(f'{folder}/{nome}',engine="xlsxwriter")
    tab_tensoes(tensoes_dict=results_dict['Tensões'],writer=writer,dados_dict=dados_dict)
    tab_fp(fp_dict=results_dict['FP'],writer=writer)
    tab_correntes(i_dict=results_dict['Correntes'],writer=writer)
    tab_fd(fd_dict=results_dict['FD'],writer=writer)
    tab_consumo(consumo_dict=results_dict['Consumo'],writer=writer)
    tab_thd(thd_dict=results_dict['THD'],writer=writer)
    writer.close()
#--------------------------------------------------------------------------------------------------------

#Criação da aba de tensões  
def tab_tensoes(tensoes_dict,dados_dict,writer):
    global merge_style
    global empty_style

    tr = tensoes_dict['Ref']
    t_eficaz = []
    lim_inf = []
    lim_sup = []
    i=0
    while i < len(dados_dict['Hora']):
        t_eficaz.append(tr)
        lim_inf.append(tr*0.9)
        lim_sup.append(tr*1.1)
        i+=1

    dados = {
        'Data': dados_dict['Data'],
        'Hora': dados_dict['Hora'],
        'V1': dados_dict['V1'],
        'V2': dados_dict['V2'],
        'V3': dados_dict['V3'],
        'Tensão Eficaz': t_eficaz,
        'Limite Superior': lim_sup,
        'Limite Inferior': lim_inf,
        'Frequência':dados_dict['Freq']
    }

    dados_df = pd.DataFrame(dados)
    start = 38
    dados_df.to_excel(writer,sheet_name='Tensões',startrow=start+1,header=False,index=False)
    workbook = writer.book
    worksheet = writer.sheets["Tensões"]
    (max_row, max_col) = dados_df.shape
    column_settings = [{"header": column} for column in dados_df.columns]
    worksheet.add_table(start, 0, max_row+start, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)

    chart = workbook.add_chart({'type':'line'})
    chart.add_series({'categories':f"='Tensões'!$B${start+2}:$B${max_row+1+start}",'name': "V1",'values':f"='Tensões'!$C${start+2}:$C${max_row+1+start}",'line':{'color':'#FFC000','width':1}})
    chart.add_series({'name': "V2",'values':f"='Tensões'!$D${start+2}:$D${max_row+1+start}",'line':{'color':'#5B9BD5','width':1}})
    chart.add_series({'name': "V3",'values':f"='Tensões'!$E${start+2}:$E${max_row+1+start}",'line':{'color':'#70AD47','width':1}})
    chart.add_series({'name': "Tensão Eficaz",'values':f"='Tensões'!$F${start+2}:$F${max_row+1+start}",'line':{'color':'#4472C4','width':1.5}})
    chart.add_series({'name': "Limite Superior",'values':f"='Tensões'!$G${start+2}:$G${max_row+1+start}",'line':{'color':'red','width':1.5,'dash_type': 'long_dash'}})
    chart.add_series({'name': "Limite Inferior",'values':f"='Tensões'!$H${start+2}:$H${max_row+1+start}",'line':{'color':'red','width':1.5,'dash_type': 'long_dash'}})
    chart.set_y_axis({'min': 150,'max': 250,'name': 'Tensão de Fase (V)'})    

    chart.set_size({'width': 1071.496063, 'height': 348.8503937})
    chart.set_legend({'position': 'bottom'})
    chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('K2', chart)

    freq_chart = workbook.add_chart({'type':'line'})
    freq_chart.add_series({'categories':f"='Tensões'!$B${start+2}:$B${max_row+1+start}",'name': "Frequência",'values':f"='Tensões'!$I${start+2}:$I${max_row+1+start}",'line':{'color':'#ED7D31','width':1}})
    freq_chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    freq_chart.set_y_axis({'min': 59.40,'max': 60.40,'name': 'Frequência Elétrica - Hz'})    

    freq_chart.set_size({'width': 1071.496063, 'height': 348.8503937})
    freq_chart.set_legend({'position': 'bottom'})
    freq_chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    freq_chart.set_title({'none':True})
    freq_chart.set_legend({'none': True})
    worksheet.insert_chart('K21', freq_chart)
    worksheet.set_column('A:I',0.1)

    merge_format = workbook.add_format(merge_style)
    valor_format = workbook.add_format({'border':1,'align':'right','num_format':'#.00 "V"'})
    pct_format = workbook.add_format({'border':1,'align':'right','num_format':'#.00"%"'})
    border_format = workbook.add_format({'border':1,'align':'right','num_format':'0.00'})

    worksheet.merge_range("AC3:AD3","Limite inferior",merge_format)
    worksheet.merge_range("AE3:AF3","Limite superior",merge_format)
    worksheet.write("AC4",'Valor',merge_format)
    worksheet.write("AE4",'Valor',merge_format)
    worksheet.write("AD4",'Desvio %',merge_format)
    worksheet.write("AF4",'Desvio %',merge_format)

    worksheet.write_column("AC5",tensoes_dict['Valor inf'],valor_format)
    worksheet.write_column("AD5",tensoes_dict['Desvio % inf'],pct_format)
    worksheet.write_column("AE5",tensoes_dict['Valor sup'],valor_format)
    worksheet.write_column("AF5",tensoes_dict['Desvio % sup'],pct_format)
    worksheet.merge_range("AC8:AF8",f"Valor de Referência: {tr}V",workbook.add_format({'border':1,'align':'center'}))

    worksheet.merge_range("AC23:AD23","Frequência [Hz]",merge_format)
    worksheet.write("AC24",'Mínimo',merge_format)
    worksheet.write("AC25",'Máximo',merge_format)
    worksheet.write("AD24",min(dados['Frequência']),border_format)
    worksheet.write("AD25",max(dados['Frequência']),border_format)

    worksheet.set_column('A:A',18)
    worksheet.set_column('AC:AF',12)
#--------------------------------------------------------------------------------------------------------

#Criação da aba de fator de potência 
def tab_fp(fp_dict,writer):
    global merge_style

    dados = {
        'Data': fp_dict['Data'],
        'Hora': fp_dict['Hora'],
        'FP1': fp_dict['FP1'],
        'FP2': fp_dict['FP2'],
        'FP3': fp_dict['FP3']
    }

    dados_df = pd.DataFrame(dados)
    start = 19
    dados_df.to_excel(writer,sheet_name='Fator de Potência',startrow=start+1,header=False,index=False)
    workbook = writer.book
    worksheet = writer.sheets["Fator de Potência"]
    (max_row, max_col) = dados_df.shape
    column_settings = [{"header": column} for column in dados_df.columns]
    worksheet.add_table(start, 0, max_row+start, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)

    chart = workbook.add_chart({'type':'line'})
    chart.add_series({'categories':f"='Fator de Potência'!$B${start+2}:$B${max_row+1+start}",'name': f"='Fator de Potência'!$C${start+1}",'values':f"='Fator de Potência'!$C${start+2}:$C${max_row+2+start}",'line':{'color':'#4472C4','width':1}})
    chart.add_series({'name': f"='Fator de Potência'!$D${start+1}",'values':f"='Fator de Potência'!$D${start+2}:$D${max_row+1+start}",'line':{'color':'#ED7D31','width':1}})
    chart.add_series({'name': f"='Fator de Potência'!$E${start+1}",'values':f"='Fator de Potência'!$E${start+2}:$E${max_row+1+start}",'line':{'color':'#A5A5A5','width':1}})
    chart.set_y_axis({'min': 0,'max': 1,'name': 'Fator de Potência'})    

    chart.set_size({'width': 1071.496063, 'height': 348.8503937})
    chart.set_legend({'position': 'bottom'})
    chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('H2', chart)
    

    merge_format = workbook.add_format(merge_style)
    border_format = workbook.add_format({'border':1,'align':'right','num_format':'0.00'})

    worksheet.write(f"Y3",'Valores',merge_format)
    worksheet.write(f"Z3",'Fase 1',merge_format)
    worksheet.write(f"AA3",'Fase 2',merge_format)
    worksheet.write(f"AB3",'Fase 3',merge_format)
    worksheet.write(f"AC3",'Total',merge_format)

    i=0
    for key in fp_dict.keys():
        if key == 'Valores':
            worksheet.write_column("Y4",fp_dict[key],merge_format)
            i+=1
        elif key == 'Fase 1' or key == 'Fase 2' or key == 'Fase 3' or key == 'Total':
            worksheet.write_column(3,24+i,fp_dict[key],border_format)
            i+=1
    worksheet.set_column('Y:AC',12)
    worksheet.set_column(1,max_col,0.1)
    worksheet.set_column('A:A',18)
#--------------------------------------------------------------------------------------------------------
        
#Criação da aba de correntes
def tab_correntes(i_dict,writer):
    global merge_style
    global empty_style

    dados = {
        'Data': i_dict['Data'],
        'Hora': i_dict['Hora'],
        'I1': i_dict['I1'],
        'I2': i_dict['I2'],
        'I3': i_dict['I3'],
        'In': i_dict['In']
    }

    dados_df = pd.DataFrame(dados)
    start = 19
    dados_df.to_excel(writer,sheet_name='Correntes',startrow=start+1,header=False,index=False)

    workbook = writer.book
    worksheet = writer.sheets["Correntes"]
    (max_row, max_col) = dados_df.shape
    column_settings = [{"header": column} for column in dados_df.columns]
    worksheet.add_table(start, 0, max_row+start, max_col - 1, {"columns": column_settings})

    chart = workbook.add_chart({'type':'line'})
    chart.add_series({'categories':f"='Correntes'!$B${start+2}:$B${max_row+1+start}",'name': "I1",'values':f"='Correntes'!$C${start+2}:$C${max_row+1+start}",'line':{'color':'#ffc30e','width':1}})
    chart.add_series({'name': "I2",'values':f"='Correntes'!$D${start+2}:$D${max_row+1+start}",'line':{'color':'#63a0d7','width':1}})
    chart.add_series({'name': "I3",'values':f"='Correntes'!$E${start+2}:$E${max_row+1+start}",'line':{'color':'#99c47c','width':1}})
    chart.add_series({'name': "In",'values':f"='Correntes'!$F${start+2}:$F${max_row+1+start}",'line':{'color':'#ff4949','width':1}})

    v_max = 5*int(max([i_dict['Fase 1'][1],i_dict['Fase 2'][1],i_dict['Fase 3'][1],i_dict['Neutro'][1]])/5)
    while v_max < max([i_dict['Fase 1'][1],i_dict['Fase 2'][1],i_dict['Fase 3'][1],i_dict['Neutro'][1]]):
        v_max+=5
    
    chart.set_size({'width': 1071.496063, 'height': 348.8503937})
    chart.set_y_axis({'min': 0,'max': v_max,'name': 'Corrente entre as fases (A)'})   
    chart.set_legend({'position': 'bottom'})
    chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('H2', chart)

    merge_format = workbook.add_format(merge_style)
    border_format = workbook.add_format({'border':1,'align':'center','num_format':'0.00'})
    empty_format = workbook.add_format(empty_style)

    worksheet.merge_range("Z3:AD3","Corrente Média Registrada",merge_format)
    worksheet.write('Z4','',empty_format)
    worksheet.write_row('AA4',['Fase 1','Fase 2','Fase 3','Neutro'],merge_format)

    i=0
    for key in i_dict.keys():
        if key == 'Valores' or key == 'Fase 1' or key == 'Fase 2' or key == 'Fase 3' or key == 'Neutro':
            worksheet.write_column(4,25+i,i_dict[key],border_format)
            i+=1
    
    worksheet.set_column('Z:AD',12)
    worksheet.set_column('A:A',18)

    merge_format.set_text_wrap()
    worksheet.merge_range("I22:M27","Para criar o gráfico de cada fase basta ocultar as demais colunas na tabela de dados",merge_format)
#--------------------------------------------------------------------------------------------------------
    
#Criação da aba de fator de distorção
def tab_fd(fd_dict,writer):
    global merge_style
    global empty_style

    dados = {
        'Data': fd_dict['Data'],
        'Hora': fd_dict['Hora'],
        'FD': fd_dict['FD'],
        'Limite': fd_dict['Limite']
    }

    dados_df = pd.DataFrame(dados)
    start = 19
    dados_df.to_excel(writer,sheet_name='Fator de Distorção',startrow=start+1,header=False,index=False)

    workbook = writer.book
    worksheet = writer.sheets["Fator de Distorção"]
    (max_row, max_col) = dados_df.shape
    column_settings = [{"header": column} for column in dados_df.columns]
    worksheet.add_table(start, 0, max_row+start, max_col - 1, {"columns": column_settings})

    chart = workbook.add_chart({'type':'line'})
    chart.add_series({'categories':f"='Fator de Distorção'!$B${start+2}:$B${max_row+1+start}",'name': "Desequilíbrio %",'values':f"='Fator de Distorção'!$C${start+2}:$C${max_row+1+start}",'line':{'color':'#70AD47','width':1}})
    chart.add_series({'name': "Limite",'values':f"='Fator de Distorção'!$D${start+2}:$D${max_row+1+start}",'line':{'color':'red','width':1.5,'dash_type': 'long_dash'}})

    v_max = 0.5*int(max(dados['FD'])/0.5)
    while v_max < max(dados['FD']):
        v_max+=0.5
    
    chart.set_size({'width': 1071.496063, 'height': 348.8503937})
    chart.set_y_axis({'min': 0,'max': v_max,'name': 'Fator de Distorção - %'})   
    chart.set_legend({'position': 'bottom'})
    chart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart('F2', chart)

    merge_format = workbook.add_format(merge_style)
    border_format = workbook.add_format({'border':1,'align':'center','num_format':'0.00'})
    empty_format = workbook.add_format(empty_style)

    worksheet.write('W4','Valores',merge_format)
    worksheet.write('X4','FD [%]',merge_format)

    i=0
    for key in fd_dict.keys():
        if key == 'Valores' or key == 'FD [%]':
            format = merge_format if key == 'Valores' else border_format
            worksheet.write_column(4,22+i,fd_dict[key],format)
            i+=1
    
    worksheet.set_column('W:X',12)
    worksheet.set_column('A:A',18)
    worksheet.set_column(1,max_col-1,0.1)
#--------------------------------------------------------------------------------------------------------

#Criação da aba de consumo
def tab_consumo(consumo_dict,writer):
    global merge_style

    dados = {
        'Data': consumo_dict['Data'],
        'Hora': consumo_dict['Hora'],
        'Potência': consumo_dict['Potência'],
        'Energia': consumo_dict['Energia']
    }

    dados_df = pd.DataFrame(dados)
    start = 19
    dados_df.to_excel(writer,sheet_name='Consumo',startrow=start+1,header=False,index=False)

    workbook = writer.book
    worksheet = writer.sheets["Consumo"]
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

def tab_thd(thd_dict,writer):
    global merge_style
    global empty_style

    dados = {
        'Data': thd_dict['Data'],
        'Hora': thd_dict['Hora'],
        'THD V1': thd_dict['THD V1'],
        'THD V2': thd_dict['THD V2'],
        'THD V3': thd_dict['THD V3'],
        'THD I1': thd_dict['THD I1'],
        'THD I2': thd_dict['THD I2'],
        'THD I3': thd_dict['THD I3'],
    }

    dados_df = pd.DataFrame(dados)
    start = 38
    dados_df.to_excel(writer,sheet_name='THD',startrow=start+1,header=False,index=False)

    workbook = writer.book
    worksheet = writer.sheets["THD"]
    (max_row, max_col) = dados_df.shape
    column_settings = [{"header": column} for column in dados_df.columns]
    worksheet.add_table(start, 0, max_row+start, max_col - 1, {"columns": column_settings})

    vchart = workbook.add_chart({'type':'line'})
    vchart.add_series({'categories':f"='THD'!$B${start+2}:$B${max_row+1+start}",'name': "thd_V1 [%]",'values':f"='THD'!$C${start+2}:$C${max_row+1+start}",'line':{'color':'#4472C4','width':1}})
    vchart.add_series({'name': "thd_V2 [%]",'values':f"='THD'!$D${start+2}:$D${max_row+1+start}",'line':{'color':'#ED7D31','width':1}})
    vchart.add_series({'name': "thd_V3 [%]",'values':f"='THD'!$E${start+2}:$E${max_row+1+start}",'line':{'color':'#A5A5A5','width':1}})

    v_max = 5*int(max([max(dados['THD V1']),max(dados['THD V2']),max(dados['THD V3'])])/5)
    while v_max < max([max(dados['THD V1']),max(dados['THD V2']),max(dados['THD V3'])]):
        v_max+=5
    
    vchart.set_size({'width': 1071.496063, 'height': 348.8503937})
    vchart.set_y_axis({'min': 0,'max': v_max,'name': 'Taxa de Distorção - %'})   
    vchart.set_legend({'position': 'bottom'})
    vchart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart(2,max_col+2, vchart)

    ichart = workbook.add_chart({'type':'line'})
    ichart.add_series({'categories':f"='THD'!$B${start+2}:$B${max_row+1+start}",'name': "thd_I1 [%]",'values':f"='THD'!$F${start+2}:$F${max_row+1+start}",'line':{'color':'#4472C4','width':1}})
    ichart.add_series({'name': "thd_I2 [%]",'values':f"='THD'!$G${start+2}:$G${max_row+1+start}",'line':{'color':'#ED7D31','width':1}})
    ichart.add_series({'name': "thd_I3 [%]",'values':f"='THD'!$H${start+2}:$H${max_row+1+start}",'line':{'color':'#A5A5A5','width':1}})

    v_max = 5*int(max([max(dados['THD I1']),max(dados['THD I2']),max(dados['THD I3'])])/5)
    while v_max < max([max(dados['THD I1']),max(dados['THD I2']),max(dados['THD I3'])]):
        v_max+=5
    
    ichart.set_size({'width': 1071.496063, 'height': 348.8503937})
    ichart.set_y_axis({'min': 0,'max': v_max,'name': 'Taxa de Distorção - %'})   
    ichart.set_legend({'position': 'bottom'})
    ichart.set_chartarea({'border':{'color': '#4472C4','width':1.25}})
    worksheet.insert_chart(21,max_col+2, ichart)

    merge_format = workbook.add_format(merge_style)
    border_format = workbook.add_format({'border':1,'align':'right','num_format':'0.00 "%"'})

    worksheet.merge_range(2,max_col+19,2,max_col+20,"Distorção Harmônica de Tensão [%]",merge_format)
    worksheet.write(3,max_col+19,'Médio',merge_format)  
    worksheet.write(4,max_col+19,'Máximo',merge_format)

    worksheet.write(3,max_col+20,thd_dict['THD V'][0],border_format)  
    worksheet.write(4,max_col+20,thd_dict['THD V'][1],border_format)

    worksheet.merge_range(23,max_col+19,23,max_col+20,"Distorção Harmônica de Corrente [%]",merge_format)
    worksheet.write(24,max_col+19,'Médio',merge_format)  
    worksheet.write(25,max_col+19,'Máximo',merge_format)

    worksheet.write(24,max_col+20,thd_dict['THD I'][0],border_format)  
    worksheet.write(25,max_col+20,thd_dict['THD I'][1],border_format)

    worksheet.set_column('A:A',18)
    worksheet.set_column(max_col+19,max_col+20,18)
    worksheet.set_column(1,max_col-1,0.1)