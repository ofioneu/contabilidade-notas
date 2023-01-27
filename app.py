import PySimpleGUI as sg
import pandas as pd
import openpyxl
from pandas import DataFrame, read_excel
from datetime import datetime

headings =['NF-e', 'Cliente', 'Valor Total', 'Custo Total','Custo NF', 'Custo Final',
           'Data de saída', 'Lucro_bruto', 'Lucro Liq', 'Fabio',' Arthur']


def calcular(faturamento, margem_de_lucro):
    custo_nf=[]
    custo_final= []    
    fabio = []
    arthur = []
    lucro_bruto = []
    lucro_liquido = []    
    faturamento_df= pd.read_excel(faturamento, engine='openpyxl', index_col=False)
    margem_de_lucro_df= pd.read_excel(margem_de_lucro, engine='openpyxl', index_col=False)
    faturamento_df.insert(3,'Custo Total', margem_de_lucro_df['Custo Total'] )
    
    #custo NF 10%
    for i in faturamento_df['Valor Total']:
      custo_nf.append(i*0.1)
    
    #custo final: custo nf + custo total prod  
    for i, j in zip(custo_nf, faturamento_df['Custo Total']):
        custo_final.append(i+j)
    
    #Lucro bruto valor nf - ((custo prod + tax nf) >> custo final)    
    for i, j in zip(faturamento_df['Valor Total'], custo_final):
        lucro_bruto.append(i-j)
        
    #comissão Fabio
    for i in lucro_bruto:
        fabio.append(i*0.1)
           
    #comissão Arthur
    for i in faturamento_df['Valor Total']:
        arthur.append(i*0.01)
    
    #lucro liquido: lucro bruto - fabio + arthur
    comissoes_somadas_fabio_arthur = [i + j for i, j in zip(fabio, arthur)]
    for i, j in zip(lucro_bruto, comissoes_somadas_fabio_arthur):
        lucro_liquido.append(i-j)
    
      
    faturamento_df.drop(['Observações internas', 'Situação', 'CFOP', 'Data Autorização', 'Data Emissão'], axis=1, inplace=True)
    
    
    faturamento_df.insert(4,'Imposto NF', custo_nf )
    faturamento_df.insert(5,'Custo Final', custo_final )
    faturamento_df.insert(6,'Lucro Bruto', lucro_bruto )
    faturamento_df.insert(7,'Lucro Liquido', lucro_liquido )
    faturamento_df.insert(8,'Fabio', fabio )
    faturamento_df.insert(9,'Arthur', arthur )
    
    
    res_table=[]
    for a, b,c,d,e,f,g,h,i,j,k in zip(faturamento_df['NF-e'],faturamento_df['Cliente'], 
            faturamento_df['Valor Total'], faturamento_df['Custo Total'], custo_nf, custo_final, faturamento_df['Data de saída'], 
            lucro_bruto, lucro_liquido, fabio, arthur):
        
        res_table.append([a, b,c,d,round(e,2),round(f,2),g,round(h,2),round(i,2),round(j,2),round(k,2)])
        
    
    return(res_table)
    
    
resultado_array=[]

sg.theme('DefaultNoMoreNagging')   # Add a touch of color
# All the stuff inside your window.
layout = [  [sg.Text('FATURAMENTO XLSX'), sg.Input(visible=True, enable_events=True, key='-FATURAMENTO-'), sg.FilesBrowse()],            
            [sg.Text('RELATORIO DE MARGEM DE LUCRO XLSX'), sg.Input(visible=True, enable_events=True, key='-MARGEM_DE_LUCRO-'), sg.FilesBrowse()],         
            [sg.Table(values=resultado_array, headings = headings, auto_size_columns=False,
                    display_row_numbers=True,
                    justification='center',
                    key='-TABLE-',
                    enable_events=True,
                    row_height=20,
                    size=(10,10))],
            [sg.Button('CALCULAR'), sg.Button('Clear', enable_events= True), sg.Button('Reset'), sg.Exit(),
             sg.Input(visible=False, enable_events=True, key='-EXPORT_XLSX-')],
            [sg.Text('Exportar Excel:', auto_size_text=False, justification='right'),sg.Input(key='-PATH_SAVE-'), sg.FolderBrowse(key='-PATH_FOLDER-', size=(10, 1)), sg.Button('Save', key='-SAVE-')]
          ]

# Create the Window
window = sg.Window('CONTABILIDADE NFS', layout)

# Event Loop to process "events" and get the "values" of the inputs
while True:
    xlsx_array =[]
    table_array=[]
    event, values = window.read()
    
    if event == sg.WIN_CLOSED or event == 'Exit':
        break   # The Event Loop
    
    faturamento = values['-FATURAMENTO-']
    margem_de_lucro = values['-MARGEM_DE_LUCRO-']
   
    
    
    if event == 'CALCULAR':
        window['Clear'].update(visible=True)
        valores=calcular(faturamento,margem_de_lucro)  
        window['-TABLE-'].update(values=valores)
    
        
        
    if event == 'Clear':
        window['-TABLE-'].Update('')
        if len(resultado_array)<=1:
            del(resultado_array[0])
        else:
            resultado_array.pop()
    
    if event == 'Reset':
        window['-TABLE-'].Update('')
        resultado_array.clear() 
    
    if event == '-SAVE-':
        valores=calcular(faturamento,margem_de_lucro)
        path_folder = values['-PATH_FOLDER-']
        print(path_folder)
        xlsx_frame = pd.DataFrame(valores, columns=headings)
        hoje = datetime.now()
        str_hoje =  hoje.strftime("%Y_%m_%d %H_%M_%S")
        xlsx_frame.to_excel(f'{path_folder}/{str_hoje}.xlsx')
        sg.SystemTray.notify('NOTIFICAÇÃO!', 'XLSX EXPORTADO COM SUCESSO', location=(500,300))


window.close()