import time
from pathlib import Path
import openpyxl
import pandas as pd
import numpy as np
from datetime import datetime
import PySimpleGUI as sg
from selenium import webdriver
import os
import pyperclip
import ctypes

logo_path = "logo.png"
ctypes.windll.user32.SetProcessDPIAware()
ctypes.windll.kernel32.SetConsoleIcon.argtypes = [ctypes.c_void_p]
ctypes.windll.kernel32.SetConsoleIcon(ctypes.windll.shell32.ShellExecuteW(None, "open", logo_path, None, None, 5))

hoje1 = datetime.today()
hoje2 = hoje1.strftime('%d/%m %H:%M')
hoje3 = hoje2.replace("/", "-")
hoje4 = hoje3.replace(":", "-")
nome = hoje4 + '.xlsx'

dir_arquivo = os.getcwd() + '\\ativos.exe'
diretorio_final = os.getcwd() + '\\StatusBolt\\'
diretorio_cnct = Path(diretorio_final) / 'conected'
diretorio_company = Path(diretorio_final) / 'company'
image_restart = diretorio_final + 'buscar.png'
arquivo_excel_empresa = diretorio_company.joinpath(nome)

icone = 'ibbx.ico'
dir_icone = Path(diretorio_final) / icone
font = ['Arial', 10, 'bold']

button_color = ('white', '#383838')
pad = (10, 10)

layout = [
         [sg.Text('...............................................................', text_color='#FFFFFF', background_color='#FFFFFF')],
         [sg.Text('........', text_color='#FFFFFF', background_color='#FFFFFF'), sg.Text('Selecione a ação desejada:', text_color='black', font=font, background_color='#FFFFFF')],
         [sg.Text('...........', text_color='#FFFFFF', background_color='#FFFFFF'), sg.Button('Comparar', button_color=button_color, pad=pad, size=(16, 2), font=font)],
         [sg.Text('...........', text_color='#FFFFFF', background_color='#FFFFFF'), sg.Button('Importar Dados', button_color=button_color, pad=pad, size=(16, 2), font=font)],
         [sg.Text('...........', text_color='#FFFFFF', background_color='#FFFFFF'), sg.Button('Mapear Empresas', button_color=button_color, pad=pad, size=(16, 2), font=font)]]


janela = sg.Window('Gestor de ativos...', background_color='#FFFFFF', icon=dir_icone, finalize=True).Layout(layout)
botao, valores = janela.Read()

if botao == 'Comparar':
    janela.close()

    layout = [[sg.Text('Selecione as planilhas para fazer as comparações:', text_color='#17202A', font=('Arial', 10, 'bold'), background_color='#FFFFFF')],
              [sg.Text('1º Planilha', text_color='#17202A', font=('Arial', 10, 'bold'), background_color='#C1C1C1')],
              [sg.Input(), sg.FileBrowse('Procurar', button_color=button_color, pad=pad, size=(9, 2), font=('Arial', 9, 'bold'))],
              [sg.Text('2º Planilha', text_color='#17202A', font=('Arial', 10, 'bold'), background_color='#C1C1C1')],
              [sg.Input(), sg.FileBrowse('Procurar', button_color=button_color, pad=pad, size=(9, 2), font=('Arial', 9, 'bold')),
               sg.OK('Next', button_color=('white', '#f79000'), pad=pad, size=(9, 2),font=('Arial', 9, 'bold'))]]

    janela = sg.Window('Selecionar Planilhas...', background_color='#FFFFFF', icon=dir_icone, finalize=True).Layout(layout)

    botao, valores = janela.Read()

    if botao == 'Next':
        arquivo1 = valores[0]
        arquivo2 = valores[1]
        dfpush = pd.read_excel(arquivo1)
        dfpush2 = pd.read_excel(arquivo2)

    janela.Close()

    texto = arquivo1
    palavras = texto.rsplit("/", 1)
    ultima_palavra = palavras[-1]
    arq1 = ultima_palavra.split(".", 1)[0]
    texto = arq1
    partes = texto.rsplit("-", 2)
    novo_texto = partes[0] + " " + partes[1] + ":" + partes[2]
    novo_texto = novo_texto.replace(' ', '-', 1)

    texto = arquivo2
    palavras = texto.rsplit("/", 1)
    ultima_palavra = palavras[-1]
    arq1 = ultima_palavra.split(".", 1)[0]
    texto = arq1
    partes = texto.rsplit("-", 2)
    novo_textoo = partes[0] + " " + partes[1] + ":" + partes[2]
    novo_textoo = novo_textoo.replace(' ', '-', 1)

    df = pd.DataFrame(dfpush)
    df2 = pd.DataFrame(dfpush2)

    planilha_a = dfpush['Conectados']
    planilha_b = dfpush2['Conectados']

    resultado = planilha_a - planilha_b
    resultado = resultado.rename('Saldo')
    df_resultado = resultado.to_frame()

    df2_5 = df2[['Ativados', 'Conectados']]
    df3 = df.join(df2_5, rsuffix=f" {novo_textoo}")
    cncttxt = "Conectados " + novo_textoo
    ativdtxt = "Ativados " + novo_textoo

    primeira_linha = df3.iloc[[0]]
    resto_df = df3.iloc[1:]
    resto_df = resto_df.fillna(0)

    resto_df = resto_df.join(resultado, rsuffix=' ').astype({'Conectados': int, cncttxt: int}).astype(str)
    df4 = pd.concat([primeira_linha, resto_df])

    primeira_linha = df4.iloc[0]
    nova_lista = primeira_linha[1:]

    data = np.array([primeira_linha])
    data1 = [list(map(str, row[1:])) for row in data]
    data1 = [['Total'] + data1[0]]

    lista_interna = data1[0]
    gap_a = float(lista_interna[1])
    gap_b = float(lista_interna[2])
    gap_total1 = int(gap_a - gap_b)

    gap_c = float(lista_interna[3])
    gap_d = float(lista_interna[4])
    gap_total2 = int(gap_c - gap_d)

    a = float(lista_interna[2])
    b = float(lista_interna[4])

    vai = a - b
    lista_interna.append(int(vai))
    lista_interna.pop(5)

    df4 = df4.drop(0)
    df4['Saldo'] = df4['Saldo'].astype(int)
    df5 = df4.sort_values('Saldo', ascending=False)
    data = df5.values.tolist()

    headers = {'Empresa': [20], f'Ativados {novo_texto}': [20], f'Conectados {novo_texto}': [20], f'Ativados {novo_textoo}': [20], f'Conectados {novo_textoo}': [20], 'Saldo': [20]}
    headers = list(headers)
    col_widht = [15, 18, 18, 18, 18, 10]

    layout = \
        [[sg.Text(f"           Gap {gap_total1}        ", text_color='black', background_color='#FFB200', font=('Arial', 10, 'bold')),
          sg.Text(f"                             {novo_texto}                               ", text_color='black', font=('Arial', 10, 'bold'), background_color='#FFB200'),
          sg.Text(f"                             {novo_textoo}                             ", text_color='white', font=('Arial', 10, 'bold'), background_color='#283B5B'),
          sg.Text(f"    Gap {gap_total2}     ", text_color='white', background_color='#283B5B', font=('Arial', 10, 'bold'))],
        [sg.Table(data, headings=headers, header_background_color='#4D4D4D', header_font=('Arial', 10, 'bold'), header_text_color='white',  header_border_width=0, col_widths=col_widht, auto_size_columns=False, enable_events=True, enable_click_events=True, justification='center', num_rows=13, background_color='#FFFFFF', alternating_row_color='#F7F7F7', text_color='#17202A', key='-CONTACT_TABLE-', row_height=50)],
        [sg.Table(coco, header_background_color='4D4D4D', header_text_color='4D4D4D', background_color='#4D4D4D', text_color='white', font=('Arial', 10, 'bold'), col_widths=col_widht, hide_vertical_scroll=True, auto_size_columns=False, enable_events=True, enable_click_events=True, justification='center', num_rows=1, key='-CONTACT_TABLE2-', row_height=25)],
        [sg.Text("Nome da empresa: ", text_color='#17202A', font=('Arial', 10, 'bold'), background_color='#FFFFFF'),
         sg.Input(size=(18, 1), background_color='#ECF0F1', key='-INPUT-'),
         sg.ReadFormButton('.', image_filename=image_restart, button_color='#FFFFFF', image_size=(25, 25), image_subsample=2, border_width=0),
         sg.Button('Abrir Diretório', font=('Arial', 10, 'bold'), button_color='#4D4D4D', size=(13, 1)),
         sg.Text(diretorio_final, font=('Arial', 10, 'bold'), text_color='black', background_color='#FFFFFF')]]

    window = sg.Window('Controle de ativos', layout, background_color='#FFFFFF', finalize=True, icon=dir_icone)
    table = window['-CONTACT_TABLE-']
    entry = window['-INPUT-']
    entry.bind('<Return>', 'RETURN-')

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break

        if event == 'Abrir Diretório':
            os.startfile(diretorio_cnct)

        elif isinstance(event, tuple) and event[:2] == ('-CONTACT_TABLE-', '+CLICKED+'):
            row, col = position = event[2]
            if None not in position and row >= 0:
                text = data[row][col]
                pyperclip.copy(text)

        if event in ('.', '-INPUT-RETURN-'):
            text = values['-INPUT-'].lower()
            if text == '':
                continue
            row_colors = []
            for row, row_data in enumerate(data):
                if text in row_data[0].lower():
                    row_colors.append((row, '#99A3A4'))
                else:
                    row_colors.append((row, '#ECF0F1'))
            table.update(row_colors=row_colors)

    window.close()

if botao == 'Importar Dados':
    janela.close()
    button_color = ('white', '#383838')

    layout = [[sg.Text('Carregando...', background_color='#FFFFFF', text_color='#17202A', font=('Arial', 10, 'bold'))],
              [sg.ProgressBar(46, orientation='h', size=(30, 20), key='progress', bar_color=('#167CE2', '#D0D2D3'))],
              [sg.Cancel('Cancel', button_color=button_color, size=(6), font=('Arial', 10, 'bold'))]]

    window = sg.Window('Verificando Empresas ...', layout, icon=dir_icone, background_color='#FFFFFF')

    for i in range(46):
        event, values = window.read(timeout=0)
        if event == 'Cancel' or event == sg.WIN_CLOSED:
            break
        if i == 0:
            options = webdriver.ChromeOptions()
            options.add_argument("--headless")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-infobars")
            options.add_argument("--disable-logging")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-plugins")
            options.add_argument("--blink-settings=imagesEnabled=false")
            options.add_argument("--disable-notifications")
            options.add_argument("--window-size=1920,1080")

            driver = webdriver.Chrome(options=options)

            driver.get("https://empresa.confidencial.com")
            time.sleep(2)

            driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/div[2]/input').send_keys('email_confidencial')
            driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/div[3]/input').send_keys('senha_confidencial!')
            driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/button').click()
            time.sleep(5)

            xpath_cnct = '//*[@id="root"]/div[1]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div[2]/span/span'
            xpath_total = '//*[@id="root"]/div[1]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div[1]/span/span'
            botao_pause = '//*[@id="root"]/div[1]/div[2]/div/div[2]/div[1]/div/button[3]'

            dorme1 = 5

        elif i == 1:
            total_conectados = driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div[2]/span/span').text
            total_conectados = int(str(total_conectados).replace('.', ''))
            total_ativados = driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div[1]/span/span').text
            ativados_empresa_confidencial = int(str(total_ativados).replace('.', ''))

        elif i == 2:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 3:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 4:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 5:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 6:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 7:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 8:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 9:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 10:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 11:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 12:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 13:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            total_ativados_empresa_confidencialunipac = driver.find_element('xpath', xpath_total).text

        elif i == 14:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 15:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 16:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 17:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 18:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 19:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 20:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 21:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 22:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 23:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 24:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 25:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 26:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 27:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 28:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 29:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 30:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 31:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 32:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 33:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 34:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 35:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 36:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 37:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 38:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 39:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 40:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 41:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 42:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 43:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 44:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        elif i == 45:
            driver.get("https://empresa.confidencial/companies/x/facilities")
            time.sleep(dorme1)
            conectados_empresa_confidencial = driver.find_element('xpath', xpath_cnct).text
            ativados_empresa_confidencial = driver.find_element('xpath', xpath_total).text

        window['progress'].update_bar(i + 1)

    window.close()
    driver.quit()

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet["A1"] = "Empresa"
    sheet["B1"] = "Ativados"
    sheet["C1"] = "Conectados"

    sheet["A2"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B2"] = conectados_empresa_confidencial
    sheet["C2"] = ativados_empresa_confidencial

    sheet["A3"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B3"] = conectados_empresa_confidencial
    sheet["C3"] = ativados_empresa_confidencial

    sheet["A4"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B4"] = conectados_empresa_confidencial
    sheet["C4"] = ativados_empresa_confidencial

    sheet["A5"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B5"] = conectados_empresa_confidencial
    sheet["C5"] = ativados_empresa_confidencial

    sheet["A6"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B6"] = conectados_empresa_confidencial
    sheet["C6"] = ativados_empresa_confidencial

    sheet["A7"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B7"] = conectados_empresa_confidencial
    sheet["C7"] = ativados_empresa_confidencial

    sheet["A8"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B8"] = conectados_empresa_confidencial
    sheet["C8"] = ativados_empresa_confidencial

    sheet["A9"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B9"] = conectados_empresa_confidencial
    sheet["C9"] = ativados_empresa_confidencial

    sheet["A10"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B10"] = conectados_empresa_confidencial
    sheet["C10"] = ativados_empresa_confidencial

    sheet["A11"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B11"] = conectados_empresa_confidencial
    sheet["C11"] = ativados_empresa_confidencial

    sheet["A12"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B12"] = conectados_empresa_confidencial
    sheet["C12"] = ativados_empresa_confidencial

    sheet["A13"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B13"] = conectados_empresa_confidencial
    sheet["C13"] = ativados_empresa_confidencial

    # Desativado
    sheet["A14"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B14"] = 0
    sheet["C14"] = 0

    sheet["A15"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B15"] = conectados_empresa_confidencial
    sheet["C15"] = ativados_empresa_confidencial

    sheet["A16"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B16"] = conectados_empresa_confidencial
    sheet["C16"] = ativados_empresa_confidencial

    sheet["A17"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B17"] = conectados_empresa_confidencial
    sheet["C17"] = ativados_empresa_confidencial

    sheet["A18"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B18"] = conectados_empresa_confidencial
    sheet["C18"] = ativados_empresa_confidencial

    sheet["A19"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B19"] = conectados_empresa_confidencial
    sheet["C19"] = ativados_empresa_confidencial

    sheet["A20"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B20"] = conectados_empresa_confidencial
    sheet["C20"] = ativados_empresa_confidencial

    # Desativado
    sheet["A21"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B21"] = 0
    sheet["C21"] = 0

    sheet["A22"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B22"] = conectados_empresa_confidencial
    sheet["C22"] = ativados_empresa_confidencial

    sheet["A23"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B23"] = conectados_empresa_confidencial
    sheet["C23"] = ativados_empresa_confidencial

    sheet["A24"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B24"] = conectados_empresa_confidencial
    sheet["C24"] = ativados_empresa_confidencial

    sheet["A25"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B25"] = conectados_empresa_confidencial
    sheet["C25"] = ativados_empresa_confidencial

    sheet["A26"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B26"] = conectados_empresa_confidencial
    sheet["C26"] = ativados_empresa_confidencial

    sheet["A27"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B27"] = conectados_empresa_confidencial
    sheet["C27"] = ativados_empresa_confidencial

    sheet["A28"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B28"] = conectados_empresa_confidencial
    sheet["C28"] = ativados_empresa_confidencial

    sheet["A29"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B29"] = conectados_empresa_confidencial
    sheet["C29"] = ativados_empresa_confidencial

    sheet["A30"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B30"] = conectados_empresa_confidencial
    sheet["C30"] = ativados_empresa_confidencial

    sheet["A31"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B31"] = conectados_empresa_confidencial
    sheet["C31"] = ativados_empresa_confidencial

    sheet["A32"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B32"] = conectados_empresa_confidencial
    sheet["C32"] = ativados_empresa_confidencial

    sheet["A33"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B33"] = conectados_empresa_confidencial
    sheet["C33"] = ativados_empresa_confidencial

    sheet["A34"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B34"] = conectados_empresa_confidencial
    sheet["C34"] = ativados_empresa_confidencial

    sheet["A35"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B35"] = conectados_empresa_confidencial
    sheet["C35"] = ativados_empresa_confidencial

    sheet["A36"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B36"] = conectados_empresa_confidencial
    sheet["C36"] = ativados_empresa_confidencial

    sheet["A37"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B37"] = conectados_empresa_confidencial
    sheet["C37"] = ativados_empresa_confidencial

    sheet["A38"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B38"] = conectados_empresa_confidencial
    sheet["C38"] = ativados_empresa_confidencial

    sheet["A39"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B39"] = conectados_empresa_confidencial
    sheet["C39"] = ativados_empresa_confidencial

    sheet["A40"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B40"] = conectados_empresa_confidencial
    sheet["C40"] = ativados_empresa_confidencial

    sheet["A41"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B41"] = conectados_empresa_confidencial
    sheet["C41"] = ativados_empresa_confidencial

    sheet["A42"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B42"] = conectados_empresa_confidencial
    sheet["C42"] = ativados_empresa_confidencial

    sheet["A43"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B43"] = conectados_empresa_confidencial
    sheet["C43"] = ativados_empresa_confidencial
    
    sheet["A44"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B44"] = conectados_empresa_confidencial
    sheet["C44"] = ativados_empresa_confidencial

    sheet["A45"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B45"] = conectados_empresa_confidencial
    sheet["C45"] = ativados_empresa_confidencial

    sheet["A46"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B46"] = conectados_empresa_confidencial
    sheet["C46"] = ativados_empresa_confidencial

    sheet["A47"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B47"] = conectados_empresa_confidencial
    sheet["C47"] = ativados_empresa_confidencial

    sheet["A48"] = 'NOME DA EMPRESA CONFIDENCIAL'
    sheet["B48"] = conectados_empresa_confidencial
    sheet["C48"] = ativados_empresa_confidencial

    arquivo_excel = diretorio_cnct.joinpath(nome)
    workbook.save(arquivo_excel)

    layout = [[sg.Text('....', text_color='white', background_color='#FFFFFF'),
               sg.Text('Dados adquiridos com sucesso!', text_color='#005421', font=('Arial', 10, 'bold'), background_color='#FFFFFF')],
              [sg.Text('Reabra o script para comparar os dados!', text_color='#17202A', font=('Arial', 10, 'bold'), background_color='#FFFFFF')],
              [sg.Text('...............', background_color='#FFFFFF', text_color='#FFFFFF'),
               sg.Text('Clique 2x no botão!', text_color='gray', font=('Arial', 9, 'bold'), background_color='#FFFFFF')],
              [sg.Text('...............', background_color='#FFFFFF', text_color='#FFFFFF'),
               sg.Button('Reabrir script', font=('Arial', 10, 'bold'), button_color='#008735', size=(13, 1))]]

    janela = sg.Window('100%', auto_close=True, background_color='#FFFFFF', auto_close_duration=500, finalize=True, icon=dir_icone).Layout(layout)
    botao, valores = janela.Read()

    while True:
        event, values = janela.read()
        if event == sg.WINDOW_CLOSED:
            break

        if event == 'Reabrir script':
            os.startfile(dir_arquivo)
            janela.close()

if botao == 'Mapear Empresas':
    janela.close()

    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-logging")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-plugins")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_argument("--disable-notifications")
    options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(options=options)

    driver.get("https://empresa.confidencial.com")
    time.sleep(2)

    driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/div[2]/input').send_keys('email_confidencial')
    driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/div[3]/input').send_keys('senha_confidencial')
    driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/form/button').click()
    time.sleep(5)
    
    numero = int(driver.find_element('xpath', '//*[@id="root"]/div[1]/div[2]/div/div[1]/div[1]/div[1]/span/span').text)
    
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    while numero > 0:
        empresa = driver.find_element('xpath', f'//*[@id="root"]/div[1]/div[2]/div/div[1]/div[2]/a[{numero}]/div[2]/div[2]/label').text
        sheet[f"A{numero}"] = empresa
        numero -= 1

    workbook.save(arquivo_excel_empresa)

    driver.quit()

    layout = [[sg.Text('', text_color='white', background_color='#FFFFFF'),
            sg.Text('Dados adquiridos com sucesso!', text_color='#256100', font=('Arial', 10, 'bold'), background_color='#FFFFFF')]]

    janela = sg.Window('100% ...', auto_close=True, background_color='#FFFFFF', auto_close_duration=500, finalize=True, icon=dir_icone).Layout(layout)
    botao, valores = janela.Read()

