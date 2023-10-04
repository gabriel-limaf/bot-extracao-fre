import pyautogui
import pyperclip
from time import sleep
import webbrowser
import csv
import re
import os


new_csv = []
with open('Empresas Listadas B3 - Novo Mercado.csv', 'r', encoding='utf-8') as arquivo_csv:
    leitor_csv = csv.reader(arquivo_csv)
    next(leitor_csv, None)
    for linha in leitor_csv:
        empresa = linha[3]
        print(empresa)
        url = 'https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/empresas-listadas.htm'
        webbrowser.open(url)
        sleep(3)
        pyautogui.hotkey('ctrl', 'f'), sleep(1)
        pyperclip.copy('Nome da Empresa')
        pyautogui.hotkey('ctrl', 'v'), sleep(1)
        pyautogui.hotkey('esc'), sleep(1)
        pyautogui.hotkey('tab'), sleep(1)
        pyperclip.copy(empresa)
        pyautogui.hotkey('ctrl', 'v'), sleep(1)
        pyautogui.hotkey('tab'), sleep(1)
        # Seleciona o segmento da B3 para acessar a empresa
        if linha[6] == 'Novo Mercado':
            pyautogui.hotkey('n'), sleep(1)
        if linha[6] == 'Nível 1 de Governança Corporativa':
            pyautogui.hotkey('n'), sleep(1)
            pyautogui.hotkey('n'), sleep(1)
            pyautogui.hotkey('n'), sleep(1)
        if linha[6] == 'Nível 2 de Governança Corporativa':
            pyautogui.hotkey('n'), sleep(1)
            pyautogui.hotkey('n'), sleep(1)
        if linha[6] == 'Tradicional - Bolsa':
            pyautogui.hotkey('t'), sleep(1)
        pyautogui.hotkey('tab'), sleep(1)
        pyautogui.hotkey('enter'), sleep(1)
        pyautogui.moveTo(345, 770), sleep(1)
        pyautogui.click(), sleep(1)
        pyautogui.hotkey('tab'), sleep(1)
        pyautogui.hotkey('r'), sleep(1)
        pyautogui.hotkey('ctrl', 'a'), sleep(1)
        pyautogui.hotkey('ctrl', 'c'), sleep(1)
        text = pyperclip.paste()
        word_to_find = 'Formulário de Referência - FRE'
        if re.search(word_to_find, text, re.IGNORECASE):
            linha[4] = f"A palavra '{word_to_find}' foi encontrada na página."
            escritor_csv = csv.writer(arquivo_csv)
            pyautogui.hotkey('ctrl', 'f'), sleep(1)
            pyperclip.copy('Formulário de Referência - FRE')
            pyautogui.hotkey('ctrl', 'v'), sleep(1)
            pyautogui.hotkey('esc'), sleep(1)
            pyautogui.hotkey('tab'), sleep(1)
            pyautogui.hotkey('enter'), sleep(2)
            # Localiza e clica no botão de salvar pdf
            pdf_button = pyautogui.locateOnScreen('pdf_button.png')
            pyautogui.moveTo(pdf_button), sleep(1)
            pyautogui.click(), sleep(2)
            # Localiza e clica no botão de download
            download_button = pyautogui.locateOnScreen('download_button.png')
            pyautogui.moveTo(download_button), sleep(1)
            pyautogui.click(), sleep(150)
            pyautogui.hotkey('ctrl', 'w'), sleep(1)
            pyautogui.hotkey('ctrl', 'w')
            downloads_folder = os.path.expanduser("~/Downloads")  # Caminho da pasta de downloads no sistema

            # Lista todos os arquivos na pasta de downloads que são arquivos PDF
            pdf_files = [arquivo for arquivo in os.listdir(downloads_folder) if arquivo.endswith('.pdf')]

            # Verifica se há arquivos PDF na lista
            if pdf_files:
                # Crie uma lista de tuplas (nome do arquivo, tempo de criação)
                pdf_files_with_timestamp = [(arquivo, os.path.getctime(os.path.join(downloads_folder, arquivo))) for
                                            arquivo in pdf_files]

                # Ordene a lista pelo tempo de criação (em ordem decrescente)
                pdf_files_with_timestamp.sort(key=lambda x: x[1], reverse=True)

                # Pegue o caminho do último arquivo PDF baixado
                arquivo_mais_recente = pdf_files_with_timestamp[0]
                caminho_do_arquivo = os.path.join(downloads_folder, arquivo_mais_recente[0])
                linha[5] = caminho_do_arquivo
                new_csv.append(linha)
            else:
                linha[5] = "Nenhum arquivo PDF foi encontrado na pasta de downloads."

        else:
            linha[4] = f"A palavra '{word_to_find}' não foi encontrada na página."
            pyautogui.hotkey('ctrl', 'w'), sleep(1)
            new_csv.append(linha)

# Salvar o arquivo
with open("results.csv", "w", newline="", encoding="utf-8") as arquivo_csv:
    escritor_csv = csv.writer(arquivo_csv, delimiter=",")
    cabecalho = ['Razão Social', 'Nome de Pregão', 'Segmento', 'Código', 'Status bot', 'ILP']
    escritor_csv.writerow(cabecalho)
    for linha in new_csv:
        escritor_csv.writerow(linha)
