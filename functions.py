
import pyautogui as py
from time import sleep
import json
from tkinter import messagebox


with open('coordenadas.json', 'r') as disc_file:
    disc = json.load(disc_file)





def EntrarNota(nota):
    global status_nota
    status_nota = f'Entrando na nota {nota} '
    sleep(1)
    py.click(disc['entrar nota']) 
    py.write(nota)
    sleep(1)
    py.click(disc['click localizar']) # click  localizar
    


def ExtrairDxf(nota,poste):
    global status_dxf1, status_dxf2
    status_dxf1 = f'Extraindo dxf da nota {nota} '
    
    py.click(disc['Arquivo']) # click  arquivo 
    sleep(1)
    py.click(disc['Exportar Arquivo CAD'], duration=0.2) # click  Exportar DXF
    py.click(disc['click janela dxf'])
    sleep(1)
    py.write(f'{nota}_{str(poste)}')
    sleep(1)  
    py.press('enter')
    sleep(1)
    py.press('enter')
    sleep(2)
    py.click(disc['fecha a janela do dxf']) # fecha a janela dos arquivos baixados após extrair o dxf
    status_dxf2 = f'Dxf da nota {nota} extraído com sucesso '
    
    

def ExtrairLista(nota):
    
    py.click(disc['click em Consultas'])  # Consultas
    sleep(1)
    py.click(disc['Consultar Lista'])  # Consultar Lista        
    py.click(disc['Excel'])  # Excel
    sleep(1)
    py.click(disc['colar nome excel']) #click no campo onde coloca o nome do arquivo excel
    sleep(1)
    py.write(nota) # nomeia o arquivo excel
    sleep(1)
    py.press('enter')
    sleep(1)
    py.click(disc['fecha janela lt']) # click  sair na janela "consultar list tecnica"
    py.click(disc['click fora do proj']) # click Fora do proj
    
    
    
def fecharNota():
    sleep(1)
    py.doubleClick(disc['click fora do proj']) # click dentro do proj 
    py.doubleClick(disc['click fora do proj']) # click dentro do proj
    py.click(disc['Arquivo']) #clickr  arquivo
    sleep(0.5)
    py.click(disc['click abrir projeto']) #click  abrir projeto
    