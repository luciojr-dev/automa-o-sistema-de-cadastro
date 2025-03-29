import pyautogui as pa
import os
import time
import pandas as pd
import openpyxl
import re
import subprocess
import tkinter as tk
from tkinter import messagebox

# 1° PARTE - ABRIR E CONECTAR USUARIO SIS
pa.FAILSAFE = False
pa.PAUSE = 1

caminho_arquivo = r'C:\\Users\\Rede Casa Nobre\\Downloads\\SIS.rdp'
os.system(f'start /B mstsc "{caminho_arquivo}"')
time.sleep(10)

janela = None
while janela is None:
    janela = pa.getWindowsWithTitle('Área de Trabalho Remota')
    time.sleep(1)
if janela:
    janela[0].activate()

# COLOCAR USUARIO E SENHA DO SIS
pa.click(918,561, duration=1) # digita usuario
pa.write("beatrizNapoli")
time.sleep(1)
pa.click(911,588, duration=1) # digita a senha
pa.write("142536")
time.sleep(1)
pa.click(934,625,duration=1) # clica em ok
time.sleep(15)

# Fechar janelas desnecessárias
for _ in range(3):
    pa.press("esc")
    time.sleep(10)

# 2° PARTE - NAVEGAR E EXPORTAR O ARQUIVO CSV
pa.click(35,179, duration=2)  # CLIENTES
time.sleep(10)
pa.click(875,197, duration=2)  # RELATORIOS
time.sleep(2)
pa.click(907,303, duration=2)  # EXPORTAR LISTAS
time.sleep(2)
pa.click(1012,747, duration=2)  # Gravar CSV
time.sleep(20)

# 3° PARTE - SALVAR O ARQUIVO CSV EM UM LUGAR ESPECIFICO
pa.click(85,376, duration=1) # entrar no disco c
time.sleep(5)
pa.doubleClick(202,241, duration=1) # entrar na pasta SIS
time.sleep(5)
pa.click(233,128, duration=1) # clicar na plhanilha para gravar csv
time.sleep(2)
pa.press("enter")
pa.click(1022,553, duration=1) # clicar sim para confirmar substituição

# CLICAR NA JANELA QUE APARECER
pa.click(1220,323, duration=1)

# TEMPO DE ESPERA POIS O SISTEMA TRAVA ATE SAIR AS JANELAS
time.sleep(300)

# FECHAR JANELA
pa.press("esc")
time.sleep(5)

#  FECHAR SISTEMA SIS 
pa.click(1878,1056, duration=1)
time.sleep(2)
pa.press("enter")
time.sleep(15)

# PROCESSO PARA ABRIR E MANIPULAR PLANILHAS

#Caminhos dos arquivos
arquivo_base2910 = r"C:\SIS\base2910.csv"
arquivo_base_contatos = r"C:\SIS\BASE-CONTATOS-27-11 (2).xlsm"

# Função para abrir o arquivo e maximizar
def abrir_planilha(caminho):
    # Abre o arquivo no Excel
    subprocess.Popen(["start", "excel", caminho], shell=True)
    time.sleep(5)  # Espera inicial para garantir que abriu

    # Pressiona ALT + ESPAÇO + X para maximizar a janela
    pa.hotkey("alt", "space")
    time.sleep(0.5)
    pa.press("x")  

# Abre a planilha base2910 e maximiza
abrir_planilha(arquivo_base2910)

# Espera 25 segundos
time.sleep(30)

# Abre a planilha base contatos e maximiza
abrir_planilha(arquivo_base_contatos)

time.sleep(30)

# pressionar enter nas janelas que entrar quando carregar a planilha
pa.click(885,592, duration=1) # clicar em sim
time.sleep(5)
pa.press("enter")
time.sleep(5)
pa.click(30,22, duration=1)# clicar em salvar planilha
time.sleep(5)

# Rolar a planilha ate o começo
pa.hotkey("ctrl", "home")
time.sleep(5)

# clicar no botao de copaire da macro
pa.click(969,273, duration=1)
time.sleep(30)

# Selecionar da coluna A ate K e colar os dados
pa.click(18,211, duration=1)
time.sleep(2)
pa.hotkey("ctrl", "c")
time.sleep(5)

# abrindo o chrome atraves do win + r
pa.hotkey("win","r")
time.sleep(1)
pa.press("enter")
pa.doubleClick(852,43, duration=1)
time.sleep(3)

# Clica no sheets na barra de favoritos
pa.click(127,105, duration=1)
time.sleep(15)

# colar os dados no sheets
pa.click(95,273, duration=1)
pa.hotkey("ctrl", "shift", "v")
time.sleep(5)
pa.hotkey("ctrl", "shift", "v")
time.sleep(60)

# Exibir mensagem ao finalizar
def mostrar_mensagem():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    messagebox.showinfo("Sucesso!", "Automação realizada com sucesso!")
    root.destroy()  # Fecha o Tkinter corretamente

# Chamar a função de exibição de mensagem ao final
mostrar_mensagem()

print("processo e automaçao realizada com sucesso!!!")