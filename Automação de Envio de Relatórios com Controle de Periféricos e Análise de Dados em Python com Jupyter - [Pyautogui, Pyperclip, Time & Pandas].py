#!/usr/bin/env python
# coding: utf-8

# In[1]:


get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pandas')


# In[44]:


import pyautogui as pg
import pyperclip as pclip
import pandas as pd
import time
pg.PAUSE = 1
# pyautogui.PAUSE cria uma Pausa de x segundos para cada Comando da Biblioteca
pg.hotkey("ctrl","t")
# pyautogui.hotkey() realiza o Controle Automático do Teclado
pclip.copy("--Link que leva até a Base de Dados--")
pg.hotkey("ctrl","v")
pg.press("enter")
time.sleep(5)
# time.sleep() dá uma Pausa de x segundos em um determinado Comando
pg.click(x=389, y=295, clicks=4)
# pyautogui.click() realiza o Controle Automático do Mouse
# Usar pg.position() para obter Posição do Mouse no momento
time.sleep(10)
pg.click(x=380, y=300, clicks=4)
time.sleep(15)
pg.click(x=380, y=300, clicks=4)
time.sleep(15)
pg.click(x=126, y=143)
pg.click(x=147, y=421)
pg.click(x=486, y=429)
time.sleep(15)
tabela = pd.read_excel(r"Caminho do Arquivo do Disco Local até a Planilha .xlsx")
# pandas.read_excel() faz a Leitura de uma Tabela de Excel
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
pg.hotkey("ctrl","t")
pclip.copy("https://mail.google.com/mail/u/0/#inbox")
# pyperclip.copy() copia texto sem alteração de Formatação
pg.hotkey("ctrl","v")
pg.press("enter")
# pyautogui.press() aperta algum botão do teclado
time.sleep(40)
pg.click(x=185,y=189)
time.sleep(10)
pg.write("fulaninhodetal@gmail.com")
time.sleep(5)
pg.press("tab")
time.sleep(5)
pg.press("tab")
time.sleep(5)
pclip.copy("Relatório de Vendas")
pg.hotkey("ctrl","v")
time.sleep(5)
pg.press("tab")
time.sleep(5)
txt = f"""
Prezados,

Segue o Relatório de Vendas:
Faturamento: R${faturamento:,.2f}
Quantidade de Produtos Vendidos: {quantidade:,}
Qualquer dúvida, estou à disposição.

Atenciosamente,
Lucas Gonçalves de Oliveira Martins
"""
pclip.copy(txt)
pg.hotkey("ctrl","v")
pg.hotkey("ctrl","enter")


# In[37]:


import pyautogui as pg
pg.position()





