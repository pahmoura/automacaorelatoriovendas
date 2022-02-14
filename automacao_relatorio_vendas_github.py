import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.hotkey("ctrl", "t" )
pyperclip.copy("")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep (4)

# Passo 2: Navegar até o local do relatório 
pyautogui.click(x=2829, y=339, clicks=2)
time.sleep (2)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=2844, y=334)
pyautogui.click(x=4544, y=215)
pyautogui.click(x=4256, y=755)

time.sleep(4)

# Passo 4: Calcular os indicadores
import pandas as pd

tabela = pd.read_excel(r"")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

# Passo 5: Entrar no e-mail
pyautogui.hotkey("ctrl", "t" )
pyperclip.copy("")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(4)

# Passo 6: Enviar por e-mail o resultado
pyautogui.click(x=2444, y=240)
pyautogui.write("")
pyautogui.press("tab")
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de:{quantidade:,}

Abs
Paloma"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
time.sleep(3)

pyautogui.hotkey("ctrl", "enter")







