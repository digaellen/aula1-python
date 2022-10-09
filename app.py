import pyautogui
import pyperclip
import time
import pandas

pyautogui.PAUSE = 1

pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')

pyautogui.hotkey('win', 'up')
pyperclip.copy ('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(5)
pyautogui.click(x=412, y=382, clicks=2)

time.sleep(5)
pyautogui.click(x=412, y=382, button='right')
pyautogui.click(x=511, y=927)

time.sleep(10)
tabela = pandas.read_excel(r"C:\Users\Ellen\Downloads\Vendas - Dez.xlsx")
faturamento = tabela ['Valor Final'].sum()
quantidade = tabela ['Quantidade'].sum()

pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(5)

pyautogui.click(x=93, y=242)
pyperclip.copy('ellencamilaa2@gmail.com')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')

texto = f'''
Prezados,
Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.

At.te
Ellen Camila
'''

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')