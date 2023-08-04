# Ler dados da planilha
# Inserir cada célula de cada linha em um campo do sistema
import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('BrunoBorges1996/projeto-bot-devaprender/vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linhas in vendas_sheet.iter_rows(min_row = 2):
    pyautogui.click(1120,624, duration = 2)
    pyautogui.write(linhas[0].value)
    pyautogui.click(1117,650, duration = 2)
    pyautogui.write(linhas[1].value)
    pyautogui.click(1132,677, duration = 2)
    pyautogui.write(str(linhas[2].value))
    pyautogui.click(1201,702, duration = 2)
    pyautogui.write(linhas[3].value)
    pyautogui.click(1075,726, duration = 2)
    pyautogui.click(938,591, duration = 2)
