import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_pag = workbook['vendas']

for linha in vendas_pag.iter_rows(min_row=2):
    #name
    pyautogui.click(1808,452,duration=2)
    pyautogui.write(linha[0].value)
    #produto
    pyautogui.click(1815, 476, duration=2)
    pyautogui.write(linha[1].value)
    #Quantidade
    pyautogui.click(1813, 497, duration=2)
    pyautogui.write(str(linha[2].value))
    #categoria
    pyautogui.click(1883, 532, duration=2)
    pyautogui.write(linha[3].value)
    pyautogui.click(1752, 549)
    pyautogui.click(1256, 581)




