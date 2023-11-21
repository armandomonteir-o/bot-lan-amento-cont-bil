import openpyxl
import pyperclip
import pyautogui
from time import sleep


# entrar na planilha

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
#copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1108,242, duration=1)
    pyautogui.hotkey("ctrl","v")
    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1088,328, duration=1)
    pyautogui.hotkey("ctrl","v")
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1097,460, duration=1)
    pyautogui.hotkey("ctrl","v")
    
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1112,556, duration=1)
    pyautogui.hotkey("ctrl","v")
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1113,642, duration=1)
    pyautogui.hotkey("ctrl","v")
    
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1114,715, duration=1)
    pyautogui.hotkey("ctrl","v")

    pyautogui.click(1087,779, duration=1)
    sleep(3)

    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1081,272, duration=1)
    pyautogui.hotkey("ctrl","v")

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1081,356, duration=1)
    pyautogui.hotkey("ctrl","v")
    
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1086,442, duration=1)
    pyautogui.hotkey("ctrl","v")
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1086,528, duration=1)
    pyautogui.hotkey("ctrl","v")
    
    tamanho = linha[10].value
    pyautogui.click(1159,606, duration=1)
    
    if tamanho == 'Pequeno':
        pyautogui.click(1166,645, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1156,671, duration=1)
    else:
        pyautogui.click(1164,691, duration=1)
    
   
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1062,681, duration=1)
    pyautogui.hotkey("ctrl","v")

    pyautogui.click(1088,757, duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1114,288, duration=1)
    pyautogui.hotkey("ctrl","v")

    pais_origem =linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1116,375, duration=1)
    pyautogui.hotkey("ctrl","v")

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1106,463, duration=1)
    pyautogui.hotkey("ctrl","v")


    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1110,593, duration=1)
    pyautogui.hotkey("ctrl","v")

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1110,682, duration=1)
    pyautogui.hotkey("ctrl","v")

    pyautogui.click(1091,742, duration=1)
    pyautogui.hotkey("enter")
    pyautogui.click(1234,527, duration=1)
    sleep(3)
    

#repetir esses passos para outros campos até preencher campos daquela pagina
#clicar em proxima
#repetir os mesmos passos e ir para a proxima pagina (pag2)
#repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
#clicar em ok para finalizar o processo
#clicar em ok novamente para mais uma confirmação de salvamento no banco de dados
#clicar em adicionar mais um e repetir o processo