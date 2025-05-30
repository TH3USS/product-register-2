import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.Workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1472,413, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1518,305, duration=1) #clikar no botão para passar para a próxima pagina
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_produto = linha[7].value
    pyperclip.copy(quantidade_produto)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(1518,305, duration=1) #abre a lista
    if tamanho == 'Pequeno':
        pyautogui.click(1518,305, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1518,305, duration=1)
    else:
        pyautogui.click(1518,305, duration=1)
    

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1518,305, duration=1) #clikar no botão para passar para a próxima pagina

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_estoque = linha[16].value
    pyperclip.copy(localizacao_estoque)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1518,305, duration=1) #clikar no botão concluir

    pyautogui.click(1518,305, duration=1) #clikar no botão confirmar
    
    pyautogui.click(1518,305, duration=1) #clikar no botão confirmar novamente

    pyautogui.click(1518,305, duration=1) #clikar no botão cadastro de novo