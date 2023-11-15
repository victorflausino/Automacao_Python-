# pacotes imoportados

import openpyxl
import pyperclip
import pyautogui
from time import sleep

# encontrar planilha com os dados
workbook = openpyxl.load_workbook('Cópia de produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# encontrar planilha produtos e começar da 2 linha
for linha in sheet_produtos.iter_rows(min_row=2):
    # encontrar valor (no caso nome) na planilha
    nome_produto = linha[0].value
    # copiar valor da planilha
    pyperclip.copy(nome_produto)
    # clicar nas coordenadas descritas pelo python > from mouseInfo import mouseInfo > mouseInfo() e  aguardar 1s.
    pyautogui.click(807, 194, duration=1)
    # colar na coordenada descrita acima
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(823, 284, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(810, 411, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(799, 499, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(818, 585, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensao = linha[5].value
    pyperclip.copy(dimensao)
    pyautogui.click(803, 676, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(831, 717, duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(828, 210, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    qt = linha[7].value
    pyperclip.copy(qt)
    pyautogui.click(828, 306, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dt_validade = linha[8].value
    pyperclip.copy(dt_validade)
    pyautogui.click(785, 390, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(769, 480, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(784, 562, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(826, 596, duration=1)
    elif tamanho == 'Medio':
        pyautogui.click(836, 615, duration=1)
    else:
        pyautogui.click(825, 639, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(787, 647, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(784, 706)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(804, 237, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(815, 317, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    obs = linha[14].value
    pyperclip.copy(obs)
    pyautogui.click(827, 412, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cod_barras = linha[1].value
    pyperclip.copy(cod_barras)
    pyautogui.click(859, 548, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    local_armazem = linha[15].value
    pyperclip.copy(local_armazem)
    pyautogui.click(803, 626, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(791, 683)
    sleep(2)

    pyautogui.click(1200, 191)
    sleep(2)

    pyautogui.click(1022, 461)
