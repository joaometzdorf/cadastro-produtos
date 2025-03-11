import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
planilha_produtos = workbook["Produtos"]


# Função que seleciona cada item da planilha
def get_item(item, x, y):
    pyperclip.copy(item)
    pyautogui.click(x, y, duration=0.5)
    pyautogui.hotkey("ctrl", "v")


# Copia cada informação pra passar pro site
for linha in planilha_produtos.iter_rows(min_row=2):
    # Primeira página
    nome_produto = linha[0].value
    get_item(nome_produto, 2245, 388)

    descricao = linha[1].value
    get_item(descricao, 2245, 458)

    categoria = linha[2].value
    get_item(categoria, 2245, 592)

    codigo_produto = linha[3].value
    get_item(codigo_produto, 2245, 677)

    peso = linha[4].value
    get_item(peso, 2245, 763)

    dimensoes = linha[5].value
    get_item(dimensoes, 2245, 848)

    pyautogui.click(2291, 917, duration=0.5)  # Botão próxima página
    sleep(2)

    # Segunda página
    preco = linha[6].value
    get_item(preco, 2241, 397)

    quantidade = linha[7].value
    get_item(quantidade, 2237, 485)

    data_validade = linha[8].value
    get_item(data_validade, 2241, 578)

    cor = linha[9].value
    get_item(cor, 2241, 658)

    tamanho = linha[10].value
    pyautogui.click(2259, 740, duration=0.5)
    if tamanho == "Pequeno":
        pyautogui.click(2251, 781, duration=0.5)
    elif tamanho == "Médio":
        pyautogui.click(2247, 808, duration=0.5)
    else:
        pyautogui.click(2248, 838, duration=0.5)

    material = linha[11].value
    get_item(material, 2247, 831)

    pyautogui.click(2285, 895, duration=0.5)  # Botão próxima página
    sleep(2)

    # Terceira página
    fabricante = linha[12].value
    get_item(fabricante, 2240, 417)

    pais_origem = linha[13].value
    get_item(pais_origem, 2242, 504)

    observacoes = linha[14].value
    get_item(observacoes, 2244, 590)

    codigo_barras = linha[15].value
    get_item(codigo_barras, 2243, 725)

    localizacao_estoque = linha[16].value
    get_item(localizacao_estoque, 2244, 812)

    pyautogui.click(2266, 872, duration=0.5)  # Botão concluir
    pyautogui.click(3043, 197, duration=0.5)  # Botão Ok
    sleep(2)
    pyautogui.click(2905, 658, duration=0.5)
    sleep(1)
