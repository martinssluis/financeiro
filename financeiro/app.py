import openpyxl
import pyperclip
import pyautogui
from time import sleep

#tamanho da página web em 805 para funcionar as posições
# entrar na planilha

workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = workbook["Produtos"]
#copiar informação e colar no campo correspondente 
for linha in sheet_produtos.iter_rows(min_row=2):

#nome do produto
    nome_poduto = linha[0].value
    pyperclip.copy(nome_poduto)
    pyautogui.click(948,248,duration=1)
    pyautogui.hotkey("ctrl","v")

#descricao
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(928,305,duration=1)
    pyautogui.hotkey("ctrl", "v")

#categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(924,417,duration=1)
    pyautogui.hotkey("ctrl", "v")

#codigo do produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(924,486,duration=1)
    pyautogui.hotkey("ctrl", "v")

#peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(924,549,duration=1)
    pyautogui.hotkey("ctrl", "v")

#dimensoes
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(920,621,duration=1)
    pyautogui.hotkey("ctrl", "v")

#botao proximo
    pyautogui.click(931,669,duration=1)
    sleep(3)
    
 #preco
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(933,263,duration=1)
    pyautogui.hotkey("ctrl", "v")

#quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(932,330,duration=1)
    pyautogui.hotkey("ctrl", "v")

#data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(932,398,duration=1)
    pyautogui.hotkey("ctrl", "v")

#cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(928,473,duration=1)
    pyautogui.hotkey("ctrl", "v")

#tamanho
    tamanho = linha[10].value
    pyautogui.click(938,538,duration=1)
    if tamanho =="Pequeno":
        pyautogui.click(929,582,duration=1)
    elif tamanho =="Médio":
        pyautogui.click(919,582,duration=1)
    else:
         pyautogui.click( 935,602,duration=1)

#material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(936,611,duration=1)
    pyautogui.hotkey("ctrl", "v")

#botao proximo
    pyautogui.click(941,658,duration=1)

#fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(932,285,duration=1)
    pyautogui.hotkey("ctrl", "v")

#pais de origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(938,347,duration=1)
    pyautogui.hotkey("ctrl", "v")

#observações
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(939,406,duration=1)
    pyautogui.hotkey("ctrl", "v")

#codigo de barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(934,524,duration=1)
    pyautogui.hotkey("ctrl", "v")

#localizacao do armazem
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(934,582,duration=1)
    pyautogui.hotkey("ctrl", "v")

#botao concluir
    pyautogui.click(935,639,duration=1)    

#botao ok do alert
    pyautogui.click(1281,164,duration=1) 

#botao adicionaar mais um
    pyautogui.click(1123,464,duration=1) 

#repetir esses passos para outros campos até preencher campos daquela pagina
#clicar em proximo
#repetir os mesmos passos e ir para a proxima pagina
#repetir os mesmos passos e finalizar o cadastro daquele produto e cliclar em concluir
#clilcar em ok, para finalizar o processo
#clicar no ok mais uma vez na mensagem de confirmação de salvamento no banco de dados
#clicar em adicionar mais um e repetir o provesso até finalizar a planilha