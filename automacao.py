import openpyxl
import pyperclip
import pyautogui
#Copiar informação de um campo e colaRua Raquel Barthou, 27r no seu camRua Raquel Barthou, 27po correspondente
#Repetir esses passos em todos os campos até preencher campos daquela página
#Clicar em limparEstância Velha
#Repetir os mesmos passos


#Entra na planilha
workbook = openpyxl.load_workbook('planilhaauto.xlsx')

#Seleciona a página da planilha que vai pegar informações
sheet_info = workbook['info']

for linha in sheet_info.iter_rows(min_row=2):
    
    nome = linha[0].value #pega informações de cada coluna e por em uma variável
    pyperclip.copy(nome) #copia a informação da célula da planilha, está sendo usado pyperclip pois o ctrl c da biblioteca não suporta acentos
    pyautogui.click(1071,160, duration=1) #determina um clic na determinada cordenada
    pyautogui.hotkey('ctrl', 'v') #cola a informação copiada no campo clicado anteriormente

    #Instalar mouseinfo: pip/pip3 install mouseinfo
    #Abrir mouseinfo: python no terminal -> from mouseinfo import mouseInfo -> mouseInfo()

    endereco = linha[1].value
    pyperclip.copy(endereco)
    pyautogui.click(1094,213, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    cidade = linha[2].value
    pyperclip.copy(cidade)
    pyautogui.click(1077,266, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')


    estado = linha[3].value
    pyperclip.copy(estado)
    pyautogui.click(1081,310, duration=0.5)
    if estado == 'Acre':
        pyautogui.click(1088,335, duration=0.5)
    elif estado == 'Alagoas':
        pyautogui.click(1115,363, duration=0.5)
    elif estado == 'Amapá':
        pyautogui.click(1085,377, duration=0.5)
    elif estado == 'Amazonas':
        pyautogui.click(1073,396, duration=0.5)
    else:
        pyautogui.click(1094,421, duration=0.5)

    natureza_do_cargo = linha[4].value
    pyperclip.copy(natureza_do_cargo)
    if natureza_do_cargo == 'Gerência':
        pyautogui.click(1005,445, duration=0.5)
    elif natureza_do_cargo == 'Financeiro':
        pyautogui.click(1112,450, duration=0.5)
    elif natureza_do_cargo == 'Recepção':
        pyautogui.click(1235,446, duration=0.5)
    elif natureza_do_cargo == 'Administrativo':
        pyautogui.click(1348,451, duration=0.5)
    else:
        pyautogui.click(1504,444, duration=0.5)

    area_de_interesse = linha[5].value
    pyperclip.copy(area_de_interesse)
    if area_de_interesse == 'Computação':
        pyautogui.click(1005,520, duration=0.5)
    elif area_de_interesse == 'Biologia':
        pyautogui.click(1141,520, duration=0.5)
    elif area_de_interesse == 'Meio Ambiente':
        pyautogui.click(1240,519, duration=0.5)
    elif area_de_interesse == 'Engenharia':
        pyautogui.click(1404,519, duration=0.5)
    else:
        pyautogui.click(1530,518, duration=0.5)

    mini_curriculo = linha[6].value
    pyperclip.copy(mini_curriculo)
    pyautogui.click(1140,615, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1063,733, duration=0.5)

    