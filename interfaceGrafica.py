import pyautogui
import win32com.client
import subprocess
import time
import os

pyautogui.alert("ATENÇÃO, FECHE TODAS AS JANELAS DO SAP ANTES DE PROSSEGUIR COM ESSE SCRIPT.") #Inicia o programa com a mensagem citada
opcao = pyautogui.confirm("Olá, selecione um das variantes e aguarde a execução:", buttons = ['iw392', 'cm333', 'SAIR']) #Seleciona um dos botões para começar ou sair para interromper o programa
pyautogui.PAUSE = 0.5

path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe" #Localiza o diretorio padrão do SAP
subprocess.Popen(path)
time.sleep(5)

SapGuiAuto = win32com.client.GetObject("SAPGUI") #Abre o SAP GUI (interface gráfica do SAP)
if not type(SapGuiAuto) == win32com.client.CDispatch:
    print("dgdfg")

application = SapGuiAuto.GetScriptingEngine #Seleciona qual a conexão do servidor SAP
connection = application.OpenConnection("#", True)
time.sleep(3)

session = connection.Children(0) #Cria a sessão de conexão atual
time.sleep(2)

if opcao == "iw392":
    session.StartTransaction(Transaction="iw39") #Executa a transação selecionada
    os.system("iw392.vbs") #Executa o script da transação correspondente

if opcao == "cm333":
    session.StartTransaction(Transaction="cm33") #Executa a transação selecionada
    os.system("cm333.vbs") #Executa o script da transação correspondente

if opcao == "SAIR":
    quit('SAIR') #Encerra o programa
