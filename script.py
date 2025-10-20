import pandas as pd
from pywinauto import Desktop
import pyautogui
from time import sleep, time
import subprocess
import os
import pyperclip
import win32com.client as win32

script = "arquivo_teste.vbs"

with open(script, 'r') as arquivo:
    backup = arquivo.read()

planilha_pedidos = pd.read_excel('pedidos_aprovados.xlsx')

path = r"C:\Users\guilherme.barbosa\OneDrive - Hospital Care Caledonia Saúde S.A\Área de Trabalho\Python\baixa_enviar_po\pedidos"

# Loop para atualizar o arquivo VBS considerando a planilha
for linha in planilha_pedidos.index:

    # Capturar número do pedido na coluna pedidos
    pedido = str(planilha_pedidos.loc[linha, 'pedidos'])

    if str(planilha_pedidos.loc[linha, 'tipo']) == "NB":
        bloco_sap = f"""
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "ME9F"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").text = "{pedido}"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/chk[1,5]").selected = true
    session.findById("wnd[0]/tbar[1]/btn[16]").press

    WScript.Sleep 8000
    """
        with open(script, 'a') as arquivo:
            arquivo.write(bloco_sap)

    else:
        bloco_sap = f"""
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "me29n"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = "{pedido}"
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[21]").press
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").setFocus
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").caretPosition = 0
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").columns.elementAt(0).width = 4
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").setFocus
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").caretPosition = 0
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/usr/chkNAST-DELET").selected = true
    session.findById("wnd[0]/usr/chkNAST-DELET").setFocus
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/btn[15]").press

    WScript.Sleep 8000
    """
        with open(script, 'a') as arquivo:
            arquivo.write(bloco_sap)
    
# Executar arquivo SAP
subprocess.Popen(["wscript", script], shell=True)

# Função para esperar a Janela de salvar pedido
def esperar_janela_salvar(timeout=45):
    inicio = time()
    while time() - inicio < timeout:
        for backend in ["uia", "win32"]:
            try:
                janelas = Desktop(backend=backend).windows()
                for janela in janelas:
                    if "Salvar" in janela.window_text() and "como" in janela.window_text():
                        try:
                            janela.set_focus()
                            return True
                        except:
                            continue
            except:
                pass
        sleep(0.5)
    return False


# Loop para escrever o diretorio e nome de cada pedido
for linha in planilha_pedidos.index:

    if esperar_janela_salvar(timeout=40):
        pyperclip.copy(os.path.join(path, f'{str(planilha_pedidos.loc[linha, 'pedidos'])}.pdf'))
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

# Fazer backup do arquivo VBS
with open(script, 'w') as arquivo:
    arquivo.write(backup)


sleep(2)

# Enviar e-mail

email_compradores = {
    "213": "guilherme.barbosa@evangelicohospital.com.br",
    "123": "compras@evangelicohospital.com.br"
}


# Caminho da pasta com os PDFs
pasta_pedidos = r"C:\Users\guilherme.barbosa\OneDrive - Hospital Care Caledonia Saúde S.A\Área de Trabalho\Python\baixa_enviar_po\pedidos"

# Caminho da planilha Excel
planilha_pedidos = "pedidos_aprovados.xlsx"

# Lê a planilha com pandas
df = pd.read_excel(planilha_pedidos)

# Inicializa o Outlook
outlook = win32.Dispatch('outlook.application')

# Percorre cada linha da planilha
for _, linha in df.iterrows():
    numero_pedido = str(linha['pedidos']).strip()

    email_comprador = email_compradores[str(linha['gcm']).strip()]


    # Monta o caminho do arquivo PDF
    caminho_pdf = os.path.join(pasta_pedidos, f"{numero_pedido}.pdf")

    if os.path.exists(caminho_pdf):
        # Cria o e-mail
        email = outlook.CreateItem(0)
        email.To = email_comprador
        email.Subject = f"Pedido de Compra {numero_pedido}"
        email.Body = f"""
Prezados,

Segue em anexo o PDF referente ao Pedido de Compra {numero_pedido}.

Atenciosamente,
Guilherme Costa Barbosa
"""

        # Anexa o arquivo
        email.Attachments.Add(caminho_pdf)
        print(f"Anexando {numero_pedido}.pdf para {email_comprador}...")

        # Envia o e-mail
        email.Send()
        print(f"E-mail enviado com sucesso para {email_comprador}\n")

    else:
        print(f"Arquivo {numero_pedido}.pdf não encontrado na pasta.\n")

print("Todos os e-mails foram processados!")

# Deletar pedidos da pasta

sleep(5)

comando_del = f'del /Q "{os.path.join(pasta_pedidos, "*.pdf")}"'

try:
    subprocess.run(comando_del, shell=True, check=True)

except subprocess.CalledProcessError as e:
    print(e)
