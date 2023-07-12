import pandas as pd
import win32com.client as win32
from selenium import webdriver


def envio_email(anexo, assunto, destinatario):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Attachments.Add(anexo)
    # for anx in anexo:
    #    anx = anx.replace("/", "\\")
    #    mail.Attachments.Add(anx)

    mail.cc = 'bbonelli@bancosofisa.onmicrosoft.com;dalves@sofisa.com.br;gelias@sofisa.com.br;wilkerf@sofisa.com.br; pcampos@sofisa.com.br'
    mail.Subject = assunto
    mail.To = destinatario
    mail.HTMLBody = 'Bom dia<br>Segue arquivo de retorno atualizado.<br>Qualquer dúvida fico a disposição.<br><br>Att,<br>Erick Martins'
    #mail.Display()  # aparece na tela
    mail.Send()  # envio direto


if __name__ == '__main__':
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # baixar_limite()

    df = pd.read_excel(r"C:\Users\esmartins\Downloads\Retorno Parcerias.xlsx")
    caminho_padrao = r'C:/Users/esmartins/Documents/leads/'
    lista_parcerias = ['addname']

    dict_email = {'addname'


    }

    for parceria in lista_parcerias:
        nome_arquivo = parceria + ' - Arquivo de retorno Sofisa'
        caminho_temp = caminho_padrao + parceria + r"/" + nome_arquivo + ".xlsx"
        df_temp = df[df["Parceria Comercial"] == parceria]
        destinatario = dict_email[parceria]
        writer = pd.ExcelWriter(caminho_temp)
        df_temp.to_excel(writer, index=False)
        writer.close()
        envio_email(anexo=caminho_temp, assunto=nome_arquivo, destinatario=destinatario)

    print("finalizou")