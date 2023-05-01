import os
import win32com.client as win32
import pandas as pd

def main():
    # Lê o arquivo Excel a partir do caminho relativo
    df = pd.read_excel('Parametros_Envio_Almox.xlsx')
    
    for i, contato in enumerate(df['EMAIL']):
        indice = df.loc[i,"ID_ARQUIVO "]
        arquivo = df.loc[i, "ANEXO"]

        # Cria uma instância do Outlook dentro de um contexto
        with win32.Dispatch('outlook.application') as outlook:
            email = outlook.CreateItem(0)

            email.To = contato
            email.Subject = f'PDF Almoxarifado {indice}'
            email.HTMLBody = "Olá, segue o PDF do seu holerite."
            email.Attachments.Add(os.path.abspath(arquivo))
            email.Send()

    print('E-mail enviado com sucesso!')

if __name__ == '__main__':
    main()
