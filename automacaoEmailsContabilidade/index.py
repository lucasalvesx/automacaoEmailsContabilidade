# instalar biblioteca -> pip install pywin32
# email rementente deve estar associado ao outlook!!

# importar biblioteca win32com 
import win32com.client as win32

# integrando python ao outlook:
outlook = win32.Dispatch('outlook.application')

# criando mensagem de email
email = outlook.CreateItem(0)

# declaracao das variaveis

#configurando as infos do email:
email.To = "d1; d2..."
email.Subject = "mensagem"
email.HTMLBody = f'''

'''

# enviar mensagem ao destinatario selecionado
email.Send()
print("Enviado com sucesso! MIC 2024")