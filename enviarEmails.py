import smtplib
import email.message
import openpyxl

def enviar_email(destinatario, nome):
    # Corpo do Email em HTML 
    corpo_email = f"""
    <p>Olá, {nome} </p>
    """

    msg = email.message.Message()
    msg['Subject'] = 'Assunto'
    msg['From'] = 'Remetente'
    msg['To'] = destinatario
    password = 'Senha gerada pelo gmail' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))

# Caminho para o arquivo da planilha
caminho_planilha = r"caminho para a planilha"


planilha = openpyxl.load_workbook(caminho_planilha)
sheet = planilha.active


destinatarios = []


for linha in sheet.iter_rows(min_row=2, values_only=True):
    nome = linha[0]
    destinatario = linha[1]
    destinatarios.append((destinatario, nome))


for destinatario, nome in destinatarios:
    enviar_email(destinatario, nome)
    print(f'E-mail enviado para {destinatario}')


planilha.close()