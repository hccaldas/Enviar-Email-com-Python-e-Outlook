# Enviar E-mail com Python

Este repositório contém um script em Python que envia um e-mail automaticamente usando o Microsoft Outlook. O e-mail inclui informações sobre faturamento, quantidade de produtos vendidos e ticket médio, além de um anexo.

## Funcionalidades

- Integração com o Microsoft Outlook para envio de e-mails.
- Envio de e-mail com corpo em HTML.
- Anexar arquivos ao e-mail.

## Pré-requisitos

- Python 3.x
- Microsoft Outlook instalado e configurado.
- Biblioteca `pywin32` instalada. Você pode instalar essa biblioteca usando o comando:
  ```bash
  pip install pywin32

  Como usar
1.Clone este repositório ou copie o código para o seu ambiente local.

2.Certifique-se de que o Microsoft Outlook está instalado e configurado no seu computador.

3.Instale a biblioteca pywin32 se ainda não estiver instalada.

4.Atualize o script com os detalhes do e-mail, como destinatários, assunto, corpo do e-mail e caminho do anexo.

5.Execute o script.

import win32com.client as win32

Exemplo de Uso:

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)

faturamento = 2394
qtde_produtos = 23
ticket_medio = faturamento / qtde_produtos

# configurar as informações do seu e-mail
email.To = "caldashh3@gmail.com; caldashh2@outlook.com"
email.Subject = "E-mail automático do Python para teste"
# definir o corpo do email a partir de HTMLBody
email.HTMLBody = f"""
<p>Olá Hc, vamos começar a montar esse portfólio, e criar o faturamento da loja</p>

<p>O faturamento da loja foi de R${faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>O ticket Médio foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>Código do Hc Pyhton</p>
"""

anexo = r"C:\Users\Usuário\Desktop\PORTIFÓLIO.png"
email.Attachments.Add(anexo)

try:
    email.Send()
    print("Email Enviado")
except Exception as e:
    print(f"Erro ao enviar email: {e}")
