import win32com.client as win32

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
