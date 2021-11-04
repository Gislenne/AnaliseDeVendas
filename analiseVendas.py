from IPython.core.display import display
import pandas as pd


# Lógica para resolução do Desafio:

#Para resolver o desafio vamos seguir a seguinte lógica:

# - Passo 1 - Importar a base de Dados
# - Passo 2 - Visualizar a Base de Dados para ver se precisamos fazer algum tratamento
# - Passo 3 - Calcular os indicadores de todas as lojas:
  # - Faturamento por Loja
  # - Quantidade de Produtos Vendidos por Loja
  # - Ticket Médio dos Produto por Loja
# - Passo 4 - Calcular os indicadores de cada loja
# - Passo 5 - Enviar e-mail para cada loja
# - Passo 6 - Enviar e-mail para a diretoria

# - Passo 1 - Importar a base de Dados
df = pd.read_excel(r'/home/gislenne/Documentos/AnaliseDeVendas/Vendas.xlsx')
display(df)

# Calculando o Faturamento por Loja (3.1)
faturamento = df[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
faturamento = faturamento.sort_values(by='Valor Final', ascending=False)
display(faturamento)

# Calculando a Quantidade Vendida por Loja (3.2)
quantidade = df[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
quantidade = quantidade.sort_values(by='ID Loja', ascending=False)
display(quantidade)

#  Calculando o Ticket Médio dos Produtos por Loja (3.3)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
ticket_medio = ticket_medio.sort_values(by='Ticket Medio', ascending=False)
display(ticket_medio)

#Criando a função de enviar e-mail
# função enviar_email
import smtplib
import email.message

def enviar_email(resumo_loja, loja, email):

  server = smtplib.SMTP('smtp.gmail.com:587')  
  email_content = f'''
  <p>Atenciosamente Gislenne Moia,</p>
  {resumo_loja.to_html()}
  <p>Fico a disposição para sanar qualquer dúvida! </p>'''

  
  msg = email.message.Message()
  msg['Subject'] = f'Analista Gislenne Moia - Loja: {loja}'
  
  senha="senha do seu email de envio"
  emailenvio = 'seu email de one vai ser enviado'
  msg['From'] = emailenvio
  msg['To'] = email
  password = senha
  msg.add_header('Content-Type', 'text/html')
  msg.set_payload(email_content)
  
  s = smtplib.SMTP('smtp.gmail.com: 587')
  s.starttls()
  # Login Credentials for sending the mail
  s.login(msg['From'], password)
  s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))

#Calculando Indicadores por Loja + Enviar E-mail para todas as lojas

lojas = df['ID Loja'].unique()

for loja in lojas:
  tabela_loja = df.loc[df['ID Loja'] == loja, ['ID Loja', 'Quantidade', 'Valor Final']]
  resumo_loja = tabela_loja.groupby('ID Loja').sum()
  resumo_loja['Ticket Médio'] = resumo_loja['Valor Final'] / resumo_loja['Quantidade']
  enviar_email(resumo_loja, loja, email)

# email para diretoria
tabela_diretoria = faturamento.join(quantidade).join(ticket_medio)
enviar_email(tabela_diretoria, 'Todas as Lojas', email='email da diretoria')