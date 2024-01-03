# Importando Bibliotecas
import pandas as pd
import win32com.client as win32

# Importando tabela
vendas = pd.read_excel("data/Vendas.xlsx")
pd.set_option('display.max_columns', None)

#Criando o Faturamento
faturamento = vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

#Criando a Quantidade
quantidade = vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

#Criando o Ticket Medio
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})

#Enviando Email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'cursosthiz@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Thiago</p>
'''

mail.Send()
print('Email Enviado')
