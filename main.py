import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns = {0: 'Ticket Médio'})


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'wagnerandrade.dev@gmail.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas de cada loja.</p>

<p>Faturamento:</p> 
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket médio dos produtos de cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}


<p>Qualquer dúvida estou a disposição, Wagner Andrade.</p>

'''

mail.Send()
print('Email enviado')