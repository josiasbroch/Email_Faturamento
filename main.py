import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)

print('-' *50)
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' *50)
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' *50)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'josias.broch@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = ''' 
Prezados Clientes,
Segue o Relatório de Vendas por cada Loja.

Faturamento:
{}

Quantidade Vendida:
{}

Ticket Médio dos Produtos em cada Loja:
{}

Qualquer dúvida estou a disposição.

Att.,
Josias
'''
mail.Send()


