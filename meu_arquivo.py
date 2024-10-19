#importar base de dados
import pandas as pd
base_dados = pd.read_excel("Vendas.xlsx")

#visualizar a base de dados
pd.set_option('display.max_columns', None)
print(base_dados)
#faturamento por loja
faturamento = base_dados[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#quantidade de produtos vendidos por loja
qtd_vendidos = base_dados[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtd_vendidos)

print('-'*50)
#ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtd_vendidos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

#enviar um email com o relatorio
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'natasha.souza@telefonica.com'
mail.Subject = 'Relatório de Vendas por Loja '
mail.HTMLBody = f'''
<p>Prezados,</p>

<h2>Segue o relatório de vendas por cada Loja.</h2>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de produtos vendidos por loja:</p>
{qtd_vendidos.to_html()}

<p>Ticket Médio por loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer Dúvida estou a disposição.</p>
<p>Att.,</p>
<p>Emerson Moura</p>

'''

mail.Send()