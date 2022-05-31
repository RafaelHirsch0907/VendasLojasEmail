import pandas as pd
import win32com.client as win32

# importar a base de dados
tabelaVendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)         # RECEBE OPÇÃO E VALOR
print(tabelaVendas)
print('-' * 50)

# faturamento por loja
faturamento = tabelaVendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()         # FILTRANDO POR COLUNAS E AGRUPANDO
print(faturamento)
print('-' * 50)

# quantidade de produtos vendidos por loja
quantidadeProdutosVendidos = tabelaVendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidadeProdutosVendidos)
print('-' * 50)

# ticket médio (faturamento/quantidade) por produto por loja
ticketMedio = (faturamento['Valor Final'] / quantidadeProdutosVendidos['Quantidade']).to_frame()        # RE-TRANSFORMAR O CÁLCULO EM TABELA
ticketMedio = ticketMedio.rename(columns = {0: 'Ticket Médio'})     # Renomeando uma coluna da tabela
print(ticketMedio)

# enviar um email com o relatório (https://stackoverflow.com/questions/6332577/send-outlook-email-via-python)
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'final.arrow92@gmail.com'
mail.Subject = 'Vendas Lojas'       # FORMATTERS: , = MARCADOR DE MILHA; . = MARCADOR DE DECIMAL; 2F = QUANTIDADE DE CASAS DECIMAIS (2)
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters = {'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidadeProdutosVendidos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticketMedio.to_html(formatters = {'Ticket Médio': 'R${:,.2f}'.format})}

<p>Att.</p>
<p>Rafael Hirsch</p>
'''

mail.Send()
print('Email eviado!!!')