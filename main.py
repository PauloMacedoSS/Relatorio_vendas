import pandas as pd
import  win32com.client as win32

# importar a base de dados
import win32com.client

tabela_vendas = pd.read_excel("Vendas.xlsx")

# visualizar a base de dados
pd.set_option('display.max_columns', None)
# No set_option devem ser informados a opção e o valor. O display.max_columns
# é usado para deifinir a quantidade de colunas a ser exibidas.
# O None vai definir que todas as colunas devem ser exibidas.

# faturamento por loja
# Os dois colchetes servem para realizar o filtro na tabela selecionando mais de um coluna
# Se for usar somente uma coluna, só é preciso usar um colchetes.
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
# O groupby vai agrupar todas as lojas e o sum() irá somar o faturamento.
print(faturamento)
print('-' * 50)

# quantidade de produtos vendidos por loja

qtdtot = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtdtot)
print('-' * 50)

# ticket médio por produto em cada loja (faturamento dividido pela quantidade deprodutos vendidos)
ticket_medio = (faturamento['Valor Final'] / qtdtot['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})# Essa linha está alterando o nome da coluna de 0 para Ticket Médio
# O to_frame() serve para transformar o resultado de operações entre colunas em uma tabela
print(ticket_medio)

# enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')# Conectando o python com o Outlook do computador.
mail = outlook.CreateItem(0)# Criando email
# Configurando o email
mail.To = 'Email para o qual você deseja enviar o relatório'
mail.Subject = 'Relatório de vendas por loja'
# O f antes de um texto no python significa fString. Isso diz ao python que o texto pode ter {} e dentro delas pode haver
# variaveis
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{qtdtot.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</p>

<p>Atenciosamente,</p>
<p>Paulo Macêdo.</p>
'''
# As três aspas simples vão permitir que você pode escrever em mais de uma linha.
mail.Send()
# O formatters={'Valor Final': 'R${:,.2f}'.format} irá realizar a formatação dos campos de valores.
print('Email enviado!')