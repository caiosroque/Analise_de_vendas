import pandas as pd
from IPython.display import display
import win32com.client as win32

db = pd.read_csv('vgsales.csv')
display(db)

total_all_groups = db[['Platform','Global_Sales']].groupby('Platform').sum()
display(total_all_groups)

total_all_genres = db[['Genre','Global_Sales']].groupby('Genre').sum()
display(total_all_genres)

print("-" * 50)
print("2006")
sales_in_2006 = db.loc[db['Year'] == 2006, ['Platform', 'Global_Sales']].groupby('Platform').sum()
display(sales_in_2006)

print("-" * 50)
print("2009")
sales_in_2009 = db.loc[db['Year'] == 2009, ['Platform', 'Global_Sales']].groupby('Platform').sum()
display(sales_in_2009)

print("-" * 50)
print("2010")
sales_in_2010 = db.loc[db['Year'] == 2010, ['Platform', 'Global_Sales']].groupby('Platform').sum()
display(sales_in_2010)

print("-" * 50)
print("2011")
sales_in_2011 = db.loc[db['Year'] == 2011, ['Platform', 'Global_Sales']].groupby('Platform').sum()
display(sales_in_2011)

print("-" * 50)
print("Mario Kart Wii in EU")
Mario_eu_wii = db.loc[db['Name'] == 'Mario Kart Wii', ['Platform', 'EU_Sales', 'Year', 'Name']]
display(Mario_eu_wii)


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'emaililustrativo@gmail.com'
mail.Subject = 'Relatorio de vendas'
mail.HTMLBody = f''' 
<p>Prezados,</p>

<p>Segue o relatório de Vendas.</p>

<p>Vendas em 2006:</p>
{sales_in_2006.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Vendas em 2009:</p>
{sales_in_2009.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Vendas em 2010:</p>
{sales_in_2010.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Vendas em 2011:</p>
{sales_in_2011.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Vendas em 2011:</p>
{Mario_eu_wii.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</p>

<p>Att.</p>
<p>Caio Roque</p>
'''

mail.Send()

#mail.Display() não envia na hora
#mail.Send() envia na hora

print("Email Enviado")

