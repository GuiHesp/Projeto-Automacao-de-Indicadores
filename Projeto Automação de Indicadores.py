#!/usr/bin/env python
# coding: utf-8

# ### Passo 1: Importar bibliotecas
# 

# In[ ]:


import pandas as pd
import pathlib


# In[ ]:


emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', sep=';', encoding="latin")
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

display(emails)
display(lojas)
display(vendas)


# ### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

# In[ ]:


vendas = vendas.merge(lojas, on="ID Loja")
display(vendas)


# In[ ]:


dict_lojas = {}
for loja in lojas["Loja"]:
    dict_lojas[loja] = vendas.loc[vendas["Loja"]==loja, :]
    
display(dict_lojas["Salvador Shopping"])


# In[ ]:


dia_indicador = vendas["Data"].max()
print(dia_indicador)
print(f'{dia_indicador.day}/{dia_indicador.month}')


# ### Passo 3 - Salvar a planilha na pasta de backup

# In[ ]:


# verificar se já existe

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dict_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

    #salvar dentro da pasta
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.day, dia_indicador.month, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dict_lojas[loja].to_excel(local_arquivo)


# ### Passo 4 - Calcular o indicador para 1 loja

# In[ ]:


meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000

meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120

meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500


# In[ ]:


for loja in dict_lojas:
    vendas_loja = dict_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja["Data"]==dia_indicador, :]
    
    #faturamento
    faturamento_ano = vendas_loja["Valor Final"].sum()
    faturamento_dia = vendas_loja_dia["Valor Final"].sum()
    
    #qtdade
    qtde_produtos_ano = len(vendas_loja["Quantidade"].unique())
    qtde_produtos_dia = len(vendas_loja_dia["Quantidade"].unique())
    
    #ticket medio
    valor_venda = vendas_loja.groupby("Código Venda").sum()
    ticket_medio_ano = valor_venda["Valor Final"].mean()
    
    valor_venda_dia = vendas_loja_dia.groupby("Código Venda").sum()
    ticket_medio_dia = valor_venda_dia["Valor Final"].mean()
    
    #enviando email
    import smtplib
    import email.message

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = "green"
    else:
        cor_fat_dia = "red"
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = "green"
    else:
        cor_fat_ano = "red"

    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = "green"
    else:
        cor_qtde_dia = "red"
    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = "green"
    else:
        cor_qtde_ano = "red"

    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = "green"
    else:
        cor_ticket_dia = "red"
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = "green"
    else:
        cor_ticket_ano = "red"

    def enviar_email():  
        corpo_email = f'''
    <p>Bom dia, <strong>{nome}</strong></p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da Loja <strong>{loja}</strong> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:,.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:,.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:,.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:,.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:,.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:,.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:,.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:,.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>

    <p>Seguem em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>

    '''

        msg = email.message.Message()
        msg['Subject'] = f"OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}"
        msg['From'] = 'guihesp@gmail.com'
        msg['To'] = 'guihesp@gmail.com'
        password = 'password' 
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email )

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('Email enviado')
        
    enviar_email()
    


# ### Passo 7 - Criar ranking para diretoria

# In[ ]:


faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)
display(faturamento_lojas_ano)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas["Data"]==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
nome_arquivo = '{}_{}_Ranking Dia Atual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

display(faturamento_lojas_dia)


# In[ ]:


import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def enviaemail(emails): 
  host = 'smtp.gmail.com'
  port = '587'
  login = 'guihesp@gmail.com'
  senha = 'password'
  
  server = smtplib.SMTP(host,port)
  
  server.ehlo()
  server.starttls()
  server.login(login,senha)
  
  
  corpo = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:,.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:,.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:,.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:,.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Lira
'''
  
  email_msg = MIMEMultipart()
  
  email_msg['From'] =  'guihesp@gmail.com' #Quem está mandando o email
  
  email_msg['To'] = emails
  
  email_msg['Subject'] = 'Ranking Anual e Diário dos Lojistas.'
  
  email_msg.attach(MIMEText(corpo,'plain'))
  
  caminho_arquivo = 'Backup Arquivos Lojas/12_26_Ranking Anual.xlsx'
  attchment = open(caminho_arquivo,'rb')
  
  
  att = MIMEBase('application', 'octet-stream')
  att.set_payload(attchment.read())
  encoders.encode_base64(att)
  
  
  att.add_header('Content-Disposition', f'attachment; filename=arquivoaSerEnviado.csv')
  attchment.close()
  
  
  email_msg.attach(att)
  
  server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
  
  server.quit()
  
  
#emails = [emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]]

for i in emails:
    enviaemail(i)

