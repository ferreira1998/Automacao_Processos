#importar Bibliotecas

import pandas as pd 
import win32com.client as win32
import pathlib


#importar bases de dados

emails = pd.read_excel(r"C:\Users\wagne\OneDrive\Documentos\Projeto Automação de Processos\Projeto Automacao Indicadores\Bases de Dados\Emails.xlsx")
lojas = pd.read_csv(r"C:\Users\wagne\OneDrive\Documentos\Projeto Automação de Processos\Projeto Automacao Indicadores\Bases de Dados\Lojas.csv", encoding='latin1', sep=';')
vendas = pd.read_excel(r"C:\Users\wagne\OneDrive\Documentos\Projeto Automação de Processos\Projeto Automacao Indicadores\Bases de Dados\Vendas.xlsx")

#print(emails)
#print(lojas)
#print(vendas)

# icluir nome da loja em vendas
#Criar uma atbela para cada Loja

vendas = vendas.merge(lojas, on='ID Loja')
#print(vendas)

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']== loja, :]

#print(dicionario_lojas['Rio Mar Recife'])
#print(dicionario_lojas['Shopping Vila Velha'])

# Definir o dia do indicador

dia_indicador = vendas['Data'].max()
#print(dia_indicador)
#print('{}/{}'.format(dia_indicador.day, dia_indicador.month))


# (identificando a existência da pasta de backup)


caminho_backup = pathlib.Path(r"C:\Users\wagne\OneDrive\Documentos\Projeto Automação de Processos\Projeto Automacao Indicadores\Backup Arquivos Lojas")

arquivo_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivo_pasta_backup]
#print(lista_nomes_backup)

for loja in dicionario_lojas : 
    if loja not in lista_nomes_backup :
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    
    #Salvar dentro da pasta

    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    #print(nome_arquivo)
    local_arquivo = caminho_backup / loja / nome_arquivo

    dicionario_lojas[loja].to_excel(local_arquivo) 


     #definição de metas 

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500




     #Calcular o indicador para 1 loja (depois aplica a todas as lojas)

for loja in dicionario_lojas:
        
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    #print(faturamento_ano)
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print(faturamento_dia)

    #Diversidade de Produtos

    qtde_produtos_ano =  len(vendas_loja['Produto'].unique())
    #print(qtde_produtos_ano)

    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    #print(qtde_produtos_dia)

    #Ticket medio

    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True )
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    #print(ticket_medio_ano)

    #ticket_medio_dia

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    #print(ticket_medio_dia)


    #Enviar E-mail para o gerente 


    outlook = win32.Dispatch('outlook.application')
    
    nome = emails.loc[emails['Loja']== loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']== loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} -  Loja {loja}'
    mail.Body = 'Texto do E-mail'

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'

    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'

    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'

    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'

    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'

    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'


    #Tabela para o Dia e Ano
    mail.HTMLBody = f'''
    <p>Bom Dia, {nome} </p>

    <p> O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi :</p>

    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>

    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>

    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>

    </tr>
    </table>
    <br>
    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>

    </tr>
    <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>

    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>

    </tr>
    </table>




    <p> Segue em anexo a planilha com todos os dados para mais detalhes. </p>

    <p> Qualquer dúvida estou á disposição. </p>
    <p>Att., Ferreira</p>



    '''
    # Anexos (pode colocar quantos quiser):
    attachment  = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()
    print('E-mail da lojas {} enviado'.format(loja))

# Criar Ranking para Diretoria 

faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas.sort_values(by = 'Valor Final', ascending=False)
print(faturamento_lojas_ano)


nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r"C:\Users\wagne\OneDrive\Documentos\Projeto Automação de Processos\Projeto Automacao Indicadores\Backup Arquivos Lojas\Ranking Anual.xlsx")

vendas_dia = vendas.loc[vendas['Data']== dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja') [['Loja','Valor Final']].sum(numeric_only=True)  
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
print(faturamento_lojas_dia)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r"C:\Users\wagne\OneDrive\Documentos\Projeto Automação de Processos\Projeto Automacao Indicadores\Backup Arquivos Lojas\Ranking Dia.xlsx")


#Enviar E-mail para diretoria (Fim do Projeto)

outlook = win32.Dispatch('outlook.application')
    
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']== loja, 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezados, Bom dia 

Melhor loja do Dia em faturamento: Loja {faturamento_lojas_dia.index[0]} com faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em faturamento: loja {faturamento_lojas_dia.index[-1]} com faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do Ano em faturamento: Loja {faturamento_lojas_ano.index[0]} com faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Ano em faturamento: Loja {faturamento_lojas_ano.index[-1]} com  faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

segue em anexo os Rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou a disposição.

Att,
Ferreira
'''
# Anexos (pode colocar quantos quiser):
attachment  = pathlib.Path.cwd() / caminho_backup / (r"C:\Users\wagne\OneDrive\Documentos\Projeto Automação de Processos\Projeto Automacao Indicadores\Backup Arquivos Lojas\Ranking Anual.xlsx")
mail.Attachments.Add(str(attachment))
attachment  = pathlib.Path.cwd() / caminho_backup / (r"C:\Users\wagne\OneDrive\Documentos\Projeto Automação de Processos\Projeto Automacao Indicadores\Backup Arquivos Lojas\Ranking Dia.xlsx")
mail.Attachments.Add(str(attachment))


mail.Send()
print('E-mail da Diretoria enviado')

