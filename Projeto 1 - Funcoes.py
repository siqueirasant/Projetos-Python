from datetime import datetime

'''
Percorre todas as planilhas do arquivo excel, 
Lê cada nome, coloca numa lista
'''
def lista_plan_nomes(arquivo):
    lista = []
    for plan in arquivo.sheetnames:
        lista.append(plan)
    return lista

def extrai_desc_aprop(texto):
    inicio = 'PAGTO A EFETUAR '
    nome_forn = ""
    for caractere in texto:
        nome_forn = nome_forn + caractere
    descricao = nome_forn[len(inicio):60]  
    return descricao

def extrai_desc_baixa(texto):
    inicio = 'PAGTO DO '
    nome_forn = ''
    for caratere in texto:
        nome_forn = nome_forn + caratere
        descricao = nome_forn[len(inicio):53]
    return descricao

def data(texto):
    if type(texto) == str:
        dataf = texto
    else:
        data = datetime.date(texto)
        dia = data.day
        mes = data.month
        ano = data.year
        dataf = f'{dia}/{mes}/{ano}'
    return dataf

def extrai_bx_cli(texto):
    inicio = "RECEBIMENTO "
    nome_cliente = texto[len(inicio):43]
    return nome_cliente

def extrai_provisao_cli(texto):
    provisao = texto[0:31]
    return provisao

def num_nd_cli(texto):
    num = ''
    for caractere in texto:
        if caractere.isdigit() == True:
            num = num + caractere
    return num
