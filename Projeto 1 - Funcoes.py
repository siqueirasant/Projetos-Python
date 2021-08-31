#Função que lê os caracteres de uma string e extrai o número da nota.

def extrai_nf(texto):
    nota = ''
    for x in texto:
        if x.isdigit() == True:
            nota = nota + x
    return nota


#Função que lê os caracteres de uma string e extrai o nome do fornecedor.
def extrai_forn(texto):
    inicio = 'PAGTO A EFETUAR DOC. NRO. DO FORN.  '
    novo_texto = ""
    for caractere in texto:
        if caractere.isdigit() == False:
            novo_texto = novo_texto + caractere
    fornecedor = novo_texto[len(inicio):56]
    return fornecedor

'''
Percorre todas as planilhas do arquivo excel, 
Le cada nome, coloca numa lista
'''
def lista_plan_nomes():
    lista = []
    for plan in arquivo.sheetnames:
        lista.append(plan)
    return lista

#Função que calcula a identificação da célula
def calcula_num(coluna, max_linha):
    num = 1
    lista_celula = []
    while num < max_linha:
        celula = str(coluna) + str(num)
        lista_celula.append(celula)
        num = num + 1
    return lista_celula