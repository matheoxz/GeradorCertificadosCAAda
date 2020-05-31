#coding: utf-8
from reportlab.pdfgen import canvas                 #importa Canvas para criação de páginas
from reportlab.pdfbase import pdfmetrics            #importa funçãoo Centred
from reportlab.lib.pagesizes import A4, landscape   #importa tamanho da página e permite que seja posta horizontalmente
import re
import os
from datetime import datetime
import Mail

#gera data hoje
def DataTexto(data):
    data = data.split('/')
    if(data[1] == '01'):
        mes = 'Janeiro'
    elif(data[1] == '02'):
        mes = 'Fevereiro'
    elif(data[1] == '03'):
        mes = 'Março'
    elif(data[1] == '04'):
        mes = 'Abril'
    elif(data[1] == '05'):
        mes = 'Maio'
    elif(data[1] == '06'):
        mes = 'Junho'
    elif(data[1] == '07'):
        mes = 'Julho'
    elif(data[1] == '08'):
        mes = 'Agosto'
    elif(data[1] == '09'):
        mes = 'Setembro'
    elif(data[1] == '10'):
        mes = 'Outubro'
    elif(data[1] == '11'):
        mes = 'Novembro'
    elif(data[1] == '12'):
        mes = 'Dezembro'
    return (data[0] + " de " + mes + " de " + data[2])



continua = True
#laço repetição do programa
while(continua):

    #Pede palestra ou minicurso
    print("Deseja gerar certificados para palestras(0), minicursos(1), visita técnica(2) ou organização de evento?")
    PouM = int(input())

    #pede nome da palestra/ minicurso 
    if(PouM == 1):
        print("Qual o tema do minicurso?")
    elif (PouM == 0):
        print("Qual o nome da palestra?")       
    elif (PouM == 2):
        print('Para onde foi a visita?')
    elif (PouM == 3):
        print("Qual o nome do evento?")
    titulo = str(input())
    #pede caminho do arquivo
    deu = False
    while(not deu):
        try:
            print("Qual o caminho do arquivo com os nomes dos alunos?")
            path = str(input())
            arq = open(path, 'r', encoding = 'utf-8')
        except:
            print("Arquivo não encontrado")
        else:
            deu = True

    #pede data
    if(PouM == 1):
        print("Qual a data do minicurso?")
    elif (PouM == 0):
        print("Qual a data da palestra?") 
    elif(PouM == 2):
        print("Qual a data da visita?")
    elif(PouM == 3):
        print("Qual a primeira data?")
    print("escreva no formato DD/MM/AAAA:")
    data = str(input())
    if(PouM == 3):
        print("Qual a última data?")
        data1 = str(input())

    # abre imagem de fundo e cria pasta
    background = 'back.png'
    try:
        os.mkdir(titulo)
    except:
        print('Diretório '+ titulo + ' não criado.')
    else:
        print('Diretório '+ titulo + ' criado com sucesso.')

    #gera certificado para cada aluno no arquivo
    for pessoa in arq:
        #limpa nome de caracteres especiais (\t, \n)
        pessoa = re.sub("\n", "", pessoa)
        pessoa = pessoa.split("\t")
        
        for a in pessoa:
            a = re.sub("\n", "", a)
            a = a.replace('\n', '')
        print(pessoa)
        try:
            nome = pessoa[0] + " " + pessoa[1]
        except:
            nome = pessoa
        
        print("Contruindo certificado de  " + nome)

        #gera nome do certificado, gera o arquivo do certificado e o título do arquivo
        filename = os.path.join(titulo, nome + ' - Certificado ' + titulo + '.pdf')
        c = canvas.Canvas(filename, pagesize = landscape(A4))
        c.setTitle(nome + " - Certificado" + titulo)

        #poe o background no fundo do certificado
        cW, cH = c._pagesize
        c.drawInlineImage(background, 0, 0, width = cW, height = cH)

        #coloca texto, linha por linha no documento, alinhados à posição (400, 380)
        if(PouM == 1):
            dec = ['Declaramos para os devidos fins que o/a discente '+ str(nome),
                   'participou do minicurso de ', str(titulo) + ' no dia ' + DataTexto(data),
                   'no evento "Ada Lovelace Week", realizado pelo Centro Acadêmico', 
                   '"Ada Lovelace" da UNESP Bauru, com duração de 1 hora.',
                   'Bauru, ' + DataTexto(datetime.now().strftime('%d/%m/%Y'))]
        elif (PouM == 0):
            dec = ['Declaramos para os devidos fins que o/a discente '+ str(nome),
                   'participou da palestra de título', str(titulo), ' no dia ' + DataTexto(data),
                   'no evento "Ada Lovelace Week", realizado pelo Centro Acadêmico', 
                   '"Ada Lovelace" da UNESP Bauru, com duração de 1 hora.',
                   'Bauru, ' + DataTexto(datetime.now().strftime('%d/%m/%Y'))]
        elif (PouM == 2):
            dec = ['Declaramos para os devidos fins que o/a discente '+ str(nome),
                   'participou da visita técnica para ' + str(titulo), ' no dia ' + DataTexto(data),
                   'no evento "Ada Lovelace Week", realizado pelo Centro Acadêmico', 
                   '"Ada Lovelace" da UNESP Bauru, com duração de 1 hora.',
                   'Bauru, ' + DataTexto(datetime.now().strftime('%d/%m/%Y'))]
        elif (PouM == 3):
            dec = ['O Centro Acadêmico \'Ada Lovelace\' de Computação da UNESP Bauru', 
                    'declara para os devidos fins que o/a discente' + str(nome),
                   'participou da comissão organizadora do evento', str(titulo), 
                   'uma semana de palestras e minicursos, participando de reuniões semanais de 1h',
                   'entre ' + DataTexto(data) + ' e ' + DataTexto(data1),
                   'dedicando um total de 20h ao projeto',
                   'Bauru, ' + DataTexto(datetime.now().strftime('%d/%m/%Y'))]
        y = 380
        for i in dec:
            c.setFont('Helvetica', 20)
            c.drawCentredString(400, y, i)
            y -= 30

        #fecha arquivo pdf e o salva
        c.showPage()
        c.save()

        nenviados = ["Pesssoas que não receberam o email: "]
        try:
            print("Enviando certificado via email para " + pessoa[2])
            Mail.send_email(pessoa[2], filename, subject="Certificado Ada Week " + titulo, message="Segue anexo o certificado da palestra/ minicurso " + titulo + " ocorrida em " + data)
        except:
            print("Email não enviado")
            nenviados.append(nome)
        
        nenvarq = open(os.path.join(titulo, 'Certificados não enviados.txt'), "w")
        for i in nenviados:
            nenvarq.write(i)
        nenvarq.close()

    #fecha arquivo de nomes
    arq.close()
    print('Deseja gerar mais certificados? (S/N)')
    resp = str(input())
    if(resp == 'N' or resp == 'n'):
        continua = False
    