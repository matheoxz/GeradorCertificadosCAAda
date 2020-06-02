from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import os
import re
from reportlab.pdfgen import canvas                 #importa Canvas para criação de páginas
from reportlab.pdfbase import pdfmetrics            #importa funçãoo Centred
from reportlab.lib.pagesizes import A4, landscape   #importa tamanho da página e permite que seja posta horizontalmente
import pandas as pd
from datetime import datetime
import Mail

class GeradorCertficados:
    arquivo = None
    imagem_fundo = None

    def __init__(self):
        self.Tela()

    def getOpcao(self):
        self.opcao = self.opcao_var.get()
        if self.opcao == '':
            messagebox.showerror(title = 'Escolha uma opção', message = 'Escolha uma opção de certificado a ser gerado!')
        if (self.opcao == 'Palestra' or self.opcao == 'Minicurso' or self.opcao == 'Visita Tecnica' or self.opcao == 'Organização') and self.evento_entry.get() == '':
            messagebox.showerror(title = 'Diga o nome do evento', message = 'Escolha um nome para o evento')
        else: 
            self.evento = self.evento_entry.get()

    def getDados(self):
        try:
            self.titulo = self.titulo_entry.get()
            self.datai = self.datai_entry.get()
            self.dataf = self.dataf_entry.get()
            self.duracao = self.duracao_entry.get()
            self.diretorio = self.diretorio_entry.get()
        except:
            messagebox.showerror(title = 'Campos não preenchidos', message = 'Preencha todos os campos')
        else:
            if(self.titulo == '' or self.datai == '' or self.dataf == '' or self.duracao == ''):
                messagebox.showerror(title = 'Campos não preenchidos', message = 'Preencha todos os campos')

    def DataTexto(self, data):
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

    def ImagemDeFundo(self):
        try:
            self.imagem_fundo_entry = filedialog.askopenfilenames(title = 'Arquivo de nomes', filetypes = (('png', '*.png'), ('jpg', '*.jpg')))
        except:
            messagebox.showerror(title = 'Não foi possivel encontrar aquivo', message = 'Tente novamente')
        else:
            if (self.imagem_fundo_entry != ''):
                self.imgaberto_label.config(text = str(self.imagem_fundo_entry))

    def AbrirArquivo(self):
        try:
            self.arquivo_entry = filedialog.askopenfilenames(title = 'Arquivo de nomes', filetypes = (('csv', '*.csv'), ('all', '*.*')))
        except:
            messagebox.showerror(title = 'Não foi possivel encontrar aquivo', message = 'Tente novamente')
        else:
            if (self.arquivo_entry != ''):
                self.arquivoaberto_label.config(text = str(self.arquivo_entry))

    def CriaDiretorio(self):
        try:
            os.mkdir(self.diretorio)
        except:
            messagebox.showerror(title = 'Não foi possivel criar diretorio', message = 'Tente novamente')
        else:
            print('Diretorio ' + self.diretorio + ' criado com sucesso')

    def AbreArquivo(self):
        try:
            self.arquivo = pd.read_csv(self.arquivo_entry[0])
        except Exception as e:
            print(e)
            messagebox.showerror(title = 'Erro ao ler arquivo', message = 'Tente novamente')
        else:
            print('Arquivo aberto com sucesso!')

    def AbreImagem(self):
        self.imagem_fundo = self.imagem_fundo_entry[0]

    def TextoOpcao(self, nome):
        if(self.opcao == 'Curto Circuito'):
            return ['Declaramos para os devidos fins que o/a discente '+ str(nome),
                   'participou do curto circuito de ' + str(self.titulo), 
                   ' no dia ' + self.DataTexto(self.datai),
                   'realizado pelo Centro Acadêmico "Ada Lovelace" da UNESP Bauru,',
                   'com duração de ' + str(self.duracao) +' hora(s).',
                   'Bauru, ' + self.DataTexto(datetime.now().strftime('%d/%m/%Y'))]

        if(self.opcao == 'Palestra'):
            return ['Declaramos para os devidos fins que o/a discente '+ str(nome),
                   'participou da palestra de título ', str(self.titulo) + ' no dia ' + self.DataTexto(self.datai),
                   'no evento "' + self.evento + '", realizado pelo Centro Acadêmico', 
                   '"Ada Lovelace" da UNESP Bauru, com duração de '+ str(self.duracao) +' hora.',
                   'Bauru, ' + self.DataTexto(datetime.now().strftime('%d/%m/%Y'))]

        if(self.opcao == 'Minicurso'):
            return ['Declaramos para os devidos fins que o/a discente '+ str(nome),
                   'participou do minicurso de ', str(self.titulo) + ' no dia ' + self.DataTexto(self.datai),
                   'no evento "' + self.evento + '", realizado pelo Centro Acadêmico', 
                   '"Ada Lovelace" da UNESP Bauru, com duração de '+ str(self.duracao) +' hora.',
                   'Bauru, ' + self.DataTexto(datetime.now().strftime('%d/%m/%Y'))]

        if(self.opcao == 'Visita Tecnica'):
            return ['Declaramos para os devidos fins que o/a discente '+ str(nome),
                   'participou da visita técnica para ', str(self.titulo) + ' no dia ' + self.DataTexto(self.datai),
                   'no evento "' + self.evento + '", realizado pelo Centro Acadêmico', 
                   '"Ada Lovelace" da UNESP Bauru, com duração de '+ str(self.duracao) +' hora.',
                   'Bauru, ' + self.DataTexto(datetime.now().strftime('%d/%m/%Y'))]

        if(self.opcao == 'Organização'):
            return ['O Centro Acadêmico \'Ada Lovelace\' de Computação da UNESP Bauru', 
                    'declara para os devidos fins que o/a discente' + str(nome),
                   'participou da comissão organizadora do evento', str(self.evento), 
                   'uma semana de palestras e minicursos, participando de reuniões semanais de 1h',
                   'entre ' + self.DataTexto(self.datai) + ' e ' + DataTexto(self.dataf),
                   'dedicando um total de ' + str(self.duracao) + ' ao projeto',
                   'Bauru, ' + self.DataTexto(datetime.now().strftime('%d/%m/%Y'))]

    def GeraCertificados(self):
        for nome in self.arquivo['Nome']:
            #gera nome do certificado, gera o arquivo do certificado e o título do arquivo
            filename = os.path.join(self.diretorio, nome + ' - Certificado ' + self.titulo + '.pdf')
            c = canvas.Canvas(filename, pagesize = landscape(A4))
            c.setTitle(nome + " - Certificado" + self.titulo)
            print('Gerando ' + filename)
            #poe o background no fundo do certificado
            cW, cH = c._pagesize
            c.drawInlineImage(self.imagem_fundo, 0, 0, width = cW, height = cH)

            texto = self.TextoOpcao(nome)

            y = 380
            for i in texto:
                c.setFont('Helvetica', 20)
                c.drawCentredString(400, y, i)
                y -= 30

            #fecha arquivo pdf e o salva
            c.showPage()
            c.save()

            if(self.email_var.get() == 1):
                email_c = self.arquivo[self.arquivo['Nome'] == nome]['email']
                for i in email_c:
                    email = i
                try:
                    print("Enviando certificado via email para " + str(email))
                    Mail.send_email(email, filename, subject='Certificado ' + self.opcao + ' de ' + self.titulo, message="Segue anexo o certificado referente à(ao) " +self.opcao+ ' de ' + self.titulo + " ocorrida em " + self.datai)
                except:
                    print("Email não enviado")
                else:
                    print('Enviado com sucesso!')

            print('Certificados gerados com sucesso!')

    
    def GerarCertificados(self):
        self.getDados()
        self.getOpcao()
        self.CriaDiretorio()
        self.AbreArquivo()
        self.AbreImagem()
        self.GeraCertificados()
        return 0

  
        
    def Tela(self):
        #inicia tela
        self.root = Tk()
        self.root.title('Gerador de Certificados')

        #frames
        geral = LabelFrame(self.root, text = 'Geral')
        evento = LabelFrame(self.root, text = 'Evento')
        self.dados = LabelFrame(self.root, text = 'Dados: ')

        #labels
        titulo_label = Label(self.dados, text = 'Título')
        datai_label = Label(self.dados, text = 'Data de início')
        dataf_label = Label(self.dados, text = 'Data de encerramento')
        duracao_label = Label(self.dados, text = 'Tempo de duração')
        diretorio_label = Label(self.dados, text = 'Nome do diretório')
        arquivo_label = Label(self.dados, text = 'Arquivo de nomes')
        img_label = Label(self.dados, text = 'Imagem de fundo')
        self.arquivoaberto_label = Label(self.dados, text = 'Sem arquivo selecionado')
        self.imgaberto_label = Label(self.dados, text = 'Sem imagem de fundo selecionada')

        evento_label = Label(evento, text = 'Título do Evento')

        #inserção dos labels
        titulo_label.grid(row = 0, column = 0, sticky = E)
        datai_label.grid(row = 1, column = 0, sticky = E)
        dataf_label.grid(row = 2, column = 0, sticky = E)
        duracao_label.grid(row = 3, column = 0, sticky = E)
        diretorio_label.grid(row = 4, column = 0, sticky = E)
        arquivo_label.grid(row = 5, column = 0, sticky = E)
        img_label.grid(row = 6, column = 0, sticky = E)
        self.arquivoaberto_label.grid(row = 5, column = 2)
        self.imgaberto_label.grid(row = 6, column = 2)

        evento_label.grid(row = 4, column = 0, sticky = E)

        #entradas
        self.titulo_entry = Entry(self.dados)
        self.datai_entry = Entry(self.dados)
        self.dataf_entry = Entry(self.dados)
        self.duracao_entry = Entry(self.dados)
        self.diretorio_entry = Entry(self.dados)
        
        self.evento_entry = Entry(evento)

        #inserção das entradas
        self.titulo_entry.grid(row = 0, column = 1, columnspan = 2, sticky = W+E)
        self.datai_entry.grid(row = 1, column = 1, columnspan = 2, sticky = W+E)
        self.dataf_entry.grid(row = 2, column = 1, columnspan = 2, sticky = W+E)
        self.duracao_entry.grid(row = 3, column = 1, columnspan = 2, sticky = W+E)
        self.diretorio_entry.grid(row = 4, column = 1, columnspan = 2, sticky = W+E)

        self.evento_entry.grid(row = 4, column = 1)

        #botoes
        arquivo_button = Button(self.dados, text = 'Abrir', command = self.AbrirArquivo)
        img_button = Button(self.dados, text = 'Abrir', command = self.ImagemDeFundo)
        gerar_button = Button(self.root, text = 'Gerar!', command = self.GerarCertificados)

        #inserção de botoes
        arquivo_button.grid(row = 5, column = 1, sticky = E)
        img_button.grid(row = 6, column = 1, sticky = E)
        gerar_button.grid(row = 2, column = 0, sticky = W+E)

        #radio buttons
        self.opcao_var = StringVar()
        curtoCircuito_rb = Radiobutton(geral, text = 'Curto Circuito', variable = self.opcao_var, value = 'Curto Circuito')
        participacao_rb = Radiobutton(geral, text = 'Participação na chapa (?)', variable = self.opcao_var, value = 'Participação na chapa')
        outro_rb = Radiobutton(geral, text = 'Outro', variable = self.opcao_var, value = 'Outro')

        palestra_rb = Radiobutton(evento, text = 'Palestra', variable = self.opcao_var, value = 'Palestra')
        minicurso_rb = Radiobutton(evento, text = 'Minicurso', variable = self.opcao_var, value = 'Minicurso')
        visitaTecnica_rb = Radiobutton(evento, text = 'Visita Tecnica', variable = self.opcao_var, value = 'Visita Tecnica')
        organizacao_rb = Radiobutton(evento, text = 'Organização', variable = self.opcao_var, value = 'Organização')

        #inserçõs radio buttons
        curtoCircuito_rb.grid(row = 0, column = 0, sticky = W)
        participacao_rb.grid(row = 1, column = 0, sticky = W)
        outro_rb.grid(row = 2, column = 0, sticky = W)

        palestra_rb.grid(row = 0, column = 0, sticky = W, columnspan = 2)
        minicurso_rb.grid(row = 1, column = 0, sticky = W, columnspan = 2)
        visitaTecnica_rb.grid(row = 2, column = 0, sticky = W, columnspan = 2)
        organizacao_rb.grid(row = 3, column = 0, sticky = W, columnspan = 2)

        #checkbox
        self.email_var = IntVar()
        email_cb = Checkbutton(self.root, text = 'Enviar por e-mail', variable = self.email_var, onvalue = 1, offvalue = 0)

        #inserções checkbox
        email_cb.grid(row = 2, column = 1)

        #inserção frames
        geral.grid(row = 0, column = 0, sticky = N+S)
        evento.grid(row = 0, column = 1, sticky = N+S)
        self.dados.grid(row = 1, column = 0, columnspan = 2, sticky = W+E)
        self.root.mainloop()


gerador = GeradorCertficados()
