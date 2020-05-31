from tkinter import *
from tkinter import filedialog
from tkinter import messagebox


class GeradorCertficados:
    arquivo = None

    def __init__(self):
        self.Tela()

    def AbrirArquivo(self):
        self.arquivo_entry = filedialog.askopenfilenames(title = 'Arquivo de nomes', filetypes = (('txt', '*.txt'), ('csv', '*.csv')))
        self.arquivoaberto_label.config(text = str(self.arquivo_entry))

    def GerarCertificados(self):
        self.getDados()
        self.getOpcao()
        return 0

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
            self.diretorio = self.duracao_entry.get()
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
        self.arquivoaberto_label = Label(self.dados, text = 'Sem arquivo selecionado')

        evento_label = Label(evento, text = 'Título do Evento')

        #inserção dos labels
        titulo_label.grid(row = 0, column = 0, sticky = E)
        datai_label.grid(row = 1, column = 0, sticky = E)
        dataf_label.grid(row = 2, column = 0, sticky = E)
        duracao_label.grid(row = 3, column = 0, sticky = E)
        diretorio_label.grid(row = 4, column = 0, sticky = E)
        arquivo_label.grid(row = 5, column = 0, sticky = E)
        self.arquivoaberto_label.grid(row = 5, column = 2)

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
        gerar_button = Button(self.root, text = 'Gerar!', command = self.GerarCertificados)

        #inserção de botoes
        arquivo_button.grid(row = 5, column = 1, sticky = E)
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
