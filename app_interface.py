import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedStyle
import sqlite3
import pandas as pd
from reportlab.pdfgen import canvas
import subprocess
import random
import time
import os
from tkcalendar import DateEntry
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *

# Função para obter o nome do usuário do banco de dados
def obter_nome_usuario():
    conn = sqlite3.connect("banco_dados.db")
    cursor = conn.cursor()
    cursor.execute("SELECT id_adm FROM db_administradores")  # Assume que há apenas um usuário
    resultado = cursor.fetchone()
    conn.close()
    if resultado:
        return resultado[0]
    else:
        return "Usuário Desconhecido"

# Função para atualizar o rodapé
def atualizar_rodape():
    data_hora = time.strftime("%d/%m/%Y %H:%M:%S")
    nome_usuario = obter_nome_usuario()
    rodape_label.config(text=f"Data e Hora: {data_hora} | Usuário: {nome_usuario}")

# Conecta com o banco de dados
banco = sqlite3.connect('banco_dados.db')
cursor = banco.cursor()

# Cria a janela principal
janela = tk.Tk()
janela.title('[Gestão de Ativos TI - GEO/CMG e TAE/TSEP] - PETROBRAS SA')

# Define o ícone da janela
icon = tk.PhotoImage(file="img//app_icon.png")
janela.iconphoto(True, icon)

# Configura a janela para abrir em tela cheia
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()
janela.geometry(f"{1366}x{760}+0+0")

# Cria o widget de notebook (guias)
notebook = ttk.Notebook(janela)
notebook.pack(fill='both', expand=True)

# Cria um rótulo para o rodapé
rodape_label = tk.Label(janela, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
rodape_label.pack(side=tk.BOTTOM, fill=tk.X)

# Atualiza o rodapé inicialmente
atualizar_rodape()

# Atualiza o rodapé a cada segundo
janela.after(1000, atualizar_rodape)

# Crie uma função para sair com confirmação
def sair_com_confirmacao():
    resultado = messagebox.askquestion("Confirmação", "Tem certeza de que deseja retornar para o menu principal?")
    if resultado == "yes":
        janela.destroy()
        os.system("python app_menu.py")

# Função para criar um documento PDF com base na tabela db_alocacoes
def criar_pdf_db_alocacoes():
    # Conecta-se ao banco de dados SQLite
    conexao = sqlite3.connect("banco_dados.db")
    cursor = conexao.cursor()

    # Executa uma consulta para obter todos os registros da tabela db_alocacoes
    cursor.execute("SELECT * FROM db_alocacao")
    registros = cursor.fetchall()

    # Cria um documento PDF com orientação horizontal
    pdf_filename = "relatorio_alocacoes.pdf"
    pdf = canvas.Canvas(pdf_filename, pagesize=(595, 842))
    pdf.rotate(90)

    # Configurações da tabela
    largura_coluna = 55
    altura_linha = 20
    linha_inicial = 50
    colunas = ["Matrícula", "Colaborador", "Lotação", "Localização", "BP",
               "Recurso", "Modelo", "Sistema Operacional", "Uso", "Status",
               "Entrega", "Devolução", "Contra-Prova"]

    # Tamanho da fonte
    pdf.setFont("Helvetica", 4)

    # Adiciona cabeçalho da tabela
    pdf.drawString(50, -linha_inicial, "Relatório de Alocações de Ativos")

    # Adiciona as colunas da tabela
    for i, coluna in enumerate(colunas):
        pdf.drawString(50 + i * largura_coluna, -linha_inicial - altura_linha, coluna)

    # Adiciona os registros na tabela
    for linha, registro in enumerate(registros, 1):
        for coluna, valor in enumerate(registro):
            pdf.drawString(50 + coluna * largura_coluna, -linha_inicial - (linha + 1) * altura_linha, str(valor))

    # Fecha o arquivo PDF
    pdf.save()

    # Fecha a conexão com o banco de dados
    conexao.close()

    # Retorna o nome do arquivo PDF criado
    return pdf_filename

# Função para imprimir o arquivo PDF usando o leitor de PDF padrão
def imprimir_pdf(pdf_filename):
    # Obtém o visualizador de PDF padrão do sistema
    visualizador_pdf = "start" if os.name == "nt" else "xdg-open"

    # Abre o arquivo PDF com o visualizador padrão
    subprocess.run([visualizador_pdf, pdf_filename], shell=True)

# Função principal para criar o PDF e imprimir
def imprimir_db_alocacoes():
    pdf_filename = criar_pdf_db_alocacoes()
    imprimir_pdf(pdf_filename)
    
# Cria um frame para cada tabela e adicionar ao notebook
frame_recursos = tk.Frame(notebook)
frame_colaboradores = tk.Frame(notebook)
frame_alienacao = tk.Frame(notebook)
frame_lotacao = tk.Frame(notebook)
frame_localizacao = tk.Frame(notebook)
frame_alocacao = tk.Frame(notebook)

# Cria abas no frame principal
notebook.add(frame_alocacao, text='Alocações')
notebook.add(frame_recursos, text='Recursos') # Retirar aba 'Recursos'.
notebook.add(frame_colaboradores, text='Colaboradores') # Retirar aba 'Recursos'.
notebook.add(frame_lotacao, text='Lotação') # Retirar aba 'Recursos'.  
notebook.add(frame_localizacao, text='Localizações') # Retirar aba 'Recursos'.
notebook.add(frame_alienacao, text='Alienações') # Retirar aba 'Recursos'.

# Altera a posição das guias "Alocação" e "Alienação" para a direita
style = ThemedStyle(janela)
style.set_theme_advanced("radiance", preserve_transparency=True)
style.configure("TNotebook.Tab", padding=(10, 5, 10, 5))  # Ajusta o espaçamento das guias

# Cria um widget Treeview para cada frame
treeview_recursos = ttk.Treeview(frame_recursos)
treeview_colaboradores = ttk.Treeview(frame_colaboradores)
treeview_alienacao = ttk.Treeview(frame_alienacao)
treeview_lotacao = ttk.Treeview(frame_lotacao)
treeview_localizacao = ttk.Treeview(frame_localizacao)
treeview_alocacao = ttk.Treeview(frame_alocacao)  # Adicione Treeview de alocação

# Define as colunas da Treeview
treeview_recursos["columns"] = ("id_bp", "tipo_recurso", "modelo", "sistema_operacional", "tipo_uso", "status")
treeview_colaboradores["columns"] = ("nome_colab", "matricula", "nome_lotacao","nome_localizacao")
treeview_alienacao["columns"] = ("id_bp", "matricula", "data_alienacao")
treeview_lotacao["columns"] = ("id_lotacao","nome_lotacao")
treeview_localizacao["columns"] = ("id_localizacao","nome_localizacao")
treeview_alocacao["columns"] = ("matricula", "nome_colab", "lotacao_nome", "nome_localizacao", "id_bp", "tipo_recurso", "modelo", "sistema_operacional", "status", "tipo_uso", "data_entrega", "data_devolucao", "foto_contraprova")

# Configura a cor para verde claro
style.configure("Treeview.Heading", background="light green", foreground="black")

# Define o cabeçalho das colunas

## Cabeçalho de Recursos
treeview_recursos.heading("id_bp", text="ID/BP")
treeview_recursos.heading("tipo_recurso", text="Tipo de Recurso")
treeview_recursos.heading("modelo", text="Modelo")
treeview_recursos.heading("sistema_operacional", text="Sistema Operacional")
treeview_recursos.heading("tipo_uso", text="Tipo de Uso")
treeview_recursos.heading("status", text="Status do Recurso")

## Cabeçalho de Colaboradores
treeview_colaboradores.heading("nome_colab", text="Colaborador")
treeview_colaboradores.heading("matricula", text="Matricula")
treeview_colaboradores.heading("nome_lotacao", text="Nome da Lotacão")
treeview_colaboradores.heading("nome_localizacao", text="Nome da Localização")

## Cabeçalho de Alienação
treeview_alienacao.heading("id_bp", text="ID/BP")
treeview_alienacao.heading("matricula", text="Matricula")
treeview_alienacao.heading("data_alienacao", text="Data da Alienação")

## Cabeçalho de Lotação
treeview_lotacao.heading("id_lotacao", text="ID da Lotação")
treeview_lotacao.heading("nome_lotacao", text="Lotação")

## Cabeçalho de Localização
treeview_localizacao.heading("id_localizacao", text="ID de Localização")
treeview_localizacao.heading("nome_localizacao", text="Nome de Localização")

## Cabeçalho de Alocação
treeview_alocacao.heading("nome_colab", text="Nome do Colaborador")
treeview_alocacao.heading("matricula", text="Matrícula")
treeview_alocacao.heading("lotacao_nome", text="Nome da Lotação")
treeview_alocacao.heading("nome_localizacao", text="Nome da Localização")
treeview_alocacao.heading("id_bp", text="ID/BP")
treeview_alocacao.heading("tipo_recurso", text="Tipo de Recurso")
treeview_alocacao.heading("modelo", text="Modelo")
treeview_alocacao.heading("sistema_operacional", text="Sistema Operacional")
treeview_alocacao.heading("tipo_uso", text="Tipo de Uso")
treeview_alocacao.heading("status", text="Status")
treeview_alocacao.heading("data_entrega", text="Data de Entrega")
treeview_alocacao.heading("data_devolucao", text="Data de Devolução")
treeview_alocacao.heading("foto_contraprova", text="Foto de Contra-Prova")

# Adiciona as colunas à Treeview
treeview_recursos.column("#0", width=0, stretch=tk.NO)
treeview_colaboradores.column("#0", width=0, stretch=tk.NO)
treeview_alienacao.column("#0", width=0, stretch=tk.NO)
treeview_lotacao.column("#0", width=0, stretch=tk.NO)
treeview_localizacao.column("#0", width=0, stretch=tk.NO)
treeview_alocacao.column("#0", width=0, stretch=tk.NO)

treeview_recursos.pack(fill='both', expand=True)
treeview_colaboradores.pack(fill='both', expand=True)
treeview_alienacao.pack(fill='both', expand=True)
treeview_lotacao.pack(fill='both', expand=True)
treeview_localizacao.pack(fill='both', expand=True)
treeview_alocacao.pack(fill='both', expand=True)

# Cria as barras de rolagem
scrollbar_vertical_recursos = ttk.Scrollbar(treeview_recursos, orient='vertical', command=treeview_recursos.yview)
scrollbar_horizontal_recursos = ttk.Scrollbar(treeview_recursos, orient='horizontal', command=treeview_recursos.xview)
scrollbar_vertical_colaboradores = ttk.Scrollbar(treeview_colaboradores, orient='vertical', command=treeview_colaboradores.yview)
scrollbar_horizontal_colaboradores = ttk.Scrollbar(treeview_colaboradores, orient='horizontal', command=treeview_colaboradores.xview)
scrollbar_vertical_alienacao = ttk.Scrollbar(treeview_alienacao, orient='vertical', command=treeview_alienacao.yview)
scrollbar_horizontal_alienacao = ttk.Scrollbar(treeview_alienacao, orient='horizontal', command=treeview_alienacao.xview)
scrollbar_vertical_lotacao = ttk.Scrollbar(treeview_lotacao, orient='vertical', command=treeview_lotacao.yview)
scrollbar_horizontal_lotacao = ttk.Scrollbar(treeview_lotacao, orient='horizontal', command=treeview_lotacao.xview)
scrollbar_vertical_localizacao = ttk.Scrollbar(treeview_localizacao, orient='vertical', command=treeview_localizacao.yview)
scrollbar_horizontal_localizacao = ttk.Scrollbar(treeview_localizacao, orient='horizontal', command=treeview_localizacao.xview)
scrollbar_vertical_alocacao = ttk.Scrollbar(treeview_alocacao, orient='vertical', command=treeview_alocacao.yview)  # Adicione a barra de rolagem para a Treeview de alocação
scrollbar_horizontal_alocacao = ttk.Scrollbar(treeview_alocacao, orient='horizontal', command=treeview_alocacao.xview)  # Adicione a barra de rolagem horizontal para a Treeview de alocação

# Configura as barras de rolagem para o Treeview
treeview_recursos.configure(yscrollcommand=scrollbar_vertical_recursos.set, xscrollcommand=scrollbar_horizontal_recursos.set)
treeview_colaboradores.configure(yscrollcommand=scrollbar_vertical_colaboradores.set, xscrollcommand=scrollbar_horizontal_colaboradores.set)
treeview_alienacao.configure(yscrollcommand=scrollbar_vertical_alienacao.set, xscrollcommand=scrollbar_horizontal_alienacao.set)
treeview_lotacao.configure(yscrollcommand=scrollbar_vertical_lotacao.set, xscrollcommand=scrollbar_horizontal_lotacao.set)
treeview_localizacao.configure(yscrollcommand=scrollbar_vertical_localizacao.set, xscrollcommand=scrollbar_horizontal_localizacao.set)
treeview_alocacao.configure(yscrollcommand=scrollbar_vertical_alocacao.set, xscrollcommand=scrollbar_horizontal_alocacao.set)  # Configure as barras de rolagem para a Treeview de alocação

# Posiciona as barras de rolagem
scrollbar_vertical_recursos.pack(side='right', fill='y')
scrollbar_horizontal_recursos.pack(side='bottom', fill='x')
scrollbar_vertical_colaboradores.pack(side='right', fill='y')
scrollbar_horizontal_colaboradores.pack(side='bottom', fill='x')
scrollbar_vertical_alienacao.pack(side='right', fill='y')
scrollbar_horizontal_alienacao.pack(side='bottom', fill='x')
scrollbar_vertical_lotacao.pack(side='right', fill='y')
scrollbar_horizontal_lotacao.pack(side='bottom', fill='x')
scrollbar_vertical_localizacao.pack(side='right', fill='y')
scrollbar_horizontal_localizacao.pack(side='bottom', fill='x')
scrollbar_vertical_alocacao.pack(side='right', fill='y')
scrollbar_horizontal_alocacao.pack(side='bottom', fill='x')

# Define a função de cadastro para cada tipo de informação
def cadastrar_recurso():
    # Carrega uma planilha do Excel de qualquer formato ou versão do software, para carregamento em lote as informações de colaboradores
    def carregar_planilha():
        file_path = filedialog.askopenfilename(filetypes=[('Planilhas Excel', '*.xls;*.xltx;*.xlsm'), ('Todos os arquivos', '*.*')])
        if file_path:
            try:
                df = pd.read_excel(file_path, header=None, names=['id_bp', 'tipo_recurso', 'modelo', 'sistema_operacional','tipo_uso','status'])

                for index, row in df.iterrows():
                    cursor.execute("INSERT INTO db_recursos (id_bp, tipo_recurso, modelo, sistema_operacional, tipo_uso, status) VALUES (?, ?, ?, ?, ?, ?)",
                                   (row['id_bp'], row['tipo_recurso'], row['modelo'], row['sistema_operacional'], row['tipo_uso'], row['status']))
                banco.commit()

                atualizar_treeview()

                messagebox.showinfo("Sucesso", "Dados carregados com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar dados da planilha: {str(e)}")

    def atualizar_treeview():
        df_recursos = pd.read_sql_query("SELECT * FROM db_recursos", banco)
        treeview_recursos.delete(*treeview_colaboradores.get_children())
        for index, row in df_recursos.iterrows():
            treeview_recursos.insert("", tk.END, text=index, values=row.tolist())

    def preencher_campos(event):
        # Obtém o id_bp digitado pelo usuário
        id_bp = campo_id_bp.get()

        # Verifica se o id_bp existe no banco de dados
        cursor.execute("SELECT * FROM db_recursos WHERE id_bp=?", (id_bp,))
        resultado = cursor.fetchone()

        if resultado:
            # Preenche os campos com as informações obtidas
            _, tipo_recurso, modelo, sistema_operacional, tipo_uso, status = resultado
            combo_tipo_recurso.set(tipo_recurso)
            campo_modelo.delete(0, tk.END)
            campo_modelo.insert(0, modelo)
            campo_sistema_operacional.delete(0, tk.END)
            campo_sistema_operacional.insert(0, sistema_operacional)
            combo_tipo_uso.set(tipo_uso)
            combo_status.set(status)

    def salvar_recurso():
        id_bp = campo_id_bp.get()

        # Verifica se o id_bp já existe no banco de dados
        cursor.execute("SELECT * FROM db_recursos WHERE id_bp=?", (id_bp,))
        if cursor.fetchone():
            # Mostra uma mensagem de erro se o id_bp já estiver cadastrado
            janela_recursos.destroy()
            messagebox.showerror("Erro", "Equipamento já cadastrado.")
            return

        tipo_recurso = combo_tipo_recurso.get()
        modelo = campo_modelo.get()
        sistema_operacional = campo_sistema_operacional.get()
        tipo_uso = combo_tipo_uso.get()
        status = combo_status.get()

        # Insere os dados nas tabelas db_recursos
        cursor.execute("INSERT INTO db_recursos (id_bp, tipo_recurso, modelo, sistema_operacional, tipo_uso, status) VALUES (?, ?, ?, ?, ?, ?)",
                       (id_bp, tipo_recurso, modelo, sistema_operacional, tipo_uso, status))
        banco.commit()

        # Atualiza a Treeview com os novos dados
        df_recursos = pd.read_sql_query("SELECT * FROM db_recursos", banco)
        treeview_recursos.delete(*treeview_recursos.get_children())
        for index, row in df_recursos.iterrows():
            treeview_recursos.insert("", tk.END, text=index, values=row.tolist())

        janela_recursos.destroy()

    def habilitar_desabilitar_sistema_operacional(event):
        if combo_tipo_recurso.get() in ['Monitor', 'Mouse', 'Teclado', 'Adaptadores', 'Cabos']:
            campo_sistema_operacional.config(state='disabled')
        else:
            campo_sistema_operacional.config(state='normal')

    # Cria uma nova janela
    janela_recursos = tk.Toplevel()
    janela_recursos.title('Cadastro de Recursos')

    # Permitir redimensionamento horizontal e vertical
    janela_recursos.resizable(True, True)

    # Configurar colunas para expansão horizontal
    janela_recursos.columnconfigure(0, weight=1)
    janela_recursos.columnconfigure(1, weight=1)

    # Cria os rótulos e os campos de entrada de texto
    rotulo_id_bp = tk.Label(janela_recursos, text='ID/BP:')
    rotulo_id_bp.grid(row=0, column=0, pady=10, padx=10, sticky="we")
    campo_id_bp = tk.Entry(janela_recursos)
    campo_id_bp.grid(row=0, column=1, pady=10, padx=10, sticky="we")
    campo_id_bp.bind('<FocusOut>', preencher_campos)

    # Cria um combobox (select box) para o Tipo de Recurso
    rotulo_tipo_recurso = tk.Label(janela_recursos, text='Tipo de Recurso:')
    rotulo_tipo_recurso.grid(row=1, column=0, pady=10, padx=10, sticky="we")
    tipos_recurso = ['Estação', 'Notebook', 'Monitor', 'Mouse', 'Teclado', 'Adaptadores', 'Cabos']
    combo_tipo_recurso = ttk.Combobox(janela_recursos, values=tipos_recurso)
    combo_tipo_recurso.grid(row=1, column=1, pady=10, padx=10, sticky="we")
    combo_tipo_recurso.bind('<FocusOut>', habilitar_desabilitar_sistema_operacional)

    # Cria um combobox (select box) para o Tipo de Modelo
    rotulo_modelo = tk.Label(janela_recursos, text='Modelo:')
    rotulo_modelo.grid(row=2, column=0, pady=10, padx=10, sticky="we")
    campo_modelo = tk.Entry(janela_recursos)
    campo_modelo.grid(row=2, column=1, pady=10, padx=10, sticky="we")

    # Cria o campo de entrada de texto para o sistema operacional
    rotulo_sistema_operacional = tk.Label(janela_recursos, text='Sistema Operacional:')
    rotulo_sistema_operacional.grid(row=3, column=0, pady=10, padx=10, sticky="we")
    campo_sistema_operacional = tk.Entry(janela_recursos)
    campo_sistema_operacional.grid(row=3, column=1, pady=10, padx=10, sticky="we")

    # Cria um combobox (select box) para o Tipo de Uso
    rotulo_tipo_uso = tk.Label(janela_recursos, text='Tipo de Uso:')
    rotulo_tipo_uso.grid(row=4, column=0, pady=10, padx=10, sticky="we")
    tipos_uso = ['Individual', 'Coletivo']
    combo_tipo_uso = ttk.Combobox(janela_recursos, values=tipos_uso)
    combo_tipo_uso.grid(row=4, column=1, pady=10, padx=10, sticky="we")

    # Cria o combobox (select box) para o Status
    rotulo_status = tk.Label(janela_recursos, text='Status:')
    rotulo_status.grid(row=5, column=0, pady=10, padx=10, sticky="we")
    tipo_status = ['Disponível','Solicitado', 'Entregue', 'Devolvido', 'Aguardando Fornecedor', 'Aguardando Entrega', 'Alienado', 'Emprestado']
    combo_status = ttk.Combobox(janela_recursos, values=tipo_status)
    combo_status.grid(row=5, column=1, pady=10, padx=10, sticky="we")

    botao_salvar = tk.Button(janela_recursos, text='Salvar', command=salvar_recurso)
    botao_salvar.grid(row=6, columnspan=2, pady=10, padx=10)

    botao_carregar_planilha = tk.Button(janela_recursos, text='Carregar Planilha', command=carregar_planilha)
    botao_carregar_planilha.grid(row=7, columnspan=2, pady=10, padx=10)

    # Função para habilitar/desabilitar o campo de Sistema Operacional
    def habilitar_desabilitar_sistema_operacional(event):
        if combo_tipo_recurso.get() in ['Monitor', 'Mouse', 'Teclado', 'Adaptadores', 'Cabos']:
            campo_sistema_operacional.config(state='disabled')
        else:
            campo_sistema_operacional.config(state='normal')

    # Adiciona o evento de <FocusOut> ao combobox de Tipo de Recurso
    combo_tipo_recurso.bind('<FocusOut>', habilitar_desabilitar_sistema_operacional)

# Define a função de cadastro para alienação
def cadastrar_alienacao():
    # Obtenha os valores dos campos de entrada de texto de alienação
    def salvar_alienacao():
        id_bp = campo_id_bp.get()
        matricula = campo_matricula.get()
        data_alienacao = campo_data_alienacao.get()

        cursor.execute("INSERT INTO db_alienacao (id_bp, matricula, data_alienacao) VALUES (?, ?, ?)",
                       (id_bp, matricula, data_alienacao))
        banco.commit()

        # Atualiza a Treeview com os novos dados
        df_alienacao = pd.read_sql_query("SELECT * FROM db_alienacao", banco)
        treeview_alienacao.delete(*treeview_alienacao.get_children())
        for index, row in df_alienacao.iterrows():
            treeview_alienacao.insert("", tk.END, text=index, values=row.tolist())

        janela_alienacao.destroy()

    # Cria novo frame chamada Cadastro de Alienação
    janela_alienacao = tk.Toplevel()
    janela_alienacao.title('Cadastro de Alienação')
    
    # Cria os rótulos e os campos de entrada de texto de Cadastro de Alienação
    rotulo_id_bp = tk.Label(janela_alienacao, text='ID/BP:')
    rotulo_id_bp.grid(row=0, column=0, pady=10, padx=10)
    campo_id_bp = tk.Entry(janela_alienacao)
    campo_id_bp.grid(row=0, column=1, pady=10, padx=10)

    rotulo_matricula = tk.Label(janela_alienacao, text='Matrícula do Colaborador:')
    rotulo_matricula.grid(row=1, column=0, pady=10, padx=10)
    campo_matricula = tk.Entry(janela_alienacao)
    campo_matricula.grid(row=1, column=1, pady=10, padx=10)

    rotulo_data_alienacao = tk.Label(janela_alienacao, text='Data da Alienação:')
    rotulo_data_alienacao.grid(row=2, column=0, pady=10, padx=10)
    campo_data_alienacao = tk.Entry(janela_alienacao)
    campo_data_alienacao.grid(row=2, column=1, pady=10, padx=10)

    botao_salvar = tk.Button(janela_alienacao, text='Salvar', command=salvar_alienacao)
    botao_salvar.grid(row=3, columnspan=2, pady=10, padx=10)

# Define a função de cadastro para colaboradores
def buscar_colaborador(nome_colab):
    cursor.execute("SELECT * FROM db_colaboradores WHERE LOWER (nome_colab) = LOWER(?)", (nome_colab,))
    resultado = cursor.fetchone()
    if resultado:
        return resultado[1:] # Retorna uma tupla com os valores dos campos, exceto o nome
    else:
        return None # Retorna None se não encontrar o colaborador

def obter_opcoes_lotacao():
    # Implement this function to fetch options from the database
    cursor.execute("SELECT nome_lotacao FROM db_lotacao")
    lotacoes = cursor.fetchall()
    return [lotacao[0] for lotacao in lotacoes]

# Define a função de cadastro de colaboradores
def cadastrar_colaboradores():
    # Carrega uma planilha do Excel de qualquer formato ou versão do software, para carregamento em lote as informações de colaboradores
    def carregar_planilha():
        file_path = filedialog.askopenfilename(filetypes=[('Planilhas Excel', '*.xls;*.xltx;*.xlsm'), ('Todos os arquivos', '*.*')])
        if file_path:
            try:
                df = pd.read_excel(file_path, header=None, names=['nome_colab', 'matricula', 'lotacao_nome', 'nome_localizacao'])

                for index, row in df.iterrows():
                    cursor.execute("INSERT INTO db_colaboradores (nome_colab, matricula, lotacao_nome, nome_localizacao) VALUES (?, ?, ?, ?)",
                                   (row['nome_colab'], row['matricula'], row['lotacao_nome'], row['nome_localizacao']))
                banco.commit()

                atualizar_treeview()

                messagebox.showinfo("Sucesso", "Dados carregados com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar dados da planilha: {str(e)}")
    
    def atualizar_treeview():
        df_colaboradores = pd.read_sql_query("SELECT * FROM db_colaboradores", banco)
        treeview_colaboradores.delete(*treeview_colaboradores.get_children())
        for index, row in df_colaboradores.iterrows():
            treeview_colaboradores.insert("", tk.END, text=index, values=row.tolist())

    # Obtém os valores dos campos de entrada de texto de colaboradores
    def salvar_colaboradores():
        nome_colab = campo_nome_colab.get()
        matricula = campo_matricula.get()
        lotacao_nome = campo_lotacao_combobox.get()
        nome_localizacao = campo_nome_localizacao_combobox.get()

    # Verifica se o colaborador já existe no banco de dados
        if buscar_colaborador(nome_colab):
            tk.messagebox.showerror("Erro", "Colaborador já cadastrado")
        else:
            cursor.execute("INSERT INTO db_colaboradores (nome_colab, matricula, lotacao_nome, nome_localizacao) VALUES (?, ?, ?, ?)",
                       (nome_colab, matricula, lotacao_nome, nome_localizacao))
            banco.commit()

        # Atualiza a Treeview com os novos dados
        df_colaboradores = pd.read_sql_query("SELECT * FROM db_colaboradores", banco)
        treeview_colaboradores.delete(*treeview_colaboradores.get_children())
        for index, row in df_colaboradores.iterrows():
            treeview_colaboradores.insert("", tk.END, text=index, values=row.tolist())

        janela_colaboradores.destroy()  # Destrói frame de cadastro de colaboradores

    # Cria novo frame chamada Cadastro de Colaboradores
    janela_colaboradores = tk.Toplevel()
    janela_colaboradores.title('Cadastro de Colaboradores')

    # Permitir redimensionamento horizontal e vertical
    janela_colaboradores.resizable(True, True)

    # Configurar colunas para expansão horizontal
    janela_colaboradores.columnconfigure(0, weight=1)
    janela_colaboradores.columnconfigure(1, weight=1)

    # Ao usuário digitar o nome completo do colaborador, ele atualiza todo os campos com o cadastro do o mesmo.
    def atualizar_campos(nome_colab):
        valores = buscar_colaborador(nome_colab.lower())
        if valores:
            # Limpa os campos anteriores
            campo_matricula.delete(0, tk.END)
            campo_lotacao_combobox.set("")  # Clear the Combobox selection
            campo_nome_localizacao_combobox.set("")  # Clear the Combobox selection

            # Insere os novos valores
            campo_matricula.insert(0, valores[0])
            campo_lotacao_combobox.set(valores[1])  # Set the Combobox value
            campo_nome_localizacao_combobox.set(valores[2])  # Set the Combobox value

    # Cria os rótulos e os campos de entrada de texto
    rotulo_nome_colab = tk.Label(janela_colaboradores, text='Colaborador:')
    rotulo_nome_colab.grid(row=1, column=0, pady=10, padx=10)
    campo_nome_colab = tk.Entry(janela_colaboradores)
    campo_nome_colab.grid(row=1, column=1, pady=10, padx=10, sticky="ew")  # Usa sticky para expandir horizontalmente

    rotulo_matricula = tk.Label(janela_colaboradores, text='Chave do Colaborador:')
    rotulo_matricula.grid(row=2, column=0, pady=10, padx=10)
    campo_matricula = tk.Entry(janela_colaboradores)
    campo_matricula.grid(row=2, column=1, pady=10, padx=10, sticky="ew")  # Usa sticky para expandir horizontalmente

    rotulo_lotacao_nome = tk.Label(janela_colaboradores, text='Lotação:')
    rotulo_lotacao_nome.grid(row=3, column=0, pady=10, padx=10)

    # Usa Combobox para lotação
    lotacao_options = obter_opcoes_lotacao()
    campo_lotacao_combobox = ttk.Combobox(janela_colaboradores, values=lotacao_options)
    campo_lotacao_combobox.grid(row=3, column=1, pady=10, padx=10, sticky="ew")  # Usa sticky para expandir horizontalmente

    rotulo_nome_localizacao = tk.Label(janela_colaboradores, text='Localização:')
    rotulo_nome_localizacao.grid(row=4, column=0, pady=10, padx=10)

    # Usa Combobox para localização
    localizacao_options = obter_opcoes_localizacao()
    campo_nome_localizacao_combobox = ttk.Combobox(janela_colaboradores, values=localizacao_options)
    campo_nome_localizacao_combobox.grid(row=4, column=1, pady=10, padx=10, sticky="ew")  # Usa sticky para expandir horizontalmente

    # Usa uma coluna fictícia para esticar verticalmente
    janela_colaboradores.grid_columnconfigure(2, weight=1)

    campo_nome_colab.bind("<KeyRelease>", lambda event: atualizar_campos(campo_nome_colab.get()))

    botao_salvar = tk.Button(janela_colaboradores, text='Salvar', command=salvar_colaboradores)
    botao_salvar.grid(row=6, columnspan=2, pady=10, padx=10)

    botao_carregar_planilha = tk.Button(janela_colaboradores, text='Carregar Planilha', command=carregar_planilha)
    botao_carregar_planilha.grid(row=5, columnspan=2, pady=10, padx=10)

def obter_opcoes_localizacao():
# Função para buscar opções do banco de dados
    cursor.execute("SELECT nome_localizacao FROM db_localizacao")
    nome_localizacoes = cursor.fetchall()
    return [nome_localizacao[0] for nome_localizacao in nome_localizacoes]

# Define a função de cadastro de lotação
def cadastrar_lotacao():
    # Carrega uma planilha do Excel de qualquer formato ou versão do software, para carregamento em lote as informações de colaboradores
    def carregar_planilha():
        file_path = filedialog.askopenfilename(filetypes=[('Planilhas Excel', '*.xls;*.xltx;*.xlsm;*.xlsx:'), ('Todos os arquivos', '*.*')])
        if file_path:
            try:
                df = pd.read_excel(file_path, header=None, names=['id_lotacao', 'nome_lotacao'])

                for index, row in df.iterrows():
                    cursor.execute("INSERT INTO db_lotacao (id_lotacao, nome_lotacao) VALUES (?, ?)",
                                   (row['id_lotacao'], row['nome_lotacao']))
                banco.commit()

                atualizar_treeview()

                messagebox.showinfo("Sucesso", "Dados carregados com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar dados da planilha: {str(e)}")
    
    def atualizar_treeview():
        df_lotacao = pd.read_sql_query("SELECT * FROM db_lotacao", banco)
        treeview_lotacao.delete(*treeview_colaboradores.get_children())
        for index, row in df_lotacao.iterrows():
            treeview_lotacao.insert("", tk.END, text=index, values=row.tolist())

    def salvar_lotacao():
        # Obtém o valor do campo nome_lotação
        nome_lotacao = campo_nome_lotacao.get()
        
        # Remove o caractere "/" do valor
        id_lotacao = nome_lotacao.replace('/', '').upper()  # Convertido para maiúsculas e sem "/"
        
        # Define o valor no campo ID_lotação
        campo_id_lotacao.delete(0, tk.END)  # Limpa o campo ID_lotação
        campo_id_lotacao.insert(0, id_lotacao)  # Define o novo valor
        
        # Resto do código para salvar a lotação no banco de dados permanece o mesmo
        cursor.execute("INSERT INTO db_lotacao (id_lotacao, nome_lotacao) VALUES (?, ?)",
                       (id_lotacao, nome_lotacao))
        banco.commit()

        # Atualiza a Treeview com os novos dados
        df_lotacao = pd.read_sql_query("SELECT * FROM db_lotacao", banco)
        treeview_lotacao.delete(*treeview_lotacao.get_children())
        for index, row in df_lotacao.iterrows():
            treeview_lotacao.insert("", tk.END, text=index, values=row.tolist())

        janela_lotacao.destroy()

    # Cria novo frame chamada Cadastro de Lotação
    janela_lotacao = tk.Toplevel()
    janela_lotacao.title('Cadastro de Lotação')
    
    # Cria os rótulos e os campos de entrada de texto de Cadastro de Lotação
    rotulo_id_lotacao = tk.Label(janela_lotacao, text='ID da Lotação:')
    rotulo_id_lotacao.grid(row=0, column=0, pady=10, padx=10)
    campo_id_lotacao = tk.Entry(janela_lotacao, state="disabled")
    campo_id_lotacao.grid(row=0, column=1, pady=10, padx=10)

    rotulo_nome_lotacao = tk.Label(janela_lotacao, text='Nome da Lotação:')
    rotulo_nome_lotacao.grid(row=1, column=0, pady=10, padx=10)
    campo_nome_lotacao = tk.Entry(janela_lotacao)
    campo_nome_lotacao.grid(row=1, column=1, pady=10, padx=10)

    botao_salvar = tk.Button(janela_lotacao, text='Salvar', command=salvar_lotacao)
    botao_salvar.grid(row=2, columnspan=2, pady=10, padx=10)

    botao_carregar_planilha = tk.Button(janela_lotacao, text='Carregar Planilha', command=carregar_planilha)
    botao_carregar_planilha.grid(row=3, columnspan=2, pady=10, padx=10)

# Cria função de cadastro de localizações
def cadastrar_localizacao():
    def salvar_localizacao():
        id_localizacao = campo_id_localizacao.get()
        nome_localizacao = campo_nome_localizacao.get()

        cursor.execute("INSERT INTO db_localizacao (id_localizacao, nome_localizacao) VALUES (?, ?)",
                       (id_localizacao, nome_localizacao))
        banco.commit()

        # Atualizar a Treeview com os novos dados
        df_localizacao = pd.read_sql_query("SELECT * FROM db_localizacao", banco)
        treeview_localizacao.delete(*treeview_localizacao.get_children())
        for index, row in df_localizacao.iterrows():
            treeview_localizacao.insert("", tk.END, text=index, values=row.tolist())

        janela_localizacao.destroy()

    # Criar uma nova janela
    janela_localizacao = tk.Toplevel()
    janela_localizacao.title('Cadastro de Localização')

    # Permitir redimensionamento horizontal e vertical
    janela_localizacao.resizable(True, True)

    # Configurar colunas para expansão horizontal
    janela_localizacao.columnconfigure(0, weight=1)
    janela_localizacao.columnconfigure(1, weight=1)

    # Gerar um número inteiro entre 1000 e 9999
    numero_aleatorio = random.randint(1000, 9999)

    # Concatena o prefixo com o número
    codigo_aleatorio = "CENPES-LOC" + str(numero_aleatorio)
    
    # Criar os rótulos e os campos de entrada de texto
    rotulo_id_localizacao = tk.Label(janela_localizacao, text='ID de Localização:')
    rotulo_id_localizacao.grid(row=0, column=0, pady=10, padx=10, stick="ew")
    campo_id_localizacao = tk.Entry(janela_localizacao)
    campo_id_localizacao.grid(row=0, column=1, pady=10, padx=10, stick="ew")

    rotulo_nome_localizacao = tk.Label(janela_localizacao, text='Nome da Localização:')
    rotulo_nome_localizacao.grid(row=1, column=0, pady=10, padx=10, stick="ew")
    
    # Use a Combobox instead of an Entry for nome_localizacao
    nome_localizacao_options = obter_opcoes_nome_localizacao()
    campo_nome_localizacao = ttk.Combobox(janela_localizacao, values=nome_localizacao_options)
    campo_nome_localizacao.grid(row=1, column=1, pady=10, padx=10, stick="ew")

    campo_id_localizacao.insert(0, codigo_aleatorio)

    botao_salvar = tk.Button(janela_localizacao, text='Salvar', command=salvar_localizacao)
    botao_salvar.grid(row=2, columnspan=2, pady=10, padx=10)
    
# Função para obter informações de localização no banco de dados
def obter_opcoes_nome_localizacao():
    cursor.execute("SELECT nome_localizacao FROM db_localizacao")
    nome_localizacoes = cursor.fetchall()
    return [nome_localizacao[0] for nome_localizacao in nome_localizacoes]

# Função de Cadastro de Alocações
def cadastrar_alocacao():
    # Função para salvar informações de alocação
    def salvar_alocacao():
        # Obtém os valores dos campos de entrada de texto
        matricula = campo_matricula.get()
        nome_colab = campo_nome_colab.get()
        lotacao_nome = campo_lotacao_nome.get()
        nome_localizacao = campo_nome_localizacao.get()
        id_bp = campo_id_bp.get()
        tipo_recurso = campo_tipo_recurso.get()
        modelo = campo_modelo.get()
        sistema_operacional = campo_sistema_operacional.get()
        status = campo_status.get()
        tipo_uso = campo_tipo_uso.get()
        data_entrega = campo_data_entrega.get_date()
        data_devolucao = campo_data_devolucao.get_date()
        foto_contraprova = campo_foto_contraprova.get()

        # Cria uma nova conexão para esta função
        with sqlite3.connect('banco_dados.db') as banco:
            cursor = banco.cursor()

            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='db_alocacao'")
            tabela_existe = cursor.fetchone()

            if tabela_existe:
                # A tabela existe, continue com a inserção de dados
                cursor.execute("INSERT INTO db_alocacao (matricula, nome_colab, lotacao_nome, nome_Localizacao, id_bp, tipo_recurso, modelo, sistema_operacional, status, tipo_uso, data_entrega, data_devolucao, foto_contraprova) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                (matricula, nome_colab, lotacao_nome, nome_localizacao, id_bp, tipo_recurso, modelo, sistema_operacional, status, tipo_uso, data_entrega, data_devolucao, foto_contraprova))
            else:
                print("A tabela db_alocacao não existe.")

            # Insere os valores na Treeview de alocação
            item = treeview_alocacao.insert("", tk.END, values=(matricula, nome_colab, lotacao_nome, nome_localizacao, id_bp, tipo_recurso, modelo, sistema_operacional,
                                                         status, tipo_uso, data_entrega,
                                                         data_devolucao, foto_contraprova))

        # Commit após fechar a janela
        banco.commit()
        janela_alocacao.destroy()

    def carregar_informacoes_colaborador(event):
        matricula = campo_matricula.get()

        # Conectar ao banco de dados
        conn = sqlite3.connect('banco_dados.db')
        cursor = conn.cursor()

        # Consultar o banco de dados para obter as informações do colaborador
        cursor.execute("SELECT nome_colab, lotacao_nome, nome_localizacao FROM db_colaboradores WHERE matricula = ?", (matricula,))
        resultado = cursor.fetchone()

        # Fechar a conexão com o banco de dados
        conn.close()

        if resultado:
            campo_nome_colab.delete(0, tk.END)
            campo_lotacao_nome.delete(0, tk.END)
            campo_nome_localizacao.delete(0, tk.END)
            campo_nome_colab.insert(0, resultado[0])
            campo_lotacao_nome.insert(0, resultado[1])
            campo_nome_localizacao.insert(0, resultado[2])
    
    def carregar_informacoes_recurso(event):
        id_bp = campo_id_bp.get()

        # Conectar ao banco de dados
        conn = sqlite3.connect('banco_dados.db')
        cursor = conn.cursor()

        # Consultar o banco de dados para obter as informações do recurso
        cursor.execute("SELECT tipo_recurso, modelo, sistema_operacional, status FROM db_recursos WHERE id_bp = ?", (id_bp,))
        resultado = cursor.fetchone()

        # Fechar a conexão com o banco de dados
        conn.close()

        if resultado:
            campo_tipo_recurso.delete(0, tk.END)
            campo_modelo.delete(0, tk.END)
            campo_sistema_operacional.delete(0, tk.END)
            campo_status.delete(0, tk.END)
            campo_tipo_recurso.insert(0, resultado[0])
            campo_modelo.insert(0, resultado[1])
            campo_sistema_operacional.insert(0, resultado[2])
            campo_status.insert(0, resultado[3])

        # Atualiza a disponibilidade do campo_data_devolucao com base no sistema_operacional e status
        habilitar_desabilitar_data_devolucao()

    def habilitar_desabilitar_data_devolucao():
        status = campo_status.get()

        if status == "Devolvido":
            campo_data_devolucao.config(state='normal')
        else:
            campo_data_devolucao.delete(0, tk.END)
            campo_data_devolucao.config(state='disabled')

    # Atualiza a disponibilidade do campo_sistema_operacional com base no tipo_recurso
    def atualizar_campos_recurso(*args):
        tipo_recurso = campo_tipo_recurso.get()
        if tipo_recurso in ["Estação", "Notebook"]:
            campo_sistema_operacional.config(state='normal')
        else:
            campo_sistema_operacional.delete(0, tk.END)
            campo_sistema_operacional.config(state='disabled')
    # Cria uma nova janela
    janela_alocacao = tk.Toplevel()
    janela_alocacao.title('Cadastro de Alocação')

    # Permitir redimensionamento horizontal e vertical
    janela_alocacao.resizable(True, True)

    # Configurar colunas para expansão horizontal
    janela_alocacao.columnconfigure(0, weight=1)
    janela_alocacao.columnconfigure(1, weight=1)
    
    # Cria os rótulos e os campos de entrada de texto
    rotulo_matricula = tk.Label(janela_alocacao, text='Chave do Colaborador:')
    rotulo_matricula.grid(row=0, column=0, pady=10, padx=10, sticky="ew")
    campo_matricula = ttk.Combobox(janela_alocacao)
    campo_matricula.grid(row=0, column=1, pady=10, padx=10, sticky="ew")
    campo_matricula.bind("<<ComboboxSelected>>", carregar_informacoes_colaborador)

    # Preencher as opções do dropdown com as matrículas existentes no banco de dados
    conn = sqlite3.connect('banco_dados.db')
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT matricula FROM db_colaboradores")
    matriculas = [matricula[0] for matricula in cursor.fetchall()]
    campo_matricula['values'] = matriculas
    conn.close()

    rotulo_nome_colab = tk.Label(janela_alocacao, text='Nome do Colaborador:')
    rotulo_nome_colab.grid(row=1, column=0, pady=10, padx=10, sticky="ew")
    campo_nome_colab = tk.Entry(janela_alocacao)
    campo_nome_colab.grid(row=1, column=1, pady=10, padx=10, sticky="ew")

    rotulo_lotacao_nome = tk.Label(janela_alocacao, text='Nome da Lotação:')
    rotulo_lotacao_nome.grid(row=3, column=0, pady=10, padx=10, sticky="ew")
    campo_lotacao_nome = tk.Entry(janela_alocacao)
    campo_lotacao_nome.grid(row=3, column=1, pady=10, padx=10, sticky="ew")

    rotulo_nome_localizacao = tk.Label(janela_alocacao, text='Nome da Localização:')
    rotulo_nome_localizacao.grid(row=5, column=0, pady=10, padx=10, sticky="ew")
    campo_nome_localizacao = tk.Entry(janela_alocacao)
    campo_nome_localizacao.grid(row=5, column=1, pady=10, padx=10, sticky="ew")

    rotulo_id_bp = tk.Label(janela_alocacao, text='ID/BP:')
    rotulo_id_bp.grid(row=6, column=0, pady=10, padx=10, sticky="ew")
    campo_id_bp = ttk.Combobox(janela_alocacao)
    campo_id_bp.grid(row=6, column=1, pady=10, padx=10, sticky="ew")
    campo_id_bp.bind("<<ComboboxSelected>>", carregar_informacoes_recurso)

    rotulo_tipo_recurso = tk.Label(janela_alocacao, text='Tipo de Recurso:')
    rotulo_tipo_recurso.grid(row=8, column=0, pady=10, padx=10, sticky="ew")
    campo_tipo_recurso = ttk.Combobox(janela_alocacao)
    campo_tipo_recurso.grid(row=8, column=1, pady=10, padx=10, sticky="ew")
    campo_tipo_recurso.bind("<<ComboboxSelected>>", atualizar_campos_recurso)
    conn = sqlite3.connect('banco_dados.db')
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT tipo_recurso FROM db_recursos")
    tipos_recurso = [tipo[0] for tipo in cursor.fetchall()]
    campo_tipo_recurso['values'] = tipos_recurso

    rotulo_modelo = tk.Label(janela_alocacao, text='Modelo:')
    rotulo_modelo.grid(row=9, column=0, pady=10, padx=10, sticky="ew")
    campo_modelo = tk.Entry(janela_alocacao)
    campo_modelo.grid(row=9, column=1, pady=10, padx=10, sticky="ew")

    rotulo_sistema_operacional = tk.Label(janela_alocacao, text='Sistema Operacional:')
    rotulo_sistema_operacional.grid(row=10, column=0, pady=10, padx=10, sticky="ew")
    campo_sistema_operacional = tk.Entry(janela_alocacao)
    campo_sistema_operacional.grid(row=10, column=1, pady=10, padx=10, sticky="ew")

    rotulo_status = tk.Label(janela_alocacao, text='Status:')
    rotulo_status.grid(row=11, column=0, pady=10, padx=10, sticky="ew")
    campo_status = tk.Entry(janela_alocacao)
    campo_status.grid(row=11, column=1, pady=10, padx=10, sticky="ew")

    rotulo_tipo_uso = tk.Label(janela_alocacao, text='Tipo de Uso:')
    rotulo_tipo_uso.grid(row=7, column=0, pady=10, padx=10, sticky="ew")
    tipos_uso = ['Individual', 'Coletivo']
    campo_tipo_uso = ttk.Combobox(janela_alocacao, values=tipos_uso)
    campo_tipo_uso.grid(row=7, column=1, pady=10, padx=10, sticky="ew")
    
    rotulo_data_entrega = tk.Label(janela_alocacao, text='Data de Entrega:')
    rotulo_data_entrega.grid(row=12, column=0, pady=10, padx=10, sticky="ew")
    campo_data_entrega = DateEntry(janela_alocacao, locale='pt_BR')
    campo_data_entrega.grid(row=12, column=1, pady=10, padx=10, sticky="ew")

    rotulo_data_devolucao = tk.Label(janela_alocacao, text='Data de Devolução:')
    rotulo_data_devolucao.grid(row=13, column=0, pady=10, padx=10, sticky="ew")

    # Adiciona opções ao Combobox campo_db_recursos
    conn = sqlite3.connect('banco_dados.db')
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT id_bp FROM db_recursos")
    ids_bp = [id_bp[0] for id_bp in cursor.fetchall()]
    campo_id_bp['values'] = ids_bp
    conn.close()
    
    # Cria o campo de entrada de texto para a data de devolução
    campo_data_devolucao = DateEntry(janela_alocacao, locale='pt_BR', state='disabled')  # O estado inicial é 'disabled'
    campo_data_devolucao.grid(row=13, column=1, pady=10, padx=10, sticky="ew")

    # Cria o campo de entrada de texto com caminho da foto contra prova
    rotulo_foto_contraprova = tk.Label(janela_alocacao, text='Caminho da Foto: ')
    rotulo_foto_contraprova.grid(row=14, column=0, pady=10, padx=10, sticky="ew")
    campo_foto_contraprova = tk.Entry(janela_alocacao)
    campo_foto_contraprova.grid(row=14, column=1, pady=10, padx=10, sticky="ew")

    botao_salvar = tk.Button(janela_alocacao, text='Salvar', command=salvar_alocacao)
    botao_salvar.grid(row=16, columnspan=2, pady=10, padx=10)

# Funções para Exportação de Dados do Banco de Dados em Planilha
def exportar_colaboradores():
    try:
        # Obtém os dados do banco de dados
        df_colaboradores = pd.read_sql_query("SELECT * FROM db_colaboradores", banco)

        # Pede ao usuário para escolher o local do arquivo Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])

        if file_path:
            # Salva os dados no arquivo Excel
            df_colaboradores.to_excel(file_path, index=False)

            messagebox.showinfo("Sucesso", "Dados exportados para o Excel com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar dados para o Excel: {str(e)}")

def exportar_lotacao():
    try:
        # Obtém os dados do banco de dados
        df_lotacao = pd.read_sql_query("SELECT * FROM db_lotacao", banco)

        # Pede ao usuário para escolher o local do arquivo Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])

        if file_path:
            # Salva os dados no arquivo Excel
            df_lotacao.to_excel(file_path, index=False)

            messagebox.showinfo("Sucesso", "Dados exportados para o Excel com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar dados para o Excel: {str(e)}")

def exportar_localizacao():
    try:
        # Obtém os dados do banco de dados
        df_localizacao = pd.read_sql_query("SELECT * FROM db_localizacao", banco)

        # Pede ao usuário para escolher o local do arquivo Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])

        if file_path:
            # Salva os dados no arquivo Excel
            df_localizacao.to_excel(file_path, index=False)

            messagebox.showinfo("Sucesso", "Dados exportados para o Excel com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar dados para o Excel: {str(e)}")

def exportar_recursos():
    try:
        # Obtém os dados do banco de dados
        df_recursos = pd.read_sql_query("SELECT * FROM db_recursos", banco)

        # Pede ao usuário para escolher o local do arquivo Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])

        if file_path:
            # Salva os dados no arquivo Excel
            df_recursos.to_excel(file_path, index=False)

            messagebox.showinfo("Sucesso", "Dados exportados para o Excel com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar dados para o Excel: {str(e)}")

def exportar_alocacoes():
    try:
        # Obtém os dados do banco de dados
        df_alocacao = pd.read_sql_query("SELECT * FROM db_recursos", banco)

        # Pede ao usuário para escolher o local do arquivo Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])

        if file_path:
            # Salva os dados no arquivo Excel
            df_recursos.to_excel(file_path, index=False)

            messagebox.showinfo("Sucesso", "Dados exportados para o Excel com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar dados para o Excel: {str(e)}")

# Função para imprimir o banco de dados
def imprimir_banco_dados(tabela):
    # Lógica para imprimir a tabela do banco de dados
    print(f"Imprimir tabela: {tabela}")

# Cria a barra de menu
barra_menu = tk.Menu(janela)

# Configura a barra de menu
janela.config(menu=barra_menu)

# Cria a barra de menu com a opção Arquivo
arquivo = Menu(barra_menu, tearoff=0)
barra_menu.add_cascade(label="Arquivo", menu=arquivo)

# Adiciona a opção Imprimir db_recursos ao menu Arquivo
arquivo.add_command(label="Gerar/Imprimir Relatório", command=imprimir_db_alocacoes)

# Atualize o comando "Sair" para usar a função com confirmação
arquivo.add_command(label="Sair", command=sair_com_confirmacao)

# Cria a barra de menu com a opção Cadastrar
cadastro = Menu(barra_menu, tearoff=0)
barra_menu.add_cascade(label="Cadastrar/Carregar", menu=cadastro)

# Comando para menu de Cadastro de Recursos
cadastro.add_command(label="Alocação", command=cadastrar_alocacao)
# Comando para menu de Cadastro de Recursos
cadastro.add_command(label="Recursos", command=cadastrar_recurso)
# Comando para menu de Cadastro de Colaboradores
cadastro.add_command(label="Colaboradores", command=cadastrar_colaboradores)
# Comando para menu de Cadastro de Lotações
cadastro.add_command(label="Lotações", command=cadastrar_lotacao)
# Comando para menu de Cadastro de Localização
cadastro.add_command(label="Localizações", command=cadastrar_localizacao)

# Cria a barra de menu com a opção Cadastrar
exportar = Menu(barra_menu, tearoff=0)
barra_menu.add_cascade(label="Exportar Dados", menu=exportar)

# Comando para menu de Exportação de Recursos
exportar.add_command(label="Alocações", command=exportar_alocacoes)
# Comando para menu de Exportação de Recursos
exportar.add_command(label="Recursos", command=exportar_recursos)
# Comando para menu de Exportação de Colaboradores
exportar.add_command(label="Colaboradores", command=exportar_colaboradores)
# Comando para menu de Exportação de Lotações
exportar.add_command(label="Lotações", command=exportar_lotacao)
# Comando para menu de Exportação de Localização
exportar.add_command(label="Localizações", command=exportar_localizacao)

# Função para 
def ajuda():
    messagebox.showinfo("Suporte", "Marcos Paulo Pereira Dias  Chave: D2ES")

# Cria a barra Sobre
sobre = Menu(barra_menu, tearoff=0)
barra_menu.add_cascade(label="Sobre", menu=sobre)

# Comando para menu de Exportação de Recursos
sobre.add_command(label="Suporte", command=ajuda)

# Configurar o protocolo para o botão "X" da janela
janela.protocol("WM_DELETE_WINDOW", sair_com_confirmacao)

# Executa o comando SQL para selecionar todos os dados das tabelas
cursor.execute("SELECT * FROM db_recursos")
df_recursos = pd.read_sql_query("SELECT * FROM db_recursos", banco)

cursor.execute("SELECT * FROM db_colaboradores")
df_colaboradores = pd.read_sql_query("SELECT * FROM db_colaboradores", banco)

cursor.execute("SELECT * FROM db_alienacao")
df_alienacao = pd.read_sql_query("SELECT * FROM db_alienacao", banco)

cursor.execute("SELECT * FROM db_lotacao")
df_lotacao = pd.read_sql_query("SELECT * FROM db_lotacao", banco)

cursor.execute("SELECT * FROM db_localizacao")
df_localizacao = pd.read_sql_query("SELECT * FROM db_localizacao", banco)

cursor.execute("SELECT * FROM db_alocacao")
df_alocacao = pd.read_sql_query("SELECT * FROM db_alocacao", banco)

# Adiciona os dados à Treeview
for index, row in df_recursos.iterrows():
    treeview_recursos.insert("", tk.END, text=index, values=row.tolist())

for index, row in df_colaboradores.iterrows():
    treeview_colaboradores.insert("", tk.END, text=index, values=row.tolist())

for index, row in df_alienacao.iterrows():
    treeview_alienacao.insert("", tk.END, text=index, values=row.tolist())

for index, row in df_lotacao.iterrows():
    treeview_lotacao.insert("", tk.END, text=index, values=row.tolist())

for index, row in df_localizacao.iterrows():
    treeview_localizacao.insert("", tk.END, text=index, values=row.tolist())

for index, row in df_alocacao.iterrows():
    treeview_alocacao.insert("", tk.END, text=index, values=row.tolist())

# Inicia o loop principal da janela
janela.mainloop()
