from tkinter import *
from functions import *
from ttkthemes import ThemedTk
from tkinter import ttk
import pandas as pd
import os
from tkinter import filedialog
from tkinter import messagebox

class Janela_principal():
    def __init__(self):
        self.window = ThemedTk(theme='breeze')
        self.window.title("Automação")
        self.window.minsize(width=500, height=600)
        # Obter coordenadas atuais do mouse
        self.posicao_x_mouse = self.window.winfo_pointerx()
        self.posicao_y_mouse = self.window.winfo_pointery()
        self.window.geometry(f"500x600+{self.posicao_x_mouse}+{self.posicao_y_mouse}")
        self.window.resizable(0,0)
        self.window.config(bg='#294380')

        # Frame do título 
        self.frame_titulo = Frame(self.window, bg='#294380')
        self.frame_titulo.pack(fill='x')
        self.titulo_label = Label(self.frame_titulo, text='Bot Fiscalização', font='Arial 20 bold', fg='white', bg='#294380')
        self.titulo_label.pack(fill='x', pady=50)
        
        # Frame para widgets 
        self.frame_widgets = Frame(self.window, bg='#294380')
        self.frame_widgets.pack(fill='x')
        self.label_notas = Label(self.frame_widgets, text='  Insira as notas: ', font='Arial 12 bold', bg='#294380', fg='white')
        self.label_notas.pack(side='left', pady=10)
        
        # Frame para escrever as notas 
        self.frame_text_notas = Frame(self.window, bg='#294380')
        self.frame_text_notas.pack()
        self.text_area_notas = Text(self.frame_text_notas, font="Arial 12", bg='white', height=10, width=45)
        self.text_area_notas.pack()
        
        # Frame Para os botões
        self.frame_buttons = Frame(self.window, bg='#294380', pady=13)
        self.frame_buttons.pack()
        self.btn_iniciar = ttk.Button(self.frame_buttons, text='Iniciar', command=self.iniciar_program)
        self.btn_iniciar.pack(side='left', padx=10)
        self.btn_DXF = ttk.Button(self.frame_buttons, text='DXF', command=lambda: ExtrairDxf(self.notas, self.poste))
        self.btn_DXF.pack(side='left', padx=10)
        self.btn_EXCEL = ttk.Button(self.frame_buttons, text='Excel', command=lambda: ExtrairLista(self.notas))
        self.btn_EXCEL.pack(side='left', padx=10)
        self.btn_relatorio = ttk.Button(self.frame_buttons, text='Relatório', command=self.create_report)
        self.btn_relatorio.pack(side='left', padx=10)
        
        # Frame status
        self.frame_status = Frame(self.window, bg='#294380', bd=1)
        self.frame_status.pack(fill='x')
        self.status = Label(self.frame_status, text= "", bg='#294380', font='Arial 12', fg='white', relief='groove', justify='left', height=5)
        self.status.pack(fill='x')
        
        self.frame_sair = Frame(self.window)
        self.frame_sair.pack(pady=10)
        self.btn_sair = ttk.Button(self.frame_sair, text='Sair', command=self.sair_programa)
        self.btn_sair.pack()

        self.sair = False  # Sinalizador para verificar se o usuário solicitou sair

        self.window.mainloop()

    def iniciar_program(self):
        
        self.lista_notas = self.text_area_notas.get(1.0, END).strip()
        if self.lista_notas == "":
            messagebox.showinfo("Atenção!", "Não tem notas para processar.")
            return
        
        # Divide as notas em uma lista
        lista = [nota.strip() for nota in self.lista_notas.splitlines()]
        
        self.total_notas = []
        self.total_postes = []
        self.observacoes = []
        self.contador = 1
        self.geral = len(lista)
        self.notas_restantes = self.geral

        for nota in lista:
            if self.sair:  # Verifica se o usuário solicitou sair
                break  # Encerra o loop se solicitou sair
            
            self.notas = nota
            self.status['text'] = f"{self.contador} de {self.geral} nota(s) processadas.\nProcessando a nota {self.notas}.\nNotas restantes: {self.notas_restantes - 1}  "
            EntrarNota(nota)
            # Abre a janela para poste
            self.abrir_janela_poste()
            self.window.wait_window(self.window_poste)
            # Espera a janela_poste ser fechada
            if not (self.status_obs == "sem lista tecnica" or self.status_obs == "bloqueada"):
                self.status["text"] = f"{self.contador} de {self.geral} nota(s) processadas.\nExtraindo Lista Técnica\n da nota {nota}."
                ExtrairLista(nota)
                self.status["text"] = f"{self.contador} de {self.geral} nota(s) processadas.\nLista Técnica da nota {nota}\n extraída com sucesso."
            self.continuar()
            self.window.wait_window(self.window_msg)
            self.status["text"] = ""
            fecharNota()
            self.total_notas.append(self.notas)
            self.total_postes.append(self.poste)
            self.observacoes.append(self.status_obs)
            self.contador +=1
        messagebox.showinfo("Concluído!","Processamento concluído")
       
            
    def abrir_janela_poste(self):
        if self.sair:  # Verifica se o usuário solicitou sair antes de abrir a janela
            return
        
        self.window_poste = Toplevel()
        self.window_poste.title('Quantidade de poste')
        self.window_poste.resizable(0, 0)
        
        self.window_poste.geometry(f'400x150+{self.posicao_x_mouse}+{self.posicao_y_mouse}')

        # Primeiro Frame (self.frame1)
        self.frame1 = Frame(self.window_poste)
        self.frame1.pack(padx=10,pady=10)
        self.label_poste = Label(self.frame1, text=f'N° de poste da nota {self.notas}', font='Arial 12 bold ', fg='black')
        self.label_poste.pack()
        self.frame2 = Frame(self.window_poste)
        self.frame2.pack()
        self.entry_poste = ttk.Entry(self.frame2, font='Arial 9')
        self.entry_poste.pack(side='left')
        
        self.btn = ttk.Button(self.frame2, text='OK', command=self.submit_poste, width=4)
        self.btn.pack(side="left")
        
        # Segundo Frame (self.frame2)
        self.frame2 = Frame(self.window_poste,  bg='#294380')
        self.frame2.pack(pady=20)
  
        # botão check
        self.observacoes_status = StringVar()
        
        self.button_check1 = ttk.Radiobutton(self.frame2, text='Sem lista técnica', value="sem lista tecnica", variable=self.observacoes_status)
        self.button_check1.pack(side='left')
        self.button_check2 = ttk.Radiobutton(self.frame2, text='Bloqueada', value="bloqueada", variable=self.observacoes_status)
        self.button_check2.pack(side='left')

    def submit_poste(self):
        self.status_obs = self.observacoes_status.get()
        self.poste = self.entry_poste.get()
        if self.poste == "":
            self.poste = 0
        self.window_poste.destroy()
        self.status["text"] = f"{self.contador} de {self.geral} nota(s) processadas.\nExtraindo o DXF da nota {self.notas} "
        if not (self.status_obs == "sem lista tecnica" or self.status_obs == "bloqueada"):
           ExtrairDxf(self.notas, self.poste)                                                              

    def sair_programa(self):
        self.sair = True  # Define o sinalizador de saída como True

        # Fecha todas as janelas abertas
        if hasattr(self, 'window_poste') and isinstance(self.window_poste, Toplevel):
            self.window_poste.destroy()
        self.window.destroy()
    
    def create_report(self):
        # Verificar se todas as listas têm o mesmo comprimento
        if len(self.total_notas) == len(self.total_postes) == len(self.observacoes):
            # Criar o dicionário com os dados
            table = {
                'Notas': self.total_notas,
                'Postes': self.total_postes,
                'Observações': self.observacoes
            }
            # Criar DataFrame
            df = pd.DataFrame(table)
            
            # Solicitar ao usuário um nome para o arquivo
            file_name = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                                    filetypes=[('Arquivo Excel', '*.xlsx')],
                                                    title='Salvar como')
            
            # Verificar se o usuário escolheu um nome de arquivo
            if file_name:
                # Salvar o arquivo Excel com o nome fornecido pelo usuário
                df.to_excel(file_name, index=False)
        else:
            messagebox.showerror("Erro", "As listas não têm o mesmo comprimento.")

    def continuar(self):
        self.window_msg = Toplevel()
        self.window_msg.title('Continuar')
        # self.window_msg.config(bg="#294380")
        self.window_msg.resizable(0,0)
        self.window_msg.geometry(f'300x100+{self.posicao_x_mouse}+{self.posicao_y_mouse}')
        text = Label(self.window_msg, text='Aperte OK para continuar', font="Arial 12", fg="black")
        text.pack(pady=10)
        button_exit = ttk.Button(self.window_msg, text='OK', command=self.window_msg.destroy)
        button_exit.pack(pady=10)
                
# Instanciando a classe principal
Janela_principal()
