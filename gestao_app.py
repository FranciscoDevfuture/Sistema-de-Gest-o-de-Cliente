import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import openpyxl
from openpyxl import Workbook
import os

ctk.set_appearance_mode('System')
ctk.set_default_color_theme('blue')

class App(ctk.CTk):
    def __init__(self, fg_color=None, **kwargs):
        super().__init__(fg_color, **kwargs)
        self.layout_config()
        self.appearance()
        self.todo_sistema()

    def layout_config(self):
        self.title('Sistema de Gestão de Clientes')
        self.geometry('700x500')

    def appearance(self):
        self.lb_apm = ctk.CTkLabel(self, text='Tema', bg_color='transparent', text_color=['#000', '#fff'])
        self.lb_apm.place(x=50, y=430)
        self.opt_apm = ctk.CTkOptionMenu(self, values=['Light', 'Dark', 'System'], command=self.change_apm)
        self.opt_apm.place(x=50, y=460)

    def todo_sistema(self):
        frame = ctk.CTkFrame(self, width=700, height=70, corner_radius=0, bg_color='teal', fg_color='teal')
        frame.place(x=0, y=10)
        title = ctk.CTkLabel(frame, text='Sistema de Gestão de Clientes', font=('Century Gothic bold', 24), text_color='#fff')
        title.place(x=20, y=20)

        spam = ctk.CTkLabel(frame, text='Por Favor, Preencha todos os campos do formulário', font=('Century Gothic bold', 16), text_color=['#000', '#fff'])
        spam.place(x=0, y=80)

        def submit():
            # Coleta os dados dos campos
            name = name_entry.get()
            contact = contact_entry.get()
            age = age_entry.get()
            address = address_entry.get()
            gender = gender_combobox.get()
            observations = obs_entry.get("1.0", "end-1c")

            # Nome do arquivo da planilha
            file_path = 'clientes.xlsx'

            # Verifica se o arquivo já existe
            if not os.path.exists(file_path):
                # Cria um novo arquivo e adiciona os cabeçalhos
                wb = Workbook()
                ws = wb.active
                ws.append(['Nome', 'Contato', 'Idade', 'Endereço', 'Gênero', 'Observações'])  # Cabeçalhos

                # Adiciona os dados
                ws.append([name, contact, age, address, gender, observations])
                wb.save(file_path)
            else:
                # Se o arquivo existir, abre e adiciona uma nova linha
                wb = openpyxl.load_workbook(file_path)
                ws = wb.active
                ws.append([name, contact, age, address, gender, observations])
                wb.save(file_path)

            # Exibe uma mensagem de sucesso
            messagebox.showinfo("Dados Salvos", "Os dados foram salvos com sucesso!")

        def clear():
            name_entry.delete(0, END)
            contact_entry.delete(0, END)
            age_entry.delete(0, END)
            address_entry.delete(0, END)
            gender_combobox.set('Masculino')
            obs_entry.delete(1.0, END)

        # Entrys
        name_entry = ctk.CTkEntry(self, width=350, font=('Century Gothic bold', 16), fg_color='transparent')
        contact_entry = ctk.CTkEntry(self, width=200, font=('Century Gothic bold', 16), fg_color='transparent')
        age_entry = ctk.CTkEntry(self, width=150, font=('Century Gothic bold', 16), fg_color='transparent')
        address_entry = ctk.CTkEntry(self, width=200, font=('Century Gothic', 16), fg_color='transparent')

        # Combobox
        gender_combobox = ctk.CTkComboBox(self, values=['Masculino', 'Feminino'], font=('Century Gothic bold', 14))
        gender_combobox.set('Masculino')

        # Entrada de observações
        obs_entry = ctk.CTkTextbox(self, width=500, height=150, font=('arial', 18), border_color='#000', border_width=2, fg_color='transparent')

        # Labels
        lb_name = ctk.CTkLabel(self, text='Nome completo', font=('Century Gothic bold', 16), text_color=['#000', '#fff'])
        lb_contact = ctk.CTkLabel(self, text='Contato', font=('Century Gothic bold', 16), text_color=['#000', '#fff'])
        lb_age = ctk.CTkLabel(self, text='Idade', font=('Century Gothic bold', 16), text_color=['#000', '#fff'])
        lb_gender = ctk.CTkLabel(self, text='Gênero', font=('Century Gothic bold', 15), text_color=['#000', '#fff'])
        lb_address = ctk.CTkLabel(self, text='Endereço', font=('Century Gothic bold', 16), text_color=['#000', '#fff'])
        lb_obs = ctk.CTkLabel(self, text='Observações', font=('Century Gothic bold', 16), text_color=['#000', '#fff'])

        btn_submit = ctk.CTkButton(self, text='Salvar dados'.upper(), command=submit, fg_color='#151', hover_color='#131')
        btn_submit.place(x=300, y=420)
        btn_clear = ctk.CTkButton(self, text='Limpar campos'.upper(), command=clear, fg_color='#151', hover_color='#131')
        btn_clear.place(x=500, y=420)

        # Positioning elements
        lb_name.place(x=50, y=120)
        name_entry.place(x=50, y=150)
        lb_contact.place(x=450, y=120)
        contact_entry.place(x=450, y=150)
        lb_age.place(x=300, y=190)
        age_entry.place(x=300, y=220)
        lb_gender.place(x=450, y=220)
        gender_combobox.place(x=500, y=219)
        lb_address.place(x=50, y=190)
        address_entry.place(x=50, y=220)
        lb_obs.place(x=50, y=260)
        obs_entry.place(x=180, y=260)

    def change_apm(self, nova_aparencia):
        ctk.set_appearance_mode(nova_aparencia)


if __name__ == '__main__':
    app = App()
    app.mainloop()
