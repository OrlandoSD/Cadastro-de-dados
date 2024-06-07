from tkinter import *
from tkinter import messagebox
import openpyxl, xlrd
import pathlib
from openpyxl import Workbook
import customtkinter as ctk

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.layout_config()
        self.appearence()
        self.todo_sistema()

    def layout_config(self):
        self.title("Sistema de Cadastro de Rádio GCM-SP")
        self.geometry("750x500")
    
    def appearence(self):
        self.lb_apm = ctk.CTkLabel(self, text="Tema", bg_color="transparent", text_color=['#000', "#fff"])
        self.lb_apm.place(x=50, y=430)
        self.opt_apm = ctk.CTkOptionMenu(self, values=["light", "dark", "system"], command=self.change_apm)
        self.opt_apm.place(x=50, y=460)
    
    def todo_sistema(self):
        frame = ctk.CTkFrame(self, width=700, height=50, corner_radius=0, bg_color="teal", fg_color="teal")
        frame.place(x=0, y=10)
        title = ctk.CTkLabel(frame, text="Sistema de Gestão de Radios GCM-SP", font=("Century Gothic bold", 24), text_color="#fff").place(x=190, y=10)
    
        span = ctk.CTkLabel(self, text="Por favor preencha todos os campos do Formulário", font=("Century Gothic bold", 16), text_color=["#000","#fff"]).place(x=50, y=70)
    
        # Variáveis de controle
        self.name_value = StringVar()
        self.contact_value = StringVar()
        self.age_value = StringVar()
        self.address_value = StringVar()

        # Entrys
        self.name_entry = ctk.CTkEntry(self, width=350, textvariable=self.name_value, font=("Century Gothic bold", 16), fg_color="transparent")
        self.contact_entry = ctk.CTkEntry(self, width=200, textvariable=self.contact_value, font=("Century Gothic bold", 16), fg_color="transparent")
        self.age_entry = ctk.CTkEntry(self, width=150, textvariable=self.age_value, font=("Century Gothic bold", 16), fg_color="transparent")
        self.address_entry = ctk.CTkEntry(self, width=200, textvariable=self.address_value, font=("Century Gothic bold", 16), fg_color="transparent")
         
        # Combobox
        self.gender_combobox = ctk.CTkComboBox(self, values=["Operando", "Em manutenção"], font=("Century Gothic bold", 14), width=150)
        self.gender_combobox.set("Operando")
         
        self.obs_entry = ctk.CTkTextbox(self, width=750, height=150, font=("arial", 18), border_color="#aaa", border_width=2, fg_color="transparent")
         
        # Labels
        self.lb_name = ctk.CTkLabel(self, text="Modelo do Rádio", font=("Century Gothic bold", 16), text_color=["#000","#fff"]).place(x=30, y=120)
        self.lb_contact = ctk.CTkLabel(self, text="Contato da Inspetoria", font=("Century Gothic bold", 16), text_color=["#000","#fff"]).place(x=55, y=170)
        self.lb_age = ctk.CTkLabel(self, text="Serial do Rádio", font=("Century Gothic bold", 16), text_color=["#000","#fff"]).place(x=50, y=220)
        self.lb_gender = ctk.CTkLabel(self, text="Status do Rádio", font=("Century Gothic bold", 16), text_color=["#000","#fff"]).place(x=100, y=270)
        self.lb__address = ctk.CTkLabel(self, text="Local da Inspetoria", font=("Century Gothic bold", 16), text_color=["#000","#fff"]).place(x=50, y=320)
        self.lb_obs = ctk.CTkLabel(self, text="Observação", font=("Century Gothic bold", 16), text_color=["#000","#fff"]).place(x=90, y=250)
        
        # Botões
        self.btn_submit = ctk.CTkButton(self, text="Salvar dados".upper(), command=self.submit, fg_color="#151", hover_color="#131").place(x=300, y=420)
        self.btn_clear = ctk.CTkButton(self, text="Limpar Campos".upper(), command=self.clear, fg_color="#555", hover_color="#333").place(x=500, y=420)
        
        # Posicionando os elementos na janela
        self.name_entry.place(x=30, y=150)
        self.contact_entry.place(x=50, y=200)
        self.age_entry.place(x=50, y=250)
        self.gender_combobox.place(x=50, y=300)
        self.address_entry.place(x=50, y=350)
        self.obs_entry.place(x=190, y=260)
    
    def submit(self):
        ficheiro = pathlib.Path("Clientes.xlsx")
        if ficheiro.exists():
            ficheiro = openpyxl.load_workbook('Clientes.xlsx')
            folha = ficheiro.active
        else:
            ficheiro = Workbook()
            folha = ficheiro.active
            folha['A1'] = "Modelo do Rádio"
            folha['B1'] = "Contato da Inspetoria"
            folha['C1'] = "Serial do Rádio"
            folha['D1'] = "Status do Rádio"
            folha['E1'] = "Local da Inspetoria"
            folha['F1'] = "Observações"
        
        # Pegando os dados dos entrys
        name = self.name_value.get()
        contact = self.contact_value.get()
        age = self.age_value.get()
        gender = self.gender_combobox.get()
        address = self.address_value.get()
        obs = self.obs_entry.get(0.0, END)
        
        folha.append([name, contact, age, gender, address, obs])
        
        ficheiro.save(r"Clientes.xlsx")
        messagebox.showinfo("Sistema", "Dados Salvos com sucesso!")
        
    def clear(self):
        self.name_value.set("")
        self.contact_value.set("")
        self.age_value.set("")
        self.address_value.set("")
        self.obs_entry.delete(0.0, END)
    
    def change_apm(self, new_appearance):
        ctk.set_appearance_mode(new_appearance)

if __name__ == "__main__":
    app = App()
    app.mainloop()
