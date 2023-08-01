#Importar bibliotecas nescessarias
import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import tkinter as tk
from tkinter import filedialog
import time


# Interface gráfica
class LoginScreen:
    def __init__(self, master):
        self.master = master
        self.master.title('Tela de Login')
        master.geometry("600x400")

        # Fundo
        self.bg_label = tk.Label(self.master, bg='#008A90', width=400, height=400)
        self.bg_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        # Título
        self.title_label = tk.Label(self.master, text='Faça o login', font=('Ubuntu', 16), bg='#008A90', fg='white')
        self.title_label.place(relx=0.5, rely=0.2, anchor=tk.CENTER)

        # Widgets da tela de login
        self.username_label = tk.Label(self.master, text='Usuário:', font=('Ubuntu', 12), bg='#008A90', fg='white')
        self.username_label.place(relx=0.2, rely=0.4, anchor=tk.CENTER)
        self.username_entry = tk.Entry(self.master, width=35, font=('Ubuntu', 12))
        self.username_entry.place(relx=0.55, rely=0.4, anchor=tk.CENTER)

        self.password_label = tk.Label(self.master, text='Senha:', font=('Ubuntu', 12), bg='#008A90', fg='white')
        self.password_label.place(relx=0.2, rely=0.5, anchor=tk.CENTER)
        self.password_entry = tk.Entry(self.master, show='*', width=35, font=('Ubuntu', 12))
        self.password_entry.place(relx=0.55, rely=0.5, anchor=tk.CENTER)

        self.login_button = tk.Button(self.master, text='Login', command=self.login, bg='#388F43', fg='white', font=('Ubuntu', 12), width=10)
        self.login_button.place(relx=0.5, rely=0.65, anchor=tk.CENTER)

    # Função de login
    def login(self):
            
        # Função de autenticação
        def autenticar(smtp_server, smtp_port, username, password):
            try:
                server = smtplib.SMTP(smtp_server, smtp_port)
                print(server)
                server.starttls()
                server.login(username, password)
                return server
            except smtplib.SMTPAuthenticationError:
                return None
        smtp_port = 587 # substitua pela porta usada pelo seu servidor SMTP
        username = self.username_entry.get()
        password = self.password_entry.get()
        if "@gmail.com" in username:
            # Usar o provedor Gmail
            smtp_server ='smtp.gmail.com'
        elif "@outlook.com" in username:
            # Usar o provedor Outlook
            smtp_server ='smtp.outlook.com'
        elif "@vetoquinol.com" in username:
            # Usar o provedor Outlook
            smtp_server ='smtp-mail.outlook.com'
        elif "@yahoo.com" in username:
            # Usar o provedor Yahoo
            smtp_server ='smtp.yahoo.com'
        else:
            # Provedor desconhecido
            tk.messagebox.showerror('Erro de login', 'Provedor desconhecido.')
        server = autenticar(smtp_server, smtp_port, username, password)
        if server:
            # Se as credenciais estiverem corretas, inicie a tela principal
            self.master.destroy()
            root = tk.Tk()
            MainScreen(root,username, password)
            root.mainloop()
        else:
            # fazer algo aqui se o login falhar
            self.username_entry.delete(0, tk.END)
            self.password_entry.delete(0, tk.END)
            tk.messagebox.showerror('Erro de login', 'Credenciais inválidas.')

class MainScreen:
    def __init__(self, master, username, password):
        self.master = master
        master.title("Envio de Emails")
        master.geometry("600x400")

        # Background
        self.bg_label = tk.Label(self.master, bg='#008A90', width=400, height=400)
        self.bg_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)


        # Criação dos widgets
        self.btn_arquivo = tk.Button(master, text="Selecionar Arquivo", command=self.selecionar_arquivo, bg='#388F43', fg='white', font=('Ubuntu', 12), width=15)
        self.btn_arquivo.place(relx=0.2, rely=0.2, anchor=tk.CENTER)

        self.lbl_enviado = tk.Label(master, text="", bg='#008A90', fg='white', font=('Ubuntu', 12), width=55)
        self.lbl_enviado.place(relx=0.5, rely=0.5, anchor=tk.CENTER)


        self.lbl_nome_arquivo = tk.Label(master, text="")
        self.lbl_nome_arquivo.place(relx=0.6, rely=0.2, anchor=tk.CENTER)

        self.btn_enviar = tk.Button(master, text="Enviar", command=self.enviar_emails, bg='#388F43', fg='white', font=('Ubuntu', 12), width=15)
        self.btn_enviar.place(relx=0.2, rely=0.6, anchor=tk.CENTER)

        self.btn_fechar = tk.Button(master, text="Sair", command=master.quit, bg='#388F43', fg='white', font=('Ubuntu', 12), width=15)
        self.btn_fechar.place(relx=0.2, rely=0.7, anchor=tk.CENTER)


        self.smtp_server = "smtp.outlook.com"
        self.smtp_port = 587

        self.smtp_user = username
        self.smtp_password = password
    def exibir_mensagem(self, email, total_email):
        mensagem = f"E-mail: {email} / {total_email}"
        self.lbl_enviado.configure(text=mensagem)
    def selecionar_arquivo(self):
        self.arquivo = filedialog.askopenfilename(title="Selecione um arquivo", filetypes=[("Arquivos Excel", "*.xlsx")])
        self.lbl_nome_arquivo.config(text=self.arquivo)

    def enviar_emails(self):
        
        if not hasattr(self, 'arquivo'):
            tk.messagebox.showerror("Erro", "Selecione um arquivo antes de enviar os emails")
            return

        # Cria o objeto SMTP e faz o login
        # Cria uma conexão SSL segura
        context = ssl.create_default_context()
        with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(self.smtp_user, self.smtp_password)
            time.sleep(1)
            # Ler a planilha
            excel_file = pd.ExcelFile(self.arquivo)
            df = pd.read_excel(excel_file)

            # Salvar cada coluna em uma lista
            colaboradores = df['colaborador'].tolist()
            responsaveis = df['resposavel'].tolist()
            versoes = df['versao'].tolist()
            codigos = df['codigo'].tolist()
            descricoes = df['descricao'].tolist()
            total_email = len(colaboradores)
            email =1
                # Enviar emails
            for i in range(len(colaboradores)):
                colaborador = colaboradores[i]
                responsavel = responsaveis[i]
                versao = versoes[i]
                codigo = codigos[i]
                descricao = descricoes[i]
                destinatarios = [colaborador, responsavel]
                subject = "Mensagem Automática"
                body = f"""            
                    Prezados,
                    
                    Por gentileza, solicito que o treinamento abaixo seja realizado o mais breve possível:
                    
                    - Versão: {versao}
                    - Código: {codigo}
                    - Descrição: {descricao}
                    
                    Atenciosamente,
                    Julio Alves
                """
                
                # Monta a mensagem do e-mail
                message = MIMEMultipart()
                message['From'] = self.smtp_user
                message['Cc'] = self.smtp_user
                message['To'] =', '.join(destinatarios)
                message['Subject'] = subject
                message.attach(MIMEText(body, 'plain'))


                server.sendmail(self.smtp_user, destinatarios, message.as_string())
            
                self.exibir_mensagem(email, total_email)
                email = email + 1
                time.sleep(0.5)

                # Fecha a conexão com o servidor SMTP
            
if __name__ == '__main__':
    root = tk.Tk()
    LoginScreen(root)
    root.mainloop()