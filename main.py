import os
import time
import re
import logging
import threading
import smtplib
import PyPDF2
import dns.resolver
import dns.exception
import customtkinter as ctk
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from tkinter import messagebox, filedialog
from PIL import Image

# Configuração do tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class LogHandler(logging.Handler):
    """Handler personalizado para exibir logs na interface"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    
    def emit(self, record):
        msg = self.format(record)
        
        def append():
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", msg + "\n")
            self.text_widget.configure(state="disabled")
            self.text_widget.see("end")
        
        self.text_widget.after(0, append)

class PDFHandler(FileSystemEventHandler):
    """Handler para processar arquivos PDF"""
    def __init__(self, smtp_server, smtp_port, email, password, monitor_folder, 
                 sent_folder, error_folder, email_template, error_template, 
                 update_stats_callback, logger):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email = email
        self.password = password
        self.monitor_folder = monitor_folder
        self.sent_folder = sent_folder
        self.error_folder = error_folder
        self.email_template = email_template
        self.error_template = error_template
        self.update_stats = update_stats_callback
        self.logger = logger
        
        # Configuração do resolvedor DNS
        self.resolver = dns.resolver.Resolver()
        self.resolver.timeout = 5
        self.resolver.lifetime = 5
        
        # Estatísticas
        self.processed_files = 0
        self.emails_sent = 0
        self.errors = 0
    
    def validate_email(self, email):
        """Valida o formato do email e verifica os registros MX do domínio"""
        # Verificação básica de formato
        if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
            return False
        
        domain = email.split('@')[-1]
        
        try:
            # Verifica registros MX
            try:
                mx_records = self.resolver.resolve(domain, 'MX')
                if len(mx_records) == 0:
                    self.logger.warning(f"Domínio {domain} não possui registros MX válidos")
                    return False
                return True
            except dns.resolver.NoAnswer:
                self.logger.warning(f"Domínio {domain} não possui registros MX")
                return False
            except dns.resolver.NXDOMAIN:
                self.logger.warning(f"Domínio {domain} não existe")
                return False
            except dns.exception.Timeout:
                self.logger.warning(f"Timeout ao verificar MX para {domain}")
                return False
            except dns.exception.DNSException as e:
                self.logger.warning(f"Erro DNS ao verificar {domain}: {str(e)}")
                return False
        except Exception as e:
            self.logger.error(f"Erro inesperado ao validar email {email}: {str(e)}")
            return False
    
    def extract_emails_from_pdf(self, pdf_path):
        """Extrai todos os emails válidos do texto do PDF"""
        try:
            with open(pdf_path, 'rb') as file:
                try:
                    reader = PyPDF2.PdfReader(file)
                    if len(reader.pages) == 0:
                        self.logger.warning(f"Arquivo PDF vazio ou corrompido: {pdf_path}")
                        return []
                except PyPDF2.PdfReadError:
                    self.logger.error(f"Arquivo PDF corrompido: {pdf_path}")
                    return []
                
                text = ""
                for page in reader.pages:
                    text += page.extract_text() or ""
                
                # Encontra e valida emails
                found_emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
                valid_emails = [email for email in found_emails if self.validate_email(email)]
                
                return valid_emails
        except Exception as e:
            self.logger.error(f"Erro ao extrair emails do PDF {pdf_path}: {str(e)}")
            return []
    
    def create_email_message(self, recipient, pdf_path, error=None):
        """Cria a mensagem de email com template apropriado"""
        msg = MIMEMultipart()
        msg['From'] = self.email
        msg['To'] = recipient
        
        filename = os.path.basename(pdf_path)
        
        if error:
            msg['Subject'] = f"Erro no processamento do arquivo: {filename}"
            body = self.error_template.format(
                nome_arquivo=filename,
                erro=error
            )
        else:
            msg['Subject'] = f"Envio automático do arquivo: {filename}"
            body = self.email_template.format(nome_arquivo=filename)
        
        msg.attach(MIMEText(body, 'plain'))
        
        if not error:
            with open(pdf_path, 'rb') as attachment:
                part = MIMEApplication(attachment.read(), Name=filename)
                part['Content-Disposition'] = f'attachment; filename="{filename}"'
                msg.attach(part)
        
        return msg
    
    def send_email(self, recipient, pdf_path, error=None):
        """Envia o email com anexo ou mensagem de erro"""
        try:
            msg = self.create_email_message(recipient, pdf_path, error)
            
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email, self.password)
                server.send_message(msg)
            
            self.logger.info(f"Email enviado com sucesso para: {recipient}")
            return True
        except Exception as e:
            self.logger.error(f"Erro ao enviar email para {recipient}: {str(e)}")
            return False
    
    def move_file(self, source, destination_folder):
        """Move o arquivo para a pasta especificada"""
        try:
            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)
            
            destination = os.path.join(destination_folder, os.path.basename(source))
            os.rename(source, destination)
            return destination
        except Exception as e:
            self.logger.error(f"Erro ao mover arquivo {source}: {str(e)}")
            return None
    
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith('.pdf'):
            self.logger.info(f"Novo arquivo PDF detectado: {event.src_path}")
            time.sleep(2)  # Espera para garantir que o arquivo esteja completamente escrito
            self.process_pdf(event.src_path)
    
    def process_pdf(self, pdf_path):
        """Processa o PDF: extrai emails, envia e move o arquivo"""
        try:
            emails = self.extract_emails_from_pdf(pdf_path)
            processed = 1
            sent = 0
            errors = 0
            
            if emails:
                success = True
                for email in emails:
                    if self.send_email(email, pdf_path):
                        sent += 1
                    else:
                        success = False
                        errors += 1
                
                if success:
                    self.move_file(pdf_path, self.sent_folder)
                    self.logger.info(f"Arquivo {pdf_path} processado com sucesso e movido para enviados")
                else:
                    self.move_file(pdf_path, self.error_folder)
                    self.logger.warning(f"Arquivo {pdf_path} movido para erros devido a falhas no envio")
                    errors += 1
            else:
                error_msg = "Nenhum email válido encontrado no documento"
                self.logger.warning(error_msg)
                self.send_email(self.email, pdf_path, error=error_msg)
                self.move_file(pdf_path, self.error_folder)
                errors += 1
            
            # Atualiza estatísticas
            self.update_stats(processed, sent, errors)
        except Exception as e:
            self.logger.error(f"Erro ao processar arquivo {pdf_path}: {str(e)}")
            self.update_stats(1, 0, 1)
            self.move_file(pdf_path, self.error_folder)

class EmailAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação de Envio de Emails")
        self.root.geometry("900x650")
        
        # Variáveis de controle
        self.monitoring = False
        self.observer = None
        
        # Layout principal
        self.create_widgets()
        self.setup_logging()
        
        # Carrega configurações padrão
        self.load_default_settings()
    
    def setup_logging(self):
        """Configura o sistema de logging"""
        self.logger = logging.getLogger('EmailAutomation')
        self.logger.setLevel(logging.INFO)
        
        # Cria um handler para exibir logs na interface
        self.log_handler = LogHandler(self.log_text)
        self.log_handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        self.log_handler.setFormatter(formatter)
        self.logger.addHandler(self.log_handler)
    
    def load_default_settings(self):
        """Carrega configurações padrão"""
        self.smtp_server_var = ctk.StringVar(value="")  # Seu servidor smtp
        self.smtp_port_var = ctk.StringVar(value="") # Sua respectiva porta 
        self.email_var = ctk.StringVar(value="")# Seu email com dominio proprio
        self.password_var = ctk.StringVar(value="")# Sua senha 
        self.folder_var = ctk.StringVar(value=os.path.join(os.path.expanduser("~"), "Documents", "PDF_Enviar"))
        self.sent_folder_var = ctk.StringVar(value=os.path.join(os.path.expanduser("~"), "Documents", "PDF_Enviados"))
        self.error_folder_var = ctk.StringVar(value=os.path.join(os.path.expanduser("~"), "Documents", "PDF_Erros"))
        
        self.email_template_var = ctk.StringVar(value="""Olá,

Segue em anexo o arquivo {nome_arquivo} conforme solicitado.

Atenciosamente,
Sistema Automático de Envio de Emails""")
        
        self.error_template_var = ctk.StringVar(value="""Olá,

Identificamos um problema ao processar o arquivo {nome_arquivo}:

{erro}

Por favor, verifique e tente novamente.

Atenciosamente,
Sistema Automático de Envio de Emails""")
        
        self.status_var = ctk.StringVar(value="Monitoramento: INATIVO")
        self.processed_var = ctk.StringVar(value="0")
        self.emails_sent_var = ctk.StringVar(value="0")
        self.errors_var = ctk.StringVar(value="0")
    
    def create_widgets(self):
        """Cria todos os widgets da interface"""
        # Frame principal
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Abas
        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Adiciona abas
        self.tabview.add("Configurações")
        self.tabview.add("Monitoramento")
        self.tabview.add("Logs")
        
        # Widgets da aba Configurações
        self.create_settings_tab()
        
        # Widgets da aba Monitoramento
        self.create_monitoring_tab()
        
        # Widgets da aba Logs
        self.create_logs_tab()
    
    def create_settings_tab(self):
        """Cria os widgets da aba de configurações"""
        tab = self.tabview.tab("Configurações")
        
        # Frame de rolagem
        scroll_frame = ctk.CTkScrollableFrame(tab)
        scroll_frame.pack(fill="both", expand=True)
        
        # Seção SMTP
        smtp_frame = ctk.CTkFrame(scroll_frame)
        smtp_frame.pack(fill="x", pady=5, padx=5)
        
        ctk.CTkLabel(smtp_frame, text="Configurações SMTP", font=("Arial", 14, "bold")).pack(pady=(5, 10))
        
        # Servidor SMTP
        ctk.CTkLabel(smtp_frame, text="Servidor SMTP:").pack(anchor="w")
        ctk.CTkEntry(smtp_frame, textvariable=self.smtp_server_var).pack(fill="x", pady=(0, 5))
        
        # Porta SMTP
        ctk.CTkLabel(smtp_frame, text="Porta SMTP:").pack(anchor="w")
        ctk.CTkEntry(smtp_frame, textvariable=self.smtp_port_var).pack(fill="x", pady=(0, 5))
        
        # Email
        ctk.CTkLabel(smtp_frame, text="Email:").pack(anchor="w")
        ctk.CTkEntry(smtp_frame, textvariable=self.email_var).pack(fill="x", pady=(0, 5))
        
        # Senha
        ctk.CTkLabel(smtp_frame, text="Senha:").pack(anchor="w")
        ctk.CTkEntry(smtp_frame, textvariable=self.password_var, show="*").pack(fill="x", pady=(0, 10))
        
        # Seção Pastas
        folder_frame = ctk.CTkFrame(scroll_frame)
        folder_frame.pack(fill="x", pady=5, padx=5)
        
        ctk.CTkLabel(folder_frame, text="Configurações de Pastas", font=("Arial", 14, "bold")).pack(pady=(5, 10))
        
        # Pasta Monitorada
        ctk.CTkLabel(folder_frame, text="Pasta Monitorada:").pack(anchor="w")
        ctk.CTkEntry(folder_frame, textvariable=self.folder_var).pack(fill="x", pady=(0, 5))
        ctk.CTkButton(folder_frame, text="Selecionar Pasta", command=lambda: self.select_folder(self.folder_var)).pack(pady=(0, 5))
        
        # Pasta Enviados
        ctk.CTkLabel(folder_frame, text="Pasta Enviados:").pack(anchor="w")
        ctk.CTkEntry(folder_frame, textvariable=self.sent_folder_var).pack(fill="x", pady=(0, 5))
        ctk.CTkButton(folder_frame, text="Selecionar Pasta", command=lambda: self.select_folder(self.sent_folder_var)).pack(pady=(0, 5))
        
        # Pasta Erros
        ctk.CTkLabel(folder_frame, text="Pasta Erros:").pack(anchor="w")
        ctk.CTkEntry(folder_frame, textvariable=self.error_folder_var).pack(fill="x", pady=(0, 5))
        ctk.CTkButton(folder_frame, text="Selecionar Pasta", command=lambda: self.select_folder(self.error_folder_var)).pack(pady=(0, 10))
        
        # Seção Templates
        template_frame = ctk.CTkFrame(scroll_frame)
        template_frame.pack(fill="x", pady=5, padx=5)
        
        ctk.CTkLabel(template_frame, text="Templates de Email", font=("Arial", 14, "bold")).pack(pady=(5, 10))
        
        # Template Email Normal
        ctk.CTkLabel(template_frame, text="Template de Email:").pack(anchor="w")
        self.email_template_text = ctk.CTkTextbox(template_frame, height=100)
        self.email_template_text.pack(fill="x", pady=(0, 5))
        self.email_template_text.insert("1.0", self.email_template_var.get())
        
        # Template Email Erro
        ctk.CTkLabel(template_frame, text="Template de Erro:").pack(anchor="w")
        self.error_template_text = ctk.CTkTextbox(template_frame, height=100)
        self.error_template_text.pack(fill="x", pady=(0, 5))
        self.error_template_text.insert("1.0", self.error_template_var.get())
        
        # Botão Salvar Configurações
        ctk.CTkButton(scroll_frame, text="Salvar Configurações", command=self.save_settings).pack(pady=10)
    
    def create_monitoring_tab(self):
        """Cria os widgets da aba de monitoramento"""
        tab = self.tabview.tab("Monitoramento")
        
        # Status do monitoramento
        ctk.CTkLabel(tab, textvariable=self.status_var, font=("Arial", 14)).pack(pady=10)
        
        # Botões de controle
        button_frame = ctk.CTkFrame(tab)
        button_frame.pack(pady=10)
        
        self.start_button = ctk.CTkButton(
            button_frame, 
            text="Iniciar Monitoramento", 
            command=self.start_monitoring,
            fg_color="green",
            hover_color="dark green"
        )
        self.start_button.pack(side="left", padx=5)
        
        self.stop_button = ctk.CTkButton(
            button_frame, 
            text="Parar Monitoramento", 
            command=self.stop_monitoring,
            fg_color="red",
            hover_color="dark red",
            state="disabled"
        )
        self.stop_button.pack(side="left", padx=5)
        
        # Botão de teste
        ctk.CTkButton(
            tab, 
            text="Testar Configurações", 
            command=self.test_settings
        ).pack(pady=5)
        
        # Estatísticas
        stats_frame = ctk.CTkFrame(tab)
        stats_frame.pack(fill="x", pady=10, padx=10)
        
        ctk.CTkLabel(stats_frame, text="Estatísticas", font=("Arial", 14, "bold")).pack(pady=(5, 10))
        
        stats_grid = ctk.CTkFrame(stats_frame)
        stats_grid.pack()
        
        # Arquivos processados
        ctk.CTkLabel(stats_grid, text="Arquivos Processados:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkLabel(stats_grid, textvariable=self.processed_var).grid(row=0, column=1, sticky="e", padx=5, pady=2)
        
        # Emails enviados
        ctk.CTkLabel(stats_grid, text="Emails Enviados:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkLabel(stats_grid, textvariable=self.emails_sent_var).grid(row=1, column=1, sticky="e", padx=5, pady=2)
        
        # Erros encontrados
        ctk.CTkLabel(stats_grid, text="Erros Encontrados:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkLabel(stats_grid, textvariable=self.errors_var).grid(row=2, column=1, sticky="e", padx=5, pady=2)
    
    def create_logs_tab(self):
        """Cria os widgets da aba de logs"""
        tab = self.tabview.tab("Logs")
        
        # Área de logs
        self.log_text = ctk.CTkTextbox(tab, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Botões de controle de logs
        button_frame = ctk.CTkFrame(tab)
        button_frame.pack(fill="x", pady=5)
        
        ctk.CTkButton(
            button_frame, 
            text="Limpar Logs", 
            command=self.clear_logs
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            button_frame, 
            text="Exportar Logs", 
            command=self.export_logs
        ).pack(side="left", padx=5)
    
    def select_folder(self, folder_var):
        """Abre diálogo para selecionar pasta"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            folder_var.set(folder_path)
    
    def save_settings(self):
        """Salva as configurações atuais"""
        # Atualiza os templates
        self.email_template_var.set(self.email_template_text.get("1.0", "end-1c"))
        self.error_template_var.set(self.error_template_text.get("1.0", "end-1c"))
        
        # Validação básica
        if not all([
            self.smtp_server_var.get(),
            self.smtp_port_var.get(),
            self.email_var.get(),
            self.folder_var.get()
        ]):
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios")
            return
        
        try:
            # Valida a porta SMTP
            port = int(self.smtp_port_var.get())
            if not (0 < port <= 65535):
                raise ValueError("Porta inválida")
            
            # Cria as pastas se não existirem
            for folder in [
                self.folder_var.get(),
                self.sent_folder_var.get(),
                self.error_folder_var.get()
            ]:
                if folder and not os.path.exists(folder):
                    os.makedirs(folder)
            
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
            self.logger.info("Configurações salvas")
        except ValueError as e:
            messagebox.showerror("Erro", f"Porta SMTP inválida: {str(e)}")
            self.logger.error(f"Porta SMTP inválida: {str(e)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar configurações: {str(e)}")
            self.logger.error(f"Erro ao salvar configurações: {str(e)}")
    
    def test_settings(self):
        """Testa as configurações atuais"""
        try:
            # Verifica se a porta é um número válido
            port = int(self.smtp_port_var.get())
            if not (0 < port <= 65535):
                raise ValueError("Porta inválida")
            
            # Testa conexão SMTP
            with smtplib.SMTP(self.smtp_server_var.get(), port) as server:
                server.starttls()
                server.login(self.email_var.get(), self.password_var.get())
            
            # Testa pastas
            for folder in [
                self.folder_var.get(),
                self.sent_folder_var.get(),
                self.error_folder_var.get()
            ]:
                if not os.path.exists(folder):
                    os.makedirs(folder)
            
            # Testa DNS
            resolver = dns.resolver.Resolver()
            resolver.timeout = 5
            resolver.lifetime = 5
            domain = self.email_var.get().split('@')[-1]
            mx_records = resolver.resolve(domain, 'MX')
            
            messagebox.showinfo("Sucesso", "Configurações testadas com sucesso!\n"
                                f"Domínio {domain} possui {len(mx_records)} registros MX.")
            self.logger.info("Configurações testadas com sucesso")
        except ValueError as e:
            messagebox.showerror("Erro", f"Porta SMTP inválida: {str(e)}")
            self.logger.error(f"Porta SMTP inválida: {str(e)}")
        except smtplib.SMTPException as e:
            error_msg = str(e).replace(self.password_var.get(), "******")
            messagebox.showerror("Erro SMTP", f"Falha na conexão SMTP: {error_msg}")
            self.logger.error(f"Erro SMTP: {error_msg}")
        except dns.resolver.NoAnswer:
            messagebox.showwarning("Aviso", "Configurações testadas, mas o domínio não possui registros MX")
            self.logger.warning("Domínio não possui registros MX")
        except dns.resolver.NXDOMAIN:
            messagebox.showerror("Erro", "Domínio do email não existe")
            self.logger.error("Domínio do email não existe")
        except dns.exception.DNSException as e:
            messagebox.showerror("Erro DNS", f"Falha ao verificar DNS: {str(e)}")
            self.logger.error(f"Erro DNS: {str(e)}")
        except Exception as e:
            error_msg = str(e).replace(self.password_var.get(), "******")
            messagebox.showerror("Erro", f"Falha ao testar configurações: {error_msg}")
            self.logger.error(f"Erro ao testar configurações: {error_msg}")
    
    def start_monitoring(self):
        """Inicia o monitoramento da pasta"""
        if self.monitoring:
            return
        
        # Atualiza os templates
        self.email_template_var.set(self.email_template_text.get("1.0", "end-1c"))
        self.error_template_var.set(self.error_template_text.get("1.0", "end-1c"))
        
        # Valida configurações
        if not all([
            self.smtp_server_var.get(),
            self.smtp_port_var.get(),
            self.email_var.get(),
            self.password_var.get(),
            self.folder_var.get()
        ]):
            messagebox.showerror("Erro", "Configure todos os campos obrigatórios antes de iniciar")
            return
        
        try:
            # Valida a porta SMTP
            port = int(self.smtp_port_var.get())
            if not (0 < port <= 65535):
                raise ValueError("Porta inválida")
            
            # Cria o observer
            event_handler = PDFHandler(
                smtp_server=self.smtp_server_var.get(),
                smtp_port=port,
                email=self.email_var.get(),
                password=self.password_var.get(),
                monitor_folder=self.folder_var.get(),
                sent_folder=self.sent_folder_var.get(),
                error_folder=self.error_folder_var.get(),
                email_template=self.email_template_var.get(),
                error_template=self.error_template_var.get(),
                update_stats_callback=self.update_stats,
                logger=self.logger
            )
            
            self.observer = Observer()
            self.observer.schedule(event_handler, self.folder_var.get(), recursive=False)
            
            # Inicia em uma thread separada
            monitoring_thread = threading.Thread(target=self.observer.start)
            monitoring_thread.daemon = True
            monitoring_thread.start()
            
            self.monitoring = True
            self.status_var.set("Monitoramento: ATIVO")
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            
            messagebox.showinfo("Sucesso", "Monitoramento iniciado com sucesso!")
            self.logger.info("Monitoramento iniciado")
        except ValueError as e:
            messagebox.showerror("Erro", f"Porta SMTP inválida: {str(e)}")
            self.logger.error(f"Porta SMTP inválida: {str(e)}")
        except Exception as e:
            error_msg = str(e).replace(self.password_var.get(), "******")
            messagebox.showerror("Erro", f"Falha ao iniciar monitoramento: {error_msg}")
            self.logger.error(f"Erro ao iniciar monitoramento: {error_msg}")
    
    def stop_monitoring(self):
        """Para o monitoramento da pasta"""
        if not self.monitoring or not self.observer:
            return
        
        try:
            self.observer.stop()
            self.observer.join(timeout=5)  # Timeout de 5 segundos
            
            if self.observer.is_alive():
                self.logger.warning("Observer não encerrou completamente")
            
            self.observer = None
            self.monitoring = False
            self.status_var.set("Monitoramento: INATIVO")
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            
            messagebox.showinfo("Sucesso", "Monitoramento parado com sucesso!")
            self.logger.info("Monitoramento parado")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao parar monitoramento: {str(e)}")
            self.logger.error(f"Erro ao parar monitoramento: {str(e)}")
    
    def update_stats(self, processed=0, sent=0, errors=0):
        """Atualiza as estatísticas na interface"""
        current_processed = int(self.processed_var.get())
        current_sent = int(self.emails_sent_var.get())
        current_errors = int(self.errors_var.get())
        
        self.processed_var.set(str(current_processed + processed))
        self.emails_sent_var.set(str(current_sent + sent))
        self.errors_var.set(str(current_errors + errors))
    
    def clear_logs(self):
        """Limpa os logs exibidos"""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
    
    def export_logs(self):
        """Exporta os logs para um arquivo"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("Arquivos de Log", "*.log"), ("Todos os arquivos", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, "w") as f:
                    f.write(self.log_text.get("1.0", "end-1c"))
                messagebox.showinfo("Sucesso", f"Logs exportados para: {file_path}")
                self.logger.info(f"Logs exportados para: {file_path}")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao exportar logs: {str(e)}")
                self.logger.error(f"Erro ao exportar logs: {str(e)}")

if __name__ == "__main__":
    # Cria a janela principal
    root = ctk.CTk()
    
    # Inicia o aplicativo
    app = EmailAutomationApp(root)
    
    # Loop principal
    root.mainloop()
