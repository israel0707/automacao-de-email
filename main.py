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
from typing import List, Optional, Tuple, Callable

# Configuração do tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class LogHandler(logging.Handler):
    """Handler personalizado para exibir logs na interface gráfica."""
    
    def __init__(self, text_widget: ctk.CTkTextbox):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record: logging.LogRecord) -> None:
        """Adiciona a mensagem de log ao widget de texto."""
        msg = self.format(record)
        
        def append():
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", msg + "\n")
            self.text_widget.configure(state="disabled")
            self.text_widget.see("end")
            
        self.text_widget.after(0, append)

class PDFHandler(FileSystemEventHandler):
    """Classe responsável por monitorar e processar arquivos PDF."""
    
    def __init__(self, smtp_server: str, smtp_port: int, email: str, password: str, 
                 monitor_folder: str, sent_folder: str, error_folder: str, 
                 email_template: str, error_template: str, 
                 update_stats_callback: Callable[[int, int, int], None], 
                 logger: logging.Logger):
        """
        Inicializa o handler de PDF com as configurações necessárias.
        
        Args:
            smtp_server: Servidor SMTP para envio de emails
            smtp_port: Porta do servidor SMTP
            email: Email do remetente
            password: Senha do email
            monitor_folder: Pasta a ser monitorada
            sent_folder: Pasta para arquivos enviados com sucesso
            error_folder: Pasta para arquivos com erro
            email_template: Template para emails normais
            error_template: Template para emails de erro
            update_stats_callback: Função para atualizar estatísticas
            logger: Objeto de logging
        """
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

    def validate_email(self, email: str) -> bool:
        """Valida um endereço de email verificando sintaxe e registros MX."""
        # Verificação básica de sintaxe
        if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
            return False
            
        domain = email.split('@')[-1]
        
        try:
            # Verificação de registros MX
            try:
                mx_records = self.resolver.resolve(domain, 'MX')
                if not mx_records:
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

    def extract_emails_from_pdf(self, pdf_path: str) -> List[str]:
        """Extrai e valida endereços de email de um arquivo PDF."""
        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ""
                
                # Extrai texto de todas as páginas
                for page in reader.pages:
                    text += page.extract_text() or ""
                
                # Encontra e valida emails no texto
                found_emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
                valid_emails = [email for email in found_emails if self.validate_email(email)]
                
                return valid_emails
        except Exception as e:
            self.logger.error(f"Erro ao extrair emails do PDF {pdf_path}: {str(e)}")
            return []

    def create_email_message(self, recipient: str, pdf_path: str, error: Optional[str] = None) -> MIMEMultipart:
        """Cria a mensagem de email com ou sem anexo, dependendo do tipo."""
        msg = MIMEMultipart()
        msg['From'] = self.email
        msg['To'] = recipient
        filename = os.path.basename(pdf_path)
        
        if error:
            # Email de erro
            msg['Subject'] = f"Erro no processamento do arquivo: {filename}"
            body = self.error_template.format(nome_arquivo=filename, erro=error)
        else:
            # Email normal com anexo
            msg['Subject'] = f"Envio automático do arquivo: {filename}"
            body = self.email_template.format(nome_arquivo=filename)
            
        msg.attach(MIMEText(body, 'plain'))
        
        if not error:
            # Adiciona o PDF como anexo para emails normais
            with open(pdf_path, 'rb') as attachment:
                part = MIMEApplication(attachment.read(), Name=filename)
                part['Content-Disposition'] = f'attachment; filename="{filename}"'
                msg.attach(part)
                
        return msg

    def send_email(self, recipient: str, pdf_path: str, error: Optional[str] = None) -> bool:
        """Envia um email com o PDF anexado ou mensagem de erro."""
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

    def move_file(self, source: str, destination_folder: str) -> Optional[str]:
        """Move um arquivo para a pasta de destino especificada."""
        try:
            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)
                
            destination = os.path.join(destination_folder, os.path.basename(source))
            os.rename(source, destination)
            
            return destination
        except Exception as e:
            self.logger.error(f"Erro ao mover arquivo {source}: {str(e)}")
            return None

    def on_created(self, event) -> None:
        """Método chamado quando um novo arquivo é detectado na pasta monitorada."""
        if not event.is_directory and event.src_path.lower().endswith('.pdf'):
            self.logger.info(f"Novo arquivo PDF detectado: {event.src_path}")
            time.sleep(2)  # Espera para garantir que o arquivo esteja completamente escrito
            self.process_pdf(event.src_path)

    def process_pdf(self, pdf_path: str) -> None:
        """Processa um arquivo PDF, extrai emails e envia mensagens."""
        try:
            emails = self.extract_emails_from_pdf(pdf_path)
            processed = 1
            sent = 0
            errors = 0
            
            if emails:
                success = True
                
                # Envia email para cada destinatário válido
                for email in emails:
                    if self.send_email(email, pdf_path):
                        sent += 1
                    else:
                        success = False
                        errors += 1
                
                # Move o arquivo para a pasta apropriada
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
                
                # Envia email de notificação de erro
                self.send_email(self.email, pdf_path, error=error_msg)
                self.move_file(pdf_path, self.error_folder)
                errors += 1
                
            self.update_stats(processed, sent, errors)
        except Exception as e:
            self.logger.error(f"Erro ao processar arquivo {pdf_path}: {str(e)}")
            self.update_stats(1, 0, 1)
            self.move_file(pdf_path, self.error_folder)

class EmailAutomationApp:
    """Classe principal da aplicação com interface gráfica."""
    
    def __init__(self, root: ctk.CTk):
        """Inicializa a aplicação com a janela principal."""
        self.root = root
        self.root.title("Automação de Envio de Emails")
        self.root.geometry("900x650")
        
        self.monitoring = False
        self.observer = None
        
        self.load_default_settings()
        self.create_widgets()
        self.setup_logging()

    def setup_logging(self) -> None:
        """Configura o sistema de logging da aplicação."""
        self.logger = logging.getLogger('EmailAutomation')
        self.logger.setLevel(logging.INFO)
        
        self.log_handler = LogHandler(self.log_text)
        self.log_handler.setLevel(logging.INFO)
        
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        self.log_handler.setFormatter(formatter)
        
        self.logger.addHandler(self.log_handler)

    def load_default_settings(self) -> None:
        """Carrega as configurações padrão da aplicação."""
        # Configurações SMTP
        self.smtp_server_var = ctk.StringVar(value="")  # Ex: smtp.gmail.com
        self.smtp_port_var = ctk.StringVar(value="587")  # Porta padrão para TLS
        self.email_var = ctk.StringVar(value="")  # Seu email
        self.password_var = ctk.StringVar(value="")  # Sua senha
        
        # Pastas padrão
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        self.folder_var = ctk.StringVar(value=os.path.join(documents_path, "PDF_Enviar"))
        self.sent_folder_var = ctk.StringVar(value=os.path.join(documents_path, "PDF_Enviados"))
        self.error_folder_var = ctk.StringVar(value=os.path.join(documents_path, "PDF_Erros"))
        
        # Templates de email
        self.email_template_var = ctk.StringVar(
            value="""Olá,\n\nSegue em anexo o arquivo {nome_arquivo} conforme solicitado.\n\nAtenciosamente,\nSistema Automático de Envio de Emails""")
        
        self.error_template_var = ctk.StringVar(
            value="""Olá,\n\nIdentificamos um problema ao processar o arquivo {nome_arquivo}:\n\n{erro}\n\nPor favor, verifique e tente novamente.\n\nAtenciosamente,\nSistema Automático de Envio de Emails""")
        
        # Status e estatísticas
        self.status_var = ctk.StringVar(value="Monitoramento: INATIVO")
        self.processed_var = ctk.StringVar(value="0")
        self.emails_sent_var = ctk.StringVar(value="0")
        self.errors_var = ctk.StringVar(value="0")

    def create_widgets(self) -> None:
        """Cria todos os widgets da interface gráfica."""
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Cria abas principais
        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.tabview.add("Configurações")
        self.tabview.add("Monitoramento")
        self.tabview.add("Logs")
        
        self.create_settings_tab()
        self.create_monitoring_tab()
        self.create_logs_tab()

    def create_settings_tab(self) -> None:
        """Cria a aba de configurações da aplicação."""
        tab = self.tabview.tab("Configurações")
        scroll_frame = ctk.CTkScrollableFrame(tab)
        scroll_frame.pack(fill="both", expand=True)
        
        # Frame de configurações SMTP
        smtp_frame = ctk.CTkFrame(scroll_frame)
        smtp_frame.pack(fill="x", pady=5, padx=5)
        
        ctk.CTkLabel(smtp_frame, text="Configurações SMTP", font=("Arial", 14, "bold")).pack(pady=(5, 10))
        
        ctk.CTkLabel(smtp_frame, text="Servidor SMTP:").pack(anchor="w")
        ctk.CTkEntry(smtp_frame, textvariable=self.smtp_server_var).pack(fill="x", pady=(0, 5))
        
        ctk.CTkLabel(smtp_frame, text="Porta SMTP:").pack(anchor="w")
        ctk.CTkEntry(smtp_frame, textvariable=self.smtp_port_var).pack(fill="x", pady=(0, 5))
        
        ctk.CTkLabel(smtp_frame, text="Email:").pack(anchor="w")
        ctk.CTkEntry(smtp_frame, textvariable=self.email_var).pack(fill="x", pady=(0, 5))
        
        ctk.CTkLabel(smtp_frame, text="Senha:").pack(anchor="w")
        ctk.CTkEntry(smtp_frame, textvariable=self.password_var, show="*").pack(fill="x", pady=(0, 10))
        
        # Frame de configurações de pastas
        folder_frame = ctk.CTkFrame(scroll_frame)
        folder_frame.pack(fill="x", pady=5, padx=5)
        
        ctk.CTkLabel(folder_frame, text="Configurações de Pastas", font=("Arial", 14, "bold")).pack(pady=(5, 10))
        
        ctk.CTkLabel(folder_frame, text="Pasta Monitorada:").pack(anchor="w")
        ctk.CTkEntry(folder_frame, textvariable=self.folder_var).pack(fill="x", pady=(0, 5))
        ctk.CTkButton(folder_frame, text="Selecionar Pasta", 
                      command=lambda: self.select_folder(self.folder_var)).pack(pady=(0, 5))
        
        ctk.CTkLabel(folder_frame, text="Pasta Enviados:").pack(anchor="w")
        ctk.CTkEntry(folder_frame, textvariable=self.sent_folder_var).pack(fill="x", pady=(0, 5))
        ctk.CTkButton(folder_frame, text="Selecionar Pasta", 
                      command=lambda: self.select_folder(self.sent_folder_var)).pack(pady=(0, 5))
        
        ctk.CTkLabel(folder_frame, text="Pasta Erros:").pack(anchor="w")
        ctk.CTkEntry(folder_frame, textvariable=self.error_folder_var).pack(fill="x", pady=(0, 5))
        ctk.CTkButton(folder_frame, text="Selecionar Pasta", 
                      command=lambda: self.select_folder(self.error_folder_var)).pack(pady=(0, 10))
        
        # Frame de templates de email
        template_frame = ctk.CTkFrame(scroll_frame)
        template_frame.pack(fill="x", pady=5, padx=5)
        
        ctk.CTkLabel(template_frame, text="Templates de Email", font=("Arial", 14, "bold")).pack(pady=(5, 10))
        
        ctk.CTkLabel(template_frame, text="Template de Email:").pack(anchor="w")
        self.email_template_text = ctk.CTkTextbox(template_frame, height=100)
        self.email_template_text.pack(fill="x", pady=(0, 5))
        self.email_template_text.insert("1.0", self.email_template_var.get())
        
        ctk.CTkLabel(template_frame, text="Template de Erro:").pack(anchor="w")
        self.error_template_text = ctk.CTkTextbox(template_frame, height=100)
        self.error_template_text.pack(fill="x", pady=(0, 5))
        self.error_template_text.insert("1.0", self.error_template_var.get())
        
        # Botão para salvar configurações
        ctk.CTkButton(scroll_frame, text="Salvar Configurações", 
                      command=self.save_settings).pack(pady=10)

    def create_monitoring_tab(self) -> None:
        """Cria a aba de monitoramento e estatísticas."""
        tab = self.tabview.tab("Monitoramento")
        
        # Status do monitoramento
        ctk.CTkLabel(tab, textvariable=self.status_var, font=("Arial", 14)).pack(pady=10)
        
        # Botões de controle
        button_frame = ctk.CTkFrame(tab)
        button_frame.pack(pady=10)
        
        self.start_button = ctk.CTkButton(
            button_frame, text="Iniciar Monitoramento", 
            command=self.start_monitoring, fg_color="green", hover_color="dark green")
        self.start_button.pack(side="left", padx=5)
        
        self.stop_button = ctk.CTkButton(
            button_frame, text="Parar Monitoramento", 
            command=self.stop_monitoring, fg_color="red", hover_color="dark red", state="disabled")
        self.stop_button.pack(side="left", padx=5)
        
        ctk.CTkButton(tab, text="Testar Configurações", 
                      command=self.test_settings).pack(pady=5)
        
        # Frame de estatísticas
        stats_frame = ctk.CTkFrame(tab)
        stats_frame.pack(fill="x", pady=10, padx=10)
        
        ctk.CTkLabel(stats_frame, text="Estatísticas", font=("Arial", 14, "bold")).pack(pady=(5, 10))
        
        stats_grid = ctk.CTkFrame(stats_frame)
        stats_grid.pack()
        
        ctk.CTkLabel(stats_grid, text="Arquivos Processados:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkLabel(stats_grid, textvariable=self.processed_var).grid(row=0, column=1, sticky="e", padx=5, pady=2)
        
        ctk.CTkLabel(stats_grid, text="Emails Enviados:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkLabel(stats_grid, textvariable=self.emails_sent_var).grid(row=1, column=1, sticky="e", padx=5, pady=2)
        
        ctk.CTkLabel(stats_grid, text="Erros Encontrados:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkLabel(stats_grid, textvariable=self.errors_var).grid(row=2, column=1, sticky="e", padx=5, pady=2)

    def create_logs_tab(self) -> None:
        """Cria a aba de logs da aplicação."""
        tab = self.tabview.tab("Logs")
        
        # Área de texto para logs
        self.log_text = ctk.CTkTextbox(tab, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Botões de ação para logs
        button_frame = ctk.CTkFrame(tab)
        button_frame.pack(fill="x", pady=5)
        
        ctk.CTkButton(button_frame, text="Limpar Logs", 
                      command=self.clear_logs).pack(side="left", padx=5)
        
        ctk.CTkButton(button_frame, text="Exportar Logs", 
                      command=self.export_logs).pack(side="left", padx=5)

    def select_folder(self, folder_var: ctk.StringVar) -> None:
        """Abre um diálogo para selecionar uma pasta e atualiza a variável."""
        folder_path = filedialog.askdirectory()
        if folder_path:
            folder_var.set(folder_path)

    def save_settings(self) -> None:
        """Salva as configurações da aplicação."""
        # Atualiza os templates dos campos de texto
        self.email_template_var.set(self.email_template_text.get("1.0", "end-1c"))
        self.error_template_var.set(self.error_template_text.get("1.0", "end-1c"))
        
        # Verifica campos obrigatórios
        if not all([self.smtp_server_var.get(), self.smtp_port_var.get(), 
                   self.email_var.get(), self.folder_var.get()]):
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios")
            return
        
        try:
            # Cria as pastas se não existirem
            for folder in [self.folder_var.get(), self.sent_folder_var.get(), 
                          self.error_folder_var.get()]:
                if folder and not os.path.exists(folder):
                    os.makedirs(folder)
                    
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
            self.logger.info("Configurações salvas")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar configurações: {str(e)}")
            self.logger.error(f"Erro ao salvar configurações: {str(e)}")

    def test_settings(self) -> None:
        """Testa as configurações SMTP e DNS."""
        try:
            # Testa conexão SMTP
            with smtplib.SMTP(self.smtp_server_var.get(), int(self.smtp_port_var.get())) as server:
                server.starttls()
                server.login(self.email_var.get(), self.password_var.get())
            
            # Verifica pastas
            for folder in [self.folder_var.get(), self.sent_folder_var.get(), 
                          self.error_folder_var.get()]:
                if not os.path.exists(folder):
                    os.makedirs(folder)
            
            # Verifica registros MX do domínio
            resolver = dns.resolver.Resolver()
            resolver.timeout = 5
            resolver.lifetime = 5
            
            domain = self.email_var.get().split('@')[-1]
            mx_records = resolver.resolve(domain, 'MX')
            
            messagebox.showinfo("Sucesso", 
                f"Configurações testadas com sucesso!\nDomínio {domain} possui {len(mx_records)} registros MX.")
            self.logger.info("Configurações testadas com sucesso")
            
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
            messagebox.showerror("Erro", f"Falha ao testar configurações: {str(e)}")
            self.logger.error(f"Erro ao testar configurações: {str(e)}")

    def start_monitoring(self) -> None:
        """Inicia o monitoramento da pasta especificada."""
        if self.monitoring:
            return
            
        # Atualiza templates
        self.email_template_var.set(self.email_template_text.get("1.0", "end-1c"))
        self.error_template_var.set(self.error_template_text.get("1.0", "end-1c"))
        
        # Verifica campos obrigatórios
        if not all([self.smtp_server_var.get(), self.smtp_port_var.get(), 
                   self.email_var.get(), self.password_var.get(), self.folder_var.get()]):
            messagebox.showerror("Erro", "Configure todos os campos obrigatórios antes de iniciar")
            return
        
        try:
            # Cria o handler para monitorar a pasta
            event_handler = PDFHandler(
                smtp_server=self.smtp_server_var.get(),
                smtp_port=int(self.smtp_port_var.get()),
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
            
            # Configura e inicia o observer
            self.observer = Observer()
            self.observer.schedule(event_handler, self.folder_var.get(), recursive=False)
            
            # Inicia o monitoramento em uma thread separada
            monitoring_thread = threading.Thread(target=self.observer.start)
            monitoring_thread.daemon = True
            monitoring_thread.start()
            
            self.monitoring = True
            self.status_var.set("Monitoramento: ATIVO")
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            
            messagebox.showinfo("Sucesso", "Monitoramento iniciado com sucesso!")
            self.logger.info("Monitoramento iniciado")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao iniciar monitoramento: {str(e)}")
            self.logger.error(f"Erro ao iniciar monitoramento: {str(e)}")

    def stop_monitoring(self) -> None:
        """Para o monitoramento da pasta."""
        if not self.monitoring or not self.observer:
            return
            
        try:
            self.observer.stop()
            self.observer.join()
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

    def update_stats(self, processed: int = 0, sent: int = 0, errors: int = 0) -> None:
        """Atualiza as estatísticas exibidas na interface."""
        current_processed = int(self.processed_var.get())
        current_sent = int(self.emails_sent_var.get())
        current_errors = int(self.errors_var.get())
        
        self.processed_var.set(str(current_processed + processed))
        self.emails_sent_var.set(str(current_sent + sent))
        self.errors_var.set(str(current_errors + errors))

    def clear_logs(self) -> None:
        """Limpa todos os logs exibidos."""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def export_logs(self) -> None:
        """Exporta os logs para um arquivo."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("Arquivos de Log", "*.log"), ("Todos os arquivos", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(self.log_text.get("1.0", "end-1c"))
                    
                messagebox.showinfo("Sucesso", f"Logs exportados para: {file_path}")
                self.logger.info(f"Logs exportados para: {file_path}")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao exportar logs: {str(e)}")
                self.logger.error(f"Erro ao exportar logs: {str(e)}")

if __name__ == "__main__":
    # Inicializa a aplicação
    root = ctk.CTk()
    app = EmailAutomationApp(root)
    root.mainloop()
