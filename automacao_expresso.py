# by Reinaldo Rodrigues ATI-PE
import os
import time
import pyautogui
import pyperclip
import pandas as pd
import smtplib
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# Configurações Visuais
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class SistemaEnvioATI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automação ATI-PE - Multi-Usuário")
        self.geometry("700x1050")
        self.arquivo_path = ""
        self.pasta_anexos = os.path.join(os.getcwd(), "anexos")
        self.anexos_selecionados = []
        self.config_file = "config.txt"
        self.modo_direto = False
        
        if not os.path.exists(self.pasta_anexos):
            os.makedirs(self.pasta_anexos)

        # --- UI: AUTENTICAÇÃO ---
        self.label_titulo = ctk.CTkLabel(self, text="Painel de Controle de Envio", font=("Arial", 22, "bold"))
        self.label_titulo.pack(pady=15)

        self.frame_auth = ctk.CTkFrame(self)
        self.frame_auth.pack(pady=10, padx=20, fill="x")

        self.input_user = ctk.CTkEntry(self.frame_auth, placeholder_text="E-mail ou Usuário Expresso", width=400)
        self.input_user.pack(pady=5)
        
        # CARREGAR UTILIZADOR GUARDADO
        self.carregar_config()

        self.input_pass = ctk.CTkEntry(self.frame_auth, placeholder_text="Senha (ou Senha de App Gmail)", show="*", width=400)
        self.input_pass.pack(pady=5)

        # --- UI: MODO DE ENVIO ---
        self.frame_modo = ctk.CTkFrame(self)
        self.frame_modo.pack(pady=10, padx=20, fill="x")
        
        self.modo_var = ctk.StringVar(value="EXPRESSO")
        self.combo_modo = ctk.CTkComboBox(self.frame_modo, values=["EXPRESSO", "GMAIL"], 
                                        variable=self.modo_var, width=200, command=self.atualizar_interface_modo)
        self.combo_modo.pack(pady=5)
        
        # Trava de domínio corrigida
        self.label_dominio = ctk.CTkLabel(self.frame_modo, 
                                        text="Expresso: qualquer @...pe.gov.br | Gmail: apenas @gmail.com", 
                                        font=("Arial", 10))
        self.label_dominio.pack(pady=5)

        # --- UI: ENVIO DIRETO ---
        self.frame_direto = ctk.CTkFrame(self)
        self.frame_direto.pack(pady=10, padx=20, fill="x")
        
        self.check_direto_var = ctk.BooleanVar(value=False)
        self.check_direto = ctk.CTkCheckBox(self.frame_direto, text="Envio Direto (sem planilha)", 
                                          variable=self.check_direto_var, command=self.toggle_envio_direto)
        self.check_direto.pack(pady=5)
        
        self.frame_destinatario = ctk.CTkFrame(self.frame_direto)
        
        self.label_destinatario = ctk.CTkLabel(self.frame_destinatario, text="Destinatário:")
        self.label_destinatario.grid(row=0, column=0, padx=5, pady=5)
        
        self.input_destinatario = ctk.CTkEntry(self.frame_destinatario, width=250, 
                                             placeholder_text="email@exemplo.com")
        self.input_destinatario.grid(row=0, column=1, padx=5, pady=5)
        
        self.label_nome_direto = ctk.CTkLabel(self.frame_destinatario, text="Nome:")
        self.label_nome_direto.grid(row=0, column=2, padx=5, pady=5)
        
        self.input_nome_direto = ctk.CTkEntry(self.frame_destinatario, width=150, 
                                            placeholder_text="Nome do Destinatário")
        self.input_nome_direto.grid(row=0, column=3, padx=5, pady=5)
        
        self.frame_destinatario.pack(pady=5)
        self.frame_destinatario.pack_forget()  # Inicialmente escondido

        # --- UI: PLANILHA ---
        self.frame_planilha = ctk.CTkFrame(self)
        self.frame_planilha.pack(pady=10, padx=20, fill="x")
        
        self.btn_planilha = ctk.CTkButton(self.frame_planilha, text="📂 Selecionar Planilha (Excel)", 
                                        command=self.escolher_arquivo)
        self.btn_planilha.pack(pady=5)

        # --- UI: ANEXOS (apenas para Gmail) ---
        self.frame_anexos = ctk.CTkFrame(self)
        self.frame_anexos.pack(pady=10, padx=20, fill="x")
        
        self.btn_anexos = ctk.CTkButton(self.frame_anexos, text="📎 Selecionar Anexos (Gmail apenas)", 
                                      command=self.selecionar_anexos)
        self.btn_anexos.pack(pady=5)
        
        self.lista_anexos = ctk.CTkTextbox(self.frame_anexos, width=500, height=80)
        self.lista_anexos.pack(pady=5)
        self.lista_anexos.insert("0.0", "Nenhum anexo selecionado")
        self.lista_anexos.configure(state="disabled")

        # --- UI: ASSUNTO E CORPO ---
        self.input_assunto = ctk.CTkEntry(self, placeholder_text="Assunto do E-mail", width=500)
        self.input_assunto.pack(pady=10)

        self.input_corpo = ctk.CTkTextbox(self, width=500, height=150)
        self.input_corpo.insert("0.0", "Prezados(as) {nome},\n\nSegue a documentação.")
        self.input_corpo.pack(pady=5)

        # --- UI: BOTÃO INICIAR ---
        self.btn_iniciar = ctk.CTkButton(self, text="🚀 INICIAR E GUARDAR UTILIZADOR", fg_color="green", 
                                        height=50, font=("Arial", 16, "bold"), command=self.iniciar_thread)
        self.btn_iniciar.pack(pady=20)
        
        # Status bar
        self.status_var = ctk.StringVar(value="Pronto")
        self.status_bar = ctk.CTkLabel(self, textvariable=self.status_var, font=("Arial", 10))
        self.status_bar.pack(pady=5)

        # Atualizar interface inicial
        self.atualizar_interface_modo()

    def carregar_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                self.input_user.insert(0, f.read().strip())

    def guardar_config(self, usuario):
        with open(self.config_file, "w") as f:
            f.write(usuario)

    def atualizar_interface_modo(self, *args):
        modo = self.modo_var.get()
        if modo == "GMAIL":
            self.btn_anexos.configure(state="normal")
        else:
            self.btn_anexos.configure(state="disabled")
            self.anexos_selecionados = []
            self.lista_anexos.configure(state="normal")
            self.lista_anexos.delete("0.0", "end")
            self.lista_anexos.insert("0.0", "Anexos disponíveis apenas para Gmail")
            self.lista_anexos.configure(state="disabled")

    def toggle_envio_direto(self):
        self.modo_direto = self.check_direto_var.get()
        if self.modo_direto:
            self.frame_destinatario.pack(pady=5)
            self.frame_planilha.pack_forget()
            self.btn_iniciar.configure(text="🚀 ENVIAR DIRETO E GUARDAR UTILIZADOR")
        else:
            self.frame_destinatario.pack_forget()
            self.frame_planilha.pack(pady=10, padx=20, fill="x")
            self.btn_iniciar.configure(text="🚀 INICIAR E GUARDAR UTILIZADOR")

    def escolher_arquivo(self):
        self.arquivo_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if self.arquivo_path:
            self.btn_planilha.configure(text=f"✅ {os.path.basename(self.arquivo_path)}", fg_color="#2b719e")

    def selecionar_anexos(self):
        if self.modo_var.get() == "GMAIL":
            arquivos = filedialog.askopenfilenames(title="Selecionar Anexos")
            if arquivos:
                self.anexos_selecionados = list(arquivos)
                self.lista_anexos.configure(state="normal")
                self.lista_anexos.delete("0.0", "end")
                for arquivo in self.anexos_selecionados:
                    self.lista_anexos.insert("end", f"• {os.path.basename(arquivo)}\n")
                self.lista_anexos.configure(state="disabled")

    def buscar_anexo_dinamico(self, nome, email):
        for arquivo in os.listdir(self.pasta_anexos):
            if nome.lower() in arquivo.lower() or email.lower() in arquivo.lower():
                return os.path.join(self.pasta_anexos, arquivo)
        return None

    def validar_dominio_email(self, email, modo):
        """Valida o domínio do email conforme o modo selecionado"""
        email_lower = email.lower().strip()
        
        if modo == "EXPRESSO":
            # Aceita qualquer email que termine com ".pe.gov.br"
            # Exemplos válidos:
            # - hugo.fgomes@ati.pe.gov.br
            # - usuario@sad.pe.gov.br
            # - nome@saude.pe.gov.br
            # - qualquer@qualquerorgao.pe.gov.br
            return email_lower.endswith(".pe.gov.br")
        
        elif modo == "GMAIL":
            # Aceita apenas emails que terminam com "@gmail.com"
            return email_lower.endswith("@gmail.com")
        
        return False

    def encontrar_coluna_email(self, df):
        """Tenta encontrar a coluna de email em diferentes formatos"""
        # Primeiro, verifica colunas padrão
        colunas_possiveis = ['EMAIL', 'E-MAIL', 'E_MAIL', 'E_MAIL', 'EMAIL_ADDRESS', 
                            'E-MAIL ADDRESS', 'CORREIO', 'MAIL', 'CONTATO']
        
        for col in df.columns:
            col_upper = str(col).upper().strip()
            if any(pos in col_upper for pos in colunas_possiveis):
                return col
        
        # Se não encontrou, tenta encontrar por padrão
        for col in df.columns:
            if '@' in str(df[col].iloc[0] if len(df) > 0 else ''):
                return col
        
        return None

    def encontrar_coluna_nome(self, df):
        """Tenta encontrar a coluna de nome em diferentes formatos"""
        colunas_possiveis = ['NOME', 'NAME', 'NOMBRE', 'CONTATO', 'PESSOA', 
                            'DESTINATARIO', 'DESTINATÁRIO', 'CLIENTE']
        
        for col in df.columns:
            col_upper = str(col).upper().strip()
            if any(pos in col_upper for pos in colunas_possiveis):
                return col
        
        return None

    def enviar_gmail(self, remetente, senha, destino, nome, assunto, mensagem, anexos=None):
        try:
            msg = MIMEMultipart()
            msg['From'] = remetente
            msg['To'] = destino
            msg['Subject'] = assunto
            msg.attach(MIMEText(mensagem.replace("{nome}", nome), 'plain'))
            
            # Adicionar anexos
            if anexos:
                for anexo_path in anexos:
                    if os.path.exists(anexo_path):
                        with open(anexo_path, "rb") as f:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", 
                                      f'attachment; filename="{os.path.basename(anexo_path)}"')
                        msg.attach(part)
            
            # Enviar email
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(remetente, senha)
            server.sendmail(remetente, destino, msg.as_string())
            server.quit()
            return True, "Sucesso"
        except Exception as e:
            return False, str(e)

    def processar_envio_direto(self):
        """Processa envio direto para um único destinatário"""
        user_logado = self.input_user.get()
        pass_logada = self.input_pass.get()
        modo = self.modo_var.get()
        
        if not user_logado or not pass_logada:
            return False, "Preencha usuário e senha!"
        
        email = self.input_destinatario.get().strip()
        nome = self.input_nome_direto.get().strip()
        assunto = self.input_assunto.get()
        mensagem_base = self.input_corpo.get("0.0", "end")
        
        if not email:
            return False, "Informe o email do destinatário!"
        
        # Validar domínio
        if not self.validar_dominio_email(email, modo):
            dominio_esperado = "qualquer .pe.gov.br" if modo == "EXPRESSO" else "@gmail.com"
            mensagem_erro = f"Email inválido para envio via {modo}!\n"
            mensagem_erro += f"Expresso espera: *@...pe.gov.br (ex: usuario@ati.pe.gov.br)\n" if modo == "EXPRESSO" else ""
            mensagem_erro += f"Gmail espera: *@gmail.com" if modo == "GMAIL" else ""
            return False, mensagem_erro
        
        msg_final = mensagem_base.replace("{nome}", nome if nome else "Prezado(a)")
        
        if modo == "EXPRESSO":
            # ENVIO VIA EXPRESSO (Selenium + PyAutoGUI) - NÃO ALTERAR
            try:
                driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
                driver.maximize_window()
                driver.get("https://expresso.pe.gov.br/login.php")
                driver.find_element(By.NAME, "user").send_keys(user_logado)
                driver.find_element(By.NAME, "passwd").send_keys(pass_logada + Keys.ENTER)
                time.sleep(8)
                driver.get("https://expresso.pe.gov.br/expressoMail1_2/index.php")
                time.sleep(5)
                
                # Clicar em Nova Mensagem
                driver.execute_script("document.querySelectorAll('a, b, span').forEach(e => { if(e.innerText.includes('Nova Mensagem')) e.click(); });")
                time.sleep(10)
                
                # Preencher dados
                largura, altura = pyautogui.size()
                pyautogui.click(largura * 0.35, altura * 0.30, clicks=2)
                pyautogui.write(email)
                pyautogui.click(largura * 0.35, altura * 0.34)
                pyautogui.write(assunto)
                pyperclip.copy(msg_final)
                pyautogui.click(largura * 0.5, altura * 0.7)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(2)
                pyautogui.click(largura * 0.12, altura * 0.26)
                time.sleep(8)
                
                driver.quit()
                return True, f"Email enviado com sucesso para {email} via Expresso!"
            except Exception as e:
                return False, f"Erro no Expresso: {str(e)}"
                
        else:  # GMAIL
            anexos = self.anexos_selecionados if self.anexos_selecionados else None
            sucesso, mensagem = self.enviar_gmail(user_logado, pass_logada, email, nome, assunto, msg_final, anexos)
            return sucesso, mensagem

    def processar_envios_planilha(self):
        """Processa envios a partir de uma planilha"""
        user_logado = self.input_user.get()
        pass_logada = self.input_pass.get()
        
        if not self.arquivo_path:
            return False, "Selecione uma planilha!"
        
        try:
            # Ler a planilha
            df = pd.read_excel(self.arquivo_path)
            
            # Exibir informações de debug
            print("Colunas encontradas na planilha:", list(df.columns))
            print("Primeiras linhas:")
            print(df.head())
            
            # Limpar e normalizar nomes das colunas
            df.columns = df.columns.astype(str).str.strip().str.upper()
            
            # Encontrar colunas de email e nome
            coluna_email = self.encontrar_coluna_email(df)
            coluna_nome = self.encontrar_coluna_nome(df)
            
            if not coluna_email:
                return False, f"Não foi possível encontrar a coluna de email na planilha.\nColunas disponíveis: {list(df.columns)}"
            
            if not coluna_nome:
                return False, f"Não foi possível encontrar a coluna de nome na planilha.\nColunas disponíveis: {list(df.columns)}"
            
            # Renomear colunas para nomes padrão
            df = df.rename(columns={coluna_email: 'EMAIL', coluna_nome: 'NOME'})
            
            # Garantir que as colunas existam
            df['EMAIL'] = df['EMAIL'].astype(str).str.strip()
            df['NOME'] = df['NOME'].astype(str).str.strip()
            
            # Adicionar coluna STATUS se não existir
            if 'STATUS' not in df.columns:
                df['STATUS'] = 'PENDENTE'
            else:
                df['STATUS'] = df['STATUS'].astype(str).str.strip().str.upper()
            
            assunto = self.input_assunto.get()
            mensagem_base = self.input_corpo.get("0.0", "end")
            modo = self.modo_var.get()
            
            if modo == "EXPRESSO":
                # INICIALIZAR EXPRESSO - NÃO ALTERAR
                driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
                driver.maximize_window()
                driver.get("https://expresso.pe.gov.br/login.php")
                driver.find_element(By.NAME, "user").send_keys(user_logado)
                driver.find_element(By.NAME, "passwd").send_keys(pass_logada + Keys.ENTER)
                time.sleep(8)
                driver.get("https://expresso.pe.gov.br/expressoMail1_2/index.php")
                time.sleep(5)

            enviados = 0
            total = len(df)
            ignorados = 0
            
            for i, linha in df.iterrows():
                email = str(linha['EMAIL']).strip()
                nome = str(linha['NOME'])
                
                # Pular se email estiver vazio
                if pd.isna(email) or email == 'nan' or not email:
                    continue
                
                # Verificar status
                if 'STATUS' in df.columns and str(linha['STATUS']) == 'ENVIADO':
                    continue
                
                # Validar domínio
                if not self.validar_dominio_email(email, modo):
                    ignorados += 1
                    df.at[i, 'STATUS'] = f'DOMÍNIO INVÁLIDO ({modo})'
                    self.status_var.set(f"Ignorado {i+1}/{total}: {email} - Domínio inválido")
                    continue
                
                # Buscar anexo dinâmico (apenas para Gmail)
                anexo = None
                if modo == "GMAIL":
                    anexo = self.buscar_anexo_dinamico(nome, email)
                    if self.anexos_selecionados:
                        anexo = self.anexos_selecionados  # Prioriza anexos selecionados
                
                msg_final = mensagem_base.replace("{nome}", nome)
                
                if modo == "EXPRESSO":
                    # ENVIO VIA EXPRESSO - NÃO ALTERAR
                    try:
                        driver.execute_script("document.querySelectorAll('a, b, span').forEach(e => { if(e.innerText.includes('Nova Mensagem')) e.click(); });")
                        time.sleep(10)
                        largura, altura = pyautogui.size()
                        pyautogui.click(largura * 0.35, altura * 0.30, clicks=2)
                        pyautogui.write(email)
                        pyautogui.click(largura * 0.35, altura * 0.34)
                        pyautogui.write(assunto)
                        pyperclip.copy(msg_final)
                        pyautogui.click(largura * 0.5, altura * 0.7)
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(2)
                        pyautogui.click(largura * 0.12, altura * 0.26)
                        time.sleep(8)
                        
                        df.at[i, 'STATUS'] = 'ENVIADO'
                        enviados += 1
                        self.status_var.set(f"Enviado {enviados}/{total}: {email}")
                        
                    except Exception as e:
                        df.at[i, 'STATUS'] = f'ERRO: {str(e)[:50]}'
                        
                else:  # GMAIL
                    anexos_envio = []
                    if anexo:
                        if isinstance(anexo, list):
                            anexos_envio = anexo
                        else:
                            anexos_envio = [anexo]
                    
                    sucesso, _ = self.enviar_gmail(user_logado, pass_logada, email, nome, assunto, msg_final, anexos_envio)
                    if sucesso:
                        df.at[i, 'STATUS'] = 'ENVIADO'
                        enviados += 1
                        self.status_var.set(f"Enviado {enviados}/{total}: {email}")
                    else:
                        df.at[i, 'STATUS'] = 'ERRO'
                
                # Salvar progresso periodicamente
                if i % 5 == 0:
                    df.to_excel(self.arquivo_path, index=False)

            # Salvar resultados finais
            df.to_excel(self.arquivo_path, index=False)
            
            if modo == "EXPRESSO":
                driver.quit()
                
            return True, f"Processamento concluído!\nEnviados: {enviados}\nIgnorados (domínio inválido): {ignorados}"
            
        except Exception as e:
            return False, f"Erro ao processar a planilha: {str(e)}\n\nCertifique-se de que:\n1. A planilha tem colunas 'EMAIL' e 'NOME'\n2. Os dados estão formatados corretamente\n3. O arquivo não está aberto em outro programa"

    def iniciar_thread(self):
        user = self.input_user.get()
        if not user or not self.input_pass.get():
            messagebox.showwarning("Atenção", "Preencha usuário e senha!")
            return
        
        if self.modo_direto and not self.input_destinatario.get():
            messagebox.showwarning("Atenção", "Informe o email do destinatário para envio direto!")
            return
        
        if not self.modo_direto and not self.arquivo_path:
            messagebox.showwarning("Atenção", "Selecione uma planilha!")
            return
        
        self.guardar_config(user)
        self.btn_iniciar.configure(state="disabled")
        self.status_var.set("Processando...")
        
        threading.Thread(target=self.executar_processamento, daemon=True).start()

    def executar_processamento(self):
        """Executa o processamento conforme o modo selecionado"""
        try:
            if self.modo_direto:
                sucesso, mensagem = self.processar_envio_direto()
            else:
                sucesso, mensagem = self.processar_envios_planilha()
            
            if sucesso:
                messagebox.showinfo("Sucesso", mensagem)
            else:
                messagebox.showerror("Erro", mensagem)
                
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        finally:
            self.btn_iniciar.configure(state="normal")
            self.status_var.set("Pronto")

if __name__ == "__main__":
    app = SistemaEnvioATI()
    app.mainloop()