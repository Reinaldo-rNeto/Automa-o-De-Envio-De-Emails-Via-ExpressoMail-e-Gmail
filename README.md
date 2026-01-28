<<<<<<< HEAD
📋 SEÇÃO 1 - VISÃO GERAL DO SISTEMA
1.1 OBJETIVO DO SISTEMA
O Sistema de Automação de Envio de E-mails ATI-PE foi desenvolvido para automatizar completamente o processo de envio de comunicações institucionais. Ele permite o envio massivo de e-mails através de duas plataformas distintas:

Expresso - Sistema de correio eletrônico oficial do Governo do Estado de Pernambuco

Gmail - Serviço de e-mail do Google para comunicações externas

1.2 CONTEXTO DE USO
Órgão: ATI-PE (Agência de Tecnologia da Informação de Pernambuco)

Setor: Administração Pública Estadual

Finalidade: Comunicação oficial, envio de documentos, circulares, convocações

Escopo: Envio para servidores públicos, órgãos parceiros, fornecedores

1.3 CARACTERÍSTICAS TÉCNICAS
Linguagem de Programação: Python 3.11.5

Interface Gráfica: CustomTkinter (Tkinter moderno)

Automação Web: Selenium WebDriver 4.15.0

Automação de Interface: PyAutoGUI 0.9.54

Manipulação de Dados: Pandas 2.1.3

Envio de E-mails: SMTP + MIME

Formato de Entrada: Microsoft Excel (.xlsx, .xls)

Sistema Operacional: Windows 10/11 (64-bit)

⚙️ SEÇÃO 2 - PRÉ-REQUISITOS E INSTALAÇÃO
2.1 REQUISITOS DO SISTEMA OPERACIONAL
Windows 10 Professional/Enterprise (versão 21H2 ou superior)

Ou Windows 11 (todas versões)

Processador: Intel Core i3 (8ª geração) ou equivalente AMD

Memória RAM: 8 GB mínimo (16 GB recomendado)

Espaço em Disco: 2 GB livres para instalação

Resolução de Tela: 1920x1080 pixels (Full HD) obrigatório para automação Expresso

Permissões de Administrador: Necessário para instalação inicial

2.2 SOFTWARE NECESSÁRIO PRÉ-INSTALADO
Google Chrome (versão 119.0.6045.200 ou superior)

Download oficial: https://www.google.com/chrome/

Instalação: Executar como administrador

Microsoft Excel (2016 ou superior) ou LibreOffice 7.6

Para edição das planilhas de destinatários

Microsoft Visual C++ Redistributable (2015-2022)

Download: https://aka.ms/vs/17/release/vc_redist.x64.exe

2.3 INSTALAÇÃO DO PYTHON
PASSO 1 - DOWNLOAD DO PYTHON
Acesse: https://www.python.org/downloads/

Clique em "Download Python 3.11.5"

Salve o arquivo python-3.11.5-amd64.exe no seu computador

PASSO 2 - INSTALAÇÃO COMPLETA
Execute o instalador como Administrador

Marque TODAS as opções:

[✓] Install launcher for all users (recommended)

[✓] Add Python to PATH (IMPORTANTE!)

Clique em "Customize installation"

Na próxima tela, marque TODAS as opções:

[✓] Documentation

[✓] pip

[✓] tcl/tk and IDLE

[✓] Python test suite

[✓] py launcher

[✓] for all users (requires elevation)

Clique em "Next"

Configure as opções avançadas:

[✓] Install for all users

[✓] Associate files with Python

[✓] Create shortcuts for installed applications

[✓] Add Python to environment variables

[✓] Precompile standard library

[✓] Download debugging symbols

[✓] Download debug binaries (requires VS 2017 or later)

Altere o diretório de instalação para: C:\Python311\

Clique em "Install" e aguarde a conclusão

PASSO 3 - VERIFICAÇÃO DA INSTALAÇÃO
Abra o Prompt de Comando como Administrador e execute:

cmd
python --version
Deve retornar: Python 3.11.5

cmd
pip --version
Deve retornar: pip 23.3.1 from C:\Python311\Lib\site-packages\pip

2.4 CONFIGURAÇÃO DO AMBIENTE PYTHON
Configurar variáveis de ambiente manualmente (se necessário):
Pressione Windows + X → "Sistema"

Clique em "Configurações avançadas do sistema"

Clique em "Variáveis de ambiente"

Em "Variáveis do sistema", edite "Path"

Adicione estas duas entradas:

C:\Python311\

C:\Python311\Scripts\

Clique em "OK" em todas as janelas

Verificar instalação completa:
cmd
python -c "import sys; print('Python', sys.version)"
python -c "import tkinter; print('Tkinter', tkinter.TkVersion)"
📦 SEÇÃO 3 - INSTALAÇÃO DAS BIBLIOTECAS PYTHON
3.1 LISTA COMPLETA DE DEPENDÊNCIAS
BIBLIOTECAS PRINCIPAIS (OBRIGATÓRIAS)
Biblioteca	Versão	Função Principal
pandas	2.1.3	Leitura e manipulação de planilhas Excel
pyautogui	0.9.54	Automação da interface gráfica do Expresso
customtkinter	5.2.0	Interface gráfica moderna do sistema
selenium	4.15.0	Automação do navegador Chrome
webdriver-manager	3.8.6	Gerenciamento automático do ChromeDriver
openpyxl	3.1.2	Suporte para arquivos Excel .xlsx
pyperclip	1.8.2	Manipulação da área de transferência
BIBLIOTECAS DE SUPORTE (INSTALADAS AUTOMATICAMENTE)
Biblioteca	Versão	Dependência de
beautifulsoup4	4.12.2	selenium
soupsieve	2.5	beautifulsoup4
certifi	2023.11.17	selenium
charset-normalizer	3.3.2	requests
click	8.1.7	customtkinter
colorama	0.4.6	selenium
et-xmlfile	1.1.0	openpyxl
idna	3.4	requests
numpy	1.26.1	pandas
packaging	23.2	selenium
Pillow	10.1.0	customtkinter
python-dateutil	2.8.2	pandas
pytz	2023.3.post1	pandas
requests	2.31.0	selenium
six	1.16.0	python-dateutil
tzdata	2023.3	pandas
urllib3	2.1.0	requests
3.2 COMANDOS DE INSTALAÇÃO PASSO A PASSO
PASSO 1 - ABRIR TERMINAL COM PRIVILÉGIOS
Pressione Windows + X

Selecione "Windows PowerShell (Administrador)"

Confirme a solicitação de controle de conta de usuário

PASSO 2 - ATUALIZAR O PIP
powershell
python -m pip install --upgrade pip
Saída esperada:

text
Successfully installed pip-23.3.1
PASSO 3 - CRIAR E ATIVAR AMBIENTE VIRTUAL (OPCIONAL MAS RECOMENDADO)
powershell
# Criar ambiente virtual
python -m venv venv_ati

# Ativar ambiente virtual
.\venv_ati\Scripts\Activate.ps1
Nota: O prompt mudará para (venv_ati) PS C:\>

PASSO 4 - INSTALAÇÃO COMPLETA DAS BIBLIOTECAS
Método A: Instalação individual (recomendado para diagnóstico)
powershell
# 1. Instalar pandas para manipulação de Excel
pip install pandas==2.1.3

# 2. Instalar pyautogui para automação de interface
pip install pyautogui==0.9.54

# 3. Instalar customtkinter para interface gráfica
pip install customtkinter==5.2.0

# 4. Instalar selenium para automação web
pip install selenium==4.15.0

# 5. Instalar webdriver-manager para gerenciar ChromeDriver
pip install webdriver-manager==3.8.6

# 6. Instalar openpyxl para suporte a Excel
pip install openpyxl==3.1.2

# 7. Instalar pyperclip para manipulação de clipboard
pip install pyperclip==1.8.2
Método B: Instalação via arquivo requirements.txt
Crie um arquivo requirements.txt com:

txt
pandas==2.1.3
pyautogui==0.9.54
customtkinter==5.2.0
selenium==4.15.0
webdriver-manager==3.8.6
openpyxl==3.1.2
pyperclip==1.8.2
Execute:

powershell
pip install -r requirements.txt
Método C: Instalação com todas as dependências
powershell
pip install pandas pyautogui customtkinter selenium webdriver-manager openpyxl pyperclip
3.3 INSTALAÇÃO INDIVIDUAL DE CADA BIBLIOTECA
3.3.1 INSTALAÇÃO DO PANDAS
powershell
pip install pandas==2.1.3
Verificação:

powershell
python -c "import pandas as pd; print(f'Pandas versão: {pd.__version__}')"
Saída esperada: Pandas versão: 2.1.3

3.3.2 INSTALAÇÃO DO PYAUTOGUI
powershell
pip install pyautogui==0.9.54
Verificação:

powershell
python -c "import pyautogui; print(f'PyAutoGUI versão: {pyautogui.__version__}')"
Nota: O PyAutoGUI pode exigir permissões especiais no Windows. Se encontrar erros, execute:

powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
3.3.3 INSTALAÇÃO DO CUSTOMTKINTER
powershell
pip install customtkinter==5.2.0
Verificação:

powershell
python -c "import customtkinter as ctk; print(f'CustomTkinter versão: {ctk.__version__}')"
3.3.4 INSTALAÇÃO DO SELENIUM
powershell
pip install selenium==4.15.0
Verificação:

powershell
python -c "from selenium import __version__; print(f'Selenium versão: {__version__}')"
3.3.5 INSTALAÇÃO DO WEBDRIVER-MANAGER
powershell
pip install webdriver-manager==3.8.6
Verificação:

powershell
python -c "from webdriver_manager.chrome import ChromeDriverManager; print('WebDriver Manager instalado')"
3.3.6 INSTALAÇÃO DO OPENPYXL
powershell
pip install openpyxl==3.1.2
Verificação:

powershell
python -c "import openpyxl; print(f'Openpyxl versão: {openpyxl.__version__}')"
3.3.7 INSTALAÇÃO DO PYPERCLIP
powershell
pip install pyperclip==1.8.2
Verificação:

powershell
python -c "import pyperclip; pyperclip.copy('teste'); print('Pyperclip funcionando')"
3.4 VERIFICAÇÃO DAS INSTALAÇÕES
Script completo de verificação:
python
# save as verificar_instalacoes.py
import sys
import pkg_resources

bibliotecas = [
    'pandas',
    'pyautogui', 
    'customtkinter',
    'selenium',
    'webdriver-manager',
    'openpyxl',
    'pyperclip'
]

print("=== VERIFICAÇÃO DE BIBLIOTECAS INSTALADAS ===")
print(f"Python versão: {sys.version}\n")

for lib in bibliotecas:
    try:
        versao = pkg_resources.get_distribution(lib).version
        print(f"✓ {lib:20} versão: {versao}")
    except pkg_resources.DistributionNotFound:
        print(f"✗ {lib:20} NÃO INSTALADA")

print("\n=== DEPENDÊNCIAS SECUNDÁRIAS ===")
deps_secundarias = ['numpy', 'Pillow', 'requests', 'beautifulsoup4']
for dep in deps_secundarias:
    try:
        versao = pkg_resources.get_distribution(dep).version
        print(f"✓ {dep:20} versão: {versao}")
    except:
        print(f"✗ {dep:20} NÃO INSTALADA")
Execute:

powershell
python verificar_instalacoes.py
🔧 SEÇÃO 4 - CONFIGURAÇÃO DO AMBIENTE
4.1 CONFIGURAÇÃO DO GMAIL
PASSO 1 - ATIVAR ACESSO DE APPS MENOS SEGUROS (ALTERNATIVA ANTIGA)
Nota: Esta opção está sendo descontinuada pelo Google.

PASSO 2 - CRIAR SENHA DE APLICATIVO (MÉTODO ATUAL)
Acesse https://myaccount.google.com/

Faça login com sua conta Gmail

No menu esquerdo, clique em "Segurança"

Role até "Como você faz login no Google"

Clique em "Senhas de app" (pode precisar verificar identidade)

Selecione "App" e escolha "Outro (Nome personalizado)"

Digite: "Sistema ATI-PE"

Clique em "Gerar"

COPIE A SENHA exibida (16 caracteres)

Use esta senha no sistema (NÃO sua senha normal)

PASSO 3 - CONFIGURAR AUTENTICAÇÃO DE 2 FATORES (OBRIGATÓRIO)
Na mesma página "Segurança", clique em "Verificação em duas etapas"

Clique em "Começar"

Siga o passo a passo para configurar

Após configurar, volte para criar senha de aplicativo

4.2 CONFIGURAÇÃO DO EXPRESSO
REQUISITOS PARA AUTOMAÇÃO:
Conta ativa no Expresso PE

Solicitar ao administrador do sistema Expresso

Usuário no formato: usuario (sem @ati.pe.gov.br)

Permissões de envio

Verificar limites de envio diário

Confirmar que a conta não está bloqueada

Configuração do navegador

Google Chrome instalado e atualizado

Permitir pop-ups para expresso.pe.gov.br

4.3 CONFIGURAÇÃO DO CHROMEDRIVER
O sistema usa webdriver-manager que gerencia automaticamente, mas se precisar manual:
Verificar versão do Chrome:

Abra Chrome

Clique nos 3 pontos → "Ajuda" → "Sobre o Google Chrome"

Anote o número da versão (ex: 119.0.6045.200)

Download manual (se necessário):

Acesse: https://chromedriver.chromium.org/

Baixe a versão correspondente ao seu Chrome

Extraia o arquivo chromedriver.exe

Coloque em C:\Windows\System32\ ou no mesmo diretório do sistema

Verificar funcionamento:

powershell
python -c "from selenium import webdriver; from webdriver_manager.chrome import ChromeDriverManager; driver = webdriver.Chrome(ChromeDriverManager().install()); print('ChromeDriver funcionando'); driver.quit()"
🖥️ SEÇÃO 5 - ESTRUTURA DO PROJETO
5.1 ARQUIVOS DO SISTEMA
text
SISTEMA_ATI_PE/                      ← Pasta principal do projeto
│
├── sistema_envio_ati.py             ← Código fonte principal
├── requirements.txt                  ← Lista de dependências
├── README_COMPLETO.md               ← Esta documentação
├── MANUAL_OPERADOR.pdf              ← Manual simplificado
├── config.txt                       ← Configurações do usuário
├── setup.py                         ← Script para criar executável
│
├── anexos/                          ← Pasta para arquivos anexos
│   ├── modelo_contrato.pdf          ← Exemplo de anexo
│   └── README_ANEXOS.txt            ← Instruções para anexos
│
├── planilhas/                       ← Pasta para planilhas de dados
│   ├── contatos.xlsx                ← Modelo de planilha
│   ├── enviados/                    ← Backup dos enviados
│   └── modelos/                     ← Modelos para diferentes usos
│
├── logs/                            ← Pasta para registros do sistema
│   ├── envios_20240115.log          ← Log diário de envios
│   └── erros/                       ← Registros de erros
│
└── backups/                         ← Backup de configurações
    ├── config_backup_20240115.txt   ← Backup automático
    └── manual/                      ← Backups manuais
5.2 ORGANIZAÇÃO DE PASTAS
CRIAR ESTRUTURA COMPLETA MANUALMENTE:
powershell
# Criar pasta principal
mkdir "C:\Sistema_ATI_PE"

# Navegar para a pasta
cd "C:\Sistema_ATI_PE"

# Criar todas as subpastas
mkdir anexos, planilhas, planilhas\enviados, planilhas\modelos, logs, logs\erros, backups, backups\manual

# Criar arquivos base
New-Item -ItemType File -Name "config.txt"
New-Item -ItemType File -Name "requirements.txt"
CONTEÚDO DO ARQUIVO CONFIG.TXT:
text
# Arquivo de configuração do Sistema ATI-PE
# Não editar manualmente - O sistema atualiza automaticamente

usuario_padrao=reinaldo.rodrigues@ati.pe.gov.br
modo_ultimo_uso=EXPRESSO
resolucao_tela=1920x1080
backup_automatico=sim
intervalo_log=5
5.3 ARQUIVOS DE CONFIGURAÇÃO
5.3.1 ARQUIVO REQUIREMENTS.TXT COMPLETO:
txt
# Sistema ATI-PE - Dependências Python
# Gerado em: 2024-01-15
# Python: 3.11.5

# Bibliotecas principais
pandas==2.1.3
pyautogui==0.9.54
customtkinter==5.2.0
selenium==4.15.0
webdriver-manager==3.8.6
openpyxl==3.1.2
pyperclip==1.8.2

# Dependências de suporte
numpy==1.26.1
Pillow==10.1.0
requests==2.31.0
beautifulsoup4==4.12.2
python-dateutil==2.8.2
pytz==2023.3.post1
tzdata==2023.3
et-xmlfile==1.1.0
charset-normalizer==3.3.2
idna==3.4
urllib3==2.1.0
certifi==2023.11.17
colorama==0.4.6
packaging==23.2
six==1.16.0
soupsieve==2.5
click==8.1.7
5.3.2 ARQUIVO SETUP.PY PARA COMPILAÇÃO:
python
# setup.py - Script para criar executável
import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": [
        "os", "sys", "time", "pyautogui", "pyperclip", 
        "pandas", "smtplib", "threading", "customtkinter",
        "tkinter", "email", "selenium", "webdriver_manager",
        "openpyxl", "json", "pathlib", "datetime"
    ],
    "includes": [
        "tkinter", "customtkinter", "pandas._libs.tslibs.timedeltas"
    ],
    "include_files": [
        ("anexos", "anexos"),
        ("planilhas", "planilhas"),
        ("logs", "logs"),
        ("config.txt", "config.txt"),
        ("README_COMPLETO.md", "README_COMPLETO.md")
    ],
    "excludes": ["unittest", "pydoc", "test"],
    "optimize": 2
}

base = "Win32GUI" if sys.platform == "win32" else None

executables = [
    Executable(
        "sistema_envio_ati.py",
        base=base,
        target_name="SistemaATI.exe",
        icon="icon.ico"  # Se tiver ícone
    )
]

setup(
    name="Sistema Envio ATI-PE",
    version="2.0.0",
    description="Sistema de automação de envio de e-mails ATI-PE",
    author="Reinaldo Rodrigues - ATI-PE",
    options={"build_exe": build_exe_options},
    executables=executables
)
🚀 SEÇÃO 6 - EXECUÇÃO DO SISTEMA
6.1 PRIMEIRA EXECUÇÃO
PASSO 1 - PREPARAR AMBIENTE:
powershell
# Navegar para pasta do projeto
cd "C:\Sistema_ATI_PE"

# Verificar se todas as bibliotecas estão instaladas
python -m pip list | Select-String "pandas|pyautogui|customtkinter|selenium"
PASSO 2 - EXECUTAR O SISTEMA:
powershell
python sistema_envio_ati.py
PASSO 3 - PRIMEIRA CONFIGURAÇÃO NA INTERFACE:
Tela inicial: Sistema de Automação ATI-PE

Campo usuário: Digite seu e-mail institucional

Campo senha: Digite sua senha (será salva localmente)

Selecionar modo: EXPRESSO ou GMAIL

Testar conexão: Botão "Testar Conexão" (opcional)

6.2 MODOS DE OPERAÇÃO
6.2.1 MODO EXPRESSO (.PE.GOV.BR)
Características:

Exclusivo para e-mails terminados em .pe.gov.br

Usa automação via navegador Chrome

Não suporta anexos via automação

Requer resolução 1920x1080

Tempo estimado por e-mail: 45-60 segundos

Fluxo do Modo Expresso:

Login automático no expresso.pe.gov.br

Navegação para caixa de mensagens

Clicar em "Nova Mensagem"

Preenchimento automático dos campos:

Para: e-mail do destinatário

Assunto: conforme configurado

Mensagem: corpo personalizado

Envio e confirmação

6.2.2 MODO GMAIL (@GMAIL.COM)
Características:

Exclusivo para e-mails terminados em @gmail.com

Usa protocolo SMTP

Suporta múltiplos anexos

Velocidade: 2-5 segundos por e-mail

Requer senha de aplicativo

Fluxo do Modo Gmail:

Conexão SMTP com smtp.gmail.com:587

Autenticação com credenciais

Construção da mensagem MIME

Anexo de arquivos (se configurado)

Envio via protocolo SMTP

Confirmação de entrega

6.3 FLUXO COMPLETO DE TRABALHO
CENÁRIO 1: ENVIO EM MASSA (PLANILHA)
text
[PREPARAÇÃO]
1. Criar/Editar planilha Excel com destinatários
2. Salvar como .xlsx na pasta "planilhas"
3. Abrir sistema ATI-PE

[CONFIGURAÇÃO]
4. Selecionar modo (Expresso ou Gmail)
5. Clicar em "Selecionar Planilha"
6. Escolher arquivo Excel
7. Preencher assunto da mensagem
8. Editar corpo da mensagem
9. (Gmail apenas) Selecionar anexos

[EXECUÇÃO]
10. Clicar em "INICIAR E GUARDAR UTILIZADOR"
11. Aguardar processamento (barra de progresso)
12. Verificar log de execução
13. Conferir planilha atualizada (status ENVIADO)

[PÓS-PROCESSAMENTO]
14. Backup automático na pasta "planilhas/enviados"
15. Log salvo na pasta "logs"
16. Relatório de execução exibido na tela
CENÁRIO 2: ENVIO DIRETO (INDIVIDUAL)
text
[PREPARAÇÃO]
1. Abrir sistema ATI-PE
2. Marcar checkbox "Envio Direto (sem planilha)"

[CONFIGURAÇÃO]
3. Digitar e-mail do destinatário
4. Digitar nome do destinatário
5. Preencher assunto
6. Editar corpo da mensagem
7. (Gmail apenas) Selecionar anexos

[EXECUÇÃO]
8. Clicar em "ENVIAR DIRETO"
9. Aguardar confirmação de envio
10. Verificar mensagem de sucesso/erro
📊 SEÇÃO 7 - PREPARAÇÃO DE DADOS
7.1 ESTRUTURA DA PLANILHA EXCEL
FORMATO OBRIGATÓRIO:
Coluna A	Coluna B	Coluna C
NOME	EMAIL	STATUS
TIPOS DE DADOS:
NOME: Texto (string) - Nome completo do destinatário

EMAIL: Texto (string) - Endereço de e-mail válido

STATUS: Texto (string) - Valores: PENDENTE, ENVIADO, ERRO

REGRAS DE FORMATAÇÃO:
Primeira linha: Deve conter os cabeçalhos exatos

Ordem das colunas: Pode variar, mas os nomes devem ser exatos

Formatação de células: Todas como "Texto" ou "Geral"

Encoding: UTF-8 (salvar normalmente no Excel)

7.2 FORMATAÇÃO CORRETA DOS DADOS
7.2.1 CAMPO NOME:
Correto: "João da Silva", "Maria Santos Oliveira"

Incorreto: "joao silva", "MARIA SANTOS", "Sr. João"

7.2.2 CAMPO EMAIL:
Para Expresso:

Correto: nome.sobrenome@ati.pe.gov.br, usuario@sad.pe.gov.br

Incorreto: nome@ati.com.br, usuario@pe.gov.br

Para Gmail:

Correto: email@gmail.com, nome.sobrenome@gmail.com

Incorreto: email@googlemail.com, nome@empresa.com.br

7.2.3 CAMPO STATUS:
Valor inicial: PENDENTE (tudo maiúsculo)

Após envio: ENVIADO (sistema atualiza automaticamente)

Em caso de erro: ERRO: [descrição]

7.3 EXEMPLOS PRÁTICOS
EXEMPLO 1: PLANILHA PARA EXPRESSO
csv
NOME,EMAIL,STATUS
João Carlos Pereira,joao.pereira@ati.pe.gov.br,PENDENTE
Maria Aparecida Silva,maria.silva@sad.pe.gov.br,PENDENTE
Carlos Eduardo Santos,carlos.santos@saude.pe.gov.br,PENDENTE
Ana Paula Oliveira,ana.oliveira@educacao.pe.gov.br,PENDENTE
EXEMPLO 2: PLANILHA PARA GMAIL
csv
NOME,EMAIL,STATUS
Pedro Almeida,pedro.almeida@gmail.com,PENDENTE
Fernanda Costa,fernanda.costa@gmail.com,PENDENTE
Ricardo Nunes,ricardo.nunes@gmail.com,PENDENTE
Juliana Mendes,juliana.mendes@gmail.com,PENDENTE
EXEMPLO 3: PLANILHA MISTA (NÃO RECOMENDADO)
csv
NOME,EMAIL,STATUS
João Pereira,joao.pereira@ati.pe.gov.br,PENDENTE
Maria Silva,maria.silva@gmail.com,PENDENTE
Nota: O sistema filtrará automaticamente conforme o modo selecionado.

🔐 SEÇÃO 8 - SEGURANÇA E PERMISSÕES
8.1 CONFIGURAÇÕES DE SEGURANÇA DO WINDOWS
CONFIGURAR POLÍTICAS DE EXECUÇÃO:
powershell
# Abrir PowerShell como Administrador
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
ADICIONAR EXCEÇÕES NO WINDOWS DEFENDER:
Abrir "Segurança do Windows"

Clique em "Proteção contra vírus e ameaças"

Em "Configurações de proteção contra vírus e ameaças", clique em "Gerenciar configurações"

Role para baixo até "Exclusões"

Clique em "Adicionar ou remover exclusões"

Adicione estas exclusões:

C:\Sistema_ATI_PE\

C:\Python311\

C:\Users\%USERNAME%\AppData\Local\Programs\Python\

CONFIGURAR CONTROLE DE CONTA DE USUÁRIO (UAC):
Pressione Windows + R, digite msconfig, Enter

Abra a aba "Ferramentas"

Selecione "Alterar configurações de UAC"

Ajuste para o segundo nível (recomendado)

8.2 PERMISSÕES DE EXECUÇÃO
CONFIGURAR PERMISSÕES PARA PASTA DO SISTEMA:
powershell
# Conceder permissões completas à pasta do sistema
$folder = "C:\Sistema_ATI_PE"
$acl = Get-Acl $folder
$permission = "BUILTIN\Users","FullControl","ContainerInherit,ObjectInherit","None","Allow"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.SetAccessRule($accessRule)
$acl | Set-Acl $folder
CONFIGURAR PERMISSÕES PARA PASTA DO PYTHON:
powershell
$pythonFolder = "C:\Python311"
$acl = Get-Acl $pythonFolder
$permission = "BUILTIN\Users","Modify","ContainerInherit,ObjectInherit","None","Allow"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.SetAccessRule($accessRule)
$acl | Set-Acl $pythonFolder
8.3 POLÍTICAS DE FIREWALL
ADICIONAR REGRAS DE FIREWALL PARA O SISTEMA:
Para Windows Defender Firewall:
powershell
# Permitir Python
New-NetFirewallRule -DisplayName "Python ATI-PE" -Direction Inbound -Program "C:\Python311\python.exe" -Action Allow

# Permitir Chrome para Expresso
New-NetFirewallRule -DisplayName "Chrome Expresso" -Direction Outbound -Program "C:\Program Files\Google\Chrome\Application\chrome.exe" -Action Allow

# Permitir conexões SMTP (Gmail)
New-NetFirewallRule -DisplayName "SMTP Gmail" -Direction Outbound -Protocol TCP -LocalPort Any -RemotePort 587 -Action Allow
Para firewalls corporativos (solicitar ao administrador):
Porta 443 (HTTPS) para expresso.pe.gov.br

Porta 587 (SMTP) para smtp.gmail.com

Portas 80 e 443 para chromedriver.chromium.org

🛠️ SEÇÃO 9 - DESENVOLVIMENTO E COMPILAÇÃO
9.1 ESTRUTURA DO CÓDIGO FONTE
ARQUITETURA DO SISTEMA:
python
# sistema_envio_ati.py - Estrutura principal
class SistemaEnvioATI:
    """
    Classe principal que gerencia todo o sistema
    Herda de customtkinter.CTk para interface gráfica
    """
    
    def __init__(self):
        # Inicialização de componentes da interface
        # Configuração de variáveis
        # Carregamento de configurações
    
    # ===== MÉTODOS PRINCIPAIS =====
    
    def iniciar_thread(self):
        """Inicia o processamento em thread separada"""
    
    def processar_envios_planilha(self):
        """Processa envios a partir de planilha Excel"""
    
    def processar_envio_direto(self):
        """Processa envio direto para único destinatário"""
    
    def enviar_gmail(self):
        """Envia e-mail via SMTP do Gmail"""
    
    # ===== MÉTODOS DE SUPORTE =====
    
    def validar_dominio_email(self):
        """Valida domínio do e-mail conforme modo selecionado"""
    
    def buscar_anexo_dinamico(self):
        """Busca anexos na pasta baseados no nome/e-mail"""
    
    def carregar_config(self):
        """Carrega configurações do arquivo config.txt"""
    
    def guardar_config(self):
        """Salva configurações no arquivo config.txt"""
    
    # ===== MÉTODOS DE INTERFACE =====
    
    def atualizar_interface_modo(self):
        """Atualiza interface conforme modo selecionado"""
    
    def toggle_envio_direto(self):
        """Alterna entre modo planilha e envio direto"""
9.2 COMPILAÇÃO PARA EXECUTÁVEL
USANDO PYINSTALLER (RECOMENDADO):
PASSO 1 - INSTALAR PYINSTALLER:
powershell
pip install pyinstaller==5.13.0
PASSO 2 - CONFIGURAR ESPECIFICAÇÕES DE BUILD:
Crie um arquivo build_spec.spec:

python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['sistema_envio_ati.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('anexos', 'anexos'),
        ('planilhas/modelos', 'planilhas/modelos'),
        ('config.txt', '.'),
        ('README_COMPLETO.md', '.')
    ],
    hiddenimports=[
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.base',
        'pandas._libs.skiplist',
        'customtkinter',
        'PIL._imaging',
        'selenium.webdriver.common.by',
        'email.mime.multipart',
        'email.mime.text',
        'email.mime.base'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'unittest',
        'pydoc',
        'test',
        'distutils'
    ],
    noarchive=False,
    optimize=2
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SistemaATI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Mudar para True para ver console de debug
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'  # Se tiver ícone
)
PASSO 3 - EXECUTAR COMPILAÇÃO:
powershell
pyinstaller build_spec.spec
PASSO 4 - VERIFICAR EXECUTÁVEL GERADO:
powershell
# O executável estará em: dist\SistemaATI\
cd dist\SistemaATI\
.\SistemaATI.exe
9.3 GERAÇÃO DE INSTALADOR MSI
USANDO INNO SETUP (MELHOR OPÇÃO):
PASSO 1 - BAIXAR E INSTALAR INNO SETUP:
Acesse: https://jrsoftware.org/isdl.php

Baixe a versão "Inno Setup QuickStart Pack"

Instale com todas as opções padrão

PASSO 2 - CRIAR SCRIPT DE INSTALAÇÃO:
Crie instalador_ati.iss:

innosetup
; Script de instalação Inno Setup para Sistema ATI-PE
; Autor: Reinaldo Rodrigues - ATI-PE

#define MyAppName "Sistema ATI-PE"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "ATI-PE - Governo do Estado de Pernambuco"
#define MyAppURL "https://www.ati.pe.gov.br"
#define MyAppExeName "SistemaATI.exe"
#define MyOutputDir "Instalador"

[Setup]
AppId={{A1B2C3D4-E5F6-4789-A0B1-C2D3E4F56789}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=LICENSE.txt
InfoBeforeFile=README_COMPLETO.md
OutputDir={#MyOutputDir}
OutputBaseFilename=Setup_Sistema_ATI_PE
SetupIconFile=icon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
ChangesEnvironment=yes

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1
Name: "addtopath"; Description: "Adicionar ao PATH do sistema"; GroupDescription: "Configurações:"

[Files]
Source: "dist\SistemaATI\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "anexos\*"; DestDir: "{app}\anexos"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "planilhas\modelos\*"; DestDir: "{app}\planilhas\modelos"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "README_COMPLETO.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "MANUAL_OPERADOR.pdf"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.txt"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:ProgramOnTheWeb,{#MyAppName}}"; Filename: "{#MyAppURL}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent unchecked

[Registry]
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"; Tasks: addtopath; Check: NeedsAddPath('{app}')

[Code]
function NeedsAddPath(Param: string): boolean;
var
  OrigPath: string;
begin
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', OrigPath)
  then begin
    Result := True;
    exit;
  end;
  Result := Pos(';' + Param + ';', ';' + OrigPath + ';') = 0;
end;
PASSO 3 - COMPILAR INSTALADOR:
Abra o Inno Setup Compiler

File → Open → selecione instalador_ati.iss

Build → Compile

O instalador será gerado em: Instalador\Setup_Sistema_ATI_PE.exe

🐛 SEÇÃO 10 - SOLUÇÃO DE PROBLEMAS
10.1 ERROS COMUNS E SOLUÇÕES
ERRO 1: "ModuleNotFoundError: No module named 'pandas'"
Causa: Biblioteca não instalada
Solução:

powershell
pip install pandas==2.1.3
python -m pip install --upgrade pip
ERRO 2: "WebDriverException: Message: unknown error: cannot find Chrome binary"
Causa: Chrome não instalado ou caminho incorreto
Solução:

Reinstalar Google Chrome

Verificar se está em C:\Program Files\Google\Chrome\Application\

Atualizar Chrome: chrome://settings/help

ERRO 3: "smtplib.SMTPAuthenticationError: (535, b'5.7.8 Username and Password not accepted')"
Causa: Senha do Gmail incorreta ou não é senha de aplicativo
Solução:

Criar nova senha de aplicativo: https://myaccount.google.com/apppasswords

Usar senha de 16 caracteres, não a senha da conta

Verificar se a verificação em 2 fatores está ativa

ERRO 4: PyAutoGUI não encontra coordenadas na tela
Causa: Resolução de tela diferente de 1920x1080
Solução:

Alterar resolução para 1920x1080

Configurar escala em 100% (não 125% ou 150%)

Windows → Configurações → Sistema → Tela → Escala e layout → 100%

ERRO 5: "PermissionError: [Errno 13] Permission denied"
Causa: Sem permissões de escrita
Solução:

powershell
# Executar como administrador
Start-Process powershell -Verb RunAs

# Ou conceder permissões
icacls "C:\Sistema_ATI_PE" /grant Users:F /T
10.2 LOGS E DIAGNÓSTICO
HABILITAR LOGS DETALHADOS:
Adicione ao início do código:

python
import logging

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/sistema_detalhado.log'),
        logging.StreamHandler()
    ]
)
SCRIPT DE DIAGNÓSTICO AUTOMÁTICO:
python
# diagnostico_sistema.py
import sys
import os
import subprocess
import platform

print("=== DIAGNÓSTICO DO SISTEMA ATI-PE ===\n")

# 1. Informações do sistema
print("1. SISTEMA OPERACIONAL:")
print(f"   OS: {platform.system()} {platform.release()}")
print(f"   Versão: {platform.version()}")
print(f"   Arquitetura: {platform.architecture()[0]}")
print(f"   Processador: {platform.processor()}")

# 2. Python
print("\n2. PYTHON:")
print(f"   Versão: {sys.version}")
print(f"   Executável: {sys.executable}")

# 3. Verificar bibliotecas
print("\n3. BIBLIOTECAS INSTALADAS:")
libs = ['pandas', 'pyautogui', 'customtkinter', 'selenium', 'openpyxl']
for lib in libs:
    try:
        exec(f"import {lib}")
        version = eval(f"{lib}.__version__")
        print(f"   ✓ {lib}: {version}")
    except ImportError:
        print(f"   ✗ {lib}: NÃO INSTALADA")
    except AttributeError:
        print(f"   ✓ {lib}: instalada (sem versão)")

# 4. Verificar Chrome
print("\n4. GOOGLE CHROME:")
try:
    import winreg
    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe")
    chrome_path = winreg.QueryValue(key, None)
    print(f"   Instalado em: {chrome_path}")
    
    # Tentar obter versão
    import win32api
    info = win32api.GetFileVersionInfo(chrome_path, '\\')
    version = "%d.%d.%d.%d" % (info['FileVersionMS'] / 16,
                                info['FileVersionMS'] % 16,
                                info['FileVersionLS'] / 16,
                                info['FileVersionLS'] % 16)
    print(f"   Versão: {version}")
except:
    print("   ✗ Chrome não encontrado ou erro na verificação")

# 5. Verificar conexão internet
print("\n5. CONEXÃO COM INTERNET:")
try:
    import urllib.request
    urllib.request.urlopen('https://www.google.com', timeout=5)
    print("   ✓ Conexão OK")
except:
    print("   ✗ Sem conexão ou problema de rede")

# 6. Verificar portas
print("\n6. PORTAS NECESSÁRIAS:")
ports = [587, 443, 80]
for port in ports:
    try:
        import socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(2)
        result = sock.connect_ex(('8.8.8.8', port))
        if result == 0:
            print(f"   ✓ Porta {port}: aberta")
        else:
            print(f"   ✗ Porta {port}: bloqueada")
        sock.close()
    except:
        print(f"   ? Porta {port}: erro na verificação")

print("\n=== FIM DO DIAGNÓSTICO ===")
input("\nPressione Enter para sair...")
10.3 CONTATO PARA SUPORTE
CANAIS DE SUPORTE:
Suporte Técnico Imediato:

E-mail: suporte.ati@ati.pe.gov.br

Telefone: (81) 3184-XXXX (Ramal técnico)

Horário: 08:00 às 17:00 (segunda a sexta)

Desenvolvedor Responsável:

Nome: Reinaldo Rodrigues

E-mail: reinaldo.rodrigues@ati.pe.gov.br

Setor: ATI-PE - Desenvolvimento de Sistemas

Documentação Oficial:

Portal: https://www.ati.pe.gov.br/suporte

Wiki interna: http://wiki.ati.pe.gov.br/sistema_envio

INFORMAÇÕES PARA REPORTAR PROBLEMAS:
Ao reportar um problema, incluir:

Versão do sistema (Help → Sobre)

Sistema operacional e versão

Modo sendo usado (Expresso ou Gmail)

Mensagem de erro completa

Printscreen da tela de erro

Logs da pasta logs\

🔄 SEÇÃO 11 - ATUALIZAÇÃO E MANUTENÇÃO
11.1 ATUALIZAÇÃO DAS BIBLIOTECAS
VERIFICAR ATUALIZAÇÕES DISPONÍVEIS:
powershell
python -m pip list --outdated
ATUALIZAR TODAS AS BIBLIOTECAS:
powershell
# Atualizar pip primeiro
python -m pip install --upgrade pip

# Atualizar bibliotecas principais
pip install --upgrade pandas pyautogui customtkinter selenium webdriver-manager openpyxl pyperclip
ATUALIZAR COM CONTROLE DE VERSÃO:
powershell
# Atualizar para versões específicas
pip install pandas==2.1.4 pyautogui==0.9.55 customtkinter==5.2.1 selenium==4.16.0
11.2 BACKUP DE CONFIGURAÇÕES
SCRIPT DE BACKUP AUTOMÁTICO:
python
# backup_configuracoes.py
import os
import shutil
import datetime
import json

def fazer_backup():
    data_atual = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    pasta_backup = f"backups/backup_{data_atual}"
    
    # Criar pasta de backup
    os.makedirs(pasta_backup, exist_ok=True)
    
    # Arquivos para backup
    arquivos = [
        "config.txt",
        "planilhas/contatos.xlsx",
        "logs/sistema.log"
    ]
    
    # Copiar arquivos
    for arquivo in arquivos:
        if os.path.exists(arquivo):
            shutil.copy2(arquivo, pasta_backup)
    
    # Salvar metadados
    metadados = {
        "data_backup": data_atual,
        "sistema": "ATI-PE Envio Emails",
        "versao": "2.0.0",
        "arquivos_backup": [a for a in arquivos if os.path.exists(a)]
    }
    
    with open(f"{pasta_backup}/metadados.json", "w") as f:
        json.dump(metadados, f, indent=4)
    
    print(f"Backup criado em: {pasta_backup}")
    return pasta_backup

if __name__ == "__main__":
    fazer_backup()
AGENDAR BACKUP AUTOMÁTICO NO WINDOWS:
Pressione Windows + R, digite taskschd.msc

Criar Tarefa Básica

Nome: "Backup Sistema ATI-PE"

Disparador: Diariamente, 17:00

Ação: Iniciar programa

Programa: C:\Python311\python.exe

Argumentos: C:\Sistema_ATI_PE\backup_configuracoes.py

11.3 MIGRAÇÃO PARA NOVAS VERSÕES
PROCESSO DE ATUALIZAÇÃO DE VERSÃO:
VERSÃO 1.x → 2.0:
Backup completo:

powershell
xcopy "C:\Sistema_ATI_PE" "C:\Backup_ATI_Antigo" /E /I /H
Desinstalar versão antiga:

powershell
pip uninstall -y pandas pyautogui customtkinter selenium
Instalar nova versão:

powershell
pip install pandas==2.1.3 pyautogui==0.9.54 customtkinter==5.2.0 selenium==4.15.0
Migrar configurações:

python
# migrar_config.py
import configparser

# Ler config antiga
with open("config_antigo.txt", "r") as f:
    dados_antigos = f.read()

# Converter para novo formato
config = configparser.ConfigParser()
config['USUARIO'] = {'email': dados_antigos.split('=')[1].strip()}
config['SISTEMA'] = {'versao': '2.0.0', 'data_migracao': '2024-01-15'}

with open('config.txt', 'w') as f:
    config.write(f)
TESTE PÓS-MIGRAÇÃO:
Verificar funcionamento básico

Testar envio de teste

Validar planilhas antigas

Confirmar backups

📞 INFORMAÇÕES FINAIS
CRÉDITOS E RESPONSABILIDADES
Desenvolvido por: Reinaldo Rodrigues

Órgão: ATI-PE - Agência de Tecnologia da Informação de Pernambuco

Supervisão: Coordenação de Desenvolvimento de Sistemas

Testes: Equipe de Qualidade ATI-PE

Documentação: Departamento de Suporte Técnico

LICENÇA DE USO
Este sistema é de uso exclusivo da Administração Pública do Estado de Pernambuco. A distribuição, modificação ou uso por terceiros não autorizados é expressamente proibida.

HISTÓRICO DE VERSÕES
v1.0.0 (2023-10-01): Versão inicial com envio básico

v1.5.0 (2023-11-15): Adicionado suporte a anexos Gmail

v2.0.0 (2024-01-15): Sistema completo com interface gráfica

PRÓXIMAS ATUALIZAÇÕES PREVISTAS
Suporte a múltiplas contas simultâneas

Sistema de templates de mensagens

Relatórios estatísticos de envios

Integração com outros provedores de e-mail

Última atualização desta documentação: 15 de Janeiro de 2024
Próxima revisão programada: 15 de Abril de 2024
Documentação mantida por: Departamento de Suporte Técnico - ATI-PE

Para sugestões ou correções nesta documentação, contactar: reinaldogithub@gmail.com
=======
# Automa-o-De-Envio-De-Emails-Via-ExpressoMail-e-Gmail
Sistema de Automação de Envio de E-mails - ATI-PE Este projeto é uma solução de automação desenvolvida para otimizar o envio de e-mails institucionais na ATI-PE (Agência de Tecnologia da Informação de Pernambuco). O sistema permite enviar e-mails em massa via planilha Excel ou envio direto individual, com suporte a duas plataformas:Expresso e Gmail
>>>>>>> f691074f20207bd222a467f2d553a6cb482434a7
