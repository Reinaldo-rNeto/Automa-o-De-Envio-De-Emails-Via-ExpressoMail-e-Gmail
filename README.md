🚀 Automação de Envio de E-mails (ATI-PE)
Sistema de automação multi-usuário desenvolvido para facilitar o envio em massa de e-mails através das plataformas ExpressoMail (.pe.gov.br) e Gmail. Desenvolvido com interface gráfica moderna e suporte a processamento em segundo plano.

✨ Funcionalidades
Dual Mode: Suporte para envio via Selenium (ExpressoMail) e SMTP (Gmail).

Envio em Massa: Leitura automática de planilhas Excel (.xlsx, .xls).

Envio Direto: Opção para envio rápido sem necessidade de planilha.

Anexos Inteligentes: * No Gmail, permite anexar arquivos fixos ou busca dinâmica na pasta /anexos baseada no nome/e-mail do destinatário.

Segurança de Domínio: Validação automática (bloqueia envios externos no modo Expresso).

Interface Moderna: Construída com CustomTkinter com suporte a Dark Mode.

Persistência: Salva o último usuário utilizado para agilizar o próximo acesso.

🛠️ Tecnologias Utilizadas
Python 3.x

Interface: CustomTkinter, pyautogui, pyperclip.

Manipulação de Dados: pandas, openpyxl.

Automação Web: selenium, webdriver-manager.

Comunicação: smtplib, email.mime.

📋 Pré-requisitos
Antes de rodar o projeto, instale as dependências necessárias:

Bash
pip install pyautogui pyperclip pandas pandas openpyxl selenium webdriver-manager customtkinter
Nota para usuários Gmail: Para usar o modo Gmail, é necessário gerar uma "Senha de App" nas configurações de segurança da sua conta Google, caso utilize autenticação de dois fatores.

🚀 Como Usar
Clone o repositório:

Bash
git clone https://github.com/Reinaldo-rNeto/Automa-o-De-Envio-De-Emails-Via-ExpressoMail-e-Gmail.git
Execute o script:

Bash
python seu_arquivo.py
Configuração da Planilha: Certifique-se de que sua planilha Excel contenha as colunas NOME e EMAIL.

Anexos Dinâmicos: Coloque os arquivos na pasta anexos/ com o nome do arquivo contendo parte do nome ou e-mail do destinatário.

📌 Observações Importantes
Modo Expresso: O sistema utiliza automação de interface (pyautogui). Evite mexer no mouse ou teclado enquanto o envio estiver em curso, pois o script simula cliques reais na tela.

Resolução: O script está configurado para calcular posições de clique baseadas na resolução da sua tela.

Desenvolvido por: Reinaldo Rodrigues - Estágiario - ATI-PE
