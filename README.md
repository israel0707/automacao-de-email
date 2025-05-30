Sistema de Automação de Envio de Emails com PDF


_______________________________________________________
📌 Visão Geral
Este é um sistema de automação que monitora uma pasta em busca de novos arquivos PDF, extrai endereços de email válidos desses arquivos e envia automaticamente o PDF como anexo para os emails encontrados. O sistema inclui uma interface gráfica intuitiva para configuração e monitoramento.

✨ Principais Funcionalidades
Monitoramento de pasta em tempo real para novos arquivos PDF

Validação de emails com verificação de registros MX no DNS

Envio automático de emails com anexos PDF

Tratamento de erros com envio de notificações

Estatísticas de processamento em tempo real

Interface gráfica amigável com abas organizadas

Sistema de logs completo com opção de exportação
_______________________________________________________

🛠️ Requisitos do Sistema
Python 3.7 ou superior

Bibliotecas listadas em requirements.txt

_______________________________________________________

📦 Instalação
Clone este repositório:

bash
git clone https://github.com/seu-usuario/email-pdf-automation.git
cd email-pdf-automation
Crie e ative um ambiente virtual (recomendado):

|bash|

python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
Instale as dependências:


|bash|

pip install -r requirements.txt
_______________________________________________________

🚀 Como Usar
Execute o aplicativo:


|bash|

python main.py

Na aba Configurações:

Configure os detalhes do servidor SMTP

Defina as pastas de monitoramento, enviados e erros

Personalize os templates de email

Na aba Monitoramento:

Clique em "Iniciar Monitoramento" para começar

Visualize estatísticas em tempo real

Na aba Logs:

Acompanhe todas as atividades do sistema

Exporte os logs se necessário

_______________________________________________________
⚙️ Configuração SMTP Recomendada

Servidor: Ultilize um fornecido pelo seu provedor 

Porta: 587/465 altere de acordo com seu uso 

Email: seudominio@dominio.com

_______________________________________________________
📂 Estrutura de Pastas
PDF_Enviar: Pasta monitorada para novos PDFs

PDF_Enviados: Arquivos processados com sucesso

PDF_Erros: Arquivos com problemas no processamento
_______________________________________________________

📝 Templates de Email
O sistema suporta dois templates:

Email normal: Para envio bem-sucedido do PDF

Email de erro: Para notificação de problemas
_______________________________________________________

🔄 Processamento de Arquivos
O sistema detecta novos PDFs na pasta monitorada

Extrai todos os emails válidos do texto do PDF

Para cada email válido:

Envia o PDF como anexo

Atualiza as estatísticas

Move o arquivo para a pasta apropriada:

Enviados: Se todos os emails foram enviados

Erros: Se ocorrerem problemas
_______________________________________________________

🛑 Parando o Monitoramento
Use o botão "Parar Monitoramento" na aba de Monitoramento para interromper o serviço com segurança.
_______________________________________________________

📊 Estatísticas
O sistema mantém contagem de:

Arquivos processados

Emails enviados

Erros encontrados
_______________________________________________________

📜 Licença
Este projeto está licenciado sob a licença MIT. Consulte o arquivo LICENSE para obter mais detalhes.
_______________________________________________________

🤝 Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests.
_______________________________________________________

Desenvolvido com ❤️ por israel salles.
Para dúvidas ou suporte entre em contato: sallesisrael66@gmail.com
