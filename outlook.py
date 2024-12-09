import win32com.client  # Importa a biblioteca para interagir com o Outlook
from datetime import datetime, timedelta  # Importa classes para manipulação de datas
import os  # Importa biblioteca para interações com o sistema operacional
import pandas as pd  # Importa biblioteca para manipulação de dados em formato de tabela
import logging  # Importa biblioteca para registro de logs
import re  # Importa biblioteca para expressões regulares

# Configuração do logging
logging.basicConfig(
    filename=(r'S:\SRC\01_Gestao_da_Receita\01-Recuperacao_Energia\2024\00 - Relatórios\Indicadores\04-TOI_Recebido\Retirados_FTP_py\logging\email_log.log'), 
    level=logging.INFO,  # Define o nível de log como INFO
    format='%(asctime)s - %(levelname)s - %(message)s',  # Formato das mensagens de log
    filemode='w'  # Modo de escrita do arquivo de log (substitui o arquivo a cada execução)
)

# Definição de caminhos como constantes
CAMINHO_DESTINATARIOS = r'S:\SRC\01_Gestao_da_Receita\01-Recuperacao_Energia\2024\00 - Relatórios\Indicadores\04-TOI_Recebido\Retirados_FTP_py\bases\Emails.xlsx'  # Caminho para o arquivo de destinatários
CAMINHO_PLANILHA_BASE = r"S:\SRC\01_Gestao_da_Receita\01-Recuperacao_Energia\03-Usuarios\ALEX GUIDONI\--- ANÁLISE TOIS FTP\Processos retirados do FTP em {}.xlsx"  # Caminho base para a planilha a ser anexada

def get_greeting():
    # Função para obter uma saudação com base na hora atual
    current_hour = datetime.now().hour  # Obtém a hora atual
    if 6 <= current_hour < 12:  # Se a hora estiver entre 6 e 12
        return "Bom dia"  # Retorna "Bom dia"
    elif 12 <= current_hour < 18:  # Se a hora estiver entre 12 e 18
        return "Boa tarde"  # Retorna "Boa tarde"
    else:  # Para qualquer outra hora
        return "Boa noite"  # Retorna "Boa noite"

def validate_email(email):
    # Função para validar o formato de um endereço de e-mail
    regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'  # Expressão regular para validação de e-mail
    return re.match(regex, email) is not None  # Retorna True se o e-mail for válido, caso contrário, False

def read_recipients(recipients_path):
    # Função para ler os destinatários de um arquivo Excel
    logging.info("Iniciando a leitura do arquivo de destinatários.")  # Log de início da leitura
    try:
        with pd.ExcelFile(recipients_path) as xls:  # Tenta abrir o arquivo Excel
            df = pd.read_excel(xls)  # Lê o conteúdo do arquivo em um DataFrame
            logging.info("Arquivo de destinatários lido com sucesso.")  # Log de sucesso na leitura
    except Exception as e:  # Captura qualquer exceção que ocorra
        logging.error(f"Erro ao ler o arquivo de destinatários: {e}")  # Log do erro
        print("Erro ao ler o arquivo de destinatários. Verifique os logs para mais detalhes.")  # Mensagem de erro para o usuário
        return [], []  # Retorna listas vazias em caso de erro

    # Extrai os destinatários principais e em cópia (CC) do DataFrame
    main_recipients = list(set(df.iloc[:, 0].dropna().tolist()))  # Destinatários principais (sem duplicatas)
    cc_recipients = df.iloc[:, 1].dropna().tolist()  # Destinatários em cópia (CC)

    return main_recipients, cc_recipients  # Retorna as listas de destinatários

def create_outlook_email():
    # Função para criar e preparar um e-mail no Outlook
    current_date = datetime.now().strftime("%d.%m.%Y")  # Obtém a data atual formatada
    previous_day = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")  # Obtém a data do dia anterior
    greeting = get_greeting()  # Obtém a saudação apropriada

    main_recipients, cc_recipients = read_recipients(CAMINHO_DESTINATARIOS)  # Lê os destinatários do arquivo Excel

    if not main_recipients:  # Verifica se não há destinatários principais
        logging.warning("Nenhum destinatário principal encontrado.")  # Log de aviso
        print("Nenhum destinatário principal encontrado.")  # Mensagem para o usuário
        return  # Sai da função se não houver destinatários

    logging.info(f"Destinatários principais encontrados: {main_recipients}")  # Log dos destinatários principais encontrados
    logging.info(f"Destinatários em cópia (CC): {cc_recipients}")  # Log dos destinatários em cópia

    outlook = win32com.client.Dispatch("Outlook.Application")  # Cria uma instância do Outlook
    email = outlook.CreateItem(0)  # Cria um novo item de e-mail

    email.Subject = f"Processos retirados do FTP em {current_date}"  # Define o assunto do e-mail

    # Define o corpo do e-mail em HTML
    email.HTMLBody = f"""
    <html>
    <body>
    <p>Boa tarde a todos!</p>
    <p>Segue em anexo a relação dos Processos de TOI recebidos no dia 08.12.2024.</p>
    <p>Processos da Energec recebidos por e-mail.<br>
    Caso tenha algum processo que vocês enviaram e não se encontra na relação anexa, favor me retornar o e-mail informando e disponibilizar novamente no FTP.</p>
    <p>Atenciosamente,<br>
    Gerencia de Gestão da Receita</p>
    </body>
    </html>
    """

    # Adiciona os destinatários principais ao e-mail
    for recipient in main_recipients:
        if validate_email(recipient):  # Valida o e-mail do destinatário
            email.Recipients.Add(recipient)  # Adiciona o destinatário ao e-mail
            logging.info(f"Destinatário adicionado: {recipient}")  # Log do destinatário adicionado
        else:  # Se o e-mail não for válido
            logging.warning(f"Endereço de e-mail inválido: {recipient}")  # Log de aviso para e-mail inválido

    # Adiciona os destinatários em cópia (CC) ao e-mail
    for cc in cc_recipients:
        if validate_email(cc):  # Valida o e-mail do destinatário em CC
            email.CC += f";{cc}"  # Adiciona o destinatário em CC
            logging.info(f"Destinatário em cópia adicionado: {cc}")  # Log do destinatário em cópia adicionado
        else:  # Se o e-mail não for válido
            logging.warning(f"Endereço de e-mail inválido em CC: {cc}")  # Log de aviso para e-mail inválido em CC

    spreadsheet_path = CAMINHO_PLANILHA_BASE.format(current_date)  # Formata o caminho da planilha com a data atual

    if os.path.exists(spreadsheet_path):  # Verifica se a planilha existe
        email.Attachments.Add(spreadsheet_path)  # Anexa a planilha ao e-mail
        logging.info(f"Anexo adicionado: {spreadsheet_path}")  # Log do anexo adicionado
    else:  # Se a planilha não for encontrada
        logging.error(f"Planilha não encontrada: {spreadsheet_path}")  # Log de erro para planilha não encontrada

    email.Display()  # Exibe o e-mail preparado para envio
    logging.info(f"E-mail preparado para envio em {current_date}.")  # Log de e-mail preparado
    print(f"E-mail preparado para envio em {current_date}.")  # Mensagem para o usuário informando que o e-mail está preparado

if __name__ == "__main__":  # Verifica se o script está sendo executado diretamente
    create_outlook_email()  # Chama a função para criar o e-mail
