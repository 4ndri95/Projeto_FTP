import ftplib 
import os  
import logging  

# Função para configurar o logging
def setup_logging():
    logging.basicConfig(
        filename=(r'\\ftp_log.log'), 
        level=logging.DEBUG,  # Define o nível de log como DEBUG
        format='%(asctime)s - %(levelname)s - %(message)s',  # Define o formato das mensagens de log
        filemode='w'  # Modo de escrita do arquivo de log (substitui o arquivo a cada execução)
    )

# Classe para gerenciar o download de arquivos via FTP
class FTPDownloader:
    def __init__(self, username, password):
        # Inicializa a classe com o nome de usuário e senha
        self.username = username
        self.password = password
        self.ftp = None  # Inicializa o objeto FTP como None

    def connect(self):
        # Método para conectar ao servidor FTP
        try:
            self.ftp = ftplib.FTP("ftp.sua_ftp.com.br")  # Cria uma instância de FTP
            self.ftp.login(self.username, self.password)  # Realiza o login com as credenciais fornecidas
            self.ftp.set_pasv(True)  # Ativa o modo passivo
            self.ftp.voidcmd("TYPE I")  # Define o tipo de transferência como binário
            logging.info(f"Conectado ao servidor FTP com usuário {self.username}")  # Log de sucesso
        except ftplib.error_perm as e:
            # Captura erros de permissão
            self.log_error(f"Erro de permissão ao conectar: {e}")
        except ftplib.error_temp as e:
            # Captura erros temporários
            self.log_error(f"Erro temporário ao conectar: {e}")
        except Exception as e:
            # Captura qualquer outro erro inesperado
            self.log_error(f"Erro inesperado ao conectar: {e}")

    def disconnect(self):
        # Método para desconectar do servidor FTP
        if self.ftp:
            self.ftp.quit()  # Encerra a sessão FTP
            logging.info(f"Desconectado do servidor FTP com usuário {self.username}")  # Log de desconexão

    def log_error(self, message):
        # Método para registrar erros
        logging.error(message)  # Registra a mensagem de erro
        raise Exception(message)  # Levanta uma exceção

    def download_files(self, folder, base_folder):
        # Método para baixar arquivos de um diretório específico
        try:
            self.ftp.cwd(folder)  # Muda para o diretório desejado
            files = self.ftp.nlst()  # Lista os arquivos no diretório
            downloaded_files, failed_downloads = self.process_files_from_list(files, base_folder)  # Processa a lista de arquivos

            logging.info(f"Número de arquivos baixados do diretório {folder}: {len(downloaded_files)}")  # Log do número de arquivos baixados
            
            self.log_failed_downloads(failed_downloads, folder)  # Registra arquivos que falharam no download
            return len(downloaded_files), downloaded_files  # Retorna a quantidade de arquivos baixados e a lista
        except Exception as e:
            self.log_error(f"Erro ao baixar arquivos do diretório {folder}: {e}")  # Registra erro ao baixar arquivos

    def process_files_from_list(self, files, base_folder):
        # Método para processar a lista de arquivos e tentar baixá-los
        downloaded_files = []  # Lista para arquivos baixados
        failed_downloads = []  # Lista para arquivos que falharam no download

        for file in files:
            file_name = self.decode_file_name(file)  # Decodifica o nome do arquivo
            if file_name and file_name.lower().endswith(".pdf"):  # Verifica se é um arquivo PDF
                local_file_path = os.path.join(base_folder, file_name)  # Define o caminho local para salvar o arquivo
                os.makedirs(os.path.dirname(local_file_path), exist_ok=True)  # Cria o diretório se não existir
                if self.download_file(file_name, local_file_path):  # Tenta baixar o arquivo
                    downloaded_files.append(file_name)  # Adiciona à lista de arquivos baixados
                else:
                    failed_downloads.append(file_name)  # Adiciona à lista de falhas

        return downloaded_files, failed_downloads  # Retorna as listas de arquivos baixados e falhados

    def process_files(self, folder, base_folder):
        # Método para processar os arquivos em um diretório específico
        if not isinstance(folder, str):
            # Verifica se a pasta é uma string, caso contrário, levanta um erro
            raise ValueError(f"A pasta esperada era uma string, conseguiu {type(folder)}")

        if not os.path.exists(base_folder):
            # Verifica se o diretório base existe
            logging.error(f"O diretório base {base_folder} não existe. Abortando o download.")  # Log de erro
            return 0, []  # Retorna 0 arquivos transferidos e lista vazia

        transferred_files, downloaded_files = self.download_files(folder, base_folder)  # Chama o método para baixar arquivos
        if transferred_files > 0:
            # Se houver arquivos transferidos, tenta deletá-los do servidor FTP
            # Adicione um '#' atrás da linha abaixo para desativar a função de deletar os documentos da FTP
            self.delete_files(folder, downloaded_files)  # Chama o método para deletar arquivos
            # Retire o '#' do 'pass' para que o script seja executado de forma correta
            # pass
                  
        return transferred_files, downloaded_files  # Retorna o número de arquivos transferidos e a lista de arquivos baixados

    def decode_file_name(self, file):
        # Método para decodificar o nome do arquivo
        try:
            return file.encode('latin1').decode('utf-8')  # Tenta decodificar o nome do arquivo
        except UnicodeDecodeError:
            # Captura erro de decodificação
            logging.warning(f"Erro ao decodificar o nome do arquivo: {file}. Ignorando.")  # Log de aviso
            return None  # Retorna None se a decodificação falhar

    def download_file(self, file_name, local_file_path):
        # Método para baixar um arquivo individual
        try:
            with open(local_file_path, "wb") as f:
                # Abre o arquivo local para escrita em modo binário
                self.ftp.retrbinary("RETR " + file_name, f.write)  # Realiza o download do arquivo
            logging.info(f"Arquivo {file_name} transferido com sucesso para {local_file_path}")  # Log de sucesso
            return True  # Retorna True se o download foi bem-sucedido
        except (OSError, ftplib.error_perm) as e:
            # Captura erros de sistema de arquivos ou de permissão
            logging.error(f"Erro ao salvar o arquivo {file_name}: {e}")  # Log de erro
            return False  # Retorna False se o download falhar

    def log_failed_downloads(self, failed_downloads, folder):
        # Método para registrar arquivos que falharam no download
        if failed_downloads:
            # Se houver arquivos que falharam
            print(f"Falha ao transferir arquivos da pasta {folder}: {', '.join(failed_downloads)}")  # Imprime os arquivos que falharam

    def delete_file(self, file):
        # Método para deletar um arquivo específico do servidor FTP
        try:
            self.ftp.delete(file)  # Tenta deletar o arquivo do servidor FTP
            logging.info(f"Arquivo {file} deletado com sucesso do servidor FTP.")  # Log de sucesso
        except ftplib.error_perm as e:
            # Captura erros de permissão
            logging.error(f"Erro de permissão ao tentar deletar o arquivo {file}: {e}")  # Log de erro
        except Exception as e:
            # Captura qualquer outro erro inesperado
            logging.error(f"Erro inesperado ao tentar deletar o arquivo {file}: {e}")  # Log de erro

    def delete_files(self, folder, files=None):
        # Método para deletar um ou múltiplos arquivos do servidor FTP
        try:
            if files is None:  # Se nenhum arquivo for especificado, assume que deve deletar tudo no diretório
                self.ftp.cwd(folder)  # Muda para o diretório especificado
                existing_files = self.ftp.nlst()  # Lista os arquivos existentes no diretório
                files = existing_files  # Prepara a lista de arquivos para deletar
            else:
                self.ftp.cwd(folder)  # Muda para o diretório especificado
                existing_files = self.ftp.nlst()  # Lista os arquivos existentes no diretório

            for file in files:
                if file in existing_files: 
                    self.delete_file(file)  # Chama o método para deletar o arquivo
                else:
                    logging.warning(f"Arquivo {file} não encontrado em {folder}. Ignorando.")  # Log de aviso se o arquivo não for encontrado
        except Exception as e:
            logging.error(f"Erro ao acessar o diretório {folder} para deletar arquivos: {e}")  # Registra erro ao acessar o diretório

    def __enter__(self):
        # Método para permitir o uso da classe como um gerenciador de contexto
        self.connect()  # Conecta ao servidor FTP
        return self  # Retorna a instância da classe para uso no contexto

    def __exit__(self, exc_type, exc_value, traceback):
        # Método chamado ao sair do bloco 'with'
        self.disconnect()  # Desconecta do servidor FTP

def main():
    # Função principal que controla o fluxo do programa
    setup_logging()  # Configura o logging

    # Lista de servidores FTP com suas credenciais
    ftp_servers = [
        {"username": "username_1", "password": "password_1"},
        {"username": "username_2", "password": "password_2"},
        {"username": "username_3", "password": "password_3"},
        {"username": "username_4", "password": "password_4"}
    ]

    # Diretórios no servidor FTP que serão acessados
    ftp_folders = ["/Auto Religacao", "/Avaria", "/processo_completo_PI", "/Operacoes"]

    # Mapeia os diretórios FTP para os diretórios locais onde os arquivos serão salvos
    base_folders = {
        "/path//ftp//1": "//path//1",
        "/path//ftp//2": "//path//2",
        "/path//ftp//3": "//path//3",
        "/path//ftp//4": "//path//4"
    }

    # Dicionário para armazenar o número total de arquivos transferidos por diretório base
    transferred_files_by_base_folder = {folder: 0 for folder in base_folders.values()} 
    total_transferred_files = 0  # Inicializa o contador total de arquivos transferidos

    # Loop através de cada servidor FTP
    for server in ftp_servers:
        with FTPDownloader(server["username"], server["password"]) as downloader:
            # Conecta ao servidor FTP usando as credenciais
            for folder in ftp_folders:
                if isinstance(folder, str):
                    # Verifica se o nome da pasta é uma string
                    transferred_files, downloaded_files = downloader.process_files(folder, base_folders[folder]) 
                    # Processa os arquivos e obtém o número de transferidos e a lista de baixados
                    transferred_files_by_base_folder[base_folders[folder]] += transferred_files  # Atualiza o contador por diretório
                    total_transferred_files += transferred_files  # Atualiza o contador total
                else:
                    logging.error(f"Folder deve ser uma string, mas recebeu {type(folder)}")  # Log de erro se a pasta não for uma string

    # Log e imprime o total de arquivos transferidos por diretório
    for base_folder, total_transferred in transferred_files_by_base_folder.items():
        logging.info(f"Transferidos {total_transferred} PDF(s) para a pasta {base_folder}")  # Log do número de arquivos transferidos
        print(f"Transferidos {total_transferred} PDF(s) para a pasta {base_folder}")  # Imprime o número de arquivos transferidos

    # Log e imprime o total de arquivos PDF transferidos
    logging.info(f"Total de arquivos PDF transferidos: {total_transferred_files}")  # Log do total de arquivos transferidos
    print(f"Total de arquivos PDF transferidos: {total_transferred_files}")  # Imprime o total de arquivos transferidos

# Verifica se o script está sendo executado diretamente
if __name__ == "__main__":
    main()  # Chama a função principal
