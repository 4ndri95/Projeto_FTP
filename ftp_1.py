import ftplib  # Importa o módulo ftplib para operações FTP
import os  # Importa o módulo os para interações com o sistema de arquivos
import logging  # Importa o módulo logging para registrar logs de eventos

def setup_logging():
    # Configura o logging para registrar mensagens em um arquivo
    logging.basicConfig(
        filename='ftp_downloader.log',  # Nome do arquivo de log
        level=logging.DEBUG,  # Nível de logging
        format='%(asctime)s - %(levelname)s - %(message)s'  # Formato das mensagens de log
    )

class FTPDownloader:
    # Inicializa a classe FTPDownloader com nome de usuário e senha
    def __init__(self, username, password):
        self.username = username  # Armazena o nome de usuário
        self.password = password  # Armazena a senha
        self.ftp = None  # Inicializa a conexão FTP como None

    def connect(self):
        # Tenta conectar ao servidor FTP
        try:
            self.ftp = ftplib.FTP("ftp.exemple.com")  # Conecta ao servidor FTP
            self.ftp.login(self.username, self.password)  # Faz login com as credenciais fornecidas
            self.ftp.set_pasv(True)  # Ativa o modo passivo para a conexão
            self.ftp.voidcmd("TYPE I")  # Define o tipo de transferência como binário
            logging.info(f"Conectado ao servidor FTP com usuário {self.username}")  # Registra a conexão bem-sucedida
        except ftplib.error_perm as e:
            logging.error(f"Erro de permissão ao conectar ao servidor FTP com usuário {self.username}: {e}")  # Registra erro de permissão
            raise
        except ftplib.error_temp as e:
            logging.error(f"Erro temporário ao conectar ao servidor FTP com usuário {self.username}: {e}")  # Registra erro temporário
            raise
        except Exception as e:
            logging.error(f"Erro inesperado ao conectar ao servidor FTP com usuário {self.username}: {e}")  # Registra erro inesperado
            raise

    def disconnect(self):
        # Desconecta do servidor FTP se a conexão estiver ativa
        if self.ftp:
            self.ftp.quit()  # Encerra a sessão FTP
            logging.info(f"Desconectado do servidor FTP com usuário {self.username}")  # Registra a desconexão

    def download_files(self, folder, base_folder):
        # Tenta baixar arquivos de um diretório específico no servidor FTP
        try:
            self.ftp.cwd(folder)  # Muda para o diretório especificado
            files = self.ftp.nlst()  # Lista os arquivos no diretório
            transferred_files = 0  # Contador de arquivos transferidos
            downloaded_files = []  # Lista de arquivos baixados
            failed_downloads = []  # Lista de arquivos que falharam ao baixar

            for file in files:  # Itera sobre cada arquivo no diretório
                try:
                    file_name = file.encode('latin1').decode('utf-8')  # Tenta decodificar o nome do arquivo
                except UnicodeDecodeError:
                    logging.warning(f"Erro ao decodificar o nome do arquivo: {file}. Ignorando.")  # Ignora arquivos com erro de codificação
                    continue 

                if file_name.lower().endswith(".pdf"):  # Verifica se o arquivo é um PDF
                    local_file_path = os.path.join(base_folder, file_name)  # Define o caminho local para salvar o arquivo
                    os.makedirs(os.path.dirname(local_file_path), exist_ok=True)  # Cria diretórios se necessário
                    try:
                        with open(local_file_path, "wb") as f:  # Abre o arquivo local para escrita em modo binário
                            self.ftp.retrbinary("RETR " + file_name, f.write)  # Baixa o arquivo do FTP
                        transferred_files += 1  # Incrementa o contador de arquivos transferidos
                        downloaded_files.append(file_name)  # Adiciona o arquivo à lista de arquivos baixados
                        logging.info(f"Arquivo {file_name} transferido com sucesso para {local_file_path}")  # Registra a transferência bem-sucedida
                    except OSError as e:
                        logging.error(f"Erro ao salvar o arquivo {file_name} na pasta {base_folder}: {e}")  # Registra erro ao salvar
                        failed_downloads.append(file_name)  # Adiciona à lista de falhas
                    except ftplib.error_perm as e:
                        logging.error(f"Erro de permissão ao tentar baixar o arquivo {file_name}: {e}")  # Registra erro de permissão ao baixar
                        failed_downloads.append(file_name)  # Adiciona à lista de falhas
                    except Exception as e:
                        logging.error(f"Erro inesperado ao tentar baixar o arquivo {file_name}: {e}")  # Registra erro inesperado ao baixar
                        failed_downloads.append(file_name)  # Adiciona à lista de falhas

            if failed_downloads:  # Se houver arquivos que falharam ao ser transferidos
                logging.error(f"Falha ao transferir os seguintes arquivos da pasta {folder} no FTP: {', '.join(failed_downloads)}")  # Registra arquivos que falharam
                print(f"Falha ao transferir os seguintes arquivos da pasta {folder} no FTP: {', '.join(failed_downloads)}")  # Exibe falhas na transferência

            return transferred_files, downloaded_files  # Retorna o número de arquivos transferidos e a lista de arquivos baixados
        except ftplib.error_perm as e:
            logging.error(f"Erro ao acessar o diretório {folder}: {e}")  # Registra erro de permissão ao acessar o diretório
            return 0, []  # Retorna zero transferidos e lista vazia
        except ftplib.error_temp as e:
            logging.error(f"Erro temporário ao acessar o diretório {folder}: {e}")  # Registra erro temporário ao acessar o diretório
            return 0, []  # Retorna zero transferidos e lista vazia
        except Exception as e:
            logging.error(f"Erro inesperado ao tentar baixar arquivos do diretório {folder}: {e}")  # Registra erro inesperado
            return 0, []  # Retorna zero transferidos e lista vazia

    def delete_files(self, folder, files):
        # Tenta deletar arquivos de um diretório específico no servidor FTP
        try:
            self.ftp.cwd(folder)  # Muda para o diretório especificado
            existing_files = self.ftp.nlst()  # Lista os arquivos existentes no diretório
            for file in files:  # Itera sobre cada arquivo que deve ser deletado
                if file in existing_files:  # Verifica se o arquivo existe no servidor
                    try:
                        self.ftp.delete(file)  # Deleta o arquivo do servidor
                        logging.info(f"Arquivo {file} deletado com sucesso de {folder}")  # Registra a deleção bem-sucedida
                    except ftplib.error_perm as e:
                        logging.error(f"Erro ao deletar o arquivo {file} em {folder}: {e}")  # Registra erro de permissão ao deletar
                    except Exception as e:
                        logging.error(f"Erro inesperado ao tentar deletar o arquivo {file} em {folder}: {e}")  # Registra erro inesperado ao deletar
                else:
                    logging.warning(f"Arquivo {file} não encontrado em {folder}. Ignorando a tentativa de deleção.")  # Registra que o arquivo não foi encontrado
        except ftplib.error_perm as e:
            logging.error(f"Erro ao acessar o diretório {folder} para deletar arquivos: {e}")  # Registra erro de permissão ao acessar o diretório
        except ftplib.error_temp as e:
            logging.error(f"Erro temporário ao acessar o diretório {folder} para deletar arquivos: {e}")  # Registra erro temporário ao acessar o diretório
        except Exception as e:
            logging.error(f"Erro inesperado ao tentar acessar o diretório {folder} para deletar arquivos: {e}")  # Registra erro inesperado

    def process_files(self, folder, base_folder):
        # Processa o download e a deleção de arquivos
        if not os.path.exists(base_folder):  # Verifica se o diretório base existe
            logging.error(f"O diretório base {base_folder} não existe. Abortando o download.")  # Registra erro se o diretório não existir
            return

        transferred_files, downloaded_files = self.download_files(folder, base_folder)  # Tenta baixar arquivos
        if transferred_files > 0:  # Se houve arquivos transferidos
            self.delete_files(folder, downloaded_files)  # Deleta os arquivos baixados do servidor

    def __enter__(self):
        # Método para permitir uso com 'with' para gerenciar a conexão automaticamente
        self.connect()  # Conecta ao servidor FTP
        return self  # Retorna a instância do downloader

    def __exit__(self, exc_type, exc_value, traceback):
        # Método para garantir que a desconexão do servidor FTP ocorra automaticamente
        self.disconnect()  # Desconecta do servidor FTP

def main():
    setup_logging()  # Configura o logging

    # Lista de servidores FTP com credenciais
    ftp_servers = [
        {"username": "username", "password": "password"},
        {"username": "username", "password": "password"},
        {"username": "username", "password": "password"},
        {"username": "username", "password": "password"}
    ]

    # Diretórios no servidor FTP a serem processados
    ftp_folders = ["/directory_1", "/directory_2", "/directory_3"]
    # Mapeia cada diretório FTP para um diretório base local
    base_folders = {
        "/directory_1": "//path//to//directory",
        "/directory_2": "//path//to//directory",
        "/directory_3": "//path//to//directory"
    }

    # Dicionário para armazenar o número total de arquivos transferidos por diretório base
    transferred_files_by_base_folder = {}

    # Itera sobre cada servidor FTP na lista
    for server in ftp_servers:
        with FTPDownloader(server["username"], server["password"]) as downloader:  # Usa o downloader com gerenciamento automático de contexto
            for folder in ftp_folders:  # Itera sobre cada diretório FTP
                transferred_files, downloaded_files = downloader.process_files(folder, base_folders[folder])  # Processa arquivos no diretório
                
                # Atualiza o total de arquivos transferidos para o diretório base correspondente
                transferred_files_by_base_folder[base_folders[folder]] = transferred_files_by_base_folder.get(base_folders[folder], 0) + transferred_files

    # Registra e imprime o total de arquivos transferidos para cada diretório base
    for base_folder, total_transferred_files in transferred_files_by_base_folder.items():
        logging.info(f"Transferidos {total_transferred_files} PDF(s) para a pasta {base_folder}")  # Registra a quantidade de arquivos transferidos
        print(f"Transferidos {total_transferred_files} PDF(s) para a pasta {base_folder}")  # Exibe a quantidade de arquivos transferidos

# Ponto de entrada principal do programa
if __name__ == "__main__":
    main()  # Chama a função principal para executar o programa
