import subprocess  # Importa o módulo subprocess para executar scripts externos
import sys  # Importa o módulo sys para acessar informações sobre o sistema
import logging  # Importa o módulo logging para registrar mensagens de erro

def run_script(script_name):
    # Função para executar um script Python dado o seu nome
    try:
        # Executa o script usando o interpretador Python atual
        subprocess.run([sys.executable, script_name], check=True)
    except subprocess.CalledProcessError as e:
        # Captura erros que ocorrem durante a execução do script
        logging.error(f"Erro ao executar {script_name}: {e}")  # Registra o erro no log
        print(f"Erro ao executar {script_name}: {e}")  # Imprime o erro no console

def main():
    # Função principal que gerencia a execução dos scripts
    scripts_to_run = [
        # Lista de scripts Python a serem executados
        r'S\\ftp.py',
        r'\\oracle.py',
        r'\\outlook.py'
    ]

    for script in scripts_to_run:
        # Itera sobre cada script na lista e chama a função run_script
        run_script(script)
    
    # Aguarda a interação do usuário antes de finalizar o processo
    input("Clique em qualquer tecla para finalizar o processo.")

if __name__ == "__main__":
    # Configura o nível de logging para ERROR
    logging.basicConfig(level=logging.ERROR)
    main()  # Chama a função principal para iniciar o processo
