import os
import subprocess
import requests
from datetime import datetime
import json

# Função para carregar configurações de um arquivo JSON
def load_config():
    try:
        with open(r'C:\Atomatização Backup\config.json', 'r') as json_file:
            config = json.load(json_file)
            return config
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

# Função para solicitar configurações do usuário
def get_config():
    config = {}
    
    # Solicita as informações ao usuário
    config['CNPJ'] = input("Digite o CNPJ da empresa: ")
    config['EMPRESA'] = input("Digite o nome da empresa: ")
    config['SERVER_NAME'] = input("Digite o nome do servidor: ")
    config['DATABASE_NAME'] = input("Digite o nome do banco de dados: ")
    config['BACKUP_PATH'] = input("Digite o caminho do backup: ")
    config['DESTINATION_PATH'] = input("Digite o destino da cópia: ")
    config['BACKUP_SCRIPT_PATH'] = input("Digite o caminho para o executável: ")
    
    # Solicitar horário da tarefa
    config['SCHEDULE_TIME'] = input("Digite o horário para agendar o backup (HH:MM, 24h): ")

    # Solicitar caminho do arquivo modelo de tarefa
    config['TEMPLATE_TASK_PATH'] = input("Digite o caminho do arquivo modelo de tarefa (.xml): ")

    # Salva as informações em um arquivo JSON
    with open(r'C:\Atomatização Backup\config.json', 'w') as json_file:
        json.dump(config, json_file, indent=4)

    print("Configurações salvas em config.json.")
    return config

# Função para importar tarefa do Agendador de Tarefas a partir de um arquivo XML
def import_task(template_task_path):
    task_name = "BackupDatabaseTask"
    command = f'schtasks /create /tn "{task_name}" /xml "{template_task_path}"'

    try:
        subprocess.run(command, shell=True, check=True)
        print(f"Tarefa '{task_name}' criada a partir do modelo.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao criar tarefa: {e}")

# Carregar configurações do arquivo JSON
config = load_config()

# Se as configurações não foram carregadas, pede ao usuário
if not config:
    config = get_config()
else:
    # Validar se cada configuração está preenchida, caso contrário solicitar
    if not config.get('CNPJ'):
        config['CNPJ'] = input("Digite o CNPJ da empresa: ")
    
    if not config.get('EMPRESA'):
        config['EMPRESA'] = input("Digite o nome da empresa: ")

    if not config.get('SERVER_NAME'):
        config['SERVER_NAME'] = input("Digite o nome do servidor: ")

    if not config.get('DATABASE_NAME'):
        config['DATABASE_NAME'] = input("Digite o nome do banco de dados: ")

    if not config.get('BACKUP_PATH'):
        config['BACKUP_PATH'] = input("Digite o caminho do backup: ")

    if not config.get('DESTINATION_PATH'):
        config['DESTINATION_PATH'] = input("Digite o destino da cópia: ")

    if not config.get('BACKUP_SCRIPT_PATH'):
        config['BACKUP_SCRIPT_PATH'] = input("Digite o caminho para o executável: ")

    if not config.get('SCHEDULE_TIME'):
        config['SCHEDULE_TIME'] = input("Digite o horário para agendar o backup (HH:MM, 24h): ")

    if not config.get('TEMPLATE_TASK_PATH'):
        config['TEMPLATE_TASK_PATH'] = input("Digite o caminho do arquivo modelo de tarefa (.xml): ")

    # Salva as informações em um arquivo JSON atualizado
    with open(r'C:\Atomatização Backup\config.json', 'w') as json_file:
        json.dump(config, json_file, indent=4)

# Configurações
CNPJ = config.get("CNPJ")
EMPRESA = config.get("EMPRESA")
SERVER_NAME = config.get("SERVER_NAME")
DATABASE_NAME = config.get("DATABASE_NAME")
BACKUP_PATH = config.get("BACKUP_PATH")
DESTINATION_PATH = config.get("DESTINATION_PATH")
SCHEDULE_TIME = config.get("SCHEDULE_TIME")
BACKUP_SCRIPT_PATH = config.get("BACKUP_SCRIPT_PATH")
TEMPLATE_TASK_PATH = config.get("TEMPLATE_TASK_PATH")
TEAMS_WEBHOOK_URL = "https://precisaosistemas.webhook.office.com/webhookb2/7a1631af-e709-491e-9fb2-688cca5e43cc@c31a0c79-c82f-45d6-abed-6a70879752ab/IncomingWebhook/4a08f529acfe4273921b22fe472bb511/2154cc37-2ff5-425d-99e9-4b4f23adc011"  # Substitua pela URL do seu webhook
SEVEN_ZIP_PATH = r"C:\REALIZA_BACKUP\7z\7za.exe"

# Função para enviar mensagem para o Microsoft Teams
def send_teams_message(message):
    payload = {
        'text': message
    }
    try:
        requests.post(TEAMS_WEBHOOK_URL, json=payload)
    except Exception as e:
        print(f"Erro ao enviar mensagem para o Teams: {e}")

# Obter data e hora atual
now = datetime.now()
day_name = now.strftime("%A")
backup_file_name = f"{DATABASE_NAME}_{day_name}.bak"

# Comando SQL para executar o backup
sql_cmd_command = (
    f"sqlcmd -S {SERVER_NAME} -d {DATABASE_NAME} -E -Q "
    f"\"BACKUP DATABASE [{DATABASE_NAME}] TO DISK='{BACKUP_PATH}\\{backup_file_name}'\""
)

# Executar comando SQL
try:
    subprocess.run(sql_cmd_command, shell=True, check=True)
    backup_status = "sucesso"
except subprocess.CalledProcessError as e:
    print(f"Erro ao executar o comando SQL: {e}")
    backup_status = "erro"

# Se o backup foi bem-sucedido, continuar com a compactação
if backup_status == "sucesso":
    # Caminho do arquivo de backup e arquivo compactado
    backup_file_path = os.path.join(BACKUP_PATH, backup_file_name)
    compressed_file_path = os.path.join(BACKUP_PATH, f"{day_name}.7z")

    # Compactar o arquivo de backup
    try:
        subprocess.run([SEVEN_ZIP_PATH, 'a', '-t7z', compressed_file_path, backup_file_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Erro ao compactar o arquivo: {e}")
        backup_status = "erro"

    # Copiar arquivo compactado para o destino
    if backup_status == "sucesso":
        try:
            os.system(f'copy "{compressed_file_path}" "{DESTINATION_PATH}"')
        except Exception as e:
            print(f"Erro ao copiar o arquivo: {e}")
            backup_status = "erro"

    # Deletar arquivo de backup
    if backup_status == "sucesso":
        try:
            os.remove(backup_file_path)
        except Exception as e:
            print(f"Erro ao deletar o arquivo de backup: {e}")

# Enviar mensagem para o Microsoft Teams
if backup_status == "sucesso":
    message = f"O backup de {DATABASE_NAME} foi realizado com sucesso e salvo em {compressed_file_path} e copiado para {DESTINATION_PATH}."
else:
    message = f"Ocorreu um erro ao realizar o backup de {DATABASE_NAME}."

send_teams_message(message)

# Importar tarefa do modelo
import_task(TEMPLATE_TASK_PATH)
