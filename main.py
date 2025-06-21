from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uuid
import os
from datetime import datetime
from ncm_validator import NCMValidator
import time

# IMPORTS NECESSÁRIOS PARA O FIREBASE
import json
from google.cloud import firestore
from google.oauth2 import service_account

app = FastAPI(
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://www.smartcont.online",
        "https://smartcont.online",
        "http://localhost:8080",
        "http://localhost:5173"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configuração Firestore CORRETA PARA O RENDER
service_account_info = json.loads(os.environ['GOOGLE_APPLICATION_CREDENTIALS_JSON'])
credentials = service_account.Credentials.from_service_account_info(service_account_info)
db = firestore.Client(credentials=credentials, project=service_account_info['project_id'])

UPLOAD_DIR = "/tmp/uploads_ncm"
RESULT_DIR = "/tmp/results_ncm"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)

# --- Endpoint para enfileirar tarefa de NCM ---
@app.post("/ncm/enfileirar")
async def enfileirar_ncm(usuario: UploadFile = File(...), user_id: str = "anonimo"):
    task_id = str(uuid.uuid4())
    filename = f"{task_id}_{usuario.filename}"
    upload_path = os.path.join(UPLOAD_DIR, filename)
    with open(upload_path, "wb") as f:
        f.write(await usuario.read())

    # Cria registro no Firestore
    db.collection("ncm_tasks").document(task_id).set({
        "user_id": user_id,
        "filename": usuario.filename,
        "status": "aguardando",
        "created_at": datetime.utcnow().isoformat(),
        "updated_at": datetime.utcnow().isoformat(),
    })

    return {"task_id": task_id, "status": "aguardando"}

# --- Endpoint para consultar status ---
@app.get("/ncm/status/{task_id}")
def status_ncm(task_id: str):
    doc = db.collection("ncm_tasks").document(task_id).get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Tarefa não encontrada")
    return doc.to_dict()

# --- Endpoint para download do resultado ---
@app.get("/ncm/download/{task_id}")
def download_ncm(task_id: str):
    doc = db.collection("ncm_tasks").document(task_id).get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Tarefa não encontrada")
    data = doc.to_dict()
    if data.get("status") != "concluido" or not data.get("result_path"):
        raise HTTPException(status_code=400, detail="Arquivo ainda não está pronto")
    return FileResponse(data["result_path"], filename="ncm_corrigido.xlsx")

# --- Endpoint de saúde ---
@app.get("/")
def health():
    return {"status": "ok"}

# --- Seus outros endpoints (PIS/COFINS, etc) continuam iguais ---

def worker_ncm():
    """
    Worker para processar tarefas de validação de NCM em segundo plano.
    Busca tarefas com status 'aguardando' no Firestore, processa-as e
    atualiza o status para 'concluido' ou 'erro'.
    """
    openai_api_key = os.getenv("OPENAI_API_KEY")
    if not openai_api_key:
        print("ERRO: A variável de ambiente OPENAI_API_KEY não está definida.")
        # O ideal é logar isso e talvez sair ou esperar.
        # Por enquanto, vamos apenas imprimir e parar o worker para evitar cobranças.
        return

    validator = NCMValidator(api_key=openai_api_key)
    ncm_table_path = "Tabela_NCM_Vigente.csv"  # Assume que o arquivo está na raiz do projeto

    print("Worker NCM iniciado...")
    while True:
        task_id_processed = None  # Para garantir que o ID da tarefa esteja disponível no bloco de exceção
        try:
            # Busca por uma tarefa que está aguardando
            tasks_ref = db.collection("ncm_tasks")
            query = tasks_ref.where("status", "==", "aguardando").limit(1)
            tasks = query.stream()

            task_doc = next(tasks, None)

            if task_doc:
                task_id_processed = task_doc.id
                task_data = task_doc.to_dict()
                print(f"Processando tarefa: {task_id_processed}")

                # Atualiza o status para 'processando'
                tasks_ref.document(task_id_processed).update({
                    "status": "processando",
                    "updated_at": datetime.utcnow().isoformat()
                })

                # Define os caminhos dos arquivos
                filename = f"{task_id_processed}_{task_data['filename']}"
                upload_path = os.path.join(UPLOAD_DIR, filename)
                result_filename = f"resultado_{filename}"
                result_path = os.path.join(RESULT_DIR, result_filename)

                # Carrega o arquivo do usuário e a tabela NCM
                df_user, df_ncm = validator.load_excel_files(upload_path, ncm_table_path)

                # Processa os dados em lotes
                df_results = validator.process_in_batches(df_user)

                # Gera o relatório final
                validator.generate_report(df_results, df_ncm, result_path)

                # Atualiza o status para 'concluido' com o caminho do resultado
                tasks_ref.document(task_id_processed).update({
                    "status": "concluido",
                    "result_path": result_path,
                    "updated_at": datetime.utcnow().isoformat()
                })
                print(f"Tarefa {task_id_processed} concluída com sucesso.")

            else:
                # Se não houver tarefas, aguarda 10 segundos antes de verificar novamente
                time.sleep(10)

        except Exception as e:
            print(f"Ocorreu um erro no worker: {e}")
            if task_id_processed:
                tasks_ref.document(task_id_processed).update({
                    "status": "erro",
                    "error_message": str(e),
                    "updated_at": datetime.utcnow().isoformat()
                })
            # Aguarda um tempo maior após um erro para evitar loops de falha rápidos
            time.sleep(30)
