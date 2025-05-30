from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from google.cloud import firestore
import redis
import uuid
import os
from datetime import datetime
from ncm_validator import NCMValidator

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

# Configuração Redis Upstash
redis_url = "rediss://default:AWuIAAIjcDFlMTE2NzdhNjhjYzI0ZTFmYmM1OGE3NzUzYzc4ZWYxN3AxMA@neutral-snake-27528.upstash.io:6379"
r = redis.from_url(redis_url)

# Configuração Firestore
db = firestore.Client()  # Certifique-se de ter as credenciais do Firebase configuradas

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

    # Adiciona na fila do Redis
    r.rpush("fila_ncm", task_id)
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

# --- Worker para processar a fila (roda em processo separado) ---
def worker_ncm():
    while True:
        task_id = r.lpop("fila_ncm")
        if not task_id:
            import time; time.sleep(2)
            continue
        task_id = task_id.decode()
        doc_ref = db.collection("ncm_tasks").document(task_id)
        doc = doc_ref.get()
        if not doc.exists:
            continue
        data = doc.to_dict()
        doc_ref.update({"status": "processando", "updated_at": datetime.utcnow().isoformat()})
        try:
            # Localiza o arquivo de upload
            filename = f"{task_id}_{data['filename']}"
            upload_path = os.path.join(UPLOAD_DIR, filename)
            ncm_path = "Tabela_NCM_Vigente.csv"  # Ajuste conforme seu projeto
            output_path = os.path.join(RESULT_DIR, f"{task_id}_corrigido.xlsx")
            API_KEY = os.environ.get("OPENAI_API_KEY", "")
            validator = NCMValidator(API_KEY)
            df_user, df_ncm = validator.load_excel_files(upload_path, ncm_path)
            df_results = validator.process_in_batches(df_user)
            validator.generate_report(df_results, df_ncm, output_path)
            os.remove(upload_path)
            doc_ref.update({
                "status": "concluido",
                "updated_at": datetime.utcnow().isoformat(),
                "result_path": output_path
            })
        except Exception as e:
            doc_ref.update({
                "status": "erro",
                "updated_at": datetime.utcnow().isoformat(),
                "error_message": str(e)
            })

# --- Endpoint de saúde ---
@app.get("/")
def health():
    return {"status": "ok"}

# --- Seus outros endpoints (PIS/COFINS, etc) continuam iguais ---
