from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import re
from datetime import datetime
import os
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
        "http://localhost:5173"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Endpoint PIS/COFINS ---
@app.post("/processar-pis-cofins/")
async def processar_pis_cofins(usuario: UploadFile = File(...)):
    try:
        tabela_df = pd.read_csv('SUBIR_P_DOMINIO.csv', sep=';', dtype=str, quotechar='"', encoding='utf-8')
        tabela_df.columns = [col.replace('\ufeff', '') for col in tabela_df.columns]
        TABELA_DOMINIO = {}
        for _, row in tabela_df.iterrows():
            ncm = str(row['NCM']).zfill(8)
            valor = f"{row['descricao']}"
            TABELA_DOMINIO[ncm] = valor

        user_path = f"/tmp/{usuario.filename}"
        with open(user_path, "wb") as f:
            f.write(await usuario.read())
        usuario_df = pd.read_excel(user_path, dtype=str)
        df = usuario_df.copy()
        df.columns = [col.strip() for col in df.columns]
        df['NCM'] = df['NCM'].astype(str).str.zfill(8)
        colunas_esperadas = ['Número Nota', 'Cliente', 'CNPJ/CPF', 'Descricao', 'NCM', 'Vlr. Cont.', 'Valor ICMS']
        for col in colunas_esperadas:
            if col not in df.columns:
                raise Exception(f"Coluna obrigatória não encontrada: {col}")
        df['ncm_formatado'] = df['NCM'].apply(lambda ncm: re.sub(r'\D', '', str(ncm)).zfill(8))
        df['tributacao_full'] = df['ncm_formatado'].map(TABELA_DOMINIO).fillna('')
        trib_split = df['tributacao_full'].str.split(';', n=2, expand=True)
        df['Tributação 1'] = trib_split[0].str.strip() if 0 in trib_split else ''
        df['Tributação 2'] = trib_split[1].str.strip() if 1 in trib_split else ''
        df['Tributação 3'] = trib_split[2].str.strip() if 2 in trib_split else ''
        df['Vlr. Cont.'] = pd.to_numeric(df['Vlr. Cont.'], errors='coerce').fillna(0)
        df = df.sort_values(['Tributação 1', 'Tributação 2', 'Tributação 3'])
        linhas = []
        for keys, grupo in df.groupby(['Tributação 1', 'Tributação 2', 'Tributação 3'], sort=False):
            for _, row in grupo.iterrows():
                linhas.append(row.to_dict())
            total = grupo['Vlr. Cont.'].sum()
            linha_total = {col: '' for col in df.columns}
            linha_total['Tributação 1'] = keys[0]
            linha_total['Tributação 2'] = keys[1]
            linha_total['Tributação 3'] = keys[2]
            linha_total['Vlr. Cont.'] = f"R$ {total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            linha_total['Descricao'] = 'TOTAL DO GRUPO'
            linhas.append(linha_total)
        colunas_finais = [
            'Número Nota', 'Cliente', 'CNPJ/CPF', 'Descricao', 'NCM', 'Vlr. Cont.', 'Valor ICMS',
            'Tributação', 'Cst de Entrada', 'Cst de Saida'
        ]
        df_organizado = pd.DataFrame(linhas, columns=df.columns)
        df_organizado = df_organizado.rename(columns={
            'Tributação 1': 'Tributação',
            'Tributação 2': 'Cst de Entrada',
            'Tributação 3': 'Cst de Saida'
        })
        for linha in linhas:
            if linha.get('Descricao') == 'TOTAL DO GRUPO':
                for col in colunas_finais:
                    if col not in ['Tributação', 'Cst de Entrada', 'Cst de Saida', 'Vlr. Cont.', 'Descricao']:
                        linha[col] = ''
        df_organizado = df_organizado[colunas_finais]
        resumo = df.groupby(['Tributação 1', 'Tributação 2', 'Tributação 3'], as_index=False)['Vlr. Cont.'].sum()
        resumo['Vlr. Cont.'] = resumo['Vlr. Cont.'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        df_organizado['__ordem__'] = df_organizado['Tributação'].apply(lambda x: 0 if str(x).strip().lower().startswith('tributado') else 1)
        df_organizado = df_organizado.sort_values(['__ordem__', 'Tributação']).drop(columns='__ordem__')

        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"/tmp/relatorio_pis_cofins_{now}.xlsx"
        with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
            df_organizado.to_excel(writer, sheet_name='Organizado', index=False)
            resumo.to_excel(writer, sheet_name='Resumo', index=False)
        os.remove(user_path)
        return FileResponse(nome_arquivo, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=nome_arquivo.split('/')[-1])
    except Exception as e:
        return JSONResponse(status_code=400, content={"erro": str(e)})

# --- Endpoint NCM ---
@app.post("/processar-ncm/")
async def processar_ncm(usuario: UploadFile = File(...)):
    try:
        user_path = f"/tmp/{usuario.filename}"
        with open(user_path, "wb") as f:
            f.write(await usuario.read())
        ncm_path = "/app/entrada/Tabela_NCM_Vigente.xlsx"  # Caminho fixo
        output_path = f"/tmp/relatorio_ncm_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        API_KEY = os.environ.get("OPENAI_API_KEY", "")
        validator = NCMValidator(API_KEY)
        df_user, df_ncm = validator.load_excel_files(user_path, ncm_path)
        df_results = validator.process_in_batches(df_user)
        validator.generate_report(df_results, df_ncm, output_path)
        os.remove(user_path)
        return FileResponse(output_path, filename=os.path.basename(output_path), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return JSONResponse(status_code=400, content={"erro": str(e)})

@app.get("/")
def health():
    return {"status": "ok"} 