import os
import re
import random
import shutil
from datetime import date, timedelta, datetime, timezone
import pandas as pd
from openpyxl import load_workbook
from fastapi import FastAPI, File, UploadFile, Depends, HTTPException, status
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
import uvicorn
from jose import JWTError, jwt

# ---------------------------------------------------------
# RUTA BASE Y CARPETAS (DECLARAR ANTES DE INICIALIZAR LA APP)
# ---------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.path.join(BASE_DIR, "base_de_datos")
WEB_DIR = os.path.join(BASE_DIR, "WEB")
PLANTILLAS_DIR = os.path.join(BASE_DIR, "Plantillas")
DICCIONARIOS_DIR = os.path.join(BASE_DIR, "diccionarios")

for folder in [DB_DIR, WEB_DIR, PLANTILLAS_DIR, DICCIONARIOS_DIR]:
    os.makedirs(folder, exist_ok=True)

# Inicializar FastAPI
app = FastAPI(title="Kosmeticos ETL API - Nube")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------
# AUTENTICACIÓN Y CREDENCIALES
# ---------------------------------------------------------
USER_CREDENTIALS = {
    "kosmeticos": "kosmeticos2026"
}

SECRET_KEY = os.environ.get("JWT_SECRET_KEY", "kosmeticos_super_secret_jwt_key_2026_cloud")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 12

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/login")

def crear_access_token(data: dict, expires_delta: timedelta = None):
    to_encode = data.copy()
    if expires_delta:
        expire = datetime.now(timezone.utc) + expires_delta
    else:
        expire = datetime.now(timezone.utc) + timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

async def obtener_usuario_actual(token: str = Depends(oauth2_scheme)):
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="No se pudieron validar las credenciales de autenticación",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username: str = payload.get("sub")
        if username is None or username not in USER_CREDENTIALS:
            raise credentials_exception
    except JWTError:
        raise credentials_exception
    return username

# ---------------------------------------------------------
# RUTAS DE FRONTEND (DEFINIDAS EXPLÍCITAMENTE)
# ---------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
@app.get("/login", response_class=HTMLResponse)
async def serve_login():
    login_path = os.path.join(WEB_DIR, "login.html")
    if os.path.exists(login_path):
        return FileResponse(login_path)
    return HTMLResponse(content="<h2>Error: El archivo login.html no existe en la carpeta WEB en Render. Verifique el repositorio.</h2>", status_code=404)

@app.get("/dashboard", response_class=HTMLResponse)
async def serve_dashboard():
    dashboard_path = os.path.join(WEB_DIR, "index.html")
    if os.path.exists(dashboard_path):
        return FileResponse(dashboard_path)
    return HTMLResponse(content="<h2>Error: El archivo index.html no existe en la carpeta WEB.</h2>", status_code=404)

# Montar estáticos al final para no interferir con las rutas raíz
if os.path.exists(WEB_DIR):
    app.mount("/static", StaticFiles(directory=WEB_DIR), name="static")

# ---------------------------------------------------------
# ENDPOINTS API (LOGIN Y SERVICIOS)
# ---------------------------------------------------------
@app.post("/api/auth/login")
async def login(form_data: OAuth2PasswordRequestForm = Depends()):
    user_password = USER_CREDENTIALS.get(form_data.username)
    if not user_password or user_password != form_data.password:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Usuario o contraseña incorrectos",
            headers={"WWW-Authenticate": "Bearer"},
        )
    access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = crear_access_token(
        data={"sub": form_data.username}, expires_delta=access_token_expires
    )
    return {"access_token": access_token, "token_type": "bearer", "username": form_data.username}

@app.get("/api/auth/me")
async def read_users_me(current_user: str = Depends(obtener_usuario_actual)):
    return {"username": current_user, "status": "authenticated"}

# ---------------------------------------------------------
# PROCESAMIENTO Y DICCIONARIOS
# ---------------------------------------------------------
COLUMNAS_OBLIGATORIAS_MAIN = [
    "SKU_Padre", "SKU_Variante", "Marca", "Nombre_General",
    "Titulo_MercadoLibre", "Precio_Regular", "Stock",
]

def _normalizar_header(valor):
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip().upper()

def _leer_un_diccionario(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    fila_header = None
    encabezados = {}
    for r in range(1, min(ws.max_row, 10) + 1):
        valores = {c: _normalizar_header(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)}
        if "SKU" in valores.values():
            fila_header = r
            encabezados = valores
            break
    if fila_header is None:
        return pd.DataFrame()

    col_sku = next((c for c, v in encabezados.items() if v == "SKU"), None)
    col_isp = next((c for c, v in encabezados.items() if v == "ISP"), None)
    col_web = next((c for c, v in encabezados.items() if "WEB" in v and "PRICE" in v), None)
    col_estipulado = next((c for c, v in encabezados.items() if "FALABELLA" in v and "RIPLEY" in v), None)

    if col_sku is None:
        return pd.DataFrame()

    filas = []
    for r in range(fila_header + 1, ws.max_row + 1):
        sku = ws.cell(row=r, column=col_sku).value
        if sku is None or str(sku).strip() == "":
            continue
        filas.append({
            "SKU": str(sku).strip(),
            "ISP": ws.cell(row=r, column=col_isp).value if col_isp else None,
            "Precio_Web": ws.cell(row=r, column=col_web).value if col_web else None,
            "Precio_Estipulado": ws.cell(row=r, column=col_estipulado).value if col_estipulado else None,
        })
    return pd.DataFrame(filas)

def obtener_datos_consolidados_diccionarios():
    archivos = [f for f in os.listdir(DICCIONARIOS_DIR) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
    if not archivos:
        return pd.DataFrame()
    dfs = []
    for archivo in archivos:
        path = os.path.join(DICCIONARIOS_DIR, archivo)
        try:
            df = _leer_un_diccionario(path)
            if not df.empty:
                dfs.append(df)
        except Exception as e:
            print(f"Error leyendo el diccionario {archivo}: {e}")
    if not dfs:
        return pd.DataFrame()
    diccionario_maestro = pd.concat(dfs, ignore_index=True)
    return diccionario_maestro.drop_duplicates(subset=['SKU'], keep='last')

def cruzar_main_con_diccionario(df_main):
    df_diccionario = obtener_datos_consolidados_diccionarios()
    if 'SKU_Variante' in df_main.columns:
        df_main['SKU_Variante'] = df_main['SKU_Variante'].astype(str).str.strip()
    if df_diccionario.empty:
        df_cruzado = df_main.copy()
    else:
        df_cruzado = pd.merge(df_main, df_diccionario, left_on='SKU_Variante', right_on='SKU', how='left')

    if 'Precio_Estipulado' in df_cruzado.columns:
        df_cruzado['Precio_Base_Final'] = df_cruzado['Precio_Estipulado'].combine_first(df_cruzado.get('Precio_Regular'))
    else:
        df_cruzado['Precio_Base_Final'] = df_cruzado.get('Precio_Regular')

    if 'Precio_Web' in df_cruzado.columns:
        df_cruzado['Precio_Oferta_Final'] = df_cruzado['Precio_Web'].combine_first(df_cruzado.get('Precio_Oferta'))
    else:
        df_cruzado['Precio_Oferta_Final'] = df_cruzado.get('Precio_Oferta')

    return df_cruzado

# ---------------------------------------------------------
# INICIO DE SERVIDOR
# ---------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    module_name = os.path.splitext(os.path.basename(__file__))[0]
    uvicorn.run(f"{module_name}:app", host="0.0.0.0", port=port, reload=False)