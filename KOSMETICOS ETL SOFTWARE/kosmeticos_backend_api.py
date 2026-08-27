import os
import re
import random
import shutil
from datetime import date, timedelta
import pandas as pd
from openpyxl import load_workbook
from fastapi import FastAPI, File, UploadFile, Depends, HTTPException, status
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
import uvicorn

# ---------------------------------------------------------
# AUTENTICACIÓN Y CREDENCIALES (AGREGADO PARA LA NUBE)
# ---------------------------------------------------------
USER_CREDENTIALS = {
    "marketing": "kosmeticos2026",  # Usuario/Clave Equipo 1
    "devweb": "kosmeticos2026"     # Usuario/Clave Equipo 2
}

# Columnas mínimas que la Plantilla Madre debe traer
COLUMNAS_OBLIGATORIAS_MAIN = [
    "SKU_Padre", "SKU_Variante", "Marca", "Nombre_General",
    "Titulo_MercadoLibre", "Precio_Regular", "Stock",
]

# Configuración de carpetas
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.path.join(BASE_DIR, "base_de_datos")
WEB_DIR = os.path.join(BASE_DIR, "WEB")
PLANTILLAS_DIR = os.path.join(BASE_DIR, "Plantillas")
DICCIONARIOS_DIR = os.path.join(BASE_DIR, "diccionarios")

for folder in [DB_DIR, WEB_DIR, PLANTILLAS_DIR, DICCIONARIOS_DIR]:
    os.makedirs(folder, exist_ok=True)

app = FastAPI(title="Kosmeticos ETL API - Nube")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------
# ENDPOINT DE LOGIN
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
    return {"access_token": form_data.username, "token_type": "bearer"}

# ---------------------------------------------------------
# DICCIONARIO DE PRECIOS
# ---------------------------------------------------------

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
    col_estipulado = next(
        (c for c, v in encabezados.items() if "FALABELLA" in v and "RIPLEY" in v), None
    )

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
    diccionario_maestro = diccionario_maestro.drop_duplicates(subset=['SKU'], keep='last')
    return diccionario_maestro


def cruzar_main_con_diccionario(df_main):
    df_diccionario = obtener_datos_consolidados_diccionarios()

    if 'SKU_Variante' in df_main.columns:
        df_main['SKU_Variante'] = df_main['SKU_Variante'].astype(str).str.strip()

    if df_diccionario.empty:
        df_cruzado = df_main.copy()
    else:
        df_cruzado = pd.merge(
            df_main,
            df_diccionario,
            left_on='SKU_Variante',
            right_on='SKU',
            how='left'
        )

    if 'Precio_Estipulado' in df_cruzado.columns:
        df_cruzado['Precio_Base_Final'] = df_cruzado['Precio_Estipulado'].combine_first(
            df_cruzado.get('Precio_Regular')
        )
    else:
        df_cruzado['Precio_Base_Final'] = df_cruzado.get('Precio_Regular')

    if 'Precio_Web' in df_cruzado.columns:
        df_cruzado['Precio_Oferta_Final'] = df_cruzado['Precio_Web'].combine_first(
            df_cruzado.get('Precio_Oferta')
        )
    else:
        df_cruzado['Precio_Oferta_Final'] = df_cruzado.get('Precio_Oferta')

    return df_cruzado


def _rango_fecha_oferta_aleatorio():
    inicio = date.today() + timedelta(days=random.randint(0, 3))
    fin = inicio + timedelta(days=random.randint(60, 120))
    return inicio.isoformat(), fin.isoformat()

# ---------------------------------------------------------
# RUTAS DEL FRONTEND Y ARCHIVOS
# ---------------------------------------------------------
app.mount("/static", StaticFiles(directory=WEB_DIR), name="static")

@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    index_path = os.path.join(WEB_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>El archivo index.html no se encuentra en la carpeta /web</h1>"


@app.get("/api/marketing/download-template/")
async def download_template():
    nombre_archivo = "plantilla madre.xlsx"
    file_path = os.path.join(PLANTILLAS_DIR, nombre_archivo)
    if not os.path.exists(file_path):
        return {"error": "Plantilla no encontrada"}
    return FileResponse(path=file_path, filename=nombre_archivo)


@app.post("/api/marketing/upload-excel/")
async def upload_main_excel(archivo: UploadFile = File(...)):
    if not archivo.filename.lower().endswith((".xlsx", ".xls")):
        return JSONResponse(
            status_code=400,
            content={"error": "El archivo debe ser un Excel (.xlsx o .xls)."},
        )

    file_location = os.path.join(DB_DIR, archivo.filename)
    with open(file_location, "wb+") as file_object:
        shutil.copyfileobj(archivo.file, file_object)

    try:
        df = pd.read_excel(file_location, sheet_name=0)
        df.columns = df.columns.str.strip()
    except Exception as e:
        if os.path.exists(file_location):
            os.remove(file_location)
        return JSONResponse(
            status_code=400,
            content={"error": f"No se pudo leer el Excel: {e}"},
        )

    faltantes = [c for c in COLUMNAS_OBLIGATORIAS_MAIN if c not in df.columns]
    if faltantes:
        if os.path.exists(file_location):
            os.remove(file_location)
        return JSONResponse(
            status_code=400,
            content={
                "error": "El archivo no respeta la estructura de la Plantilla Madre.",
                "columnas_faltantes": faltantes,
            },
        )

    if df.dropna(how="all").empty:
        if os.path.exists(file_location):
            os.remove(file_location)
        return JSONResponse(
            status_code=400, content={"error": "El archivo no tiene filas de datos."}
        )

    return {
        "mensaje": "Archivo validado y guardado",
        "archivo": archivo.filename,
        "filas": int(len(df.dropna(how="all"))),
    }


@app.get("/api/dev/archivos-disponibles/")
async def listar_archivos():
    archivos = [f for f in os.listdir(DB_DIR) if f.endswith(('.xlsx', '.xls', '.csv'))]
    return {"archivos": archivos}


@app.post("/api/dev/upload-diccionario/")
async def upload_diccionario(archivo: UploadFile = File(...)):
    if not archivo.filename.lower().endswith((".xlsx", ".xls")):
        return JSONResponse(
            status_code=400,
            content={"error": "El archivo debe ser un Excel (.xlsx o .xls)."},
        )

    file_location = os.path.join(DICCIONARIOS_DIR, archivo.filename)
    with open(file_location, "wb+") as file_object:
        shutil.copyfileobj(archivo.file, file_object)

    try:
        df = _leer_un_diccionario(file_location)
    except Exception as e:
        if os.path.exists(file_location):
            os.remove(file_location)
        return JSONResponse(
            status_code=400,
            content={"error": f"No se pudo leer el Excel: {e}"},
        )

    if df.empty:
        if os.path.exists(file_location):
            os.remove(file_location)
        return JSONResponse(
            status_code=400,
            content={
                "error": "No se encontró una columna 'SKU' en el archivo, o no tiene filas de datos."
            },
        )

    for f in os.listdir(DICCIONARIOS_DIR):
        if f != archivo.filename:
            try:
                os.remove(os.path.join(DICCIONARIOS_DIR, f))
            except OSError:
                pass

    return {
        "mensaje": "Diccionario guardado y vinculado por SKU a la Plantilla Madre",
        "archivo": archivo.filename,
        "productos_leidos": int(len(df)),
    }


@app.get("/api/dev/diccionario-actual/")
async def diccionario_actual():
    archivos = [
        f
        for f in os.listdir(DICCIONARIOS_DIR)
        if f.endswith((".xlsx", ".xls")) and not f.startswith("~")
    ]
    if not archivos:
        return {"archivo": None}
    return {"archivo": archivos[0]}


@app.delete("/api/dev/eliminar-diccionario/")
async def eliminar_diccionario():
    eliminados = 0
    for f in os.listdir(DICCIONARIOS_DIR):
        if f.endswith((".xlsx", ".xls")) and not f.startswith("~"):
            os.remove(os.path.join(DICCIONARIOS_DIR, f))
            eliminados += 1
    return {"mensaje": "Diccionario eliminado", "archivos_eliminados": eliminados}

# ---------------------------------------------------------
# MERCADO LIBRE: Clasificación Bilingüe y Dinámica de Pestañas
# ---------------------------------------------------------

MAPEO_CATEGORIAS_ML = {
    "Cuidado Facial": [
        "crema", "cream", "moisturizer", "moisturising", "calming cream", "gel", "lotion", "locion",
        "serum", "sérum", "ampoule", "ampolla", "esencia", "essence", "pdrn", "exosomas", "treatment",
        "tónico", "tonico", "toner", "bruma", "mist", "pad", "pads", "discos",
        "exfoliante", "limpiador", "cleanser", "cleansing", "peeling", "gommage", "scrub", 
        "gel de limpieza", "espuma limpiadora", "desmaquillante", "aceite limpiador", "wash",
        "bloqueador", "protector solar", "sunscreen", "sun cream", "sun stick", "sunblock",
        "mascarilla", "sheet mask", "mask", "cuidado facial", "facial", "skincare"
    ],
    "Maquillaje": [
        "rubor", "rubores", "blush", "cheek",
        "delineador", "eyeliner", "pen liner", "pencil liner", "tinta de ojos",
        "sombra", "eyeshadow", "paleta de sombras", "shadow",
        "mascara de pestañas", "máscara de pestañas", "rimel", "mascara", "lash",
        "bb cream", "cc cream", "base de maquillaje", "foundation", "tone up", "cushion", "corrector", "concealer",
        "labial", "tint", "tinta de labios", "gloss", "bálsamo labial", "balsamo labial", "lip balm", "lipstick"
    ],
    "Cuidado Corporal": [
        "loción corporal", "crema corporal", "jabón corporal", "body wash", 
        "body lotion", "body cream", "body scrub", "corporal", "body"
    ],
    "Cuidado Capilar": [
        "shampoo", "champú", "acondicionador", "hair", "tratamiento capilar", 
        "mascarilla capilar", "aceite capilar", "scalp", "hair care"
    ],
    "Ambientadores y Difusores": [
        "ambientador", "ambientadores", "difusor", "difusores", 
        "home fragrance", "velas", "candle", "diffuser"
    ]
}

HOJA_ML_POR_DEFECTO = "Otros"
FILA_HEADER_ML = 3


def _fila_inicio_datos_ml(ws):
    max_row_combinada = FILA_HEADER_ML
    for rango in ws.merged_cells.ranges:
        if rango.max_row > max_row_combinada:
            max_row_combinada = rango.max_row
    return max_row_combinada + 1

MAPEO_ML = {
    "GTIN_EAN": ["Código universal de producto"],
    "SKU_Variante": ["SKU"],
    "Stock": ["Stock"],
    "Marca": ["Marca"],
    "Imagenes_URL": ["Fotos"],
    "Descripcion_Parrafo": ["Descripción"],
    "Ancho_cm": ["Ancho (cm)"],
    "Alto_cm": ["Alto (cm)"],
    "Largo_cm": ["Profundidad (cm)"],
    "Apto_Para_Piel": ["Tipo de piel"],
    "Vegano": ["Vegano"],
    "Cruelty_Free": ["Libre de crueldad"],
    "Ingredientes_Completos_INCI": ["Ingredientes"],
}


def _mapear_a_hoja_existente(nombre_categoria_detectada, lista_hojas):
    for hoja in lista_hojas:
        if nombre_categoria_detectada.lower() in hoja.lower() or hoja.lower() in nombre_categoria_detectada.lower():
            return hoja
            
    palabras_clave_hoja = {
        "Cuidado Facial": ["facial", "rostro", "skincare", "piel"],
        "Maquillaje": ["maquillaje", "makeup"],
        "Cuidado Corporal": ["corporal", "cuerpo", "body"],
        "Cuidado Capilar": ["capilar", "cabello", "hair"]
    }
    
    keywords_familia = palabras_clave_hoja.get(nombre_categoria_detectada, [])
    for hoja in lista_hojas:
        if any(kw in hoja.lower() for kw in keywords_familia):
            return hoja

    for hoja in lista_hojas:
        if "otro" in hoja.lower():
            return hoja

    return lista_hojas[0]


def _elegir_hoja_ml(fila, hojas_disponibles):
    tipo_prod = str(fila.get('Tipo_Producto', '')).lower()
    cat_orig = str(fila.get('Categorias', '')).lower()
    titulo = str(fila.get('Titulo_MercadoLibre', '')).lower()
    nombre_gen = str(fila.get('Nombre_General', '')).lower()

    texto_completo = f"{tipo_prod} {cat_orig} {titulo} {nombre_gen}"

    for cat_familia, keywords in MAPEO_CATEGORIAS_ML.items():
        if any(re.search(r'\b' + re.escape(kw) + r'\b', texto_completo) for kw in keywords):
            return _mapear_a_hoja_existente(cat_familia, hojas_disponibles)

    return _mapear_a_hoja_existente(HOJA_ML_POR_DEFECTO, hojas_disponibles)


def procesar_mercado_libre(df_datos, archivo):
    dir_ml = os.path.join(PLANTILLAS_DIR, "mercado libre")
    archivos_ml = [f for f in os.listdir(dir_ml) if f.endswith('.xlsx') and not f.startswith('~')]
    if not archivos_ml:
        return None, {"error": "No se encontró ninguna plantilla en la carpeta Plantillas/mercado libre"}

    ruta_plantilla = os.path.join(dir_ml, archivos_ml[0])
    wb = load_workbook(ruta_plantilla)
    
    hojas_categoria = list(wb.sheetnames)
    if not hojas_categoria:
        return None, {"error": "La plantilla de Mercado Libre no tiene hojas reconocibles"}

    filas_por_hoja = {}
    for _, fila in df_datos.iterrows():
        hoja = _elegir_hoja_ml(fila, hojas_categoria)
        filas_por_hoja.setdefault(hoja, []).append(fila)

    for hoja, filas in filas_por_hoja.items():
        ws = wb[hoja]

        col_indices = {}
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=FILA_HEADER_ML, column=col).value
            if not val:
                continue
            val = str(val).strip()
            if val.startswith("Título"):
                col_indices.setdefault("Titulo_MercadoLibre", col)
                continue
            if val == "Precio":
                col_indices.setdefault("Precio_Base_Final", col)
                continue
            for col_main, variantes in MAPEO_ML.items():
                if val in variantes:
                    col_indices.setdefault(col_main, col)

        fila_actual = _fila_inicio_datos_ml(ws)
        for fila in filas:
            for col_main, col_dest in col_indices.items():
                valor = fila.get(col_main) if col_main in fila else None
                if col_main == "Peso_g" and pd.notna(valor):
                    valor = round(float(valor) / 1000, 3)
                if valor is not None and not (isinstance(valor, float) and pd.isna(valor)):
                    ws.cell(row=fila_actual, column=col_dest).value = valor
            fila_actual += 1

    ruta_salida = os.path.join(DB_DIR, f"MercadoLibre_Listo_{archivo}")
    wb.save(ruta_salida)
    return ruta_salida, None

# ---------------------------------------------------------
# MOTOR DE INYECCIÓN INTELIGENTE
# ---------------------------------------------------------
@app.get("/api/dev/procesar-tienda/{tienda_id}")
async def procesar_tienda(tienda_id: str, archivo: str):
    file_path = os.path.join(DB_DIR, archivo)
    if not os.path.exists(file_path):
        return {"error": "El archivo no existe en la BD"}
    
    try:
        df_main = pd.read_excel(file_path)
        df_datos = cruzar_main_con_diccionario(df_main)
    except Exception as e:
        return {"error": f"Error leyendo el archivo Main: {e}"}

    ruta_salida = ""

    try:
        # MERCADO LIBRE
        if tienda_id in ("ml", "mercadolibre"):
            ruta_salida, error = procesar_mercado_libre(df_datos, archivo)
            if error:
                return error

        # FALABELLA
        elif tienda_id == "falabella":
            dir_falabella = os.path.join(PLANTILLAS_DIR, "falabella")
            archivos_fala = [f for f in os.listdir(dir_falabella) if f.endswith('.xlsx') and not f.startswith('~')]
            if not archivos_fala:
                return {"error": "No se encontró ninguna plantilla en la carpeta Plantillas/falabella"}
            
            ruta_plantilla = os.path.join(dir_falabella, archivos_fala[0])
            wb = load_workbook(ruta_plantilla)
            ws = wb['Subir plantilla']
            
            mapeo = {
                'Nombre_General': 'Nombre #39',
                'Marca': 'Marca #26',
                'Descripcion_Parrafo': 'Descripción #53',
                'SKU_Variante': 'SKU del vendedor #29',
                'GTIN_EAN': 'Código de barras #56',
                'Stock': 'QuantityFalabella #25',
                'Ingredientes_Completos_INCI': 'Ingredientes #36104',
                'Sugerencia_de_Uso': 'InstruccionesDeUso #36103',
                'Apto_Para_Piel': 'TipoDePiel #391',
                'Contenido': 'MedidaVolumen #405',
                'Peso_g': 'Peso del paquete #8',
                'Largo_cm': 'Largo del paquete #33',
                'Ancho_cm': 'Ancho del paquete #60',
                'Alto_cm': 'Alto del paquete #47',
                'SKU_Padre': 'SKU Padre #3',
                'Precio_Base_Final': 'PriceFalabella #52',
                'Precio_Oferta_Final': 'SalePriceFalabella #18',
            }
            if 'ISP' in df_datos.columns:
                mapeo['ISP'] = 'ResolucionIsp #402'

            col_indices = {}
            col_fecha_inicio_oferta = None
            col_fecha_fin_oferta = None
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row=4, column=col).value
                if val:
                    val = str(val).strip()
                    if val == 'SaleStartDateFalabella #45':
                        col_fecha_inicio_oferta = col
                        continue
                    if val == 'SaleEndDateFalabella #31':
                        col_fecha_fin_oferta = col
                        continue
                    for col_main, col_fala in mapeo.items():
                        if val == col_fala:
                            col_indices[col_main] = col

            fecha_inicio_oferta, fecha_fin_oferta = _rango_fecha_oferta_aleatorio()

            fila_inicio = 5
            for _, fila in df_datos.iterrows():
                for col_main, col_dest in col_indices.items():
                    if col_main in fila and not pd.isna(fila[col_main]):
                        ws.cell(row=fila_inicio, column=col_dest).value = fila[col_main]
                tiene_oferta = 'Precio_Oferta_Final' in fila and pd.notna(fila['Precio_Oferta_Final'])
                if tiene_oferta:
                    if col_fecha_inicio_oferta:
                        ws.cell(row=fila_inicio, column=col_fecha_inicio_oferta).value = fecha_inicio_oferta
                    if col_fecha_fin_oferta:
                        ws.cell(row=fila_inicio, column=col_fecha_fin_oferta).value = fecha_fin_oferta
                fila_inicio += 1
                
            ruta_salida = os.path.join(DB_DIR, f"Falabella_Listo_{archivo}")
            wb.save(ruta_salida)

        # RIPLEY
        elif tienda_id == "ripley":
            dir_ripley = os.path.join(PLANTILLAS_DIR, "ripley")
            archivos_ripley = [f for f in os.listdir(dir_ripley) if f.endswith('.xlsx') and not f.startswith('~')]
            if not archivos_ripley:
                return {"error": "No se encontró ninguna plantilla en la carpeta Plantillas/ripley"}

            ruta_plantilla = os.path.join(dir_ripley, archivos_ripley[0])
            wb = load_workbook(ruta_plantilla)
            ws = wb['Data'] 
            
            mapeo = {
                'SKU_Variante': ['sku_seller', 'SKU de oferta', 'ID de producto'],
                'Titulo_MercadoLibre': 'Titulo',
                'Descripcion_Parrafo': 'Descripcion',
                'Marca': 'Marca',
                'Tipo_Producto': 'Tipo de Producto',
                'Contenido': ['Contenido', 'Contenido Producto'],
                'Apto_Para_Piel': 'Tipo de piel',
                'Nombre_General': 'Nombre del Producto',
                'Ingredientes_Principales': 'Activo Cosmético',
                'Sugerencia_de_Uso': 'Modo de Empleo',
                'Ingredientes_Completos_INCI': 'Ingredientes',
                'Precio_Base_Final': 'Precio de la oferta',
                'Precio_Oferta_Final': 'Precio de descuento',
            }
            if 'ISP' in df_datos.columns:
                mapeo['ISP'] = 'N° de Registro ISP'
            
            col_indices = {}
            col_fecha_inicio_desc = None
            col_fecha_fin_desc = None
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row=1, column=col).value
                if val:
                    val = str(val).strip()
                    if val == 'Fecha de inicio del descuento':
                        col_fecha_inicio_desc = col
                        continue
                    if val == 'Fecha de finalización del descuento':
                        col_fecha_fin_desc = col
                        continue
                    for col_main, col_rips in mapeo.items():
                        if isinstance(col_rips, list):
                            if val in col_rips:
                                if col_main not in col_indices:
                                    col_indices[col_main] = []
                                col_indices[col_main].append(col)
                        else:
                            if val == col_rips:
                                if col_main not in col_indices:
                                    col_indices[col_main] = []
                                col_indices[col_main].append(col)

            fecha_inicio_desc, fecha_fin_desc = _rango_fecha_oferta_aleatorio()

            fila_inicio = 2
            for _, fila in df_datos.iterrows():
                for col_main, columnas_dest in col_indices.items():
                    if col_main in fila and not pd.isna(fila[col_main]):
                        for col_dest in columnas_dest:
                            ws.cell(row=fila_inicio, column=col_dest).value = fila[col_main]
                tiene_descuento = 'Precio_Oferta_Final' in fila and pd.notna(fila['Precio_Oferta_Final'])
                if tiene_descuento:
                    if col_fecha_inicio_desc:
                        ws.cell(row=fila_inicio, column=col_fecha_inicio_desc).value = fecha_inicio_desc
                    if col_fecha_fin_desc:
                        ws.cell(row=fila_inicio, column=col_fecha_fin_desc).value = fecha_fin_desc
                fila_inicio += 1
                
            ruta_salida = os.path.join(DB_DIR, f"Ripley_Listo_{archivo}")
            wb.save(ruta_salida)

        # WOOCOMMERCE
        elif tienda_id == "woocommerce":
            def col_o_vacio(nombre):
                return df_datos[nombre] if nombre in df_datos.columns else ""

            df_wp_plantilla = pd.DataFrame(columns=['SKU', 'Name', 'Short description', 'Description', 'Regular price'])
            df_wp_plantilla['SKU'] = col_o_vacio('SKU_Variante')
            df_wp_plantilla['Name'] = col_o_vacio('Nombre_General')
            df_wp_plantilla['Short description'] = col_o_vacio('Descripcion_BulletPoints')
            df_wp_plantilla['Description'] = col_o_vacio('Descripcion_Parrafo')
            df_wp_plantilla['Regular price'] = col_o_vacio('Precio_Oferta_Final')
            ruta_salida = os.path.join(DB_DIR, f"WooCommerce_Listo_{archivo.replace('.xlsx', '.csv')}")
            df_wp_plantilla.to_csv(ruta_salida, index=False)
            
        else:
            return {"error": "Lógica para esta tienda en construcción."}
            
    except Exception as e:
        return {"error": f"Error al procesar plantilla: {e}"}

    return FileResponse(path=ruta_salida, filename=os.path.basename(ruta_salida))

# ---------------------------------------------------------
# ALIAS /export/<canal>
# ---------------------------------------------------------
TIENDA_ID_POR_CANAL = {
    "mercadolibre": "ml",
    "falabella": "falabella",
    "ripley": "ripley",
    "woocommerce": "woocommerce",
}

for _canal, _tienda_id in TIENDA_ID_POR_CANAL.items():
    def _make_export(t_id):
        async def _export(archivo: str):
            return await procesar_tienda(t_id, archivo)
        return _export

    _handler = _make_export(_tienda_id)
    app.get(f"/export/{_canal}")(_handler)
    app.post(f"/export/{_canal}")(_handler)

# ---------------------------------------------------------
# EJECUCIÓN (Ajustado para puerto de servidor Cloud)
# ---------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("kosmeticos_backend_api:app", host="0.0.0.0", port=port, reload=False)