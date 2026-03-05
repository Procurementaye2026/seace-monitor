import requests, pandas as pd, smtplib, os, json
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==== CONFIGURACION (variables de entorno) ====
CORREO_REMITENTE     = os.getenv("CORREO_REMITENTE", "")
CORREO_CONTRASENA    = os.getenv("CORREO_CONTRASENA", "")
CORREO_DESTINO       = os.getenv("CORREO_DESTINO", "")
REGION_ACTIVA        = os.getenv("REGION_ACTIVA", "LIMA")
MONTO_MINIMO         = int(os.getenv("MONTO_MINIMO", "0"))
ENTIDAD_ACTIVA       = os.getenv("ENTIDAD_ACTIVA", "TODAS")
TIPO_PROCESO_ACTIVO  = os.getenv("TIPO_PROCESO_ACTIVO", "TODAS")

REGIONES = {
    "LIMA": "15", "CALLAO": "07", "AREQUIPA": "04",
    "CUSCO": "08", "LA LIBERTAD": "13", "PIURA": "20", "TODAS": ""
}

RUBROS = {
    "Tecnologia": [
        "laptop", "computadora", "pc", "impresora", "monitor", "teclado", "mouse",
        "disco duro", "memoria ram", "tablet", "proyector", "servidor", "ups",
        "switch", "router", "scanner", "toner", "cpu", "equipo informatico",
        "equipo de computo", "equipo de computo", "equipo tecnologico",
        "hardware", "componente", "suministro informatico"
    ],
    "Limpieza": [
        "limpieza", "desinfeccion", "aseo", "utiles de limpieza",
        "mantenimiento de limpieza", "higiene", "saneamiento",
        "fumigacion", "desratizacion", "servicios de limpieza",
        "limpieza y mantenimiento", "servicio de limpieza"
    ],
    "Computo y TI": [
        "mantenimiento de computo", "soporte tecnico", "mantenimiento informatico",
        "reparacion de equipos", "instalacion de software", "redes", "cableado",
        "mantenimiento preventivo", "servicio informatico", "sistema", "software",
        "infraestructura ti", "soporte de sistemas", "administracion de redes",
        "servicios de tecnologia", "tecnologia de la informacion", "ti "
    ],
    "Ferreteria": [
        "ferreteria", "herramientas", "materiales de construccion", "pintura",
        "tuberias", "cables electricos", "candado", "cerradura", "tornillos",
        "taladro", "llave", "soldadura", "pegamento", "sellador",
        "material ferretero", "equipos de seguridad industrial",
        "suministros de ferreteria", "gasfiteria", "plomeria"
    ]
}

ESTADOS_VALIDOS = [
    "CONVOCADO", "EN CONVOCATORIA", "ABIERTO", "VIGENTE",
    "REGISTRO DE PARTICIPANTES", "EN PROCESO", "PUBLICADO", "ACTIVO",
    "PRESENTACION DE OFERTAS"
]

# ==== HELPERS ====
def limpiar_monto(valor):
    if valor in [None, "N/A", "", "S/N"]: return 0
    try:
        return float(str(valor).replace(",", "").replace("S/.", "").strip())
    except:
        return 0

def parsear_fecha(fecha_str):
    if not fecha_str or fecha_str in ["N/A", ""]: return None
    for fmt in ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"]:
        try:
            return datetime.strptime(str(fecha_str).strip(), fmt)
        except:
            pass
    return None

def calcular_dias_restantes(fecha_str):
    fecha = parsear_fecha(fecha_str)
    if not fecha: return "N/A"
    dias = (fecha - datetime.now()).days
    if dias < 0: return "VENCIDO"
    if dias == 0: return "HOY"
    return f"1 dia" if dias == 1 else f"{dias} dias"

def aplicar_filtros(item):
    # Filtro por entidad
    if ENTIDAD_ACTIVA != "TODAS":
        entidad = str(item.get("nombreEntidad", item.get("entidad", ""))).upper()
        if ENTIDAD_ACTIVA not in entidad:
            return False
    # Filtro por monto
    monto = limpiar_monto(item.get("valorReferencial", item.get("vrVeCuantia", 0)))
    if monto < MONTO_MINIMO and monto != 0:
        return False
    # Filtro por tipo de proceso
    if TIPO_PROCESO_ACTIVO != "TODAS":
        tipo = str(item.get("tipoProceso", item.get("nomTipoProceso", ""))).upper()
        if TIPO_PROCESO_ACTIVO not in tipo:
            return False
    return True

def buscar_texto_en_item(item):
    """Concatena TODOS los campos de texto del item para busqueda robusta"""
    texto_total = []
    for key, value in item.items():
        if isinstance(value, str):
            texto_total.append(value.lower())
    return " ".join(texto_total)

def extraer_campo(item, *campos_posibles):
    """Busca el primer campo disponible entre los posibles nombres"""
    for campo in campos_posibles:
        val = item.get(campo)
        if val and str(val).strip() not in ["", "N/A", "None"]:
            return str(val)
    return "N/A"

# ==== BUSQUEDA SEACE ====
def buscar_en_seace():
    fecha_inicio = (datetime.now() - timedelta(days=30)).strftime("%d/%m/%Y")
    fecha_fin = datetime.now().strftime("%d/%m/%Y")
    print(f"\n{'='*55}")
    print(f"   MONITOR SEACE - {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"{'='*55}")
    print(f"   Rango de fechas : {fecha_inicio} al {fecha_fin}")
    print(f"   Region          : {REGION_ACTIVA}")
    print(f"   Monto minimo    : S/. {MONTO_MINIMO:,}")
    print(f"{'='*55}")

    resultados = {rubro: [] for rubro in RUBROS}
    resultados["Todos los Rubros"] = []

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": "https://prod4.seace.gob.pe",
        "Referer": "https://prod4.seace.gob.pe/openegocio/"
    }

    # ── FUENTE 1: API Oportunidades de Negocio (prod4) ──
    try:
        url_api = "https://prod4.seace.gob.pe/openegocio/api/v1/proceso/listar"
        payload = {
            "codigoDepartamento": REGIONES.get(REGION_ACTIVA, "15"),
            "pagina": 1,
            "cantidad": 500,
            "fechaInicio": fecha_inicio,
            "fechaFin": fecha_fin
        }
        resp = requests.post(url_api, json=payload, headers=headers, timeout=25)
        print(f"\n[API prod4] Status: {resp.status_code}")

        if resp.status_code == 200:
            data = resp.json()
            # Detectar la lista en la respuesta (prueba varias claves)
            items = (data.get("data") or data.get("lista") or
                     data.get("items") or data.get("result") or
                     data.get("content") or [])
            if isinstance(data, list):
                items = data

            print(f"[API prod4] {len(items)} procesos recibidos")

            # Mostrar campos del primer item para debug
            if items:
                print(f"[DEBUG] Campos disponibles: {list(items[0].keys())}")
                # Mostrar 2 ejemplos
                for i, it in enumerate(items[:2]):
                    print(f"[DEBUG] Ejemplo {i+1}: {json.dumps(it, ensure_ascii=False)[:300]}")

            for item in items:
                if not aplicar_filtros(item):
                    continue
                # Buscar en TODO el texto del item (robusto ante cambios de campos)
                texto_completo = buscar_texto_en_item(item)
                monto = limpiar_monto(extraer_campo(item, "valorReferencial", "vrVeCuantia", "monto", "valor"))
                fecha_cierre = extraer_campo(
                    item,
                    "fechaLimitePresentacion", "fechaFinRegParticipantes",
                    "fechaFinRegistro", "fecFinRegPart", "fechaCierre",
                    "fechaLimite", "fechaFin", "fechaPublicacion"
                )
                registro = {
                    "Entidad":       extraer_campo(item, "nombreEntidad", "entidad", "nomEntidad"),
                    "Descripcion":   extraer_campo(item, "sintesProceso", "descripcionObjeto",
                                                   "descripcion", "objeto", "detalle",
                                                   "objetoContratacion", "nombreProceso"),
                    "Tipo Proceso":  extraer_campo(item, "nomTipoProceso", "tipoProceso", "tipo"),
                    "Valor (S/.)":   f"S/. {monto:,.2f}" if monto > 0 else "N/A",
                    "Moneda":        extraer_campo(item, "moneda", "nomMoneda"),
                    "Fecha Inicio":  extraer_campo(item, "fechaConvocatoria", "fechaPublicacion",
                                                   "fecConvocatoria"),
                    "Fecha Cierre":  fecha_cierre,
                    "Dias Restantes": calcular_dias_restantes(fecha_cierre),
                    "Estado":        extraer_campo(item, "estadoProceso", "estado", "nomEstado"),
                    "Fuente":        "SEACE - OpenNegocio"
                }
                for rubro, palabras in RUBROS.items():
                    if any(p in texto_completo for p in palabras):
                        registro["Rubro"] = rubro
                        registro["Palabra Clave"] = next(p for p in palabras if p in texto_completo)
                        resultados[rubro].append(registro)
                        resultados["Todos los Rubros"].append(registro)
                        break

    except Exception as e:
        print(f"[API prod4] Error: {e}")

    # ── FUENTE 2: Buscador Publico SEACE (prod2) ──
    try:
        for palabra in ["computadora", "laptop", "limpieza", "ferreteria",
                        "soporte tecnico", "mantenimiento", "impresora",
                        "herramientas", "software", "servidor"]:
            url_busq = "https://prod2.seace.gob.pe/seacebus-uiwd-pub/buscadorPublico/buscadorPublico.xhtml"
            payload_busq = {
                "javax.faces.partial.ajax": "true",
                "javax.faces.partial.execute": "@all",
                "javax.faces.partial.render": "@all",
                "frmBuscarProceso:btnBuscar": "frmBuscarProceso:btnBuscar",
                "frmBuscarProceso": "frmBuscarProceso",
                "frmBuscarProceso:txtDescripcion": palabra,
                "frmBuscarProceso:ddlAnio": str(datetime.now().year),
                "frmBuscarProceso:ddlDpto": "15"
            }
            r2 = requests.post(
                url_busq, data=payload_busq,
                headers={"User-Agent": "Mozilla/5.0",
                         "Content-Type": "application/x-www-form-urlencoded"},
                timeout=15
            )
            if r2.status_code == 200:
                print(f"[BuscadorPublico] '{palabra}' -> OK ({len(r2.text)} bytes)")
    except Exception as e:
        print(f"[BuscadorPublico] Error: {e}")

    # ── RESUMEN ──
    print(f"\n{'='*55}")
    print("   RESUMEN POR RUBRO")
    print(f"{'='*55}")
    for rubro, lista in resultados.items():
        if rubro != "Todos los Rubros":
            icon = {"Tecnologia":"💻","Limpieza":"🧹","Computo y TI":"🖥️","Ferreteria":"🔧"}.get(rubro,"📁")
            print(f"   {icon} {rubro:<20}: {len(lista)} oportunidades")
    total = len(resultados["Todos los Rubros"])
    print(f"   {'📊 TOTAL':<22}: {total} oportunidades")
    print(f"{'='*55}")
    return resultados

# ==== EXCEL ====
def aplicar_estilo_excel(ws, color_header):
    fill   = PatternFill("solid", fgColor=color_header)
    font   = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    for cell in ws[1]:
        cell.fill = fill; cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="center")
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 20

def guardar_excel_pestanas(resultados):
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    ruta = os.path.join("/tmp", f"SEACE_Oportunidades_{fecha_hoy}.xlsx")
    colores = {
        "Todos los Rubros": "1a73e8",
        "Tecnologia":       "0f9d58",
        "Limpieza":         "f4b400",
        "Computo y TI":     "4285f4",
        "Ferreteria":       "db4437"
    }

    def ordenar_df(df):
        if "Fecha Cierre" not in df.columns: return df
        def peso(v):
            p = parsear_fecha(str(v))
            return p if p else datetime(9999, 12, 31)
        df["_sort"] = df["Fecha Cierre"].apply(peso)
        return df.sort_values("_sort", ascending=True).drop(columns=["_sort"])

    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        # Hoja resumen
        resumen = []
        for rubro, lista in resultados.items():
            if rubro != "Todos los Rubros":
                resumen.append({
                    "Rubro":        rubro,
                    "Oportunidades": len(lista),
                    "Mayor Monto":   max([limpiar_monto(r.get("Valor (S/.)", 0)) for r in lista], default=0),
                    "Fecha Reporte": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "Region":        REGION_ACTIVA
                })
        pd.DataFrame(resumen).to_excel(writer, sheet_name="Resumen", index=False)

        # Hoja por rubro
        for rubro, lista in resultados.items():
            nombre_hoja = rubro[:31]
            if lista:
                df = pd.DataFrame(lista).drop_duplicates(subset=["Descripcion"])
                df = ordenar_df(df)
            else:
                df = pd.DataFrame([{"Mensaje": f"Sin oportunidades para {rubro}"}])
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)

    # Aplicar estilos
    wb = load_workbook(ruta)
    for sheet_name in wb.sheetnames:
        color = next((v for k, v in colores.items() if k in sheet_name), "1a73e8")
        aplicar_estilo_excel(wb[sheet_name], color)
    wb.save(ruta)
    print(f"\n[Excel] Guardado en: {ruta}")
    return ruta

# ==== CORREO ====
def enviar_correo(resultados, ruta_excel):
    total = len(resultados.get("Todos los Rubros", []))
    fecha = datetime.now().strftime("%d/%m/%Y")
    asunto = f"SEACE Lima - {total} Oportunidades ({fecha})"

    filas_resumen = ""
    for rubro, lista in resultados.items():
        if rubro != "Todos los Rubros":
            icon = {"Tecnologia":"💻","Limpieza":"🧹","Computo y TI":"🖥️","Ferreteria":"🔧"}.get(rubro,"📁")
            color_fondo = {"Tecnologia":"#e8f5e9","Limpieza":"#fff9c4","Computo y TI":"#e3f2fd","Ferreteria":"#fce4ec"}.get(rubro,"#f5f5f5")
            filas_resumen += f"""
            <tr style="background:{color_fondo}">
                <td style="padding:8px;font-weight:bold">{icon} {rubro}</td>
                <td style="padding:8px;text-align:center;font-size:18px;color:#1a73e8"><b>{len(lista)}</b></td>
            </tr>"""

    filas_top = ""
    for op in resultados.get("Todos los Rubros", [])[:8]:
        dias = op.get("Dias Restantes", "N/A")
        color_dias = "#c62828" if dias in ["HOY","VENCIDO"] else ("#f57f17" if "1" in str(dias) else "#2e7d32")
        filas_top += f"""
        <tr>
            <td style="padding:6px;border-bottom:1px solid #eee">{op.get('Entidad','')[:40]}</td>
            <td style="padding:6px;border-bottom:1px solid #eee">{op.get('Descripcion','')[:60]}...</td>
            <td style="padding:6px;border-bottom:1px solid #eee">{op.get('Valor (S/.)','')}</td>
            <td style="padding:6px;border-bottom:1px solid #eee;color:{color_dias};font-weight:bold">{dias}</td>
            <td style="padding:6px;border-bottom:1px solid #eee">{op.get('Rubro','')}</td>
        </tr>"""

    html = f"""
    <html><body style="font-family:Arial,sans-serif;margin:0;padding:20px;background:#f5f5f5">
    <div style="max-width:800px;margin:0 auto;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1)">
        <div style="background:linear-gradient(135deg,#1a73e8,#0d47a1);padding:30px;text-align:center;color:white">
            <h1 style="margin:0;font-size:24px">🏛️ Monitor SEACE Lima</h1>
            <p style="margin:10px 0 0">Reporte Diario - {fecha} | Ultimos 30 dias</p>
        </div>
        <div style="padding:20px">
            <h2 style="color:#1a73e8">📊 Resumen por Rubro</h2>
            <table style="width:100%;border-collapse:collapse">
                <tr style="background:#1a73e8;color:white">
                    <th style="padding:10px">Rubro</th>
                    <th style="padding:10px">Oportunidades</th>
                </tr>
                {filas_resumen}
                <tr style="background:#e8eaf6;font-weight:bold">
                    <td style="padding:10px">📊 TOTAL</td>
                    <td style="padding:10px;text-align:center;font-size:20px;color:#1a73e8">{total}</td>
                </tr>
            </table>
            <h2 style="color:#1a73e8;margin-top:30px">🔝 Top 8 Oportunidades (por cierre mas proximo)</h2>
            <table style="width:100%;border-collapse:collapse;font-size:13px">
                <tr style="background:#1a73e8;color:white">
                    <th style="padding:8px">Entidad</th>
                    <th style="padding:8px">Descripcion</th>
                    <th style="padding:8px">Monto</th>
                    <th style="padding:8px">Cierre</th>
                    <th style="padding:8px">Rubro</th>
                </tr>
                {filas_top if filas_top else '<tr><td colspan="5" style="text-align:center;padding:20px;color:#666">Sin oportunidades encontradas hoy</td></tr>'}
            </table>
            <p style="color:#666;font-size:12px;margin-top:20px">
                📎 El Excel adjunto contiene el detalle completo con todas las oportunidades ordenadas por fecha de cierre.<br>
                🔗 Ver SEACE: <a href="https://prod4.seace.gob.pe/openegocio/">prod4.seace.gob.pe/openegocio</a>
            </p>
        </div>
    </div>
    </body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"]    = CORREO_REMITENTE
    msg["To"]      = CORREO_DESTINO
    msg.attach(MIMEText(html, "html"))

    if ruta_excel and os.path.exists(ruta_excel):
        with open(ruta_excel, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename={os.path.basename(ruta_excel)}")
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(CORREO_REMITENTE, CORREO_CONTRASENA)
            server.sendmail(CORREO_REMITENTE, CORREO_DESTINO, msg.as_string())
        print(f"[Correo] Enviado exitosamente a {CORREO_DESTINO}")
    except Exception as e:
        print(f"[Correo] Error al enviar: {e}")

# ==== EJECUCION ====
if __name__ == "__main__":
    print("\n" + "="*55)
    print("   INICIANDO MONITOR SEACE")
    print("="*55)
    datos    = buscar_en_seace()
    archivo  = guardar_excel_pestanas(datos)
    enviar_correo(datos, archivo)
    print("\n[FIN] Ejecucion completada.")
