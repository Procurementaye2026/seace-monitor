import requests
import pandas as pd
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

CORREO_REMITENTE    = os.getenv("CORREO_REMITENTE",    "")
CORREO_CONTRASENA   = os.getenv("CORREO_CONTRASENA",   "")
CORREO_DESTINO      = os.getenv("CORREO_DESTINO",      "")
REGION_ACTIVA       = os.getenv("REGION_ACTIVA",       "LIMA")
MONTO_MINIMO        = int(os.getenv("MONTO_MINIMO",    "0"))
ENTIDAD_ACTIVA      = os.getenv("ENTIDAD_ACTIVA",      "TODAS")
TIPO_PROCESO_ACTIVO = os.getenv("TIPO_PROCESO_ACTIVO", "TODAS")

REGIONES = {
    "LIMA": "15", "CALLAO": "07", "AREQUIPA": "04",
    "CUSCO": "08", "LA LIBERTAD": "13", "PIURA": "20", "TODAS": ""
}

RUBROS = {
    "Tecnologia": [
        "laptop", "computadora", "computador", "pc", "impresora",
        "monitor", "teclado", "mouse", "disco duro", "memoria ram",
        "tablet", "proyector", "servidor", "ups", "switch",
        "router", "scanner", "toner", "cartucho", "cpu",
        "equipo informatico", "equipo de computo"
    ],
    "Limpieza": [
        "servicio de limpieza", "limpieza", "desinfeccion",
        "aseo", "utiles de limpieza", "mantenimiento de limpieza",
        "higiene", "saneamiento", "fumigacion", "desratizacion"
    ],
    "Computo y TI": [
        "mantenimiento de computo", "soporte tecnico",
        "mantenimiento informatico", "reparacion de equipos",
        "instalacion de software", "redes", "cableado",
        "mantenimiento preventivo", "servicio informatico",
        "sistema", "software", "hardware", "infraestructura ti"
    ],
    "Ferreteria": [
        "ferreteria", "herramientas", "materiales de construccion",
        "pintura", "tuberias", "cables electricos", "candado",
        "cerradura", "tornillos", "taladro", "llave", "soldadura",
        "pegamento", "sellador", "material ferretero",
        "equipos de seguridad industrial"
    ]
}

def limpiar_monto(valor):
    try:
        if valor in [None, "N/A", "", "S/N"]:
            return 0
        return float(str(valor).replace(",", "").replace("S/.", "").strip())
    except:
        return 0

def aplicar_filtros(item):
    if ENTIDAD_ACTIVA != "TODAS":
        if ENTIDAD_ACTIVA not in str(item.get("entidad", "")).upper():
            return False
    monto = limpiar_monto(item.get("valorReferencial", 0))
    if monto < MONTO_MINIMO and monto != 0:
        return False
    if TIPO_PROCESO_ACTIVO != "TODAS":
        if TIPO_PROCESO_ACTIVO not in str(item.get("tipoProceso", "")).upper():
            return False
    return True

def buscar_en_seace():
    print(f"\nBuscando... {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    todos_resultados = {rubro: [] for rubro in RUBROS.keys()}
    todos_resultados["Todos los Rubros"] = []

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "application/json, text/plain, */*",
    }

    try:
        print(f"Consultando SEACE - Region: {REGION_ACTIVA}...")
        url = "https://prod4.seace.gob.pe/openegocio/api/v1/proceso/listar"
        payload = {
            "descripcion": "",
            "codigoDepartamento": REGIONES.get(REGION_ACTIVA, "15"),
            "pagina": 1,
            "cantidad": 500
        }
        resp = requests.post(url, json=payload, headers=headers, timeout=20)

        if resp.status_code == 200:
            data = resp.json()
            items = data.get("data", data.get("lista",
                    data.get("items", data.get("result", []))))

            if isinstance(items, list):
                print(f"  {len(items)} procesos encontrados")
                for item in items:
                    if not aplicar_filtros(item):
                        continue
                    descripcion = str(
                        item.get("descripcionObjeto", "") or
                        item.get("descripcion", "") or
                        item.get("objeto", "")
                    ).lower()
                    monto = limpiar_monto(item.get("valorReferencial", 0))
                    registro = {
                        "Entidad": item.get("entidad",
                                   item.get("nombreEntidad", "N/A")),
                        "Descripcion": item.get("descripcionObjeto",
                                       item.get("descripcion", "N/A")),
                        "Valor S/.": f"S/. {monto:,.2f}" if monto > 0 else "N/A",
                        "Fecha": item.get("fechaConvocatoria",
                                 item.get("fecha", "N/A")),
                        "Region": REGION_ACTIVA,
                        "Tipo Proceso": item.get("tipoProceso",
                                        item.get("tipo", "N/A")),
                        "Estado": item.get("estadoProceso",
                                  item.get("estado", "N/A")),
                        "Fuente": "SEACE - Oportunidades de Negocio"
                    }
                    encontrado = False
                    for rubro, palabras in RUBROS.items():
                        for palabra in palabras:
                            if palabra.lower() in descripcion:
                                registro["Palabra Clave"] = palabra
                                registro["Rubro"] = rubro
                                todos_resultados[rubro].append(registro)
                                todos_resultados["Todos los Rubros"].append(registro)
                                encontrado = True
                                break
                        if encontrado:
                            break
        else:
            print(f"  Codigo: {resp.status_code}")
    except Exception as e:
        print(f"  Error: {e}")

    print("\nRESUMEN POR RUBRO:")
    print("-" * 40)
    for rubro, lista in todos_resultados.items():
        if rubro != "Todos los Rubros":
            print(f"  {rubro}: {len(lista)} oportunidades")
    print(f"  TOTAL: {len(todos_resultados['Todos los Rubros'])} oportunidades")
    print("-" * 40)
    return todos_resultados

def aplicar_estilo_excel(ws, color_header):
    fill = PatternFill("solid", fgColor=color_header)
    fuente = Font(bold=True, color="FFFFFF", size=11)
    borde = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    for cell in ws[1]:
        cell.fill = fill
        cell.font = fuente
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borde
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = borde
            cell.alignment = Alignment(horizontal="left", vertical="center")
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 20

def guardar_excel_pestanas(resultados):
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    nombre_archivo = f"SEACE_Oportunidades_{fecha_hoy}.xlsx"
    ruta = os.path.join("/tmp", nombre_archivo)

    colores = {
        "Todos los Rubros": "1a73e8",
        "Tecnologia"      : "0f9d58",
        "Limpieza"        : "f4b400",
        "Computo y TI"    : "4285f4",
        "Ferreteria"      : "db4437"
    }

    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        resumen_data = []
        for rubro, lista in resultados.items():
            if rubro != "Todos los Rubros":
                resumen_data.append({
                    "Rubro"        : rubro,
                    "Oportunidades": len(lista),
                    "Mayor Monto"  : max(
                        [limpiar_monto(r.get("Valor S/.", 0))
                         for r in lista], default=0),
                    "Fecha Reporte": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "Region"       : REGION_ACTIVA,
                })
        pd.DataFrame(resumen_data).to_excel(
            writer, sheet_name="Resumen", index=False)

        for rubro, lista in resultados.items():
            nombre_hoja = rubro[:31]
            if lista:
                df = pd.DataFrame(lista)
                df = df.drop_duplicates(subset=["Descripcion"])
                df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            else:
                pd.DataFrame([{
                    "Mensaje": f"Sin oportunidades para {rubro} hoy"
                }]).to_excel(writer, sheet_name=nombre_hoja, index=False)

    wb = load_workbook(ruta)
    for nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
        color = colores.get(nombre_hoja, "1a73e8")
        aplicar_estilo_excel(ws, color)
    wb.save(ruta)
    print(f"\nExcel guardado: {ruta}")
    return ruta

def enviar_correo(resultados, ruta_excel):
    print("\nEnviando correo...")
    try:
        total = len(resultados.get("Todos los Rubros", []))
        msg = MIMEMultipart("alternative")
        msg["Subject"] = (
            f"SEACE - {total} Oportunidades "
            f"{datetime.now().strftime('%d/%m/%Y')}"
        )
        msg["From"] = CORREO_REMITENTE
        msg["To"] = CORREO_DESTINO

        filas_resumen = ""
        for rubro, lista in resultados.items():
            if rubro != "Todos los Rubros":
                filas_resumen += f"""
                <tr>
                    <td style='padding:8px;border:1px solid #ddd'>{rubro}</td>
                    <td style='padding:8px;border:1px solid #ddd;
                               text-align:center;font-weight:bold;
                               color:#1a73e8'>{len(lista)}</td>
                </tr>"""

        top_oportunidades = ""
        for op in resultados.get("Todos los Rubros", [])[:5]:
            top_oportunidades += f"""
            <tr>
                <td style='padding:8px;border:1px solid #ddd'>
                    {op.get('Entidad','N/A')}</td>
                <td style='padding:8px;border:1px solid #ddd'>
                    {str(op.get('Descripcion','N/A'))[:60]}...</td>
                <td style='padding:8px;border:1px solid #ddd'>
                    {op.get('Valor S/.','N/A')}</td>
                <td style='padding:8px;border:1px solid #ddd'>
                    {op.get('Rubro','N/A')}</td>
            </tr>"""

        cuerpo_html = f"""
        <html><body style='font-family:Arial;max-width:800px;margin:auto'>
            <div style='background:#1a73e8;padding:20px;border-radius:8px'>
                <h2 style='color:white;margin:0'>
                    Monitor SEACE - Reporte Diario</h2>
                <p style='color:white;margin:5px 0'>
                    Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')} |
                    Region: {REGION_ACTIVA} |
                    Monto minimo: S/. {MONTO_MINIMO:,}
                </p>
            </div><br>
            <h3>Resumen por Rubro</h3>
            <table style='border-collapse:collapse;width:100%'>
                <thead>
                    <tr style='background:#f8f9fa'>
                        <th style='padding:10px;border:1px solid #ddd;
                                   text-align:left'>Rubro</th>
                        <th style='padding:10px;border:1px solid #ddd;
                                   text-align:center'>Oportunidades</th>
                    </tr>
                </thead>
                <tbody>{filas_resumen}</tbody>
                <tfoot>
                    <tr style='background:#1a73e8;color:white;font-weight:bold'>
                        <td style='padding:10px;border:1px solid #ddd'>TOTAL</td>
                        <td style='padding:10px;border:1px solid #ddd;
                                   text-align:center'>{total}</td>
                    </tr>
                </tfoot>
            </table><br>
            <h3>Top 5 Oportunidades</h3>
            <table style='border-collapse:collapse;width:100%'>
                <thead>
                    <tr style='background:#1a73e8;color:white'>
                        <th style='padding:10px'>Entidad</th>
                        <th style='padding:10px'>Descripcion</th>
                        <th style='padding:10px'>Monto</th>
                        <th style='padding:10px'>Rubro</th>
                    </tr>
                </thead>
                <tbody>{top_oportunidades}</tbody>
            </table><br>
            <div style='background:#f8f9fa;padding:15px;border-radius:8px'>
                <p style='margin:0;color:#555'>
                    Excel adjunto con pestanas por rubro<br>
                    <a href='https://prod4.seace.gob.pe/openegocio'>
                        Ver SEACE directamente</a><br>
                    Generado automaticamente con GitHub Actions
                </p>
            </div>
        </body></html>"""

        msg.attach(MIMEText(cuerpo_html, "html"))

        if ruta_excel and os.path.exists(ruta_excel):
            with open(ruta_excel, "rb") as f:
                adjunto = MIMEBase("application", "octet-stream")
                adjunto.set_payload(f.read())
                encoders.encode_base64(adjunto)
                adjunto.add_header(
                    "Content-Disposition",
                    f"attachment; filename={os.path.basename(ruta_excel)}"
                )
                msg.attach(adjunto)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(CORREO_REMITENTE, CORREO_CONTRASENA)
            server.sendmail(CORREO_REMITENTE, CORREO_DESTINO, msg.as_string())
        print("Correo enviado exitosamente!")

    except Exception as e:
        print(f"Error enviando correo: {e}")

if __name__ == "__main__":
    print("=" * 55)
    print("MONITOR SEACE - GITHUB ACTIONS")
    print("=" * 55)
    print(f"Alertas a    : {CORREO_DESTINO}")
    print(f"Region       : {REGION_ACTIVA}")
    print(f"Monto minimo : S/. {MONTO_MINIMO:,}")
    print(f"Tipo entidad : {ENTIDAD_ACTIVA}")
    print(f"Tipo proceso : {TIPO_PROCESO_ACTIVO}")
    print(f"Rubros       : {len(RUBROS)}")
    print("=" * 55)

    resultados = buscar_en_seace()
    ruta_excel = guardar_excel_pestanas(resultados)
    enviar_correo(resultados, ruta_excel)
    print("\nEjecucion completada exitosamente!")
