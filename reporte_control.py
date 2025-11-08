print("Iniciando análisis de todas las hojas...")

import os
import pandas as pd
from datetime import datetime
import yagmail
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC


# CONFIGURACIÓN DEL CORREO

EMAIL = os.getenv("EMAIL_APP")
PASSWORD = os.getenv("EMAIL_PASS")

if not EMAIL or not PASSWORD:
    raise ValueError("Faltan las variables de entorno EMAIL_APP o EMAIL_PASS.")

ASUNTO_GENERAL = "Reporte de Novedades y Sisbén"

# Correos de los profesores por salón
PROFESORES = {
    "CUNAS 1": "prof_cuna1@correo.com",
    "CUNAS 2": "prof_cuna2@correo.com",
    "CUNAS 3": "prof_cuna3@correo.com",
    "CUNAS 4": "prof_cuna4@correo.com",
    "CUNAS 5": "prof_cuna5@correo.com"
}


# LECTURA DE ARCHIVO EXCEL

print("Cargando archivo Excel con todas las hojas...")
archivo = r"D:\dev\Control_vacunas\DATA_FAKE.xlsx"
excel = pd.read_excel(archivo, sheet_name=None)  # Cargar todas las hojas


# CONFIGURACIÓN DEL DRIVER (SISBEN)

print("Inicializando navegador para consultas SISBÉN...")
url = "https://www.sisben.gov.co/paginas/consulta-tu-grupo.html"
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 15)

def consultar_sisben(doc):
    try:
        driver.get(url)
        iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)

        Select(driver.find_element(By.ID, "TipoID")).select_by_value("1")
        campo_doc = driver.find_element(By.ID, "documento")
        campo_doc.clear()
        campo_doc.send_keys(str(doc))

        boton = wait.until(EC.element_to_be_clickable((By.ID, "botonenvio")))
        driver.execute_script("arguments[0].click();", boton)
        time.sleep(3)

        if len(driver.find_elements(By.CLASS_NAME, "swal2-html-container")) > 0:
            driver.find_element(By.CLASS_NAME, "swal2-confirm").click()
            driver.switch_to.default_content()
            return None

        municipio = driver.find_elements(By.CSS_SELECTOR, "p.campo1")[3].text.strip()
        nivel = driver.find_element(By.CSS_SELECTOR, "p.text-uppercase.font-weight-bold.text-white").text.strip()
        driver.switch_to.default_content()
        return f"SI {municipio} / {nivel}"

    except Exception as e:
        driver.switch_to.default_content()
        print(f"Error consultando {doc}: {e}")
        return None


# PROCESAR CADA HOJA

yag = yagmail.SMTP(EMAIL, PASSWORD)
hoy = pd.Timestamp.today()
columnas_fecha = [
    "PROX. CONTROL VALORACION INTEGRAL",
    "PROXIMA CITA ODONTOLOGO",
    "PROX. CONTROL FLUORIZACION",
    "PROX. CONTROL DESPARASITACION",
]
columnas_excluidas = {"PORTABILIDAD", "COMPROMISO", "SEGUIMIENTO 1", "SEGUIMIENTO 2"}

for nombre_hoja, df in excel.items():
    if nombre_hoja not in PROFESORES:
        print(f"Saltando hoja {nombre_hoja} (sin profesor asignado)")
        continue

    print(f"\n📋 Analizando hoja: {nombre_hoja}...")
    df.columns = df.columns.str.strip().str.upper()

    proximas, vencidas, pendientes = [], [], []
    sisben_actualizados, sisben_pendientes = [], []

    # === Revisión de fechas y campos ===
    for idx, fila in df.iterrows():
        nombre = str(fila.get("NOMBRES Y APELLIDOS", "SIN NOMBRE")).strip()
        correo_acudiente = str(fila.get("CORREO ACUDIENTE", "")).strip()

        # Revisión de fechas
        for col in columnas_fecha:
            if col in df.columns:
                valor = fila[col]
                if pd.isnull(valor):
                    continue
                if isinstance(valor, str) and valor.strip().upper() == "NO":
                    pendientes.append((nombre, col))
                else:
                    fecha = pd.to_datetime(valor, errors='coerce')
                    if pd.notnull(fecha):
                        dias = (fecha - hoy).days
                        if 0 <= dias <= 8:
                            proximas.append((nombre, col, fecha.strftime("%d/%m/%Y")))
                        elif dias < 0:
                            vencidas.append((nombre, col, fecha.strftime("%d/%m/%Y")))

        # Revisión de "NO" en otros campos
        for col in df.columns:
            if col in columnas_excluidas:
                continue
            valor = fila[col]
            if isinstance(valor, str) and valor.strip().upper() == "NO":
                pendientes.append((nombre, col))

        # === Consulta SISBEN ===
        if "SISBEN" in df.columns and str(fila["SISBEN"]).strip().upper() == "NO":
            doc = fila.get("ID", "")
            if doc:
                resultado = consultar_sisben(doc)
                if resultado:
                    df.at[idx, "SISBEN"] = resultado
                    sisben_actualizados.append((nombre, doc, resultado))
                else:
                    sisben_pendientes.append((nombre, doc))

        # === Si el niño tiene novedades, enviar correo al acudiente ===
        if correo_acudiente:
            cuerpo_padre = f"Estimado acudiente de {nombre},\n\nSe encontraron las siguientes novedades:\n"
            novedades = []

            for n, c, f in proximas:
                if n == nombre:
                    novedades.append(f"🟡 Próximo {c} el {f}")
            for n, c, f in vencidas:
                if n == nombre:
                    novedades.append(f"🔴 {c} vencido el {f}")
            for n, c in pendientes:
                if n == nombre:
                    novedades.append(f"⚠️ Campo pendiente: {c}")
            for n, d, r in sisben_actualizados:
                if n == nombre:
                    novedades.append(f"✅ SISBEN actualizado: {r}")
            for n, d in sisben_pendientes:
                if n == nombre:
                    novedades.append(f"⚠️ SISBEN pendiente para documento {d}")

            if novedades:
                cuerpo_padre += "\n".join(novedades) + "\n\nPor favor, ponerse al día lo antes posible."
                try:
                    yag.send(
                        to=correo_acudiente,
                        subject=f"Novedades de {nombre}",
                        contents=cuerpo_padre
                    )
                    print(f"📧 Correo enviado a acudiente de {nombre}")
                except Exception as e:
                    print(f"⚠️ Error enviando correo a {correo_acudiente}: {e}")

    # Correo al profesor con resumen 
    cuerpo_profesor = f"Estimado profesor/a {nombre_hoja},\n\nResumen de novedades:\n\n"
    if proximas:
        cuerpo_profesor += "🟡 Próximos controles:\n" + "\n".join([f"- {n} ({c}: {f})" for n, c, f in proximas]) + "\n\n"
    if vencidas:
        cuerpo_profesor += "🔴 Vencidas:\n" + "\n".join([f"- {n} ({c}: {f})" for n, c, f in vencidas]) + "\n\n"
    if pendientes:
        cuerpo_profesor += "⚠️ Pendientes:\n" + "\n".join([f"- {n} ({c})" for n, c in pendientes]) + "\n\n"
    if sisben_actualizados:
        cuerpo_profesor += "✅ Sisbén actualizados:\n" + "\n".join([f"- {n} ({r})" for n, d, r in sisben_actualizados]) + "\n\n"
    if sisben_pendientes:
        cuerpo_profesor += "⚠️ Sisbén pendientes:\n" + "\n".join([f"- {n}" for n, d in sisben_pendientes]) + "\n\n"

    try:
        yag.send(
            to=PROFESORES[nombre_hoja],
            subject=f"{ASUNTO_GENERAL} - {nombre_hoja}",
            contents=cuerpo_profesor
        )
        print(f"📧 Resumen enviado a profesor de {nombre_hoja}")
    except Exception as e:
        print(f"⚠️ Error enviando correo al profesor de {nombre_hoja}: {e}")

driver.quit()


# GUARDAR SOLO LOS NIÑOS CON NOVEDADES


print("Generando archivo con solo las novedades...")

# Creamos un conjunto con los nombres que tienen alguna novedad
nombres_novedad = set()

for lista in [proximas, vencidas]:
    for nombre, _, _ in lista:
        nombres_novedad.add(nombre)

for lista in [pendientes]:
    for nombre, _ in lista:
        nombres_novedad.add(nombre)

for lista in [sisben_actualizados, sisben_pendientes]:
    for nombre, *_ in lista:
        nombres_novedad.add(nombre)

# Filtramos el DataFrame solo con los nombres con novedades
df_novedades = df[df["NOMBRES Y APELLIDOS"].isin(nombres_novedad)].copy()

# Guardamos el archivo filtrado con fecha
fecha_hoy = datetime.today().strftime("%Y%m%d")
nuevo_archivo = archivo.replace(".xlsx", f"_NOVEDADES_{fecha_hoy}.xlsx")

if not df_novedades.empty:
    df_novedades.to_excel(nuevo_archivo, index=False)
    print(f"✅ Archivo de novedades guardado en: {nuevo_archivo}")
else:
    print("No se encontraron novedades, no se generó archivo nuevo.")