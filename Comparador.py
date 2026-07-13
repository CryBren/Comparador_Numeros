import re
import time
import html
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager


URL = "https://sns.ift.org.mx/sns-frontend/consulta-numeracion/numeracion-geografica.xhtml"


# =====================================================
# NORMALIZACIÓN
# =====================================================

def normalizar(proveedor):

    if pd.isna(proveedor):
        return ""

    txt = html.unescape(str(proveedor))
    txt = txt.upper().strip()

    equivalencias = {

        # TELCEL
        "RADIOMOVIL DIPSA": "TELCEL",
        "RADIOMÓVIL DIPSA": "TELCEL",
        "TELCEL GSM": "TELCEL",
        "AMERICA MOVIL": "TELCEL",
        "AMÉRICA MÓVIL": "TELCEL",
        "FREEDOMPOP": "TELCEL",

        # AT&T
        "ATT": "AT&T",
        "AT T": "AT&T",
        "AT&T": "AT&T",
        "YONDER": "AT&T",
        "IUSACELL": "AT&T",

        # ALTÁN
        "ALTAN": "ALTAN",
        "ALTÁN": "ALTAN",
        "ALTAN REDES": "ALTAN",
        "ALTÁN REDES": "ALTAN",
        "BAIT": "ALTAN",
        "WAL-MART": "ALTAN",
        "CORPORACIÓN NOVAVISIÓN": "ALTAN",
        "MEGACABLE": "ALTAN",
        "TELECOMUNICACIONES 360": "ALTAN",
        "ORIÓNIDAS": "ALTAN",
        "CFE": "ALTAN",
        "GRUPO TECNOLOGÍA Y COMUNICACIÓN": "ALTAN",
        "DALE FON": "ALTAN",
        "ABSOLUTETECK": "ALTAN",

        # MOVISTAR
        "MOVISTAR": "TELEFONICA",
        "TELEFÓNICA": "TELEFONICA",
        "TELEFONICA": "TELEFONICA",

        # IZZI
        "IZZI MOBILE": "TVI"

    }

    return equivalencias.get(txt, txt)

# =====================================================
# VALIDAR NÚMERO
# =====================================================

def validar_numero(numero):

    numero = str(numero).strip()

    # Quitar .0 que agrega Excel
    if numero.endswith(".0"):
        numero = numero[:-2]

    # Dejar solo dígitos
    numero = re.sub(r"\D", "", numero)

    # Debe tener exactamente 10 dígitos
    if len(numero) != 10:
        return None

    return numero


# =====================================================
# CONSULTA IFT
# =====================================================

def obtener_proveedor(driver, numero):

    for intento in range(3):

        try:

            driver.get(URL)

            wait = WebDriverWait(driver, 20)

            campo = wait.until(
                EC.presence_of_element_located(
                    (By.ID, "FORM_myform:TXT_NationalNumber")
                )
            )

            campo.clear()
            campo.send_keys(numero)

            time.sleep(1)

            boton = wait.until(
                EC.element_to_be_clickable(
                    (By.ID, "FORM_myform:BTN_publicSearch")
                )
            )

            driver.execute_script(
                "arguments[0].click();",
                boton
            )

            # Esperar resultado
            try:
                wait.until(
                    lambda d:
                    "Proveedor que atiende el número"
                    in d.page_source
                )
            except:
                pass

            html_page = driver.page_source

            # ==========================================
            # NÚMERO NO ENCONTRADO
            # ==========================================

            if "Número no encontrado" in html_page:

                print(f"IFT: Número no encontrado ({numero})")

                return "NO ENCONTRADO"

            # Extraer proveedor
            patron = (
                r'Proveedor que atiende el número\.'
                r'</div>\s*'
                r'<div class="ui-panelgrid-cell ui-grid-col-6">'
                r'(.*?)'
                r'</div>'
            )

            m = re.search(
                patron,
                html_page,
                re.IGNORECASE | re.DOTALL
            )

            if m:

                proveedor = html.unescape(
                    m.group(1).strip()
                )

                proveedor = re.sub(
                    r"<.*?>",
                    "",
                    proveedor
                ).strip()

                if proveedor:

                    print(
                        f"IFT: {proveedor}"
                    )

                    return proveedor

            # respaldo
            texto = re.sub(
                r"\s+",
                " ",
                html_page
            )

            m = re.search(
                r"Proveedor que atiende el número\.</div>\s*<div[^>]*>(.*?)</div>",
                texto,
                re.IGNORECASE
            )

            if m:

                proveedor = html.unescape(
                    m.group(1).strip()
                )

                proveedor = re.sub(
                    r"<.*?>",
                    "",
                    proveedor
                ).strip()

                if proveedor:

                    print(
                        f"IFT: {proveedor}"
                    )

                    return proveedor

            print(
                f"Proveedor no encontrado para {numero}"
            )

        except Exception as e:

            print(
                f"Reintentando {numero} "
                f"(intento {intento + 1}) - "
                f"{type(e).__name__}"
            )

            time.sleep(2)

    try:

        with open(
            f"error_{numero}.html",
            "w",
            encoding="utf-8"
        ) as f:

            f.write(driver.page_source)

    except:
        pass

    return ""


# =====================================================
# CONFIGURACIÓN
# =====================================================

archivo_entrada = "numeros.xlsx"

col_numero = "Cel"
col_proveedor = "Proveedor"


# =====================================================
# LEER EXCEL
# =====================================================

df = pd.read_excel(archivo_entrada)

print()
print(
    f"Registros encontrados: {len(df)}"
)
print()


# =====================================================
# CHROME
# =====================================================

options = webdriver.ChromeOptions()

options.add_argument("--start-maximized")

# Si quieres que no abra ventana:
# options.add_argument("--headless=new")

driver = webdriver.Chrome(
    service=Service(
        ChromeDriverManager().install()
    ),
    options=options
)


# =====================================================
# PROCESAMIENTO
# =====================================================

try:

    proveedores_ift = []
    coincidencias = []
    observaciones = []

    total = len(df)

    # ==========================
    # CONTADORES
    # ==========================

    validos = 0
    invalidos = 0
    coinciden = 0
    no_coinciden = 0
    errores_consulta = 0

    for i, row in df.iterrows():

        numero = validar_numero(row[col_numero])

        print()
        print(f"[{i+1}/{total}]")

        # ==========================
        # VALIDACIÓN
        # ==========================

        if numero is None:

            numero_original = str(row[col_numero])

            print(f"Número inválido: {numero_original}")

            invalidos += 1

            proveedores_ift.append("")
            coincidencias.append("NO")
            observaciones.append("Número inválido")

            continue

        print(numero)
        validos += 1

        # ==========================
        # CONSULTAR IFT
        # ==========================

        proveedor_ift = obtener_proveedor(
            driver,
            numero
        )

        if proveedor_ift == "":
            errores_consulta += 1

        proveedores_ift.append(
            proveedor_ift
        )

        proveedor_excel = normalizar(
            row[col_proveedor]
        )

        proveedor_ift_norm = normalizar(
            proveedor_ift
        )

        coincide = (
            "SI"
            if proveedor_excel == proveedor_ift_norm
            else "NO"
        )

        if coincide == "SI":
            coinciden += 1
        else:
            no_coinciden += 1

        coincidencias.append(
            coincide
        )

        if proveedor_ift == "":
            observaciones.append("No fue posible obtener proveedor")
        else:
            observaciones.append("")

        print(
            f"{numero} -> '{proveedor_ift}'"
        )

    # ==========================
    # EXPORTAR
    # ==========================

    df["Proveedor_IFT"] = proveedores_ift
    df["Coincide"] = coincidencias
    df["Observaciones"] = observaciones

    archivo_salida = "comparacion_ift.xlsx"

    df.to_excel(
        archivo_salida,
        index=False
    )

    print()
    print("=" * 45)
    print("        RESUMEN DEL PROCESO")
    print("=" * 45)
    print(f"Total de registros        : {total}")
    print(f"Números válidos           : {validos}")
    print(f"Números inválidos         : {invalidos}")
    print(f"Consultas sin respuesta   : {errores_consulta}")
    print(f"Coincidencias             : {coinciden}")
    print(f"No coinciden              : {no_coinciden}")
    print("=" * 45)
    print(f"Archivo generado: {archivo_salida}")
    print("=" * 45)

finally:

    driver.quit()
