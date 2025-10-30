import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
import os

# --- CONFIGURACIÓN GENERAL ---
excel_file = "/Users/andres/Documents/Certificados automaticos/Datos/datos-2.xlsx"
plantilla = "/Users/andres/Documents/Certificados automaticos/Plantilla/Diploma Certificado Título Curso Clase Cursillo Profesional Llamativo Elegante Dorado .png"
carpeta_salida = "Certificados"

os.makedirs(carpeta_salida, exist_ok=True)

# --- CARGA DE DATOS ---
df = pd.read_excel(excel_file)

# --- FUENTES ---
def cargar_fuente(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()

fuente_titulo = cargar_fuente("/Library/Fonts/Arial Bold.ttf", 58)
fuente_texto = cargar_fuente("/Library/Fonts/Arial.ttf", 34)
fuente_nombre = cargar_fuente("/Library/Fonts/Arial Bold.ttf", 40)
fuente_fecha = cargar_fuente("/Library/Fonts/Arial Italic.ttf", 28)

# --- FECHA EN ESPAÑOL ---
meses = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}
hoy = datetime.now()
fecha_actual = f"{hoy.day} de {meses[hoy.month]} de {hoy.year}"

# --- FUNCIONES AUXILIARES ---
def medir_texto(draw, texto, fuente):
    if not texto:
        return (0, 0)
    bbox = draw.textbbox((0, 0), texto, font=fuente)
    return (bbox[2] - bbox[0], bbox[3] - bbox[1])

def centrar_texto(draw, imagen, texto, fuente, y, fill=(0, 0, 0), interlineado=12):
    """Dibuja texto centrado y devuelve nueva coordenada Y con más interlineado."""
    w, h = medir_texto(draw, texto, fuente)
    x = (imagen.width - w) / 2
    draw.text((x, y), texto, font=fuente, fill=fill)
    return y + h + interlineado  # Aumentamos el espacio entre líneas

def wrap_text(draw, texto, fuente, max_width):
    palabras = texto.split()
    lineas, actual = [], palabras[0] if palabras else ""
    for palabra in palabras[1:]:
        prueba = actual + " " + palabra
        w, _ = medir_texto(draw, prueba, fuente)
        if w <= max_width:
            actual = prueba
        else:
            lineas.append(actual)
            actual = palabra
    if actual:
        lineas.append(actual)
    return lineas

# --- RECORRER CADA FILA DEL EXCEL ---
for idx, fila in df.iterrows():
    proyecto = str(fila.get("Proyecto", "")).strip()
    espacio = str(fila.get("Espacio academico", "")).strip()

    for n in range(1, 5):
        nombre_col = f"Nombre {n}"
        codigo_col = f"Código {n}" if n != 2 else "Código  2"

        nombre = str(fila.get(nombre_col, "")).strip()
        codigo = fila.get(codigo_col)

        # Validar que tenga nombre Y código válidos
        if not nombre or pd.isna(codigo) or str(codigo).strip() == "":
            continue

        # Normalizar código (sin .0)
        if isinstance(codigo, float) and codigo.is_integer():
            codigo = str(int(codigo))
        else:
            codigo = str(codigo).strip()

        # --- CREAR CERTIFICADO INDIVIDUAL ---
        imagen = Image.open(plantilla).convert("RGB")
        draw = ImageDraw.Draw(imagen)
        max_w = int(imagen.width * 0.80)
        y = int(imagen.height * 0.22)
        inter = 16  # interlineado general más amplio

        # --- TÍTULO ---
        y = centrar_texto(draw, imagen, "CONSTANCIA DE PARTICIPACIÓN", fuente_titulo, y, interlineado=inter)
        y = centrar_texto(draw, imagen, "EN LA VIII MUESTRA DE INGENIERÍA INDUSTRIAL", fuente_titulo, y, interlineado=inter + 4)
        y += 20

        # --- TEXTO GENERAL ---
        texto_intro = "La Facultad de Ingeniería Industrial Seccional Villavicencio hace constar que el estudiante:"
        for linea in wrap_text(draw, texto_intro, fuente_texto, max_w):
            y = centrar_texto(draw, imagen, linea, fuente_texto, y, interlineado=inter)
        y += 20

        # --- NOMBRE DEL ESTUDIANTE ---
        y = centrar_texto(draw, imagen, f"{nombre} ({codigo})", fuente_nombre, y, interlineado=inter + 6)
        y += 25

        # --- PROYECTO ---
        y = centrar_texto(draw, imagen, "Participó como ponente en modalidad Póster con el Proyecto:", fuente_texto, y, interlineado=inter)
        y += 12
        for linea in wrap_text(draw, f"\"{proyecto}\"", fuente_texto, max_w):
            y = centrar_texto(draw, imagen, linea, fuente_texto, y, interlineado=inter)
        y += 20

        # --- ESPACIO ACADÉMICO ---
        y = centrar_texto(draw, imagen, "estudio realizado en el espacio académico de", fuente_texto, y, interlineado=inter)
        y += 12
        y = centrar_texto(draw, imagen, espacio.upper(), fuente_nombre, y, interlineado=inter + 4)
        y += 25

        # --- FECHA ---
        y = centrar_texto(draw, imagen, f"El certificado es otorgado el día {fecha_actual}", fuente_fecha, y, interlineado=inter)

        # --- GUARDAR COMO PDF ---
        nombre_archivo = f"{nombre.replace(' ', '_')}_{codigo}.pdf"
        ruta = os.path.join(carpeta_salida, nombre_archivo)
        imagen.save(ruta, "PDF", resolution=300.0)
        print(f"✅ Guardado: {ruta}")

print("\n🎉 Proceso finalizado. Certificados PDF individuales generados correctamente en:", carpeta_salida)
