"""
Presentación: Discount data for Retention (textos en inglés)
Necesidad, brecha, definición de la transacción y plan

Genera un fichero .pptx con python-pptx.

Cómo está organizado el script:
    1. Configuración: colores, fuentes y nombre del fichero de salida.
    2. Funciones de ayuda: texto, rectángulos, flechas, barras de precio y transiciones.
    3. Una función por diapositiva (diapositiva_portada, diapositiva_resumen, ...).
    4. La función principal, que crea la presentación y llama a cada diapositiva en orden.

Para cambiar el contenido de una diapositiva basta con editar su función.
Para cambiar el orden, quitar o añadir diapositivas, editar la lista en generar_presentacion().

Requisito:  pip install python-pptx
Ejecución:  python presentacion_descuentos_msrp.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree


# =============================================================================
# 1. CONFIGURACIÓN
# =============================================================================

NOMBRE_FICHERO_SALIDA = "discount_data_for_retention.pptx"

ANCHO_DIAPOSITIVA_PULGADAS = 13.333
ALTO_DIAPOSITIVA_PULGADAS = 7.5
MARGEN_IZQUIERDO = 0.6
ANCHO_UTIL = ANCHO_DIAPOSITIVA_PULGADAS - 2 * MARGEN_IZQUIERDO

# Fuente corporativa. Si no está instalada en el equipo que abre el fichero,
# PowerPoint la sustituye por otra. Para probar sin ella, cambiar a "Calibri" o "Arial".
FUENTE_CORPORATIVA = "Kaspersky Sans Display"
FUENTE_TITULOS = FUENTE_CORPORATIVA
FUENTE_CUERPO = FUENTE_CORPORATIVA

# Dos temas de color con los mismos significados:
#   - "oscuro": paleta sacada del gráfico de Power BI "First Payment Failure reason category".
#   - "claro":  fondo blanco y los mismos colores algo más intensos, para que se lean sobre blanco.
# Las bandas oscuras (COLOR_FONDO_PORTADA) son iguales en los dos temas, como tus tablas
# oscuras de Power BI dentro de diapositivas blancas.
# Cada diapositiva elige su tema en generar_presentacion(); aplicar_tema() cambia los colores.

TEMA_OSCURO = {
    # Fondos y texto
    "COLOR_FONDO_PORTADA": RGBColor(0x14, 0x17, 0x1C),
    "COLOR_FONDO_CONTENIDO": RGBColor(0x1E, 0x21, 0x28),
    "COLOR_TARJETA": RGBColor(0x2A, 0x2E, 0x37),
    "COLOR_TEXTO": RGBColor(0xEC, 0xEE, 0xF2),
    "COLOR_TEXTO_SECUNDARIO": RGBColor(0xA9, 0xB1, 0xBF),
    # Significado de los importes
    "COLOR_MSRP": RGBColor(0x6F, 0x8F, 0xF0),                  # azul
    "COLOR_PAGADO": RGBColor(0x6C, 0xC9, 0xEC),                # azul claro
    "COLOR_DESCUENTO": RGBColor(0xF7, 0x9A, 0x6A),             # naranja
    "COLOR_DESCUENTO_SECUNDARIO": RGBColor(0xEE, 0xDB, 0x5C),  # amarillo
    "COLOR_MORADO": RGBColor(0xA3, 0x7B, 0xE6),                # morado
    # Resultado de las comprobaciones
    "COLOR_CORRECTO": RGBColor(0x3D, 0xC4, 0xA2),              # verde
    "COLOR_CORRECTO_SUAVE": RGBColor(0x1E, 0x3A, 0x36),
    "COLOR_ERROR": RGBColor(0xF2, 0x66, 0x7F),                 # rosa
    "COLOR_ERROR_SUAVE": RGBColor(0x45, 0x27, 0x30),
}

TEMA_CLARO = {
    # Fondos y texto
    "COLOR_FONDO_PORTADA": RGBColor(0x14, 0x17, 0x1C),
    "COLOR_FONDO_CONTENIDO": RGBColor(0xFF, 0xFF, 0xFF),
    "COLOR_TARJETA": RGBColor(0xF1, 0xF3, 0xF6),
    "COLOR_TEXTO": RGBColor(0x1F, 0x29, 0x37),
    "COLOR_TEXTO_SECUNDARIO": RGBColor(0x5B, 0x64, 0x72),
    # Significado de los importes
    "COLOR_MSRP": RGBColor(0x3D, 0x6F, 0xD9),                  # azul
    "COLOR_PAGADO": RGBColor(0x3B, 0xA8, 0xDE),                # azul claro
    "COLOR_DESCUENTO": RGBColor(0xEC, 0x7A, 0x3C),             # naranja
    "COLOR_DESCUENTO_SECUNDARIO": RGBColor(0xE5, 0xB4, 0x00),  # amarillo
    "COLOR_MORADO": RGBColor(0x7B, 0x47, 0xD1),                # morado
    # Resultado de las comprobaciones
    "COLOR_CORRECTO": RGBColor(0x14, 0x9A, 0x7A),              # verde
    "COLOR_CORRECTO_SUAVE": RGBColor(0xDD, 0xF3, 0xEC),
    "COLOR_ERROR": RGBColor(0xD4, 0x2A, 0x4F),                 # rojo
    "COLOR_ERROR_SUAVE": RGBColor(0xFB, 0xE3, 0xE7),
}

# Colores que no cambian con el tema
COLOR_TEXTO_SOBRE_OSCURO = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO = RGBColor(0xB8, 0xC2, 0xD0)
COLOR_TEXTO_SOBRE_COLOR_CLARO = RGBColor(0x14, 0x17, 0x1C)


def aplicar_tema(tema):
    """Cambia los colores globales al tema indicado. Afecta a las diapositivas que se dibujen después."""
    globals().update(tema)


# Al cargar el script se parte del tema oscuro
aplicar_tema(TEMA_OSCURO)

SIMBOLO_CORRECTO = "✓"
SIMBOLO_ERROR = "✗"


# =============================================================================
# 2. FUNCIONES DE AYUDA
# =============================================================================

def formatear_importe(valor):
    """Devuelve un importe en formato inglés: 84.7 -> '€84.70', 50 -> '€50'."""
    if float(valor).is_integer():
        texto_numero = str(int(valor))
    else:
        texto_numero = f"{valor:.2f}"
    return f"€{texto_numero}"


def formatear_numero(valor):
    """Número con dos decimales y separador de miles: 1250.1 -> '1,250.10', 36 -> '36.00'."""
    return f"{valor:,.2f}"


def formatear_euros(valor):
    """Importe en euros siempre con dos decimales: 36 -> '€36.00'."""
    return f"€{formatear_numero(valor)}"


def dibujar_linea(diapositiva, x_inicio, y_inicio, x_fin, y_fin, color, grosor_en_puntos=1.25):
    """Línea recta entre dos puntos (en pulgadas)."""
    linea = diapositiva.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x_inicio), Inches(y_inicio),
                                             Inches(x_fin), Inches(y_fin))
    linea.line.color.rgb = color
    linea.line.width = Pt(grosor_en_puntos)
    return linea


def formatear_porcentaje(valor):
    """Porcentaje sin decimales si es entero y con uno si no: 20.0 -> '20%', 16.666 -> '16.7%'."""
    redondeado_a_un_decimal = round(valor, 1)
    if float(redondeado_a_un_decimal).is_integer():
        return f"{int(redondeado_a_un_decimal)}%"
    return f"{redondeado_a_un_decimal}%"


def fragmento(texto, tamaño=16, color=None, negrita=False, fuente=FUENTE_CUERPO, cursiva=False):
    """
    Un trozo de texto con su formato. Un párrafo es una lista de fragmentos.
    Sin color indicado, usa el color de texto del tema activo.
    """
    if color is None:
        color = COLOR_TEXTO
    return {
        "texto": texto,
        "tamaño": tamaño,
        "color": color,
        "negrita": negrita,
        "fuente": fuente,
        "cursiva": cursiva,
    }


def rellenar_marco_de_texto(marco_de_texto, parrafos, alineacion, anclaje, espacio_despues_de_parrafo):
    """
    Escribe una lista de párrafos dentro de un marco de texto (de un cuadro de texto o de una forma).
    Cada párrafo es una lista de fragmentos creados con fragmento().
    """
    marco_de_texto.clear()
    marco_de_texto.word_wrap = True
    marco_de_texto.vertical_anchor = anclaje
    marco_de_texto.margin_left = 0
    marco_de_texto.margin_right = 0
    marco_de_texto.margin_top = 0
    marco_de_texto.margin_bottom = 0

    for numero_de_parrafo, fragmentos_del_parrafo in enumerate(parrafos):
        if numero_de_parrafo == 0:
            parrafo = marco_de_texto.paragraphs[0]
        else:
            parrafo = marco_de_texto.add_paragraph()

        parrafo.alignment = alineacion
        parrafo.space_after = Pt(espacio_despues_de_parrafo)

        for datos_del_fragmento in fragmentos_del_parrafo:
            trozo = parrafo.add_run()
            trozo.text = datos_del_fragmento["texto"]
            trozo.font.size = Pt(datos_del_fragmento["tamaño"])
            trozo.font.color.rgb = datos_del_fragmento["color"]
            trozo.font.bold = datos_del_fragmento["negrita"]
            trozo.font.italic = datos_del_fragmento["cursiva"]
            trozo.font.name = datos_del_fragmento["fuente"]


def escribir_parrafos(diapositiva, x, y, ancho, alto, parrafos,
                      alineacion=PP_ALIGN.LEFT, anclaje=MSO_ANCHOR.TOP, espacio_despues_de_parrafo=6):
    """Crea un cuadro de texto con varios párrafos, cada uno con su propio formato."""
    cuadro_de_texto = diapositiva.shapes.add_textbox(Inches(x), Inches(y), Inches(ancho), Inches(alto))
    rellenar_marco_de_texto(cuadro_de_texto.text_frame, parrafos, alineacion, anclaje, espacio_despues_de_parrafo)
    return cuadro_de_texto


def escribir_texto(diapositiva, x, y, ancho, alto, texto, tamaño=16, color=None, negrita=False,
                   fuente=FUENTE_CUERPO, alineacion=PP_ALIGN.LEFT, anclaje=MSO_ANCHOR.TOP, cursiva=False):
    """
    Atajo para un cuadro de texto con un único formato. Cada salto de línea crea un párrafo.
    Lo que vaya entre **dobles asteriscos** se escribe en negrita.
    """
    parrafos = []
    for linea in texto.split("\n"):
        fragmentos_de_la_linea = []
        trozos = linea.split("**")
        for numero_de_trozo, trozo in enumerate(trozos):
            if trozo == "":
                continue
            es_un_trozo_entre_asteriscos = numero_de_trozo % 2 == 1
            fragmentos_de_la_linea.append(
                fragmento(trozo, tamaño, color, negrita or es_un_trozo_entre_asteriscos, fuente, cursiva))
        if len(fragmentos_de_la_linea) == 0:
            fragmentos_de_la_linea.append(fragmento("", tamaño, color, negrita, fuente, cursiva))
        parrafos.append(fragmentos_de_la_linea)
    return escribir_parrafos(diapositiva, x, y, ancho, alto, parrafos, alineacion, anclaje)


# Nombres de tablas y fuentes de datos que se resaltan en negrita en los textos explicativos
ENTIDADES_DE_DATOS = ["WebPriceList", "eComm", "API"]


def resaltar_entidades(texto):
    """Pone en negrita (con **) los nombres de tablas y fuentes de datos que aparecen en el texto."""
    for entidad in ENTIDADES_DE_DATOS:
        texto = texto.replace(entidad, f"**{entidad}**")
    return texto


def dibujar_rectangulo(diapositiva, x, y, ancho, alto, color_relleno, redondeado=False, color_borde=None):
    """Dibuja un rectángulo sin sombra. Si no se indica borde, no lleva borde."""
    tipo_de_forma = MSO_SHAPE.ROUNDED_RECTANGLE if redondeado else MSO_SHAPE.RECTANGLE
    forma = diapositiva.shapes.add_shape(tipo_de_forma, Inches(x), Inches(y), Inches(ancho), Inches(alto))
    forma.shadow.inherit = False

    forma.fill.solid()
    forma.fill.fore_color.rgb = color_relleno

    if color_borde is None:
        forma.line.fill.background()
    else:
        forma.line.color.rgb = color_borde
        forma.line.width = Pt(1.5)

    if redondeado:
        forma.adjustments[0] = 0.12

    return forma


def calcular_luminosidad(color):
    """Luminosidad relativa de un color RGB (fórmula WCAG), entre 0 (negro) y 1 (blanco)."""
    componentes_lineales = []
    for componente in (color[0], color[1], color[2]):
        valor = componente / 255
        if valor <= 0.03928:
            componentes_lineales.append(valor / 12.92)
        else:
            componentes_lineales.append(((valor + 0.055) / 1.055) ** 2.4)
    rojo, verde, azul = componentes_lineales
    return 0.2126 * rojo + 0.7152 * verde + 0.0722 * azul


def elegir_color_de_texto_para_fondo(color_de_fondo):
    """Devuelve texto oscuro sobre colores claros de la paleta y texto blanco sobre colores oscuros."""
    if calcular_luminosidad(color_de_fondo) > 0.3:
        return COLOR_TEXTO_SOBRE_COLOR_CLARO
    return COLOR_TEXTO_SOBRE_OSCURO


def escribir_dentro_de_forma(forma, texto, tamaño=14, color=None, negrita=True,
                             fuente=FUENTE_CUERPO, alineacion=PP_ALIGN.CENTER):
    """
    Escribe texto centrado dentro de una forma. Cada salto de línea crea un párrafo.
    Si no se indica color, se elige según el relleno de la forma para que se lea bien.
    """
    if color is None:
        color = elegir_color_de_texto_para_fondo(forma.fill.fore_color.rgb)

    parrafos = []
    for linea in texto.split("\n"):
        parrafos.append([fragmento(linea, tamaño, color, negrita, fuente)])
    rellenar_marco_de_texto(forma.text_frame, parrafos, alineacion, MSO_ANCHOR.MIDDLE, 0)


def dibujar_flecha(diapositiva, x, y, ancho, alto, color=None):
    """Flecha de bloque apuntando a la derecha. Sin color indicado, usa el color de texto del tema activo."""
    if color is None:
        color = COLOR_TEXTO
    flecha = diapositiva.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(ancho), Inches(alto))
    flecha.shadow.inherit = False
    flecha.fill.solid()
    flecha.fill.fore_color.rgb = color
    flecha.line.fill.background()
    return flecha


def dibujar_circulo_con_texto(diapositiva, x, y, diametro, texto, color_relleno, tamaño=14):
    circulo = diapositiva.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(diametro), Inches(diametro))
    circulo.shadow.inherit = False
    circulo.fill.solid()
    circulo.fill.fore_color.rgb = color_relleno
    circulo.line.fill.background()
    escribir_dentro_de_forma(circulo, texto, tamaño=tamaño)
    return circulo


def dibujar_corchete_msrp(diapositiva, x, y, ancho, texto, color_linea=None, color_texto=None):
    """
    Línea horizontal con dos marcas en los extremos y una etiqueta encima.
    Marca el ancho que corresponde al MSRP vigente, por encima de una barra de precio.
    'y' es la altura de la línea horizontal.
    Sin colores indicados, usa los del tema activo.
    """
    if color_linea is None:
        color_linea = COLOR_TEXTO_SECUNDARIO
    if color_texto is None:
        color_texto = COLOR_TEXTO
    grosor = 0.03
    alto_marca = 0.16

    dibujar_rectangulo(diapositiva, x, y, ancho, grosor, color_linea)
    dibujar_rectangulo(diapositiva, x, y, grosor, alto_marca, color_linea)
    dibujar_rectangulo(diapositiva, x + ancho - grosor, y, grosor, alto_marca, color_linea)

    escribir_texto(diapositiva, x, y - 0.38, ancho, 0.32, texto, tamaño=13, color=color_texto,
                   negrita=True, alineacion=PP_ALIGN.CENTER, anclaje=MSO_ANCHOR.BOTTOM)


def dibujar_barra_de_precio(diapositiva, x, y, alto, pulgadas_por_unidad, segmentos, tamaño_etiqueta=12,
                            ancho_minimo_para_etiqueta_dentro=1.15):
    """
    Dibuja una barra horizontal partida en segmentos proporcionales a su importe.
    Es el elemento visual que se repite en toda la presentación: pagado + descuentos = MSRP.

    segmentos: lista de diccionarios con las claves
        "nombre"  -> texto de la etiqueta (por ejemplo "Pagado")
        "texto"   -> (opcional) etiqueta completa que sustituye a "nombre + importe"
        "texto_exterior" -> (opcional) etiqueta cuando el segmento es estrecho y va debajo
        "importe" -> valor numérico del segmento
        "color"   -> color de relleno

    Si el segmento es más ancho que ancho_minimo_para_etiqueta_dentro, la etiqueta va dentro.
    Si es más estrecho, la etiqueta va debajo, en el color del segmento.

    Devuelve la coordenada x donde termina la barra.
    """
    posicion_x_actual = x

    for segmento in segmentos:
        ancho_del_segmento = segmento["importe"] * pulgadas_por_unidad
        # Por defecto la etiqueta es "nombre + importe"; con la clave "texto" se puede poner otra
        texto_de_la_etiqueta = segmento.get("texto", f"{segmento['nombre']}\n{formatear_importe(segmento['importe'])}")

        forma_del_segmento = dibujar_rectangulo(diapositiva, posicion_x_actual, y, ancho_del_segmento, alto,
                                                segmento["color"])

        if ancho_del_segmento >= ancho_minimo_para_etiqueta_dentro:
            escribir_dentro_de_forma(forma_del_segmento, texto_de_la_etiqueta, tamaño=tamaño_etiqueta)
        else:
            # Etiqueta exterior en una sola línea, para que no invada lo que haya debajo
            texto_de_la_etiqueta_exterior = segmento.get(
                "texto_exterior", f"{segmento['nombre']} {formatear_importe(segmento['importe'])}")
            ancho_etiqueta_exterior = 1.6
            centro_del_segmento = posicion_x_actual + ancho_del_segmento / 2
            escribir_texto(diapositiva, centro_del_segmento - ancho_etiqueta_exterior / 2, y + alto + 0.04,
                           ancho_etiqueta_exterior, 0.25, texto_de_la_etiqueta_exterior,
                           tamaño=tamaño_etiqueta - 1, color=segmento["color"], negrita=True,
                           alineacion=PP_ALIGN.CENTER)

        posicion_x_actual += ancho_del_segmento

    return posicion_x_actual


def poner_titulo(diapositiva, titulo, subtitulo=None):
    """Título de las diapositivas de contenido. Siempre en la misma posición y con el mismo formato."""
    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 0.4, ANCHO_UTIL, 0.75, titulo, tamaño=30,
                   color=COLOR_TEXTO, negrita=True, fuente=FUENTE_TITULOS, anclaje=MSO_ANCHOR.BOTTOM)
    if subtitulo is not None:
        escribir_texto(diapositiva, MARGEN_IZQUIERDO, 1.18, ANCHO_UTIL, 0.4, subtitulo, tamaño=16,
                       color=COLOR_TEXTO_SECUNDARIO)


def poner_notas(diapositiva, texto_de_las_notas):
    """Notas del orador (no se ven en la diapositiva)."""
    diapositiva.notes_slide.notes_text_frame.text = texto_de_las_notas


def añadir_transicion_de_fundido(diapositiva):
    """
    python-pptx no tiene una función para transiciones, así que se añade el XML a mano:
        <p:transition spd="med"><p:fade/></p:transition>
    Se coloca al final del elemento de la diapositiva, que es donde lo espera PowerPoint.
    """
    elemento_diapositiva = diapositiva._element
    transicion = etree.SubElement(elemento_diapositiva, qn("p:transition"))
    transicion.set("spd", "med")
    etree.SubElement(transicion, qn("p:fade"))


# Las tres partes de la presentación (la proyección va dentro de los gaps), con el mismo color que en el resumen y la agenda
PARTES_DE_LA_PRESENTACION = {
    "need": {"texto": "1 · THE NEED", "nombre_del_color": "COLOR_PAGADO"},
    "gap": {"texto": "2 · THE GAPS", "nombre_del_color": "COLOR_ERROR"},
    "plan": {"texto": "3 · THE PLAN", "nombre_del_color": "COLOR_DESCUENTO"},
}


def poner_parte_y_numero(diapositiva, parte, numero_de_diapositiva):
    """
    Arriba a la derecha: la parte a la que pertenece la diapositiva (en su color)
    y el número de diapositiva. Sirve de "estás aquí" sin repetir la agenda.
    """
    ancho_numero = 0.45
    x_numero = ANCHO_DIAPOSITIVA_PULGADAS - MARGEN_IZQUIERDO - ancho_numero
    escribir_texto(diapositiva, x_numero, 0.3, ancho_numero, 0.32, str(numero_de_diapositiva), tamaño=13,
                   color=COLOR_TEXTO_SECUNDARIO, negrita=True, alineacion=PP_ALIGN.RIGHT,
                   anclaje=MSO_ANCHOR.MIDDLE)

    if parte is None:
        return
    datos_de_la_parte = PARTES_DE_LA_PRESENTACION[parte]
    color_de_la_parte = globals()[datos_de_la_parte["nombre_del_color"]]
    ancho_etiqueta = 1.9
    etiqueta = dibujar_rectangulo(diapositiva, x_numero - 0.15 - ancho_etiqueta, 0.3, ancho_etiqueta, 0.32,
                                  COLOR_FONDO_CONTENIDO, redondeado=True, color_borde=color_de_la_parte)
    escribir_dentro_de_forma(etiqueta, datos_de_la_parte["texto"], tamaño=11, color=color_de_la_parte)


def crear_diapositiva(presentacion, color_fondo):
    """Diapositiva en blanco con fondo de color y transición de fundido."""
    diseño_en_blanco = presentacion.slide_layouts[6]
    diapositiva = presentacion.slides.add_slide(diseño_en_blanco)
    diapositiva.background.fill.solid()
    diapositiva.background.fill.fore_color.rgb = color_fondo
    añadir_transicion_de_fundido(diapositiva)
    return diapositiva


# =============================================================================
# 3. DIAPOSITIVAS
# =============================================================================

# Transacción real usada como ejemplo del modelo (consulta del correo a Dmitry).
# Todos los importes vienen de la tabla de transacciones eComm, excepto el MSRP,
# que es CustomerFullPrice de WebPriceList: no pertenece a la transacción.
MODELO_PEDIDO = "4K7019VAUH4"
MODELO_SKU = "KL1042GDAFF"
MODELO_MASTER_PARTNER_KEY = "294"
MODELO_FECHA = "3 May 2026"
MODELO_MONEDA = "EUR"
# País deducido del SKU (región DACH, como KL1048GDAFS) y del IVA del 19%. Verificable en WebPriceList
# (price_list_region e idioma) o con el país de facturación del pedido.
MODELO_PAIS = "Germany"


# Transacción teórica. Dado el modelo de la transacción que define Retención, estos son los datos
# conocidos (qué vendemos, dónde, con qué descuento, con qué comisión y con qué impuestos) y los
# importes que salen de ellos. Cada importe se redondea a céntimos, como en una factura.
TEORICA_SKU = "X"
TEORICA_FECHA = "2 May 2026"
TEORICA_PAIS = "Germany"
TEORICA_MONEDA = "EUR"
TEORICA_PROVEEDOR = "Nexway"

TEORICA_MSRP_CON_IMPUESTOS = 59.99
TEORICA_PORCENTAJE_DESCUENTO = 0.60
TEORICA_TIPO_IMPOSITIVO = 0.19        # IVA de Alemania
TEORICA_TIPO_PARTNER_FEE = 0.075      # comisión de Nexway sobre el importe sin impuestos

TEORICA_DESCUENTO = round(TEORICA_MSRP_CON_IMPUESTOS * TEORICA_PORCENTAJE_DESCUENTO, 2)
TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS = round(TEORICA_MSRP_CON_IMPUESTOS - TEORICA_DESCUENTO, 2)
TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS = round(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS / (1 + TEORICA_TIPO_IMPOSITIVO), 2)
TEORICA_IMPUESTOS = round(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS - TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS, 2)
TEORICA_PARTNER_FEE = round(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS * TEORICA_TIPO_PARTNER_FEE, 2)
TEORICA_DUE_TO_KL = round(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS - TEORICA_PARTNER_FEE, 2)
TEORICA_DESCUENTO_GUARDADO_SIN_IMPUESTOS = round(TEORICA_DESCUENTO / (1 + TEORICA_TIPO_IMPOSITIVO), 2)

MODELO_IMPORTE_CLIENTE_SIN_IMPUESTOS = 20.16          # CalculatedTransactionCustomerAmountWithoutTaxes
MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS = 23.99          # CalculatedTransactionCustomerAmountWithTaxes
MODELO_DUE_TO_KL = 18.65                              # TransactionDueToKLAmount
MODELO_PARTNER_FEE = 1.51                             # TransactionPartnerAmount
MODELO_IMPUESTOS = 3.83                               # TransactionTaxAmount
MODELO_DESCUENTO_GUARDADO_SIN_IMPUESTOS = 30.25       # TransactionDiscountAmount (sin impuestos)
MODELO_MSRP_CON_IMPUESTOS = 59.99                     # CustomerFullPrice, de WebPriceList
MODELO_DESCUENTO = round(MODELO_MSRP_CON_IMPUESTOS - MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS, 2)
# Tipo de IVA conocido. Se aplica sobre el importe sin impuestos: 20.16 × 19% = 3.83
MODELO_TIPO_IMPOSITIVO = 0.19


def dibujar_barra_del_modelo_con_zoom(diapositiva, x_barra, y_msrp, y_detalle, alto, segmentos_arriba,
                                      segmentos_detalle, texto_del_corchete, color_corchete, color_texto_corchete):
    """
    Dibuja el modelo de la transacción en dos niveles:
        arriba:  MSRP con impuestos = importe del cliente con impuestos + descuento
        debajo:  el importe del cliente ampliado a todo el ancho (due to KL + partner fee + impuestos)
    Dos líneas unen el importe del cliente con su ampliación, como un zoom.
    """
    ancho_total = 7.0
    escala_arriba = ancho_total / TEORICA_MSRP_CON_IMPUESTOS
    escala_detalle = ancho_total / TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS

    dibujar_corchete_msrp(diapositiva, x_barra, y_msrp - 0.2, ancho_total, texto_del_corchete,
                          color_linea=color_corchete, color_texto=color_texto_corchete)
    dibujar_barra_de_precio(diapositiva, x_barra, y_msrp, alto, escala_arriba, segmentos_arriba, tamaño_etiqueta=12)

    ancho_importe_cliente = TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS * escala_arriba
    dibujar_linea(diapositiva, x_barra, y_msrp + alto, x_barra, y_detalle, color_corchete)
    dibujar_linea(diapositiva, x_barra + ancho_importe_cliente, y_msrp + alto, x_barra + ancho_total, y_detalle,
                  color_corchete)

    dibujar_barra_de_precio(diapositiva, x_barra, y_detalle, alto, escala_detalle, segmentos_detalle,
                            tamaño_etiqueta=11, ancho_minimo_para_etiqueta_dentro=0.6)
    return escala_detalle


def diapositiva_portada(presentacion):
    """
    Modelo de la transacción con una transacción real. Es la vista ideal: funciona porque
    hemos añadido el MSRP (CustomerFullPrice) desde WebPriceList.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_PORTADA)

    escribir_texto(diapositiva, 0.8, 0.8, 7.0, 0.9, "Transaction model", tamaño=44,
                   color=COLOR_TEXTO_SOBRE_OSCURO, negrita=True, fuente=FUENTE_TITULOS)
    etiqueta_ideal = dibujar_rectangulo(diapositiva, 8.2, 1.05, 4.53, 0.5, COLOR_FONDO_PORTADA, redondeado=True,
                                        color_borde=COLOR_DESCUENTO_SECUNDARIO)
    escribir_dentro_de_forma(etiqueta_ideal, "THEORETICAL TRANSACTION", tamaño=14,
                             color=COLOR_DESCUENTO_SECUNDARIO)

    escribir_parrafos(diapositiva, 0.8, 1.8, 11.9, 1.8, [
        [fragmento("A transaction is unique: one product, bought or renewed, at one moment.", 20,
                   COLOR_TEXTO_SOBRE_OSCURO)],
        [fragmento("Known inputs: ", 14, COLOR_DESCUENTO_SECUNDARIO, negrita=True),
         fragmento(f"an order of SKU {TEORICA_SKU} on {TEORICA_FECHA} in {TEORICA_PAIS} · MSRP "
                   f"{formatear_euros(TEORICA_MSRP_CON_IMPUESTOS)} · discount "
                   f"{formatear_porcentaje(100 * TEORICA_PORCENTAJE_DESCUENTO)} · VAT "
                   f"{formatear_porcentaje(100 * TEORICA_TIPO_IMPOSITIVO)} · {TEORICA_PROVEEDOR} fee "
                   f"{formatear_porcentaje(100 * TEORICA_TIPO_PARTNER_FEE)}", 14,
                   COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO)],
        [fragmento("From what the customer pays, ", 15, COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO),
         fragmento("taxes", 15, COLOR_TEXTO_SOBRE_OSCURO, negrita=True),
         fragmento(" go to the State, the ", 15, COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO),
         fragmento("partner fee", 15, COLOR_TEXTO_SOBRE_OSCURO, negrita=True),
         fragmento(" goes to the e-commerce provider (also in our P&L), and ", 15,
                   COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO),
         fragmento("Due to KL", 15, COLOR_TEXTO_SOBRE_OSCURO, negrita=True),
         fragmento(" is what we earn net.", 15, COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO)],
    ], espacio_despues_de_parrafo=7)

    dibujar_barra_del_modelo_con_zoom(
        diapositiva, x_barra=0.8, y_msrp=4.45, y_detalle=5.4, alto=0.6,
        segmentos_arriba=[
            {"nombre": "Customer amount", "importe": TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS, "color": COLOR_PAGADO,
             "texto": f"Customer amount\nwith taxes {formatear_euros(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)}"},
            {"nombre": "Discount", "importe": TEORICA_DESCUENTO, "color": COLOR_DESCUENTO,
             "texto": f"Discount amount\n{formatear_euros(TEORICA_DESCUENTO)}"},
        ],
        segmentos_detalle=[
            {"nombre": "Due to KL", "importe": TEORICA_DUE_TO_KL, "color": COLOR_MSRP,
             "texto": f"Due to KL\n{formatear_euros(TEORICA_DUE_TO_KL)}"},
            {"nombre": "Partner fee", "importe": TEORICA_PARTNER_FEE, "color": COLOR_MORADO,
             "texto_exterior": f"Partner fee {formatear_euros(TEORICA_PARTNER_FEE)}"},
            {"nombre": "Taxes", "importe": TEORICA_IMPUESTOS, "color": COLOR_TEXTO_SECUNDARIO,
             "texto": f"Taxes\n{formatear_euros(TEORICA_IMPUESTOS)}"},
        ],
        texto_del_corchete=f"MSRP with taxes: {formatear_euros(TEORICA_MSRP_CON_IMPUESTOS)}",
        color_corchete=COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO, color_texto_corchete=COLOR_TEXTO_SOBRE_OSCURO)


    # ---- Tarjeta con las fórmulas: son a la vez el modelo y la comprobación ----
    x_tarjeta = 8.2
    y_tarjeta = 3.75
    ancho_tarjeta = 4.53
    dibujar_rectangulo(diapositiva, x_tarjeta, y_tarjeta, ancho_tarjeta, 3.35, COLOR_TARJETA, redondeado=True)

    porcentaje_descuento = 100 * TEORICA_DESCUENTO / TEORICA_MSRP_CON_IMPUESTOS
    formulas = [
        ("Due to KL + Partner fee + Taxes = Customer amount with taxes",
         f"{formatear_numero(TEORICA_DUE_TO_KL)} + {formatear_numero(TEORICA_PARTNER_FEE)} + "
         f"{formatear_numero(TEORICA_IMPUESTOS)} = {formatear_numero(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)}"),
        ("Customer amount with taxes + Discount amount = MSRP with taxes",
         f"{formatear_numero(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)} + {formatear_numero(TEORICA_DESCUENTO)} = "
         f"{formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)}"),
        ("MSRP with taxes − Customer amount with taxes = Discount amount",
         f"{formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)} − {formatear_numero(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)} = "
         f"{formatear_numero(TEORICA_DESCUENTO)}"),
        ("Discount amount / MSRP with taxes = % discount",
         f"{formatear_numero(TEORICA_DESCUENTO)} / {formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)} = "
         f"{formatear_porcentaje(porcentaje_descuento)}"),
    ]

    parrafos_de_la_tarjeta = [[
        fragmento("Transaction model", 16, COLOR_TEXTO_SOBRE_OSCURO, negrita=True, fuente=FUENTE_TITULOS),
        fragmento("  ·  these equalities are the check", 12, COLOR_DESCUENTO_SECUNDARIO),
    ]]
    # Cada fórmula en una línea y debajo su comprobación con los importes de la transacción
    for numero_formula, (formula, ejemplo_numerico) in enumerate(formulas, start=1):
        parrafos_de_la_tarjeta.append([
            fragmento(f"{numero_formula}  ", 12, COLOR_DESCUENTO_SECUNDARIO, negrita=True),
            fragmento(formula, 12, COLOR_TEXTO_SOBRE_OSCURO),
        ])
        parrafos_de_la_tarjeta.append([
            fragmento(f"     {ejemplo_numerico}  {SIMBOLO_CORRECTO}", 12, COLOR_CORRECTO, negrita=True),
        ])
    escribir_parrafos(diapositiva, x_tarjeta + 0.25, y_tarjeta + 0.2, ancho_tarjeta - 0.45, 3.0,
                      parrafos_de_la_tarjeta, espacio_despues_de_parrafo=4)

    poner_notas(diapositiva, "Given the transaction model that Retention defines, an order with known inputs: what "
                             "we sell, where, with which discount, which fee and which taxes. The amounts follow from "
                             "them and all four equalities close. Later we look for a real one.")


def diapositiva_modelo_en_porcentajes(presentacion):
    """
    La misma transacción real en porcentajes. Cada porcentaje tiene su referencia:
        descuento   -> MSRP con impuestos
        impuestos   -> importe del cliente sin impuestos (así sale el tipo de IVA conocido, 19%)
        partner fee -> importe del cliente sin impuestos (el proveedor cobra de lo que cobramos nosotros)
        due to KL   -> importe del cliente sin impuestos
    Además se muestra la cuenta de cómo se quita el IVA del importe con impuestos.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_PORTADA)
    poner_titulo(diapositiva, "The same transaction in percentages",
                 "Same theoretical transaction. Each percentage has its own reference")

    porcentaje_descuento = 100 * TEORICA_DESCUENTO / TEORICA_MSRP_CON_IMPUESTOS
    porcentaje_impuestos = 100 * TEORICA_IMPUESTOS / TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS
    porcentaje_partner_fee = 100 * TEORICA_PARTNER_FEE / TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS
    porcentaje_due_to_kl = 100 * TEORICA_DUE_TO_KL / TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS
    factor_de_iva = 1 + TEORICA_TIPO_IMPOSITIVO

    x_barra = MARGEN_IZQUIERDO
    escala_detalle = dibujar_barra_del_modelo_con_zoom(
        diapositiva, x_barra=x_barra, y_msrp=2.5, y_detalle=3.5, alto=0.7,
        segmentos_arriba=[
            {"nombre": "Customer amount", "importe": TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS, "color": COLOR_PAGADO,
             "texto": f"Customer amount\nwith taxes {formatear_euros(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)}"},
            {"nombre": "Discount", "importe": TEORICA_DESCUENTO, "color": COLOR_DESCUENTO,
             "texto": f"Discount {formatear_euros(TEORICA_DESCUENTO)}"
                      f"\n{formatear_porcentaje(porcentaje_descuento)} of MSRP"},
        ],
        segmentos_detalle=[
            {"nombre": "Due to KL", "importe": TEORICA_DUE_TO_KL, "color": COLOR_MSRP,
             "texto": f"Due to KL {formatear_euros(TEORICA_DUE_TO_KL)}\n{formatear_porcentaje(porcentaje_due_to_kl)}"},
            {"nombre": "Partner fee", "importe": TEORICA_PARTNER_FEE, "color": COLOR_MORADO,
             "texto_exterior": f"Partner fee {formatear_euros(TEORICA_PARTNER_FEE)} · "
                               f"{formatear_porcentaje(porcentaje_partner_fee)}"},
            {"nombre": "Taxes", "importe": TEORICA_IMPUESTOS, "color": COLOR_TEXTO_SECUNDARIO,
             "texto": f"VAT {formatear_euros(TEORICA_IMPUESTOS)}\n{formatear_porcentaje(porcentaje_impuestos)}"},
        ],
        texto_del_corchete=f"MSRP with taxes: {formatear_euros(TEORICA_MSRP_CON_IMPUESTOS)}  ·  reference for the discount",
        color_corchete=COLOR_TEXTO_SECUNDARIO, color_texto_corchete=COLOR_TEXTO)

    # Corchete debajo: importe del cliente sin impuestos, referencia de impuestos, fee y Due to KL
    ancho_sin_impuestos = TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS * escala_detalle
    y_linea = 4.75
    grosor = 0.03
    dibujar_rectangulo(diapositiva, x_barra, y_linea, ancho_sin_impuestos, grosor, COLOR_TEXTO_SECUNDARIO)
    dibujar_rectangulo(diapositiva, x_barra, y_linea - 0.16, grosor, 0.16 + grosor, COLOR_TEXTO_SECUNDARIO)
    dibujar_rectangulo(diapositiva, x_barra + ancho_sin_impuestos - grosor, y_linea - 0.16, grosor, 0.16 + grosor,
                       COLOR_TEXTO_SECUNDARIO)
    escribir_texto(diapositiva, x_barra, y_linea + 0.08, 7.0, 0.32,
                   f"Customer amount without taxes: {formatear_euros(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)}"
                   f"  ·  reference for VAT, partner fee and Due to KL", tamaño=12, negrita=True)

    # Cómo se quita el IVA: la cuenta, no el peso del IVA en el total
    importe_sin_impuestos_calculado = TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS / factor_de_iva
    escribir_parrafos(diapositiva, x_barra, 5.35, 7.0, 1.75, [
        [fragmento(f"How VAT is removed ({TEORICA_PAIS}, {formatear_porcentaje(100 * TEORICA_TIPO_IMPOSITIVO)} VAT)", 15,
                   negrita=True)],
        [fragmento("Without taxes: ", 14, negrita=True),
         fragmento(f"{formatear_numero(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)} / {factor_de_iva:.2f} = "
                   f"{formatear_numero(importe_sin_impuestos_calculado)}", 14, COLOR_TEXTO_SECUNDARIO)],
        [fragmento("VAT: ", 14, negrita=True),
         fragmento(f"{formatear_numero(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)} − "
                   f"{formatear_numero(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)} = "
                   f"{formatear_numero(TEORICA_IMPUESTOS)} = {formatear_numero(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)} × "
                   f"{formatear_porcentaje(100 * TEORICA_TIPO_IMPOSITIVO)}", 14, COLOR_TEXTO_SECUNDARIO)],
        [fragmento("Partner fee and Due to KL: ", 14, negrita=True),
         fragmento("the provider charges on what we collect, not on the taxes.", 14, COLOR_TEXTO_SECUNDARIO)],
    ], espacio_despues_de_parrafo=4)

    # ---- Tarjeta con cada porcentaje, su cálculo y su referencia ----
    x_tarjeta = 8.0
    ancho_tarjeta = ANCHO_DIAPOSITIVA_PULGADAS - MARGEN_IZQUIERDO - x_tarjeta
    dibujar_rectangulo(diapositiva, x_tarjeta, 1.75, ancho_tarjeta, 5.25, COLOR_TARJETA, redondeado=True)

    filas = [
        ("Discount", porcentaje_descuento,
         f"{formatear_numero(TEORICA_DESCUENTO)} / {formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)}",
         "of MSRP with taxes", COLOR_DESCUENTO),
        ("VAT", porcentaje_impuestos,
         f"{formatear_numero(TEORICA_IMPUESTOS)} / {formatear_numero(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)}",
         "of customer amount without taxes", COLOR_TEXTO_SECUNDARIO),
        ("Partner fee", porcentaje_partner_fee,
         f"{formatear_numero(TEORICA_PARTNER_FEE)} / {formatear_numero(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)}",
         "of customer amount without taxes", COLOR_MORADO),
        ("Due to KL", porcentaje_due_to_kl,
         f"{formatear_numero(TEORICA_DUE_TO_KL)} / {formatear_numero(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)}",
         "of customer amount without taxes", COLOR_MSRP),
    ]
    for numero_fila, (nombre, porcentaje, calculo, referencia, color) in enumerate(filas):
        y_fila = 1.95 + numero_fila * 0.85
        escribir_texto(diapositiva, x_tarjeta + 0.3, y_fila, 2.5, 0.4, nombre, tamaño=16, negrita=True)
        escribir_texto(diapositiva, x_tarjeta + ancho_tarjeta - 1.55, y_fila, 1.3, 0.4,
                       formatear_porcentaje(porcentaje), tamaño=20, color=color, negrita=True,
                       alineacion=PP_ALIGN.RIGHT)
        escribir_texto(diapositiva, x_tarjeta + 0.3, y_fila + 0.4, ancho_tarjeta - 0.55, 0.35,
                       f"{calculo}  {referencia}", tamaño=12, color=COLOR_TEXTO_SECUNDARIO)

    # Las tres partes del importe del cliente sobre la misma referencia: el importe sin impuestos = 100%
    porcentaje_con_impuestos = 100 + porcentaje_impuestos
    importe_con_impuestos_calculado = TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS * factor_de_iva
    dibujar_rectangulo(diapositiva, x_tarjeta + 0.2, 5.3, ancho_tarjeta - 0.4, 1.55, COLOR_FONDO_PORTADA,
                       redondeado=True)
    escribir_parrafos(diapositiva, x_tarjeta + 0.4, 5.4, ancho_tarjeta - 0.75, 1.4, [
        [fragmento("Same reference: amount without taxes = 100%", 12, COLOR_DESCUENTO_SECUNDARIO, negrita=True)],
        [fragmento(f"Due to KL {formatear_porcentaje(porcentaje_due_to_kl)} + partner fee "
                   f"{formatear_porcentaje(porcentaje_partner_fee)} = 100%", 13, negrita=True)],
        [fragmento(f"+ VAT {formatear_porcentaje(porcentaje_impuestos)} = "
                   f"{formatear_porcentaje(porcentaje_con_impuestos)} = amount with taxes", 13, negrita=True)],
        [fragmento(f"{formatear_numero(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)} × {factor_de_iva:.2f} = "
                   f"{formatear_numero(importe_con_impuestos_calculado)}", 13, COLOR_CORRECTO, negrita=True)],
    ], espacio_despues_de_parrafo=3)

    poner_notas(diapositiva, "Same theoretical transaction in percentages. Discount: 60% of the MSRP. VAT: 19% of "
                             "the amount without taxes, the known German rate. Partner fee 7.5% (Nexway) and Due "
                             "to KL 92.5% of the amount without taxes. The three use the same reference, so they add up to "
                             "119%: the amount without taxes × 1.19 is the amount with taxes.")


CONSULTA_DEL_CORREO = [
    ("SELECT [OnlineSalesTransactionDateKey], [ExternalOrderId], base.[SKU]", ""),
    ("      ,[CalculatedTransactionCustomerAmountWithTaxes]", ""),
    ("      ,[TransactionDiscountAmount], [TransactionTaxAmount]", ""),
    ("      ,price.[CustomerFullPrice]", "1"),
    ("FROM #example AS base", ""),
    ("LEFT JOIN [DWH].[sales].[WebPriceLists] AS price", "1"),
    ("  ON  price.[MasterPartnerKey] = base.[MasterPartnerKey]", ""),
    ("  AND price.SKU = 'KL1042GDAFG'", "2"),
    ("  AND CONVERT(DATE, CAST(base.[OnlineSalesTransactionDateKey]", "3"),
    ("                         AS varchar(8)), 112)", ""),
    ("      BETWEEN price.[PriceListEffectiveTimestamp]", ""),
    ("          AND price.[PriceListExpirationTimestamp]", ""),
    ("  AND CONVERT(DATE, CAST(base.[OnlineSalesTransactionDateKey]", ""),
    ("                         AS varchar(8)), 112)", ""),
    ("      BETWEEN price.[PriceEffectiveTimestamp]", ""),
    ("          AND price.[PriceExpirationTimestamp]", ""),
    ("  AND eStore = '[N/A]'", ""),
    ("  AND price.[CurrencyKey] = base.TransactionCurrencyKey", ""),
]


def diapositiva_definiciones_del_modelo(presentacion):
    """
    Definiciones del modelo de la transacción, en orden de cálculo, con los importes de la transacción
    teórica, y las igualdades que se tienen que cumplir.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_PORTADA)
    poner_titulo(diapositiva, "Transaction model: definitions and checks",
                 "Each amount, how it is calculated, and the equalities that must always hold")

    factor_de_iva = 1 + TEORICA_TIPO_IMPOSITIVO
    definiciones = [
        ("MSRP with taxes", "Full price for product × country × currency at that moment",
         formatear_euros(TEORICA_MSRP_CON_IMPUESTOS)),
        ("Discount %", "Effective discount of the transaction (demand level, coupon, ladder, offer)",
         formatear_porcentaje(100 * TEORICA_PORCENTAJE_DESCUENTO)),
        ("Discount amount", "MSRP with taxes × discount %", formatear_euros(TEORICA_DESCUENTO)),
        ("Customer amount with taxes", "MSRP with taxes − discount amount",
         formatear_euros(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)),
        ("VAT rate", f"Known rate of the country ({TEORICA_PAIS})",
         formatear_porcentaje(100 * TEORICA_TIPO_IMPOSITIVO)),
        ("Customer amount without taxes", f"Customer amount with taxes / {factor_de_iva:.2f}",
         formatear_euros(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)),
        ("VAT amount", "Customer amount without taxes × VAT rate", formatear_euros(TEORICA_IMPUESTOS)),
        ("Partner fee", f"Customer amount without taxes × fee rate "
                        f"({formatear_porcentaje(100 * TEORICA_TIPO_PARTNER_FEE)}, {TEORICA_PROVEEDOR})",
         formatear_euros(TEORICA_PARTNER_FEE)),
        ("Due to KL", "Customer amount without taxes − partner fee", formatear_euros(TEORICA_DUE_TO_KL)),
    ]

    # ---- Izquierda: definiciones en orden de cálculo ----
    x_tabla = MARGEN_IZQUIERDO
    ancho_tabla = 7.6
    alto_fila = 0.52
    dibujar_rectangulo(diapositiva, x_tabla, 1.75, ancho_tabla, 0.25 + alto_fila * len(definiciones), COLOR_TARJETA,
                       redondeado=True)
    for numero_fila, (concepto, calculo, valor) in enumerate(definiciones):
        y_fila = 1.87 + numero_fila * alto_fila
        escribir_parrafos(diapositiva, x_tabla + 0.25, y_fila, ancho_tabla - 1.7, alto_fila, [
            [fragmento(concepto, 13, COLOR_TEXTO_SOBRE_OSCURO, negrita=True)],
            [fragmento(calculo, 11, COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO)],
        ], anclaje=MSO_ANCHOR.MIDDLE, espacio_despues_de_parrafo=0)
        escribir_texto(diapositiva, x_tabla + ancho_tabla - 1.45, y_fila, 1.2, alto_fila, valor, tamaño=14,
                       color=COLOR_DESCUENTO_SECUNDARIO, negrita=True, alineacion=PP_ALIGN.RIGHT,
                       anclaje=MSO_ANCHOR.MIDDLE)

    # ---- Derecha: comprobaciones que se tienen que cumplir ----
    x_tarjeta = x_tabla + ancho_tabla + 0.25
    ancho_tarjeta = ANCHO_DIAPOSITIVA_PULGADAS - MARGEN_IZQUIERDO - x_tarjeta
    dibujar_rectangulo(diapositiva, x_tarjeta, 1.75, ancho_tarjeta, 0.25 + alto_fila * len(definiciones),
                       COLOR_TARJETA, redondeado=True, color_borde=COLOR_CORRECTO)

    porcentaje_descuento = 100 * TEORICA_DESCUENTO / TEORICA_MSRP_CON_IMPUESTOS
    comprobaciones = [
        ("Due to KL + partner fee + VAT = customer amount with taxes",
         f"{formatear_numero(TEORICA_DUE_TO_KL)} + {formatear_numero(TEORICA_PARTNER_FEE)} + "
         f"{formatear_numero(TEORICA_IMPUESTOS)} = {formatear_numero(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)}"),
        ("Customer amount + discount = MSRP (same tax basis)",
         f"{formatear_numero(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)} + {formatear_numero(TEORICA_DESCUENTO)} = "
         f"{formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)}"),
        ("Discount / MSRP = discount %",
         f"{formatear_numero(TEORICA_DESCUENTO)} / {formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)} = "
         f"{formatear_porcentaje(porcentaje_descuento)}"),
        ("VAT / amount without taxes = country VAT rate",
         f"{formatear_numero(TEORICA_IMPUESTOS)} / {formatear_numero(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)} = "
         f"{formatear_porcentaje(100 * TEORICA_IMPUESTOS / TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)}"),
        ("Partner fee / amount without taxes = agreed fee rate",
         f"{formatear_numero(TEORICA_PARTNER_FEE)} / {formatear_numero(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)} = "
         f"{formatear_porcentaje(100 * TEORICA_PARTNER_FEE / TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)}"),
    ]
    parrafos_de_comprobaciones = [[fragmento("Must always hold", 16, COLOR_CORRECTO, negrita=True,
                                             fuente=FUENTE_TITULOS)]]
    for numero, (regla, cuenta) in enumerate(comprobaciones, start=1):
        parrafos_de_comprobaciones.append([
            fragmento(f"{numero}  ", 12, COLOR_DESCUENTO_SECUNDARIO, negrita=True),
            fragmento(regla, 12, COLOR_TEXTO_SOBRE_OSCURO, negrita=True),
        ])
        parrafos_de_comprobaciones.append([
            fragmento(f"     {cuenta}  {SIMBOLO_CORRECTO}", 12, COLOR_CORRECTO, negrita=True),
        ])
    parrafos_de_comprobaciones.append([
        fragmento("Rounding tolerance ±0.01. If one fails, the record is wrong.", 11,
                  COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO, cursiva=True),
    ])
    escribir_parrafos(diapositiva, x_tarjeta + 0.25, 1.9, ancho_tarjeta - 0.4, 0.25 + alto_fila * len(definiciones),
                      parrafos_de_comprobaciones, espacio_despues_de_parrafo=3)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 6.75, ANCHO_UTIL, 0.4,
                   "Known inputs: what we sell, where, with which discount, fee and taxes. Everything else is "
                   "calculated.", tamaño=14, color=COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO, cursiva=True)

    poner_notas(diapositiva, "The definitions in calculation order, with the theoretical transaction. The equalities "
                             "on the right are the checks: they must hold for every transaction. Slides on columns "
                             "and checks translate them into data.")


# Relación entre los conceptos del modelo y los campos de la tabla de transacciones eComm.
# Dos conceptos son los anclajes: se toman como correctos y el resto se reconstruye a partir de ellos.
#   - importe del cliente con impuestos: lo que pagó el cliente
#   - Due to KL: lo que recibimos
# Cada fila: concepto del modelo, campo en la tabla (o dónde está), cómo se comprueba, si es anclaje.
MAPPING_DE_CONCEPTOS = [
    ("Customer amount with taxes", "CalculatedTransactionCustomerAmountWithTaxes",
     "Anchor: taken as correct (what the customer paid)", True),
    ("Due to KL", "TransactionDueToKLAmount",
     "Anchor: taken as correct (what we receive)", True),
    ("Customer amount without taxes", "CalculatedTransactionCustomerAmountWithoutTaxes",
     "= with taxes / (1 + VAT rate)", False),
    ("VAT amount", "TransactionTaxAmount",
     "= with taxes − without taxes", False),
    ("VAT rate", "not stored",
     "= with taxes / without taxes − 1, compared with the country rate", False),
    ("Partner fee", "TransactionPartnerAmount",
     "= without taxes − Due to KL, compared with the fee rate", False),
    ("MSRP with taxes", "not stored · CustomerFullPrice in WebPriceList",
     "Needed for the discount. Must come with the transaction", False),
    ("Discount amount (with taxes)", "TransactionDiscountAmount (stored without taxes)",
     "= MSRP with taxes − customer amount with taxes", False),
    ("Discount %", "not stored",
     "= discount amount / MSRP with taxes", False),
]


def diapositiva_mapping_de_conceptos(presentacion):
    """
    Cómo se comprueba una transacción real: cada concepto del modelo con su campo en la tabla.
    Los dos anclajes se toman como correctos y el resto se reconstruye con las reglas del modelo.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_PORTADA)
    poner_titulo(diapositiva, "Checking real data, concept by concept",
                 "Each concept of the model has its field in the table. Two of them are the anchors")

    cabeceras = ["Concept (model)", "Field in the table", "How we check it"]
    anchos = [3.0, 4.45, 4.68]
    x_tabla = MARGEN_IZQUIERDO
    ancho_tabla = sum(anchos)

    def x_de_columna(numero_columna):
        return x_tabla + sum(anchos[:numero_columna])

    dibujar_rectangulo(diapositiva, x_tabla, 1.75, ancho_tabla, 0.42, COLOR_TARJETA, redondeado=True)
    for numero_columna, cabecera in enumerate(cabeceras):
        escribir_texto(diapositiva, x_de_columna(numero_columna) + 0.15, 1.75, anchos[numero_columna] - 0.2, 0.42,
                       cabecera, tamaño=12, color=COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO, negrita=True,
                       anclaje=MSO_ANCHOR.MIDDLE)

    alto_fila = 0.47
    for numero_fila, (concepto, campo, como_se_comprueba, es_anclaje) in enumerate(MAPPING_DE_CONCEPTOS):
        y_fila = 2.22 + numero_fila * alto_fila
        if es_anclaje:
            dibujar_rectangulo(diapositiva, x_tabla, y_fila + 0.03, ancho_tabla, alto_fila - 0.06,
                               COLOR_CORRECTO_SUAVE, redondeado=True)
        no_esta_en_la_tabla = campo.startswith("not stored")
        escribir_texto(diapositiva, x_de_columna(0) + 0.15, y_fila, anchos[0] - 0.2, alto_fila, concepto, tamaño=13,
                       color=COLOR_TEXTO_SOBRE_OSCURO, negrita=True, anclaje=MSO_ANCHOR.MIDDLE)
        escribir_texto(diapositiva, x_de_columna(1) + 0.15, y_fila, anchos[1] - 0.2, alto_fila,
                       resaltar_entidades(campo), tamaño=11,
                       color=COLOR_ERROR if no_esta_en_la_tabla else COLOR_PAGADO, anclaje=MSO_ANCHOR.MIDDLE)
        escribir_texto(diapositiva, x_de_columna(2) + 0.15, y_fila, anchos[2] - 0.2, alto_fila, como_se_comprueba,
                       tamaño=12, color=COLOR_CORRECTO if es_anclaje else COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO,
                       negrita=es_anclaje, anclaje=MSO_ANCHOR.MIDDLE)
        if numero_fila < len(MAPPING_DE_CONCEPTOS) - 1 and not es_anclaje:
            dibujar_rectangulo(diapositiva, x_tabla, y_fila + alto_fila - 0.01, ancho_tabla, 0.01, COLOR_TARJETA)

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 6.55, ANCHO_UTIL, 0.7, [
        [fragmento("Anchors: ", 14, COLOR_CORRECTO, negrita=True),
         fragmento("what the customer paid and what we receive. Every other amount must follow from them.", 14,
                   COLOR_TEXTO_SOBRE_OSCURO)],
        [fragmento("Open point: ", 14, COLOR_DESCUENTO_SECUNDARIO, negrita=True),
         fragmento("some transactions have no VAT, and the table does not say it.", 14,
                   COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO)],
    ], espacio_despues_de_parrafo=3)

    poner_notas(diapositiva, "We take two amounts as correct: what the customer paid, with taxes, and Due to KL. "
                             "With the model rules we rebuild the other concepts and compare them, one by one, with "
                             "the fields of the table.")


def comparar_con_la_tabla(valor_calculado, valor_en_la_tabla, tolerancia=0.01):
    """True si el valor calculado y el de la tabla coinciden dentro de la tolerancia de redondeo."""
    return abs(round(valor_calculado - valor_en_la_tabla, 2)) <= tolerancia


def diapositiva_caso_1_el_ejemplo_en_los_datos(presentacion):
    """
    Caso 1: una transacción real, comprobada concepto a concepto con el mapping.
    Anclajes: importe del cliente con impuestos y Due to KL. El resto se reconstruye con las reglas del modelo.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_PORTADA)
    poner_titulo(diapositiva, "Case 1: we looked for a real one. Does it match the model?",
                 f"Real order {MODELO_PEDIDO} · SKU {MODELO_SKU} · {MODELO_FECHA} · {MODELO_PAIS} · "
                 f"VAT {formatear_porcentaje(100 * MODELO_TIPO_IMPOSITIVO)}")

    con_impuestos = MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS
    due_to_kl = MODELO_DUE_TO_KL
    factor_de_iva = 1 + MODELO_TIPO_IMPOSITIVO

    # Reconstrucción a partir de los anclajes
    sin_impuestos = round(con_impuestos / factor_de_iva, 2)
    impuestos = round(con_impuestos - sin_impuestos, 2)
    tipo_de_iva_deducido = 100 * (MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS / MODELO_IMPORTE_CLIENTE_SIN_IMPUESTOS - 1)
    partner_fee = round(sin_impuestos - due_to_kl, 2)
    tipo_de_fee = 100 * partner_fee / sin_impuestos
    descuento = round(MODELO_MSRP_CON_IMPUESTOS - con_impuestos, 2)
    descuento_de_la_tabla_con_impuestos = round(MODELO_DESCUENTO_GUARDADO_SIN_IMPUESTOS * factor_de_iva, 2)
    porcentaje_descuento = 100 * descuento / MODELO_MSRP_CON_IMPUESTOS

    def resultado_de_la_comparacion(calculado, en_la_tabla):
        if comparar_con_la_tabla(calculado, en_la_tabla):
            return SIMBOLO_CORRECTO, COLOR_CORRECTO
        return f"{SIMBOLO_ERROR}  differs", COLOR_ERROR

    # Cada fila: concepto, cálculo desde los anclajes, valor en la tabla, (texto, color) del resultado
    filas = [
        ("Customer amount with taxes", "anchor", formatear_numero(con_impuestos), ("Anchor", COLOR_PAGADO)),
        ("Due to KL", "anchor", formatear_numero(due_to_kl), ("Anchor", COLOR_PAGADO)),
        ("Customer amount without taxes",
         f"{formatear_numero(con_impuestos)} / {factor_de_iva:.2f} = {formatear_numero(sin_impuestos)}",
         formatear_numero(MODELO_IMPORTE_CLIENTE_SIN_IMPUESTOS),
         resultado_de_la_comparacion(sin_impuestos, MODELO_IMPORTE_CLIENTE_SIN_IMPUESTOS)),
        ("VAT amount",
         f"{formatear_numero(con_impuestos)} − {formatear_numero(sin_impuestos)} = {formatear_numero(impuestos)}",
         formatear_numero(MODELO_IMPUESTOS), resultado_de_la_comparacion(impuestos, MODELO_IMPUESTOS)),
        ("VAT rate",
         f"{formatear_numero(MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS)} / "
         f"{formatear_numero(MODELO_IMPORTE_CLIENTE_SIN_IMPUESTOS)} − 1 = {formatear_porcentaje(tipo_de_iva_deducido)}",
         "not stored", (f"{SIMBOLO_CORRECTO}  = {MODELO_PAIS}, inferred", COLOR_CORRECTO)),
        ("Partner fee",
         f"{formatear_numero(sin_impuestos)} − {formatear_numero(due_to_kl)} = {formatear_numero(partner_fee)} "
         f"({formatear_porcentaje(tipo_de_fee)})",
         formatear_numero(MODELO_PARTNER_FEE), resultado_de_la_comparacion(partner_fee, MODELO_PARTNER_FEE)),
        ("MSRP with taxes", f"{formatear_numero(MODELO_MSRP_CON_IMPUESTOS)} only from WebPriceList",
         "not stored", (f"{SIMBOLO_ERROR}  missing", COLOR_ERROR)),
        ("Discount amount (with taxes)",
         f"{formatear_numero(MODELO_MSRP_CON_IMPUESTOS)} − {formatear_numero(con_impuestos)} = "
         f"{formatear_numero(descuento)}",
         f"{formatear_numero(MODELO_DESCUENTO_GUARDADO_SIN_IMPUESTOS)} (no tax)",
         (f"{SIMBOLO_ERROR}  other tax basis", COLOR_ERROR)),
        ("Discount %",
         f"{formatear_numero(descuento)} / {formatear_numero(MODELO_MSRP_CON_IMPUESTOS)} = "
         f"{formatear_porcentaje(porcentaje_descuento)}",
         "not stored", (f"{SIMBOLO_ERROR}  missing", COLOR_ERROR)),
    ]

    cabeceras = ["Concept", "Rebuilt from the anchors", "In the table", "Result"]
    anchos = [2.45, 3.15, 1.35, 1.75]
    x_tabla = MARGEN_IZQUIERDO
    ancho_tabla = sum(anchos)

    def x_de_columna(numero_columna):
        return x_tabla + sum(anchos[:numero_columna])

    dibujar_rectangulo(diapositiva, x_tabla, 1.75, ancho_tabla, 0.42, COLOR_TARJETA, redondeado=True)
    for numero_columna, cabecera in enumerate(cabeceras):
        escribir_texto(diapositiva, x_de_columna(numero_columna) + 0.12, 1.75, anchos[numero_columna] - 0.17, 0.42,
                       cabecera, tamaño=12, color=COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO, negrita=True,
                       anclaje=MSO_ANCHOR.MIDDLE)

    alto_fila = 0.44
    for numero_fila, (concepto, calculo, en_la_tabla, (texto_resultado, color_resultado)) in enumerate(filas):
        y_fila = 2.22 + numero_fila * alto_fila
        es_anclaje = calculo == "anchor"
        if es_anclaje:
            dibujar_rectangulo(diapositiva, x_tabla, y_fila + 0.03, ancho_tabla, alto_fila - 0.06,
                               COLOR_TARJETA, redondeado=True)
        no_esta_en_la_tabla = en_la_tabla == "not stored" or "(no tax)" in en_la_tabla
        celdas = [
            (concepto, COLOR_TEXTO_SOBRE_OSCURO, True, 12),
            ("taken as correct" if es_anclaje else calculo, COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO, False, 12),
            (en_la_tabla, COLOR_ERROR if no_esta_en_la_tabla else COLOR_TEXTO_SOBRE_OSCURO, True, 12),
            (texto_resultado, color_resultado, True, 12),
        ]
        for numero_columna, (texto, color, negrita, tamaño) in enumerate(celdas):
            escribir_texto(diapositiva, x_de_columna(numero_columna) + 0.12, y_fila, anchos[numero_columna] - 0.17,
                           alto_fila, resaltar_entidades(texto) if numero_columna == 1 else texto, tamaño=tamaño,
                           color=color, negrita=negrita, anclaje=MSO_ANCHOR.MIDDLE)
        if numero_fila < len(filas) - 1 and not es_anclaje:
            dibujar_rectangulo(diapositiva, x_tabla, y_fila + alto_fila - 0.01, ancho_tabla, 0.01, COLOR_TARJETA)

    # ---- Derecha: lectura ----
    x_derecha = x_tabla + ancho_tabla + 0.3
    ancho_derecha = ANCHO_DIAPOSITIVA_PULGADAS - MARGEN_IZQUIERDO - x_derecha
    importe_mal_sumado = MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS + MODELO_DESCUENTO_GUARDADO_SIN_IMPUESTOS

    escribir_parrafos(diapositiva, x_derecha, 1.75, ancho_derecha, 1.3, [
        [fragmento("What matches", 15, COLOR_CORRECTO, negrita=True)],
        [fragmento("From the two anchors we rebuild the amount without taxes, the VAT, the 19% rate and the "
                   "partner fee, to the cent.", 13, COLOR_TEXTO_SOBRE_OSCURO)],
    ], espacio_despues_de_parrafo=3)

    escribir_parrafos(diapositiva, x_derecha, 3.15, ancho_derecha, 1.75, [
        [fragmento("What does not", 15, COLOR_ERROR, negrita=True)],
        [fragmento(f"No MSRP, so no discount %. The discount is stored without taxes: "
                   f"{formatear_numero(MODELO_DESCUENTO_GUARDADO_SIN_IMPUESTOS)} × {factor_de_iva:.2f} = "
                   f"{formatear_numero(descuento_de_la_tabla_con_impuestos)}.", 13, COLOR_TEXTO_SOBRE_OSCURO)],
        [fragmento(f"{formatear_numero(MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS)} + "
                   f"{formatear_numero(MODELO_DESCUENTO_GUARDADO_SIN_IMPUESTOS)} = "
                   f"{formatear_euros(importe_mal_sumado)} ≠ {formatear_euros(MODELO_MSRP_CON_IMPUESTOS)}  "
                   f"{SIMBOLO_ERROR}", 14, COLOR_ERROR, negrita=True)],
    ], espacio_despues_de_parrafo=3)

    escribir_parrafos(diapositiva, x_derecha, 5.0, ancho_derecha, 1.2, [
        [fragmento("Open point: no VAT", 15, COLOR_DESCUENTO_SECUNDARIO, negrita=True)],
        [fragmento("If with taxes = without taxes, is the VAT 0, or a missing value? The table does not say it.",
                   13, COLOR_TEXTO_SOBRE_OSCURO)],
    ], espacio_despues_de_parrafo=3)

    banda = dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 6.45, ANCHO_UTIL, 0.6, COLOR_TARJETA, redondeado=True,
                               color_borde=COLOR_ERROR)
    rellenar_marco_de_texto(banda.text_frame, [[
        fragmento("The split follows the model. The MSRP and the discount on its tax basis are missing.", 16,
                  COLOR_TEXTO_SOBRE_OSCURO, negrita=True),
    ]], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, 0)

    poner_notas(diapositiva, "We take the paid amount with taxes and Due to KL as correct. From them, the amount "
                             "without taxes, the VAT, the 19% rate and the 7.5% fee all match the table. The MSRP is "
                             "not in the table, and the discount is stored without taxes: 30.25 × 1.19 = 36.00, the "
                             "discount with taxes. When there is no VAT, the table does not tell us.")


def diapositiva_realidad_hoy(presentacion):
    """
    La misma transacción real, como está hoy en los datos: la tabla de transacciones no lleva el MSRP.
    Se muestra la consulta del correo a Dmitry, con las tres suposiciones que hace.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_PORTADA)
    poner_titulo(diapositiva, "Case 2: the workaround, an external MSRP",
                 f"To find {formatear_euros(MODELO_MSRP_CON_IMPUESTOS)} for order {MODELO_PEDIDO} we join "
                 f"WebPriceList. It works here, with assumptions")

    # ---- Columna izquierda: de dónde sale el 59.99 y qué pasa cuando lo tenemos ----
    x_izquierda = MARGEN_IZQUIERDO
    ancho_izquierda = 5.2

    escribir_parrafos(diapositiva, x_izquierda, 1.75, ancho_izquierda, 1.6, [
        [fragmento("1 · Where 59.99 comes from", 14, COLOR_TEXTO_SECUNDARIO, negrita=True)],
        [fragmento("CustomerFullPrice", 14, negrita=True),
         fragmento(" from ", 14), fragmento("WebPriceList", 14, negrita=True),
         fragmento(", not from the transaction. If it belonged to the transaction, it would be called "
                   "TransactionCustomerFullPrice.", 14)],
    ], espacio_despues_de_parrafo=4)

    escribir_parrafos(diapositiva, x_izquierda, 3.25, ancho_izquierda, 1.3, [
        [fragmento("2 · With it, the example closes", 14, COLOR_TEXTO_SECUNDARIO, negrita=True)],
        [fragmento(f"{formatear_numero(MODELO_MSRP_CON_IMPUESTOS)} − "
                   f"{formatear_numero(MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS)} = "
                   f"{formatear_euros(MODELO_DESCUENTO)}  {SIMBOLO_CORRECTO}", 18, COLOR_CORRECTO, negrita=True)],
        [fragmento(f"= {formatear_numero(MODELO_DESCUENTO_GUARDADO_SIN_IMPUESTOS)} × "
                   f"{1 + MODELO_TIPO_IMPOSITIVO:.2f}, the stored discount with taxes", 13, COLOR_TEXTO_SECUNDARIO)],
    ], espacio_despues_de_parrafo=4)

    escribir_parrafos(diapositiva, x_izquierda, 4.65, ancho_izquierda, 1.5, [
        [fragmento("3 · But the link depends on assumptions", 14, COLOR_TEXTO_SECUNDARIO, negrita=True)],
        [fragmento("Three of them, marked in the query. We do not know yet how many transactions the link "
                   "finds correctly.", 14)],
    ], espacio_despues_de_parrafo=4)

    # ---- Columna derecha: la consulta, con las suposiciones numeradas ----
    x_codigo = 6.05
    ancho_codigo = ANCHO_DIAPOSITIVA_PULGADAS - MARGEN_IZQUIERDO - x_codigo
    dibujar_rectangulo(diapositiva, x_codigo, 1.75, ancho_codigo, 3.45, COLOR_TARJETA, redondeado=True)
    escribir_texto(diapositiva, x_codigo + 0.25, 1.85, ancho_codigo - 0.5, 0.3,
                   "Query from the email to Dmitry (select list shortened)", tamaño=11,
                   color=COLOR_TEXTO_SECUNDARIO, cursiva=True)

    parrafos_del_codigo = []
    for linea, marca in CONSULTA_DEL_CORREO:
        fragmentos_de_la_linea = [fragmento(linea, 9.5, COLOR_TEXTO, fuente="Courier New")]
        if marca != "":
            fragmentos_de_la_linea.append(fragmento(f"   -- {marca}", 9.5, COLOR_DESCUENTO_SECUNDARIO,
                                                    negrita=True, fuente="Courier New"))
        parrafos_del_codigo.append(fragmentos_de_la_linea)
    escribir_parrafos(diapositiva, x_codigo + 0.25, 2.2, ancho_codigo - 0.4, 3.2, parrafos_del_codigo,
                      espacio_despues_de_parrafo=0)

    suposiciones = [
        ("1", "The MSRP comes from another source (WebPriceList), not from the transaction."),
        ("2", "The renewal SKU (…F) is joined as the purchase SKU (…G): same MSRP assumed. "
              "Not true for products with only renewals."),
        ("3", "The join uses the date, not the time of the transaction."),
    ]
    parrafos_de_suposiciones = []
    for marca, texto in suposiciones:
        parrafos_de_suposiciones.append([
            fragmento(f"{marca}  ", 13, COLOR_DESCUENTO_SECUNDARIO, negrita=True),
            fragmento(resaltar_entidades(texto).replace("**", ""), 13),
        ])
    escribir_parrafos(diapositiva, x_codigo, 5.3, ancho_codigo, 1.2, parrafos_de_suposiciones,
                      espacio_despues_de_parrafo=3)

    banda = dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 6.62, ANCHO_UTIL, 0.55, COLOR_TARJETA, redondeado=True,
                               color_borde=COLOR_ERROR)
    rellenar_marco_de_texto(banda.text_frame, [[
        fragmento("It works for this example. ", 15, COLOR_TEXTO, negrita=True),
        fragmento("How often does the link work? We do not know yet.", 15, COLOR_ERROR, negrita=True),
    ]], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, 0)

    poner_notas(diapositiva, "The workaround: the 59.99 comes from WebPriceList, with three assumptions: another "
                             "source, renewal SKU mapped to the purchase SKU, and a join by date. It works here; "
                             "the next slide shows where it does not.")


# Cada caso de "The gap" termina en un acuerdo con Dmitry.
# En el plan, cada acuerdo se resuelve dentro de uno de los tres retos (RETOS, "acuerdos_del_gap").
ACUERDOS = [
    {"numero": 1, "caso": "Case 1", "texto": "We need the MSRP in the transaction."},
    {"numero": 2, "caso": "Case 2",
     "texto": "We hand over how we link to WebPriceList. Your team reviews it and measures the coverage of the link."},
    {"numero": 3, "caso": "Case 3",
     "texto": "Clarify what each existing field means. New columns with a guaranteed meaning."},
    {"numero": 4, "caso": "Case 4",
     "texto": "Amounts on the same tax basis, and the VAT rate stored, also when it is 0."},
    {"numero": 5, "caso": "Case 5",
     "texto": "Review the values we cannot explain with the owners of the API."},
    {"numero": 6, "caso": "The projection",
     "texto": "A price list we can query at any date: at least a snapshot, ideally the full history."},
]


def buscar_acuerdo(numero):
    for acuerdo in ACUERDOS:
        if acuerdo["numero"] == numero:
            return acuerdo
    raise ValueError(f"No existe el acuerdo {numero}")


def poner_banda_de_acuerdo(diapositiva, numero_de_acuerdo, y=6.88, alto=0.47):
    """Banda al pie de la diapositiva que cierra un caso: el acuerdo que sale de él."""
    acuerdo = buscar_acuerdo(numero_de_acuerdo)
    banda = dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, y, ANCHO_UTIL, alto, COLOR_TARJETA, redondeado=True,
                               color_borde=COLOR_CORRECTO)
    rellenar_marco_de_texto(banda.text_frame, [[
        fragmento(f"→  Agreement {acuerdo['numero']}:  ", 15, COLOR_CORRECTO, negrita=True),
        fragmento(acuerdo["texto"], 15, COLOR_TEXTO, negrita=True),
    ]], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, 0)


def diapositiva_resumen(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "In one slide")

    bloques = [
        {"etiqueta": "THE NEED", "color": COLOR_PAGADO,
         "texto": "Retention needs the discount of every transaction, measured against the full valid MSRP."},
        {"etiqueta": "THE GAPS", "color": COLOR_ERROR,
         "texto": "The transaction does not carry the MSRP, some fields do not mean what we need, and the "
                  "projection needs the MSRP at any date."},
        {"etiqueta": "THE PLAN", "color": COLOR_DESCUENTO,
         "texto": "Three challenges. Experts fix the sources; Dmitry's team delivers the data daily, with Retention's criteria."},
    ]

    separacion_con_flecha = 0.5
    ancho_bloque = (ANCHO_UTIL - separacion_con_flecha * (len(bloques) - 1)) / len(bloques)
    for posicion, bloque in enumerate(bloques):
        x_bloque = MARGEN_IZQUIERDO + posicion * (ancho_bloque + separacion_con_flecha)
        dibujar_rectangulo(diapositiva, x_bloque, 1.9, ancho_bloque, 3.1, COLOR_TARJETA, redondeado=True,
                           color_borde=bloque["color"])
        dibujar_circulo_con_texto(diapositiva, x_bloque + 0.25, 2.15, 0.6, str(posicion + 1), bloque["color"],
                                  tamaño=18)
        escribir_texto(diapositiva, x_bloque + 0.95, 2.15, ancho_bloque - 1.05, 0.6, bloque["etiqueta"], tamaño=14,
                       color=bloque["color"], negrita=True, anclaje=MSO_ANCHOR.MIDDLE)
        escribir_texto(diapositiva, x_bloque + 0.25, 3.0, ancho_bloque - 0.45, 1.9, bloque["texto"], tamaño=15)

        es_el_ultimo = posicion == len(bloques) - 1
        if not es_el_ultimo:
            dibujar_flecha(diapositiva, x_bloque + ancho_bloque + 0.08, 3.25, 0.34, 0.4)

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 5.4, ANCHO_UTIL, 0.9, [[
        fragmento("Today: ", 18, COLOR_TEXTO, negrita=True),
        fragmento("agree that the three challenges are well defined, and the path to solve each one.", 18,
                  COLOR_TEXTO_SECUNDARIO),
    ]])

    poner_notas(diapositiva, "Three messages. What Retention needs; why the data does not give it today, also for "
                             "the projection; and the plan, with other teams.")


def diapositiva_indice(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Agenda")

    # Mismos colores que los bloques de la diapositiva resumen
    columnas = [
        {"titulo": "1 · The need", "color": COLOR_PAGADO, "temas": [
            "Theoretical transaction",
            "The same in %",
            "Definitions and checks",
            "What Retention needs",
            "Why per transaction",
            "One price, many discounts",
            "Discount of a transaction",
            "Configuration ≠ discount",
            "The rule",
            "Columns and checks",
        ]},
        {"titulo": "2 · The gaps", "color": COLOR_ERROR, "temas": [
            "Checking real data",
            "A real one: does it match?",
            "Store or infer?",
            "External MSRP",
            "'Full price' ≠ MSRP",
            "Mixed tax bases",
            "Unexplained values",
            "All 9 cases",
        ], "temas_de_la_proyeccion": [
            "Projection: MSRP at another moment",
            "Same contract, new amount",
            "Renewed vs churned",
            "Price cap",
            "MSRP at any date",
        ]},
        {"titulo": "3 · The plan", "color": COLOR_DESCUENTO, "temas": [
            "Three challenges",
            "How each is solved",
            "1 · MSRP in the transaction",
            "2 · Coverage",
            "3 · Reference price source",
            "Retention's criteria",
            "Plan and status",
            "What we ask today",
        ]},
    ]

    separacion = 0.2
    ancho_columna = (ANCHO_UTIL - separacion * (len(columnas) - 1)) / len(columnas)
    numero_de_tema = 1
    for posicion, columna in enumerate(columnas):
        x_columna = MARGEN_IZQUIERDO + posicion * (ancho_columna + separacion)
        dibujar_rectangulo(diapositiva, x_columna, 1.6, ancho_columna, 5.2, COLOR_TARJETA, redondeado=True,
                           color_borde=columna["color"])
        escribir_texto(diapositiva, x_columna + 0.22, 1.82, ancho_columna - 0.35, 0.5, columna["titulo"],
                       tamaño=18, color=columna["color"], negrita=True, fuente=FUENTE_TITULOS)

        # Los temas de la proyección van en morado dentro de su columna
        temas_con_color = [(tema, columna["color"]) for tema in columna["temas"]]
        temas_con_color += [(tema, COLOR_MORADO) for tema in columna.get("temas_de_la_proyeccion", [])]
        espacio_por_tema = min(0.42, 4.25 / len(temas_con_color))
        for posicion_tema, (tema, color_del_numero) in enumerate(temas_con_color):
            y_del_tema = 2.45 + posicion_tema * espacio_por_tema
            escribir_parrafos(diapositiva, x_columna + 0.22, y_del_tema, ancho_columna - 0.3, espacio_por_tema, [[
                fragmento(f"{numero_de_tema:02d}  ", 13, color_del_numero, negrita=True),
                fragmento(tema, 13, COLOR_TEXTO),
            ]])
            numero_de_tema += 1

    poner_notas(diapositiva, "Three parts, same colours as the summary. The gaps end in agreements; the projection, "
                             "in purple, adds one requirement on top.")


def diapositiva_lo_que_necesita_retencion(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "What Retention needs to know",
                 "Questions we must answer for each customer, with precision")

    grupos = [
        {"titulo": "Acquisition", "y": 1.7, "preguntas": [
            ("Q1", "What discount did each customer have at acquisition?"),
            ("Q2", "Across many customers and over time: which discounts brought more acquisitions?"),
        ]},
        {"titulo": "Renewal or churn", "y": 3.4, "preguntas": [
            ("Q3", "What discount did they have, and what was the renewal tactic?"),
            ("Q4", "No discount plan: the discount is lost. How big is the price delta, in % and in money?"),
            ("Q5", "With a discount plan (price ladder): what discount does the plan set?"),
            ("Q6", "Are they price capped (zero discount)? Did an MSRP increase give room to grow?"),
        ]},
    ]
    ancho_texto = 6.55
    for grupo in grupos:
        escribir_texto(diapositiva, MARGEN_IZQUIERDO, grupo["y"], ancho_texto, 0.35, grupo["titulo"], tamaño=15,
                       color=COLOR_PAGADO, negrita=True)
        for posicion, (codigo, pregunta) in enumerate(grupo["preguntas"]):
            y_pregunta = grupo["y"] + 0.42 + posicion * 0.6
            dibujar_circulo_con_texto(diapositiva, MARGEN_IZQUIERDO, y_pregunta, 0.5, codigo, COLOR_MSRP, tamaño=12)
            escribir_texto(diapositiva, MARGEN_IZQUIERDO + 0.7, y_pregunta - 0.03, ancho_texto, 0.56, pregunta,
                           tamaño=14, anclaje=MSO_ANCHOR.MIDDLE)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 6.55, 7.0, 0.4, "All amounts with and without taxes.", tamaño=14,
                   color=COLOR_TEXTO_SECUNDARIO)

    # ---- Panel derecho: los dos precios, cada uno en su parte de la presentación ----
    x_panel = 8.2
    ancho_panel = 4.53
    dibujar_rectangulo(diapositiva, x_panel, 1.7, ancho_panel, 3.6, COLOR_TARJETA, redondeado=True)
    escribir_texto(diapositiva, x_panel + 0.3, 1.9, ancho_panel - 0.6, 0.4, "We need two prices", tamaño=17,
                   negrita=True, fuente=FUENTE_TITULOS)

    caja_precio_transaccion = dibujar_rectangulo(diapositiva, x_panel + 0.3, 2.45, ancho_panel - 0.6, 0.6,
                                                 COLOR_MSRP, redondeado=True)
    escribir_dentro_de_forma(caja_precio_transaccion, "MSRP valid at the transaction", tamaño=14)
    escribir_texto(diapositiva, x_panel + 0.3, 3.08, ancho_panel - 0.6, 0.3,
                   "Q1 to Q3 · the need and the gaps", tamaño=12, color=COLOR_TEXTO_SECUNDARIO)

    caja_precio_hoy = dibujar_rectangulo(diapositiva, x_panel + 0.3, 3.6, ancho_panel - 0.6, 0.6, COLOR_MORADO,
                                         redondeado=True)
    escribir_dentro_de_forma(caja_precio_hoy, "MSRP at expiration, or today", tamaño=14)
    escribir_texto(diapositiva, x_panel + 0.3, 4.23, ancho_panel - 0.6, 0.6,
                   "Q4 to Q6 · the gaps: the projection", tamaño=12, color=COLOR_TEXTO_SECUNDARIO)

    poner_notas(diapositiva, "Acquisition: discount per customer, and which discounts bring more customers. "
                             "Renewal or churn: the discount and the renewal tactic. Without a plan, the default "
                             "tactic is to lose the discount; the jump is the price delta, in % and in money. "
                             "With a plan, the plan sets the discount.")


def diapositiva_inferir_el_msrp(presentacion):
    """
    ¿Hace falta guardar el MSRP o se puede inferir de la transacción?
    Mensaje: no inventamos nada, es aritmética básica; lo que importa es que los valores sean correctos.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Do we need to store the MSRP, or can we infer it?",
                 "In theory yes, with basic arithmetic. In practice, the fields we have today do not allow it")

    # ---- Columna izquierda: sí, con una suma ----
    x_izquierda = MARGEN_IZQUIERDO
    ancho_columna = 5.9

    escribir_texto(diapositiva, x_izquierda, 1.75, ancho_columna, 0.4, "In theory: yes, basic arithmetic", tamaño=19,
                   color=COLOR_CORRECTO, negrita=True, fuente=FUENTE_TITULOS)

    escala = 0.05
    dibujar_corchete_msrp(diapositiva, x_izquierda, 2.6, 100 * escala, "MSRP with taxes = ?   →   70 + 30 = €100")
    dibujar_barra_de_precio(diapositiva, x_izquierda, 2.8, 0.65, escala, [
        {"nombre": "Customer amount with taxes", "importe": 70, "color": COLOR_PAGADO},
        {"nombre": "Discount amount", "importe": 30, "color": COLOR_DESCUENTO},
    ], tamaño_etiqueta=12)

    escribir_texto(diapositiva, x_izquierda, 3.7, ancho_columna, 0.4, "It only works if the values are right:",
                   tamaño=15, negrita=True)

    condiciones = [
        "The discount amount is the total discount: demand level + coupon + ladder, not only one part.",
        "Both amounts are on the same tax basis.",
        "The values are populated: no 0 or empty field when there is a discount.",
    ]
    for numero_condicion, condicion in enumerate(condiciones, start=1):
        y_condicion = 4.2 + (numero_condicion - 1) * 0.62
        dibujar_circulo_con_texto(diapositiva, x_izquierda, y_condicion, 0.45, str(numero_condicion),
                                  COLOR_CORRECTO, tamaño=13)
        escribir_texto(diapositiva, x_izquierda + 0.65, y_condicion - 0.03, ancho_columna - 0.65, 0.55, condicion,
                       tamaño=14, anclaje=MSO_ANCHOR.MIDDLE)

    # ---- Columna derecha: lo que sale cuando un valor está mal (casos reales) ----
    x_derecha = 6.83
    escribir_texto(diapositiva, x_derecha, 1.75, ancho_columna, 0.4, "In practice: what we infer today",
                   tamaño=19, color=COLOR_ERROR, negrita=True, fuente=FUENTE_TITULOS)

    casos = [
        {"etiqueta": "Ex 3", "problema": "Demand level not stored",
         "cuenta": "20.99 + 0.00 = 20.99",
         "lectura": "We infer the price after demand level, not the MSRP (23.99)."},
        {"etiqueta": "Ex 2", "problema": "Discount without tax",
         "cuenta": "4.24 + 0.63 = 4.87",
         "lectura": "The MSRP is 4.99. The discount with tax (0.75) is not stored."},
        {"etiqueta": "Ex 5", "problema": "Empty value",
         "cuenta": "API grossMsrp = 0",
         "lectura": "The stored MSRP is empty: nothing to compare against."},
    ]
    alto_caso = 1.15
    for numero_caso, caso in enumerate(casos):
        y_caso = 2.3 + numero_caso * (alto_caso + 0.15)
        dibujar_rectangulo(diapositiva, x_derecha, y_caso, ancho_columna, alto_caso, COLOR_ERROR_SUAVE,
                           redondeado=True)
        chip = dibujar_rectangulo(diapositiva, x_derecha + 0.2, y_caso + 0.2, 0.8, 0.38, COLOR_ERROR,
                                  redondeado=True)
        escribir_dentro_de_forma(chip, caso["etiqueta"], tamaño=12)
        escribir_texto(diapositiva, x_derecha + 1.15, y_caso + 0.17, 2.3, 0.42, caso["problema"], tamaño=14,
                       negrita=True, anclaje=MSO_ANCHOR.MIDDLE)
        escribir_texto(diapositiva, x_derecha + 3.4, y_caso + 0.17, ancho_columna - 3.6, 0.42,
                       f"{caso['cuenta']}  {SIMBOLO_ERROR}", tamaño=14, color=COLOR_ERROR, negrita=True,
                       alineacion=PP_ALIGN.RIGHT, anclaje=MSO_ANCHOR.MIDDLE)
        escribir_texto(diapositiva, x_derecha + 0.2, y_caso + 0.68, ancho_columna - 0.4, 0.4,
                       resaltar_entidades(caso["lectura"]), tamaño=13, color=COLOR_TEXTO_SECUNDARIO)

    # ---- Banda inferior: el mensaje ----
    banda = dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 6.25, ANCHO_UTIL, 0.85, COLOR_FONDO_PORTADA,
                               redondeado=True)
    rellenar_marco_de_texto(banda.text_frame, [
        [fragmento("We are not inventing a new model. ", 17, COLOR_TEXTO_SOBRE_OSCURO, negrita=True),
         fragmento("The formula is simple: the values must mean what they say.", 17,
                   COLOR_DESCUENTO_SECUNDARIO, negrita=True)],
        [fragmento("→  Agreement 1: ", 14, COLOR_CORRECTO, negrita=True),
         fragmento(buscar_acuerdo(1)["texto"] + " Then the same equality becomes the check.", 14,
                   COLOR_TEXTO_SOBRE_OSCURO, negrita=True)],
    ], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, 2)

    poner_notas(diapositiva, "The MSRP can be inferred: customer amount with taxes plus the total discount with "
                             "taxes. Nothing new. The real requirement is that the stored values are populated "
                             "and mean what the business means. Storing the MSRP as well turns the equality "
                             "into a check.")





def diapositiva_por_que_por_transaccion(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Why the model is needed per transaction")

    # ---- Banda: cada transacción tiene su propia configuración ----
    dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 1.5, ANCHO_UTIL, 0.95, COLOR_TARJETA, redondeado=True,
                       color_borde=COLOR_PAGADO)
    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO + 0.3, 1.5, ANCHO_UTIL - 0.6, 0.95, [[
        fragmento("Each transaction has its own configuration. ", 16, negrita=True),
        fragmento("A customer acquired today can have a different discount than one acquired tomorrow, because "
                  "e-commerce changed the demand level. The MSRP stays the same until a price increase; "
                  "everything else is relative to it.", 15, COLOR_TEXTO_SECUNDARIO),
    ]], anclaje=MSO_ANCHOR.MIDDLE)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 2.6, 6.0, 0.35, "It unlocks two studies", tamaño=15,
                   color=COLOR_TEXTO_SECUNDARIO, negrita=True)

    y_tarjeta = 3.0
    alto_tarjeta = 3.95

    # ---- Tarjeta 1: sensibilidad individual al precio ----
    x_tarjeta_1 = MARGEN_IZQUIERDO
    dibujar_rectangulo(diapositiva, x_tarjeta_1, y_tarjeta, 5.9, alto_tarjeta, COLOR_TARJETA, redondeado=True)
    dibujar_circulo_con_texto(diapositiva, x_tarjeta_1 + 0.3, y_tarjeta + 0.2, 0.5, "1", COLOR_PAGADO, tamaño=16)
    escribir_texto(diapositiva, x_tarjeta_1 + 0.95, y_tarjeta + 0.2, 4.7, 0.5, "Individual price sensitivity",
                   tamaño=18, negrita=True, fuente=FUENTE_TITULOS, anclaje=MSO_ANCHOR.MIDDLE)
    escribir_parrafos(diapositiva, x_tarjeta_1 + 0.3, y_tarjeta + 0.85, 5.3, 1.5, [
        [fragmento("We need positive and negative cases: who renews and who does not, each with its discount.",
                   14)],
        [fragmento("Then we can simulate the tactic, at acquisition and at renewal.", 14)],
        [fragmento("Discount is a key driver of renewal. How much, we can only measure with all transactions.",
                   14, COLOR_DESCUENTO, negrita=True, cursiva=True)],
    ], espacio_despues_de_parrafo=5)

    # Mini diagrama: adquisición -> ¿renueva? -> sí / no
    y_diagrama = y_tarjeta + 2.55
    caja_adquisicion = dibujar_rectangulo(diapositiva, 0.9, y_diagrama, 1.55, 0.7, COLOR_MSRP, redondeado=True)
    escribir_dentro_de_forma(caja_adquisicion, "Acquisition\nwith discount", tamaño=12)
    dibujar_flecha(diapositiva, 2.55, y_diagrama + 0.18, 0.45, 0.35)
    caja_renovacion = dibujar_rectangulo(diapositiva, 3.1, y_diagrama, 1.35, 0.7, COLOR_MSRP, redondeado=True)
    escribir_dentro_de_forma(caja_renovacion, "Renews?", tamaño=13)
    dibujar_flecha(diapositiva, 4.55, y_diagrama + 0.18, 0.4, 0.35)
    caja_si = dibujar_rectangulo(diapositiva, 5.05, y_diagrama - 0.15, 1.2, 0.47, COLOR_CORRECTO_SUAVE,
                                 redondeado=True)
    escribir_dentro_de_forma(caja_si, "Renews", tamaño=12, color=COLOR_CORRECTO)
    caja_no = dibujar_rectangulo(diapositiva, 5.05, y_diagrama + 0.4, 1.2, 0.47, COLOR_ERROR_SUAVE, redondeado=True)
    escribir_dentro_de_forma(caja_no, "Does not renew", tamaño=11, color=COLOR_ERROR)

    # ---- Tarjeta 2: potencial de crecimiento de la pipeline ----
    x_tarjeta_2 = 6.83
    dibujar_rectangulo(diapositiva, x_tarjeta_2, y_tarjeta, 5.9, alto_tarjeta, COLOR_TARJETA, redondeado=True)
    dibujar_circulo_con_texto(diapositiva, x_tarjeta_2 + 0.3, y_tarjeta + 0.2, 0.5, "2", COLOR_DESCUENTO, tamaño=16)
    escribir_texto(diapositiva, x_tarjeta_2 + 0.95, y_tarjeta + 0.2, 4.7, 0.5, "Growth potential of the pipeline",
                   tamaño=18, negrita=True, fuente=FUENTE_TITULOS, anclaje=MSO_ANCHOR.MIDDLE)
    escribir_parrafos(diapositiva, x_tarjeta_2 + 0.3, y_tarjeta + 0.85, 5.3, 1.5, [
        [fragmento("Discounts decrease with each renewal until they reach zero: ", 14),
         fragmento("price cap.", 14, COLOR_ERROR, negrita=True)],
        [fragmento("Transaction by transaction we measure the growth potential of the whole pipeline, and "
                   "who leaves the price cap after an MSRP increase.", 14)],
    ], espacio_despues_de_parrafo=5)

    # Mini gráfico: descuento por transacción a lo largo de las renovaciones
    descuentos_por_transaccion = [("Acquisition", 50), ("Renewal 1", 30), ("Renewal 2", 15), ("Renewal 3", 0)]
    base_de_las_barras = y_tarjeta + 3.45
    alto_maximo = 0.8
    ancho_de_barra = 0.85
    for posicion, (nombre, porcentaje) in enumerate(descuentos_por_transaccion):
        x_barra = x_tarjeta_2 + 0.45 + posicion * 1.3
        alto_barra = max(alto_maximo * porcentaje / 50, 0.04)
        color_barra = COLOR_DESCUENTO if porcentaje > 0 else COLOR_ERROR
        dibujar_rectangulo(diapositiva, x_barra, base_de_las_barras - alto_barra, ancho_de_barra, alto_barra,
                           color_barra)
        if porcentaje > 0:
            texto_valor = f"{porcentaje}%"
            color_valor = COLOR_DESCUENTO
        else:
            texto_valor = "0%  price cap"
            color_valor = COLOR_ERROR
        escribir_texto(diapositiva, x_barra - 0.4, base_de_las_barras - alto_barra - 0.32, ancho_de_barra + 0.8,
                       0.3, texto_valor, tamaño=13, color=color_valor, negrita=True, alineacion=PP_ALIGN.CENTER)
        escribir_texto(diapositiva, x_barra - 0.3, base_de_las_barras + 0.06, ancho_de_barra + 0.6, 0.3,
                       nombre, tamaño=12, color=COLOR_TEXTO_SECUNDARIO, alineacion=PP_ALIGN.CENTER)

    poner_notas(diapositiva, "What we need per transaction is the transaction model. Each transaction has its own "
                             "configuration. It unlocks two studies: individual price sensitivity, and the growth "
                             "potential of the pipeline, transaction by transaction. Discount is important; how "
                             "important, we can only measure with all transactions.")


def diapositiva_msrp_unico(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Premise 1: one price, many discounts",
                 "At a given moment the MSRP is the same for everyone. The operation changes the discount")

    # Cabeceras de las columnas
    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 1.65, 2.6, 0.35, "What defines the MSRP", tamaño=14,
                   color=COLOR_TEXTO_SECUNDARIO, negrita=True, alineacion=PP_ALIGN.CENTER)
    escribir_parrafos(diapositiva, 9.4, 1.55, 3.3, 0.5, [
        [fragmento("Discount modifiers", 14, COLOR_DESCUENTO, negrita=True)],
        [fragmento("they change the discount, not the MSRP", 11, COLOR_TEXTO_SECUNDARIO)],
    ], alineacion=PP_ALIGN.CENTER, espacio_despues_de_parrafo=0)

    # Claves que determinan el MSRP
    claves = ["Product", "Region", "Currency", "Validity period"]
    for posicion, clave in enumerate(claves):
        caja_clave = dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 2.15 + posicion * 0.8, 2.6, 0.62,
                                        COLOR_TARJETA, redondeado=True)
        escribir_dentro_de_forma(caja_clave, clave, tamaño=15, color=COLOR_TEXTO)

    dibujar_flecha(diapositiva, 3.45, 3.3, 0.8, 0.6)

    caja_msrp = dibujar_rectangulo(diapositiva, 4.5, 2.4, 3.6, 2.4, COLOR_MSRP, redondeado=True)
    rellenar_marco_de_texto(caja_msrp.text_frame, [
        [fragmento("Valid MSRP", 28, elegir_color_de_texto_para_fondo(COLOR_MSRP), negrita=True,
                   fuente=FUENTE_TITULOS)],
        [fragmento("one price for everyone", 16, elegir_color_de_texto_para_fondo(COLOR_MSRP))],
    ], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, 4)

    dibujar_flecha(diapositiva, 8.35, 3.3, 0.8, 0.6)
    escribir_texto(diapositiva, 8.2, 3.95, 1.1, 0.3, "− discount", tamaño=12, color=COLOR_DESCUENTO, negrita=True,
                   alineacion=PP_ALIGN.CENTER)

    # Modificadores del descuento
    operaciones = ["Acquisition", "Renewal", "Stay Subscribed", "Winback"]
    for posicion, operacion in enumerate(operaciones):
        caja_operacion = dibujar_rectangulo(diapositiva, 9.4, 2.15 + posicion * 0.8, 3.3, 0.62, COLOR_TARJETA,
                                            redondeado=True, color_borde=COLOR_DESCUENTO)
        escribir_dentro_de_forma(caja_operacion, operacion, tamaño=15, color=COLOR_TEXTO)

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 5.6, ANCHO_UTIL, 1.3, [
        [fragmento("One price for everyone at a given moment. ", 16, negrita=True),
         fragmento("Acquisition, Renewal, Stay Subscribed and Winback change the discount, never the MSRP.", 16)],
        [fragmento("The MSRP changes only with a price increase (by region and product). That is why the "
                   "transaction time is part of the key.", 16, COLOR_TEXTO_SECUNDARIO)],
    ], espacio_despues_de_parrafo=8)

    poner_notas(diapositiva, "Product, region, currency and period define one MSRP. The type of operation is a "
                             "modifier of the discount, not of the price.")


def diapositiva_descuento_de_la_transaccion(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Premise 2: the discount belongs to the transaction")

    escala = 0.065
    x_barra = 3.9
    dibujar_corchete_msrp(diapositiva, x_barra, 2.05, 100 * escala, "Same valid MSRP for all: €100")

    casos = [
        {"nombre": "Acquisition", "detalle": "demand level", "porcentaje": 50, "segmentos": [
            {"nombre": "Paid", "importe": 50, "color": COLOR_PAGADO},
            {"nombre": "Demand level", "importe": 50, "color": COLOR_DESCUENTO}]},
        {"nombre": "Acquisition + coupon", "detalle": "demand level + additional coupon", "porcentaje": 55,
         "segmentos": [
            {"nombre": "Paid", "importe": 45, "color": COLOR_PAGADO},
            {"nombre": "Demand level", "importe": 50, "color": COLOR_DESCUENTO},
            {"nombre": "Coupon", "importe": 5, "color": COLOR_DESCUENTO_SECUNDARIO}]},
        {"nombre": "Renewal at 100%", "detalle": "no discount", "porcentaje": 0, "segmentos": [
            {"nombre": "Paid", "importe": 100, "color": COLOR_PAGADO}]},
        {"nombre": "Renewal with ladder", "detalle": "agreed ladder: 50% → 30%", "porcentaje": 30,
         "segmentos": [
            {"nombre": "Paid", "importe": 70, "color": COLOR_PAGADO},
            {"nombre": "Ladder", "importe": 30, "color": COLOR_DESCUENTO}]},
        {"nombre": "Stay Subscribed", "detalle": "offer to keep the subscription", "porcentaje": 40,
         "segmentos": [
            {"nombre": "Paid", "importe": 60, "color": COLOR_PAGADO},
            {"nombre": "Stay Subscribed", "importe": 40, "color": COLOR_DESCUENTO}]},
    ]

    for posicion, caso in enumerate(casos):
        y_fila = 2.3 + posicion * 0.92
        alto_barra = 0.55

        escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, y_fila - 0.02, 3.2, alto_barra + 0.1, [
            [fragmento(caso["nombre"], 15, negrita=True)],
            [fragmento(caso["detalle"], 12, COLOR_TEXTO_SECUNDARIO)],
        ], anclaje=MSO_ANCHOR.MIDDLE, espacio_despues_de_parrafo=0)

        dibujar_barra_de_precio(diapositiva, x_barra, y_fila, alto_barra, escala, caso["segmentos"],
                                tamaño_etiqueta=11)

        escribir_parrafos(diapositiva, 10.75, y_fila, 2.0, alto_barra, [[
            fragmento("Discount ", 13, COLOR_TEXTO_SECUNDARIO),
            fragmento(f"{caso['porcentaje']}%", 17, COLOR_DESCUENTO, negrita=True),
        ]], anclaje=MSO_ANCHOR.MIDDLE)

    poner_notas(diapositiva, "The MSRP is the same in all five rows. The discount changes with the operation: "
                             "demand level, coupon, renewal at 100%, agreed ladder or Stay Subscribed.")


def diapositiva_configuracion_no_es_descuento(presentacion):
    """
    Demand level, ladder step, Stay Subscribed o Winback son configuraciones, no importes.
    La misma etiqueta puede significar descuentos distintos según el momento y el MSRP vigente.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "A discount configuration is not a discount",
                 "The same label can mean different discounts. Only the amounts at transaction time are reliable")

    momentos = [
        {"momento": "January", "porcentaje": 50, "msrp": 100, "nota": "Demand level 1 means 50%"},
        {"momento": "March", "porcentaje": 60, "msrp": 100, "nota": "E-commerce moves DL 1 to 60%"},
        {"momento": "June", "porcentaje": 50, "msrp": 110, "nota": "Price increase, DL 1 back to 50%"},
    ]

    escala = 0.028
    ancho_tarjeta = 3.9
    separacion = 0.2
    y_tarjeta = 1.75
    for posicion, momento in enumerate(momentos):
        x_tarjeta = MARGEN_IZQUIERDO + posicion * (ancho_tarjeta + separacion)
        importe_descuento = momento["msrp"] * momento["porcentaje"] / 100
        importe_pagado = momento["msrp"] - importe_descuento

        dibujar_rectangulo(diapositiva, x_tarjeta, y_tarjeta, ancho_tarjeta, 3.05, COLOR_TARJETA, redondeado=True)
        escribir_texto(diapositiva, x_tarjeta + 0.3, y_tarjeta + 0.2, 1.8, 0.4, momento["momento"], tamaño=17,
                       negrita=True, fuente=FUENTE_TITULOS)
        chip = dibujar_rectangulo(diapositiva, x_tarjeta + ancho_tarjeta - 1.85, y_tarjeta + 0.22, 1.55, 0.38,
                                  COLOR_DESCUENTO, redondeado=True)
        escribir_dentro_de_forma(chip, "Demand level 1", tamaño=12)
        escribir_texto(diapositiva, x_tarjeta + 0.3, y_tarjeta + 0.7, ancho_tarjeta - 0.6, 0.35, momento["nota"],
                       tamaño=13, color=COLOR_TEXTO_SECUNDARIO)

        dibujar_corchete_msrp(diapositiva, x_tarjeta + 0.3, y_tarjeta + 1.5, momento["msrp"] * escala,
                              f"MSRP {formatear_importe(momento['msrp'])}")
        dibujar_barra_de_precio(diapositiva, x_tarjeta + 0.3, y_tarjeta + 1.7, 0.55, escala, [
            {"nombre": "Paid", "importe": importe_pagado, "color": COLOR_PAGADO},
            {"nombre": "Discount", "importe": importe_descuento, "color": COLOR_DESCUENTO},
        ], tamaño_etiqueta=11, ancho_minimo_para_etiqueta_dentro=1.0)

        escribir_parrafos(diapositiva, x_tarjeta + 0.3, y_tarjeta + 2.42, ancho_tarjeta - 0.6, 0.45, [[
            fragmento("Discount: ", 14),
            fragmento(f"{momento['porcentaje']}% = {formatear_importe(importe_descuento)}", 16, COLOR_DESCUENTO,
                      negrita=True),
        ]])

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 5.0, ANCHO_UTIL, 1.0, [
        [fragmento("The same happens with price ladder steps, Stay Subscribed and Winback offers: ", 16),
         fragmento("sometimes 40%, sometimes 30%.", 16, negrita=True)],
        [fragmento("There is no dictionary from a configuration to a discount. ", 16, COLOR_ERROR, negrita=True),
         fragmento("Keep the configuration as context; it does not replace the amount.", 16,
                   COLOR_TEXTO_SECUNDARIO)],
    ], espacio_despues_de_parrafo=6)

    banda = dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 6.1, ANCHO_UTIL, 0.9, COLOR_FONDO_PORTADA,
                               redondeado=True)
    rellenar_marco_de_texto(banda.text_frame, [
        [fragmento("At transaction time: ", 17, COLOR_TEXTO_SOBRE_OSCURO, negrita=True),
         fragmento("effective discount = valid MSRP − customer amount", 17, COLOR_DESCUENTO_SECUNDARIO,
                   negrita=True)],
        [fragmento("Later, the configuration and the MSRP may have changed, and we lose the trace to validate "
                   "the whole.", 13, COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO)],
    ], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, 2)

    poner_notas(diapositiva, "Demand levels are discount levels that e-commerce changes over time. Knowing the "
                             "demand level, the ladder step or the offer does not tell you the amount. The effective "
                             "discount must be calculated at transaction time with the valid MSRP and what the "
                             "customer paid. This leads to the rule.")


def diapositiva_la_regla(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "The non-negotiable rule")

    caja_regla = dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 1.6, ANCHO_UTIL, 1.0, COLOR_FONDO_PORTADA,
                                    redondeado=True)
    rellenar_marco_de_texto(caja_regla.text_frame, [[
        fragmento("The discount is always relative to the ", 22, COLOR_TEXTO_SOBRE_OSCURO, fuente=FUENTE_TITULOS),
        fragmento("full valid MSRP", 22, COLOR_DESCUENTO_SECUNDARIO, negrita=True, fuente=FUENTE_TITULOS),
        fragmento(" at the time of the transaction", 22, COLOR_TEXTO_SOBRE_OSCURO, fuente=FUENTE_TITULOS),
    ]], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, 0)

    # Barra de ejemplo
    escala = 0.07
    dibujar_corchete_msrp(diapositiva, MARGEN_IZQUIERDO, 3.45, 100 * escala, "Valid MSRP: €100")
    dibujar_barra_de_precio(diapositiva, MARGEN_IZQUIERDO, 3.65, 0.85, escala, [
        {"nombre": "Paid", "importe": 70, "color": COLOR_PAGADO},
        {"nombre": "Discount", "importe": 30, "color": COLOR_DESCUENTO},
    ], tamaño_etiqueta=14)

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 4.75, 7.0, 1.2, [
        [fragmento(f"30 + 70 = 100   {SIMBOLO_CORRECTO}", 22, COLOR_CORRECTO, negrita=True)],
        [fragmento("30 / 100 = 30% discount", 18, COLOR_TEXTO)],
    ], espacio_despues_de_parrafo=4)

    # Las tres condiciones
    condiciones = [
        ("Amounts reconcile", "discount + paid = valid MSRP"),
        ("Relative discount", "discount % = discount / valid MSRP"),
        ("Same tax basis", "everything tax-inclusive or everything tax-exclusive"),
    ]
    for posicion, (nombre, formula) in enumerate(condiciones):
        escribir_parrafos(diapositiva, 8.3, 3.05 + posicion * 0.95, 4.4, 0.85, [
            [fragmento(nombre, 16, COLOR_TEXTO, negrita=True)],
            [fragmento(formula, 15, COLOR_TEXTO_SECUNDARIO)],
        ], espacio_despues_de_parrafo=2)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 5.95, ANCHO_UTIL, 0.45,
                   "If the sum does not reconcile, the record is wrong.", tamaño=20, color=COLOR_ERROR, negrita=True)


    poner_notas(diapositiva, "The check uses the discount amount, not the percentage: "
                             "discount plus amount paid must equal the MSRP, on the same tax basis.")


def diapositiva_ejemplo_cupon(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Case 3: the 'full price' is not the MSRP",
                 "MSRP €100 · demand level 50% · 10% coupon applied on the already discounted price")

    escala = 0.05

    # ---- Columna izquierda: registro correcto ----
    x_izquierda = MARGEN_IZQUIERDO
    cabecera_correcta = dibujar_rectangulo(diapositiva, x_izquierda, 1.8, 5.9, 0.5, COLOR_CORRECTO_SUAVE,
                                           redondeado=True)
    escribir_dentro_de_forma(cabecera_correcta, "What we need: everything relative to the MSRP", tamaño=15,
                             color=COLOR_CORRECTO)

    dibujar_corchete_msrp(diapositiva, x_izquierda, 2.85, 100 * escala, "Valid MSRP: €100")
    dibujar_barra_de_precio(diapositiva, x_izquierda, 3.05, 0.7, escala, [
        {"nombre": "Paid", "importe": 45, "color": COLOR_PAGADO},
        {"nombre": "Demand level", "importe": 50, "color": COLOR_DESCUENTO},
        {"nombre": "Coupon", "importe": 5, "color": COLOR_DESCUENTO_SECUNDARIO},
    ])

    escribir_parrafos(diapositiva, x_izquierda, 4.35, 5.9, 2.0, [
        [fragmento("Demand level: €50 → 50% of MSRP", 14)],
        [fragmento("Coupon: 10% of €50 = €5 → 5% of MSRP", 14)],
        [fragmento("Total discount: €55 → 55%", 14, negrita=True)],
        [fragmento(f"55 + 45 = 100   {SIMBOLO_CORRECTO}", 20, COLOR_CORRECTO, negrita=True)],
    ], espacio_despues_de_parrafo=5)

    # ---- Columna derecha: lo que dan los datos ----
    x_derecha = 6.83
    cabecera_error = dibujar_rectangulo(diapositiva, x_derecha, 1.8, 5.9, 0.5, COLOR_ERROR_SUAVE, redondeado=True)
    escribir_dentro_de_forma(cabecera_error, "What the data gives: reference = price after demand level",
                             tamaño=15, color=COLOR_ERROR)

    dibujar_corchete_msrp(diapositiva, x_derecha, 2.85, 50 * escala, "'Full' price in data: €50",
                          color_linea=COLOR_ERROR, color_texto=COLOR_ERROR)
    x_fin_de_lo_registrado = dibujar_barra_de_precio(diapositiva, x_derecha, 3.05, 0.7, escala, [
        {"nombre": "Paid", "importe": 45, "color": COLOR_PAGADO},
        {"nombre": "Coupon", "importe": 5, "color": COLOR_DESCUENTO_SECUNDARIO},
    ])

    # Hueco donde debería estar el demand level: no está en los datos
    caja_no_registrada = dibujar_rectangulo(diapositiva, x_fin_de_lo_registrado, 3.05, 50 * escala, 0.7,
                                            COLOR_FONDO_CONTENIDO, color_borde=COLOR_ERROR)
    escribir_dentro_de_forma(caja_no_registrada, "Demand level €50\nnot stored", tamaño=11, color=COLOR_ERROR)

    escribir_parrafos(diapositiva, x_derecha, 4.35, 5.9, 2.0, [
        [fragmento("Coupon: €5 → 10% of €50", 14)],
        [fragmento("The demand level (€50) is not in the data", 14)],
        [fragmento("Discount shown: 10%. Real discount: 55%.", 14, negrita=True)],
        [fragmento(f"5 + 45 = 50 {SIMBOLO_CORRECTO}   but €50 is not the MSRP {SIMBOLO_ERROR}", 20, COLOR_ERROR,
                   negrita=True)],
    ], espacio_despues_de_parrafo=5)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 6.5, ANCHO_UTIL, 0.45,
                   "The sum can close against the wrong reference. Next: the same in real data (Ex 3 and Ex 4).",
                   tamaño=16, negrita=True)

    poner_notas(diapositiva, "This is what we see in real example 4: the data closes against the price after "
                             "demand level, and the demand level discount is invisible.")


def diapositiva_ejemplo_cambio_de_msrp(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Same contract, different amount",
                 "The ladder keeps the agreed 30%. The price increase changes what the customer pays")

    escala = 0.06

    # Fila 1: adquisición con el MSRP antiguo
    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 1.8, 7.0, 0.35, "Acquisition · 2025 · valid MSRP €100",
                   tamaño=15, negrita=True)
    dibujar_barra_de_precio(diapositiva, MARGEN_IZQUIERDO, 2.2, 0.7, escala, [
        {"nombre": "Paid", "importe": 50, "color": COLOR_PAGADO},
        {"nombre": "Demand level", "importe": 50, "color": COLOR_DESCUENTO},
    ])

    # Fila 2: renovación con el MSRP nuevo
    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 3.25, 7.0, 0.35,
                   "Renewal · 2026 · valid MSRP €110 (price increase)", tamaño=15, negrita=True)
    dibujar_barra_de_precio(diapositiva, MARGEN_IZQUIERDO, 3.65, 0.7, escala, [
        {"nombre": "Paid", "importe": 77, "color": COLOR_PAGADO},
        {"nombre": "Ladder", "importe": 33, "color": COLOR_DESCUENTO},
    ])

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 4.65, 7.0, 1.3, [
        [fragmento(f"33 + 77 = 110   {SIMBOLO_CORRECTO}", 20, COLOR_CORRECTO, negrita=True)],
        [fragmento("33 / 110 = 30%  →  matches the agreed ladder", 16)],
    ], espacio_despues_de_parrafo=4)

    # Panel derecho: qué pasa con el MSRP equivocado
    dibujar_rectangulo(diapositiva, 7.9, 1.8, 4.8, 3.6, COLOR_ERROR_SUAVE, redondeado=True)
    escribir_parrafos(diapositiva, 8.2, 2.05, 4.2, 3.2, [
        [fragmento("What if we use the acquisition MSRP?", 18, COLOR_ERROR, negrita=True,
                   fuente=FUENTE_TITULOS)],
        [fragmento("100 − 77 = 23  →  23% discount", 15)],
        [fragmento("23 + 77 = 100: the sum also reconciles", 15)],
        [fragmento("But the agreed ladder was 30%, not 23%.", 15, negrita=True)],
        [fragmento("The check is only valid if the MSRP is the one valid for product × region × currency "
                   "on the transaction date.", 15, COLOR_TEXTO)],
    ], espacio_despues_de_parrafo=10)

    poner_notas(diapositiva, "The sum reconciles against any MSRP we use as reference. "
                             "That is why the reference must be the one valid at the time of the transaction.")


def diapositiva_ejemplo_impuestos(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Case 4: mixed tax bases",
                 "MSRP €100 tax-exclusive · €121 tax-inclusive (21%, illustrative) · 30% ladder")

    escala = 0.027
    ancho_tarjeta = 3.9
    posiciones_x = [MARGEN_IZQUIERDO, MARGEN_IZQUIERDO + ancho_tarjeta + 0.2, MARGEN_IZQUIERDO + 2 * (ancho_tarjeta + 0.2)]

    tarjetas_correctas = [
        {"titulo": "Tax-exclusive", "msrp": 100, "pagado": 70, "descuento": 30},
        {"titulo": "Tax-inclusive", "msrp": 121, "pagado": 84.70, "descuento": 36.30},
    ]

    for posicion, datos in enumerate(tarjetas_correctas):
        x_tarjeta = posiciones_x[posicion]
        dibujar_rectangulo(diapositiva, x_tarjeta, 1.8, ancho_tarjeta, 3.4, COLOR_TARJETA, redondeado=True)
        escribir_texto(diapositiva, x_tarjeta + 0.3, 2.0, 3.3, 0.4, datos["titulo"], tamaño=18, negrita=True,
                       fuente=FUENTE_TITULOS)
        dibujar_corchete_msrp(diapositiva, x_tarjeta + 0.3, 2.9, datos["msrp"] * escala,
                              f"MSRP {formatear_importe(datos['msrp'])}")
        dibujar_barra_de_precio(diapositiva, x_tarjeta + 0.3, 3.1, 0.6, escala, [
            {"nombre": "Paid", "importe": datos["pagado"], "color": COLOR_PAGADO},
            {"nombre": "Discount", "importe": datos["descuento"], "color": COLOR_DESCUENTO},
        ], tamaño_etiqueta=12)
        texto_suma = (f"{formatear_importe(datos['descuento'])} + {formatear_importe(datos['pagado'])}"
                      f" = {formatear_importe(datos['msrp'])}")
        escribir_parrafos(diapositiva, x_tarjeta + 0.3, 4.15, 3.3, 1.0, [
            [fragmento(texto_suma, 17, COLOR_CORRECTO, negrita=True)],
            [fragmento(f"Discount 30%   {SIMBOLO_CORRECTO}", 15, COLOR_CORRECTO)],
        ], espacio_despues_de_parrafo=4)

    # Tercera tarjeta: bases mezcladas
    x_tarjeta = posiciones_x[2]
    dibujar_rectangulo(diapositiva, x_tarjeta, 1.8, ancho_tarjeta, 3.4, COLOR_ERROR_SUAVE, redondeado=True)
    escribir_texto(diapositiva, x_tarjeta + 0.3, 2.0, 3.3, 0.4, "Mixed bases", tamaño=18, negrita=True,
                   fuente=FUENTE_TITULOS, color=COLOR_ERROR)
    escribir_parrafos(diapositiva, x_tarjeta + 0.3, 2.7, 3.3, 2.9, [
        [fragmento("Tax-exclusive discount + tax-inclusive amount paid", 15)],
        [fragmento(f"€30 + €84.70 = €114.70   {SIMBOLO_ERROR}", 17, COLOR_ERROR, negrita=True)],
        [fragmento("Reconciles with neither €100 nor €121.", 15)],
    ], espacio_despues_de_parrafo=12)

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 5.55, ANCHO_UTIL, 1.0, [
        [fragmento("The discount % is the same on both bases; the amounts are not. ", 16, negrita=True),
         fragmento("Each amount is stored on both bases and each basis reconciles separately. "
                   "The tax rate depends on the country.", 16, COLOR_TEXTO_SECUNDARIO)],
    ])

    poner_notas(diapositiva, "Tax-inclusive and tax-exclusive are two independent sums. Mixing them never reconciles.")


def diapositiva_columnas(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "What we want for each transaction",
                 "New columns with a guaranteed meaning (provisional names)")

    filas = [
        ("Column", "Meaning", "Rule"),
        ("transaction_type", "Acquisition, renewal, Stay Subscribed, Winback…", "Defines which discounts are possible"),
        ("msrp_at_transaction", "Full MSRP for product × region × currency, valid at transaction time",
         "Single reference for the discount"),
        ("customer_amount", "What the customer pays in the transaction",
         "= due_to_kl + partner_fee + tax"),
        ("tax_amount", "Taxes for the State", "= customer_amount without taxes × tax rate"),
        ("partner_fee_amount", "Fee of the e-commerce provider", "% on customer_amount without taxes"),
        ("due_to_kl_amount", "What we earn net", "% on customer_amount without taxes"),
        ("discount_amount", "Total discount as an amount", "= msrp_at_transaction − customer_amount"),
        ("discount_pct", "Total discount as a ratio", "= discount_amount / msrp_at_transaction"),
        ("discount_amount_<component>", "Breakdown: demand level, coupon, ladder, offers",
         "Σ components = discount_amount"),
    ]
    anchos_de_columna = [3.3, 5.1, 3.73]
    alto_de_fila = 0.43

    forma_tabla = diapositiva.shapes.add_table(len(filas), 3, Inches(MARGEN_IZQUIERDO), Inches(1.75),
                                               Inches(sum(anchos_de_columna)), Inches(alto_de_fila * len(filas)))
    tabla = forma_tabla.table
    for numero_columna, ancho in enumerate(anchos_de_columna):
        tabla.columns[numero_columna].width = Inches(ancho)

    for numero_fila, fila in enumerate(filas):
        tabla.rows[numero_fila].height = Inches(alto_de_fila)
        es_cabecera = numero_fila == 0

        for numero_columna, texto_celda in enumerate(fila):
            celda = tabla.cell(numero_fila, numero_columna)
            celda.fill.solid()
            if es_cabecera:
                celda.fill.fore_color.rgb = COLOR_FONDO_PORTADA
                color_texto = COLOR_TEXTO_SOBRE_OSCURO
            elif numero_fila % 2 == 0:
                celda.fill.fore_color.rgb = COLOR_TARJETA
                color_texto = COLOR_TEXTO
            else:
                celda.fill.fore_color.rgb = COLOR_FONDO_CONTENIDO
                color_texto = COLOR_TEXTO

            # La columna con el nombre técnico va en negrita y en el color del MSRP
            negrita = es_cabecera or numero_columna == 0
            if numero_columna == 0 and not es_cabecera:
                color_texto = COLOR_MSRP

            rellenar_marco_de_texto(celda.text_frame, [[fragmento(texto_celda, 12, color_texto, negrita)]],
                                    PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, 0)
            celda.margin_left = Inches(0.12)
            celda.margin_right = Inches(0.08)
            celda.vertical_anchor = MSO_ANCHOR.MIDDLE

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 6.35, ANCHO_UTIL, 0.8, [
        [fragmento("Every amount in two variants: ", 14, negrita=True),
         fragmento("_tax_incl and _tax_excl.", 14, COLOR_MSRP, negrita=True)],
        [fragmento("These are new columns: existing ones have similar names but do not guarantee this meaning.",
                   14, COLOR_TEXTO_SECUNDARIO)],
    ], espacio_despues_de_parrafo=3)

    poner_notas(diapositiva, "The columns follow the transaction model: how the customer amount is split, and how "
                             "the discount relates to the MSRP. Once populated, their meaning is exactly this.")


def diapositiva_comprobaciones(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "How we validate it",
                 "Seven checks for every transaction, applied to the theoretical transaction")

    importe_sin_impuestos_mas_descuento = TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS + TEORICA_DESCUENTO_GUARDADO_SIN_IMPUESTOS
    porcentaje_descuento = 100 * TEORICA_DESCUENTO / TEORICA_MSRP_CON_IMPUESTOS

    # (código, nombre, regla, comprobación con el ejemplo, ¿se puede comprobar con el ejemplo?)
    comprobaciones = [
        ("C0", "Amount breakdown", "due_to_kl_amount + partner_fee_amount + tax_amount = customer_amount",
         f"{formatear_numero(TEORICA_DUE_TO_KL)} + {formatear_numero(TEORICA_PARTNER_FEE)} + "
         f"{formatear_numero(TEORICA_IMPUESTOS)} = {formatear_numero(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)}", True),
        ("C1", "Amounts reconcile", "discount_amount + customer_amount = msrp_at_transaction",
         f"{formatear_numero(TEORICA_DESCUENTO)} + {formatear_numero(TEORICA_IMPORTE_CLIENTE_CON_IMPUESTOS)} = "
         f"{formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)}", True),
        ("C2", "Relative discount", "discount_pct = discount_amount / msrp_at_transaction",
         f"{formatear_numero(TEORICA_DESCUENTO)} / {formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)} = "
         f"{formatear_porcentaje(porcentaje_descuento)}", True),
        ("C3", "Complete breakdown", "Σ discount components = discount_amount",
         "Components not shown in this example", False),
        ("C4", "Valid MSRP", "msrp_at_transaction = MSRP for product × region × currency at transaction time",
         f"{formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)} on {TEORICA_FECHA}", True),
        ("C5", "Same tax basis", "C0 to C3 hold on _tax_incl and on _tax_excl separately",
         f"{formatear_numero(TEORICA_IMPORTE_CLIENTE_SIN_IMPUESTOS)} + "
         f"{formatear_numero(TEORICA_DESCUENTO_GUARDADO_SIN_IMPUESTOS)} = "
         f"{formatear_numero(importe_sin_impuestos_mas_descuento)} = "
         f"{formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)} / {1 + TEORICA_TIPO_IMPOSITIVO:.2f}", True),
        ("C6", "Valid range", "0 ≤ discount_amount ≤ msrp_at_transaction",
         f"0 ≤ {formatear_numero(TEORICA_DESCUENTO)} ≤ {formatear_numero(TEORICA_MSRP_CON_IMPUESTOS)}", True),
    ]

    columnas_de_la_cuadricula = 4
    separacion = 0.2
    ancho_tarjeta = (ANCHO_UTIL - separacion * (columnas_de_la_cuadricula - 1)) / columnas_de_la_cuadricula
    alto_tarjeta = 2.3

    for posicion, (codigo, nombre, regla, comprobacion_ejemplo, se_puede_comprobar) in enumerate(comprobaciones):
        numero_columna = posicion % columnas_de_la_cuadricula
        numero_fila = posicion // columnas_de_la_cuadricula
        x_tarjeta = MARGEN_IZQUIERDO + numero_columna * (ancho_tarjeta + separacion)
        y_tarjeta = 1.75 + numero_fila * (alto_tarjeta + 0.2)

        dibujar_rectangulo(diapositiva, x_tarjeta, y_tarjeta, ancho_tarjeta, alto_tarjeta, COLOR_TARJETA,
                           redondeado=True)
        dibujar_circulo_con_texto(diapositiva, x_tarjeta + 0.22, y_tarjeta + 0.2, 0.55, codigo, COLOR_CORRECTO,
                                  tamaño=12)
        escribir_texto(diapositiva, x_tarjeta + 0.9, y_tarjeta + 0.18, ancho_tarjeta - 1.05, 0.6, nombre,
                       tamaño=15, negrita=True, fuente=FUENTE_TITULOS, anclaje=MSO_ANCHOR.MIDDLE)
        escribir_texto(diapositiva, x_tarjeta + 0.22, y_tarjeta + 0.88, ancho_tarjeta - 0.4, 0.85, regla,
                       tamaño=12, color=COLOR_TEXTO_SECUNDARIO)
        if se_puede_comprobar:
            texto_ejemplo = f"{comprobacion_ejemplo}  {SIMBOLO_CORRECTO}"
            color_ejemplo = COLOR_CORRECTO
        else:
            texto_ejemplo = comprobacion_ejemplo
            color_ejemplo = COLOR_TEXTO_SECUNDARIO
        escribir_texto(diapositiva, x_tarjeta + 0.22, y_tarjeta + 1.72, ancho_tarjeta - 0.4, 0.5, texto_ejemplo,
                       tamaño=12, color=color_ejemplo, negrita=se_puede_comprobar, cursiva=not se_puede_comprobar)

    # Octavo hueco de la cuadrícula: tolerancia y consecuencia
    x_hueco = MARGEN_IZQUIERDO + 3 * (ancho_tarjeta + separacion)
    y_hueco = 1.75 + alto_tarjeta + 0.2
    dibujar_rectangulo(diapositiva, x_hueco, y_hueco, ancho_tarjeta, alto_tarjeta, COLOR_ERROR_SUAVE, redondeado=True)
    escribir_parrafos(diapositiva, x_hueco + 0.22, y_hueco + 0.25, ancho_tarjeta - 0.4, alto_tarjeta - 0.4, [
        [fragmento("Rounding tolerance: ±0.01 in local currency.", 13, COLOR_TEXTO_SECUNDARIO)],
        [fragmento("A transaction that fails a check is recorded wrongly.", 15, COLOR_ERROR, negrita=True)],
    ], espacio_despues_de_parrafo=10)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 6.65, ANCHO_UTIL, 0.4,
                   "The theoretical transaction passes every check we can test. This is what we expect from every "
                   "transaction.",
                   tamaño=15, negrita=True)

    poner_notas(diapositiva, "The example passes the checks, including the same tax basis: 20.16 + 30.25 = 50.41, "
                             "which is 59.99 without taxes. C3 cannot be tested: the example does not show the "
                             "discount components.")


# =============================================================================
# EVIDENCIA: EJEMPLOS REALES
# Transcritos de la página de Confluence "eCommerce, MSRPs and Discounts".
# Fuentes: API de Nexway (JSON del pedido), tabla eComm y tabla WebPriceList.
# =============================================================================

EJEMPLOS_REALES = [
    {
        "etiqueta": "Ex 1", "caso": "No demand level · no coupon", "moneda": "EUR", "pedido": "4C9M4H0370D",
        "resultado": "correcto",
        "valores": [
            ("eComm · paid (with tax)", "6.99"),
            ("eComm · discount", "0.00"),
            ("WebPriceList · MSRP", "6.99"),
        ],
        "comprobacion": [
            ("6.99 + 0.00 = 6.99 = MSRP  " + SIMBOLO_CORRECTO, "correcto"),
        ],
        "lectura": "The only clean case: no demand level and no coupon. The paid amount is the MSRP.",
    },
    {
        "etiqueta": "Ex 2", "caso": "No demand level · coupon", "moneda": "USD", "pedido": "4CAHHEP0WNE",
        "resultado": "error",
        "valores": [
            ("eComm · paid (with tax)", "4.24"),
            ("eComm · discount (no tax)", "0.63"),
            ("API · grossPrice = MSRP", "4.99"),
            ("API · discount with tax", "0.75"),
        ],
        "comprobacion": [
            ("4.24 + 0.63 = 4.87 ≠ 4.99  " + SIMBOLO_ERROR, "error"),
            ("0.63 × 4.24 / 3.56 = 0.75  →  4.24 + 0.75 = 4.99", "neutral"),
        ],
        "lectura": "The discount is stored without tax and the paid amount with tax. "
                   "The discount with tax (0.75) is not in eComm: we must rebuild it.",
    },
    {
        "etiqueta": "Ex 3", "caso": "Demand level · no coupon", "moneda": "EUR", "pedido": "4AKQYNNG6R4",
        "resultado": "error",
        "valores": [
            ("eComm · paid (with tax)", "20.99"),
            ("eComm · discount", "0.00"),
            ("API · grossMsrp", "23.99"),
        ],
        "comprobacion": [
            ("20.99 + 0.00 = 20.99 ≠ 23.99  " + SIMBOLO_ERROR, "error"),
            ("Real discount: 23.99 − 20.99 = 3.00 (12.5%)", "neutral"),
        ],
        "lectura": "The demand level discount is not stored: eComm shows 0%. The MSRP is only in the API "
                   "or in WebPriceList, and renewal SKUs are not in WebPriceList.",
    },
    {
        "etiqueta": "Ex 4", "caso": "Demand level · coupon", "moneda": "INR", "pedido": "4CA0PVURQGA",
        "resultado": "error",
        "valores": [
            ("eComm · paid (with tax)", "1,250.10"),
            ("eComm · discount (no tax)", "117.71"),
            ("API · grossPrice", "1,389.00"),
            ("API · grossMsrp", "1,439.00"),
        ],
        "comprobacion": [
            ("1,250.10 + 117.71 = 1,367.81  " + SIMBOLO_ERROR, "error"),
            ("Coupon with tax 138.90  →  1,389.00 = price after DL", "neutral"),
            ("Real discount: 1,439.00 − 1,250.10 = 188.90 (13.1%)", "neutral"),
        ],
        "lectura": "Both problems: the discount has no tax, and the demand level (50.00) is not stored. "
                   "The coupon is 10% of the price after DL.",
    },
    {
        "etiqueta": "Ex 5", "caso": "International SKU · coupon", "moneda": "USD", "pedido": "4EAHPHV8P00",
        "resultado": "parcial",
        "valores": [
            ("eComm · paid with / no tax", "157.47 / 140.60"),
            ("eComm · discount (no tax)", "7.40"),
            ("WebPriceList · MSRP", "148.00"),
            ("API · grossMsrp", "0"),
        ],
        "comprobacion": [
            ("With tax: 157.47 + 7.40 = 164.87 ≠ 148.00  " + SIMBOLO_ERROR, "error"),
            ("No tax: 140.60 + 7.40 = 148.00  " + SIMBOLO_CORRECTO, "correcto"),
        ],
        "lectura": "Here the MSRP is without tax. The API sends grossMsrp = 0, so we cannot confirm "
                   "that 148.00 is the MSRP and not a price after demand level.",
    },
    {
        "etiqueta": "Ex 6", "caso": "No coupon", "moneda": "TRY", "pedido": "4EPFNR3MJNC",
        "resultado": "error",
        "valores": [
            ("eComm · paid (with tax)", "129.99"),
            ("eComm · discount", "0.00"),
            ("WebPriceList · MSRP", "99.99"),
        ],
        "comprobacion": [
            ("129.99 + 0.00 = 129.99 > 99.99  " + SIMBOLO_ERROR, "error"),
            ("Discount would be −30.00", "neutral"),
        ],
        "lectura": "The price changed on the transaction date. The join takes the old price.",
    },
    {
        "etiqueta": "Ex 7", "caso": "Coupon", "moneda": "ARS", "pedido": "4D1TNCDYH8P",
        "resultado": "error",
        "valores": [
            ("eComm · paid", "2,099.30"),
            ("eComm · discount", "899.70"),
            ("API · grossPrice", "2,999.00"),
            ("WebPriceList · MSRP", "839.00"),
        ],
        "comprobacion": [
            ("2,099.30 + 899.70 = 2,999.00", "neutral"),
            ("2,999.00 ≠ 839.00  " + SIMBOLO_ERROR, "error"),
        ],
        "lectura": "The sum closes with the API price, not with WebPriceList. "
                   "Price increase on that date: wrong row in the join.",
    },
    {
        "etiqueta": "Ex 8a", "caso": "Demand level only", "moneda": "EUR", "pedido": "4C8P9ATTPEN",
        "resultado": "error",
        "valores": [
            ("eComm · paid (with tax)", "15.99"),
            ("eComm · discount", "0.00"),
            ("API · grossMsrp", "19.99"),
            ("WebPriceList · after DL", "14.99"),
        ],
        "comprobacion": [
            ("15.99 + 0.00 = 15.99 ≠ 19.99  " + SIMBOLO_ERROR, "error"),
            ("Paid 1.00 above the price of that day", "neutral"),
        ],
        "lectura": "New price 14.99 valid from 01-10-2025 08:00, order on 02-10-2025. "
                   "Real discount 4.00 (20%) is not stored.",
    },
    {
        "etiqueta": "Ex 8b", "caso": "Demand level · coupon", "moneda": "GBP", "pedido": "4CC518GR7D3",
        "resultado": "error",
        "valores": [
            ("eComm · paid (with tax)", "22.20"),
            ("eComm · discount (no tax)", "43.16"),
            ("API · grossPrice", "73.99"),
            ("WebPriceList · MSRP", "72.99"),
        ],
        "comprobacion": [
            ("eComm: 22.20 + 43.16 = 65.36  " + SIMBOLO_ERROR, "error"),
            ("API: 22.20 + 51.79 = 73.99 ≠ 72.99  " + SIMBOLO_ERROR, "error"),
        ],
        "lectura": "73.99 does not exist in WebPriceList for this SKU and currency, neither as MSRP "
                   "nor as price after demand level. We cannot explain the source.",
    },
]


def buscar_ejemplo(etiqueta):
    for ejemplo in EJEMPLOS_REALES:
        if ejemplo["etiqueta"] == etiqueta:
            return ejemplo
    raise ValueError(f"No existe el ejemplo {etiqueta}")


def color_del_resultado(resultado):
    if resultado == "correcto":
        return COLOR_CORRECTO
    if resultado == "error":
        return COLOR_ERROR
    if resultado == "parcial":
        return COLOR_DESCUENTO_SECUNDARIO
    return COLOR_TEXTO


def dibujar_tarjeta_de_ejemplo(diapositiva, x, y, ancho, alto, ejemplo):
    """
    Tarjeta de un ejemplo real:
        cabecera (etiqueta + caso + moneda y pedido),
        valores de las fuentes,
        banda con la suma de comprobación,
        lectura en una o dos frases.
    """
    color_resultado = color_del_resultado(ejemplo["resultado"])
    es_tarjeta_estrecha = ancho < 5

    dibujar_rectangulo(diapositiva, x, y, ancho, alto, COLOR_TARJETA, redondeado=True)

    # Cabecera
    chip = dibujar_rectangulo(diapositiva, x + 0.25, y + 0.25, 0.95, 0.45, color_resultado, redondeado=True)
    escribir_dentro_de_forma(chip, ejemplo["etiqueta"], tamaño=13)
    escribir_texto(diapositiva, x + 1.35, y + 0.18, ancho - 1.55, 0.35, ejemplo["caso"], tamaño=16,
                   negrita=True, fuente=FUENTE_TITULOS)
    escribir_texto(diapositiva, x + 1.35, y + 0.52, ancho - 1.55, 0.3,
                   f"{ejemplo['moneda']} · order {ejemplo['pedido']}", tamaño=12, color=COLOR_TEXTO_SECUNDARIO)

    # Valores de las fuentes
    ancho_del_valor = 1.35 if es_tarjeta_estrecha else 1.8
    y_primer_valor = y + 1.05
    alto_de_fila = 0.36
    for numero_de_fila, (etiqueta_valor, valor) in enumerate(ejemplo["valores"]):
        y_fila = y_primer_valor + numero_de_fila * alto_de_fila
        escribir_texto(diapositiva, x + 0.25, y_fila, ancho - 0.5 - ancho_del_valor, 0.33, etiqueta_valor,
                       tamaño=13, color=COLOR_TEXTO_SECUNDARIO, anclaje=MSO_ANCHOR.MIDDLE)
        escribir_texto(diapositiva, x + ancho - 0.25 - ancho_del_valor, y_fila, ancho_del_valor, 0.33, valor,
                       tamaño=14, negrita=True, alineacion=PP_ALIGN.RIGHT, anclaje=MSO_ANCHOR.MIDDLE)

    # Banda con la suma de comprobación
    lineas = ejemplo["comprobacion"]
    tamaño_comprobacion = 13 if es_tarjeta_estrecha else 15
    alto_por_linea = 0.32 if es_tarjeta_estrecha else 0.36
    y_banda = y_primer_valor + 4 * alto_de_fila + 0.15
    alto_banda = alto_por_linea * len(lineas) + 0.3
    dibujar_rectangulo(diapositiva, x + 0.2, y_banda, ancho - 0.4, alto_banda, COLOR_FONDO_PORTADA, redondeado=True)

    parrafos_de_la_banda = []
    for texto_linea, estado_linea in lineas:
        if estado_linea == "neutral":
            parrafos_de_la_banda.append([fragmento(texto_linea, tamaño_comprobacion - 1, COLOR_TEXTO_SECUNDARIO)])
        else:
            parrafos_de_la_banda.append([fragmento(texto_linea, tamaño_comprobacion,
                                                   color_del_resultado(estado_linea), negrita=True)])
    escribir_parrafos(diapositiva, x + 0.4, y_banda + 0.15, ancho - 0.8, alto_banda - 0.3, parrafos_de_la_banda,
                      anclaje=MSO_ANCHOR.MIDDLE, espacio_despues_de_parrafo=3)

    # Lectura
    y_lectura = y_banda + alto_banda + 0.2
    escribir_texto(diapositiva, x + 0.25, y_lectura, ancho - 0.5, y + alto - y_lectura - 0.15,
                   resaltar_entidades(ejemplo["lectura"]),
                   tamaño=13 if es_tarjeta_estrecha else 14)


def diapositiva_evidencia_resumen(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "9 real transactions: the check closes only once",
                 "Check C1 (paid + discount = valid MSRP, with tax) with the fields we have today")

    etiquetas = [ejemplo["etiqueta"] for ejemplo in EJEMPLOS_REALES]
    causas = [
        {"nombre": "Discount stored without tax", "color": COLOR_DESCUENTO,
         "ejemplos": ["Ex 2", "Ex 4", "Ex 5", "Ex 8b"]},
        {"nombre": "Demand level not stored: MSRP lost", "color": COLOR_DESCUENTO_SECUNDARIO,
         "ejemplos": ["Ex 3", "Ex 4", "Ex 8a"]},
        {"nombre": "WebPriceList gives the wrong price", "color": COLOR_MSRP,
         "ejemplos": ["Ex 6", "Ex 7", "Ex 8a"]},
        {"nombre": "MSRP without tax, API grossMsrp = 0", "color": COLOR_MORADO,
         "ejemplos": ["Ex 5"]},
        {"nombre": "No source matches the amount", "color": COLOR_ERROR,
         "ejemplos": ["Ex 8b"]},
    ]

    ancho_etiqueta_fila = 3.95
    ancho_columna = (ANCHO_UTIL - ancho_etiqueta_fila) / len(etiquetas)
    y_cabecera = 1.75
    alto_fila = 0.6

    def centro_de_columna(numero_columna):
        return MARGEN_IZQUIERDO + ancho_etiqueta_fila + numero_columna * ancho_columna + ancho_columna / 2

    # Cabecera con las etiquetas y la moneda
    for numero_columna, ejemplo in enumerate(EJEMPLOS_REALES):
        escribir_parrafos(diapositiva, centro_de_columna(numero_columna) - ancho_columna / 2, y_cabecera,
                          ancho_columna, 0.55, [
                              [fragmento(ejemplo["etiqueta"], 14, COLOR_TEXTO, negrita=True)],
                              [fragmento(ejemplo["moneda"], 11, COLOR_TEXTO_SECUNDARIO)],
                          ], alineacion=PP_ALIGN.CENTER, espacio_despues_de_parrafo=0)

    # Fila de resultado
    y_resultado = y_cabecera + 0.7
    dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, y_resultado, ANCHO_UTIL, alto_fila, COLOR_TARJETA,
                       redondeado=True)
    escribir_texto(diapositiva, MARGEN_IZQUIERDO + 0.2, y_resultado, ancho_etiqueta_fila - 0.3, alto_fila,
                   "Check C1 result", tamaño=15, negrita=True, anclaje=MSO_ANCHOR.MIDDLE)
    for numero_columna, ejemplo in enumerate(EJEMPLOS_REALES):
        if ejemplo["resultado"] == "correcto":
            texto_resultado = SIMBOLO_CORRECTO
        elif ejemplo["resultado"] == "parcial":
            texto_resultado = "no tax"
        else:
            texto_resultado = SIMBOLO_ERROR
        ancho_chip = 0.72
        chip = dibujar_rectangulo(diapositiva, centro_de_columna(numero_columna) - ancho_chip / 2,
                                  y_resultado + 0.12, ancho_chip, alto_fila - 0.24,
                                  color_del_resultado(ejemplo["resultado"]), redondeado=True)
        escribir_dentro_de_forma(chip, texto_resultado, tamaño=12 if ejemplo["resultado"] == "parcial" else 16)

    # Filas de causas
    y_primera_causa = y_resultado + alto_fila + 0.25
    for numero_fila, causa in enumerate(causas):
        y_fila = y_primera_causa + numero_fila * alto_fila
        if numero_fila % 2 == 0:
            dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, y_fila, ANCHO_UTIL, alto_fila, COLOR_TARJETA)
        escribir_texto(diapositiva, MARGEN_IZQUIERDO + 0.2, y_fila, ancho_etiqueta_fila - 0.3, alto_fila,
                       causa["nombre"], tamaño=14, anclaje=MSO_ANCHOR.MIDDLE)

        diametro = 0.3
        for numero_columna, etiqueta in enumerate(etiquetas):
            if etiqueta in causa["ejemplos"]:
                punto = diapositiva.shapes.add_shape(
                    MSO_SHAPE.OVAL, Inches(centro_de_columna(numero_columna) - diametro / 2),
                    Inches(y_fila + (alto_fila - diametro) / 2), Inches(diametro), Inches(diametro))
                punto.shadow.inherit = False
                punto.fill.solid()
                punto.fill.fore_color.rgb = causa["color"]
                punto.line.fill.background()

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 6.45, ANCHO_UTIL, 0.6, [[
        fragmento("Examples chosen to cover each case, not a random sample. ", 14, COLOR_TEXTO_SECUNDARIO,
                  negrita=True),
        fragmento("How often each cause happens is still to be measured.", 14, COLOR_TEXTO_SECUNDARIO),
    ]])

    poner_notas(diapositiva, "Source: Confluence page eCommerce, MSRPs and Discounts. Nexway API order JSON, "
                             "eComm table and WebPriceList. Most examples have more than one cause.")


def diapositiva_evidencia_pareja(presentacion, titulo, subtitulo, etiquetas, notas, numero_de_acuerdo=None):
    """
    Diapositiva con dos o tres tarjetas de ejemplos reales, una al lado de otra.
    Si se indica un acuerdo, las tarjetas son algo más bajas y el acuerdo va en una banda al pie.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, titulo, subtitulo)

    if numero_de_acuerdo is None:
        alto_tarjeta = 5.25
    else:
        alto_tarjeta = 4.98

    separacion = 0.2
    ancho_tarjeta = (ANCHO_UTIL - separacion * (len(etiquetas) - 1)) / len(etiquetas)
    for posicion, etiqueta in enumerate(etiquetas):
        x_tarjeta = MARGEN_IZQUIERDO + posicion * (ancho_tarjeta + separacion)
        dibujar_tarjeta_de_ejemplo(diapositiva, x_tarjeta, 1.75, ancho_tarjeta, alto_tarjeta, buscar_ejemplo(etiqueta))

    if numero_de_acuerdo is not None:
        poner_banda_de_acuerdo(diapositiva, numero_de_acuerdo)

    poner_notas(diapositiva, notas)


def diapositiva_evidencia_sin_demand_level(presentacion):
    diapositiva_evidencia_pareja(
        presentacion,
        "Mixed tax bases in real data",
        "Case 4 · Ex 1 closes. Ex 2 fails only because the discount is stored without tax",
        ["Ex 1", "Ex 2"],
        "Same problem as our example: the discount is stored without taxes. Even without demand level, a coupon "
        "breaks the sum.",
        numero_de_acuerdo=4)


def diapositiva_evidencia_con_demand_level(presentacion):
    diapositiva_evidencia_pareja(
        presentacion,
        "With demand level, the MSRP is lost",
        "Case 3 · The data only knows the price after demand level",
        ["Ex 3", "Ex 4"],
        "Ex 4 is the toy coupon example with real numbers. The 'full price' fields do not mean the MSRP.",
        numero_de_acuerdo=3)


def diapositiva_evidencia_cambios_de_precio(presentacion):
    diapositiva_evidencia_pareja(
        presentacion,
        "Price changes: WebPriceList gives the wrong price",
        "Case 2 · Where the link fails: the MSRP is looked up after the transaction, not stored with it",
        ["Ex 6", "Ex 7", "Ex 8a"],
        "The link works for our example, but not on price change dates. We need to know how often.",
        numero_de_acuerdo=2)


def diapositiva_evidencia_sin_explicacion(presentacion):
    diapositiva_evidencia_pareja(
        presentacion,
        "Values we cannot explain",
        "Case 5 · International SKU without gross MSRP, and a price that exists in no source",
        ["Ex 5", "Ex 8b"],
        "These need the owners of the API and of WebPriceList.",
        numero_de_acuerdo=5)


# =============================================================================
# PARTE 3: LA PROYECCIÓN
# Potencial de cada cliente, price cap y proyecciones. Necesita el MSRP en otro momento:
# el de la expiración (análisis del pasado) o el de hoy (proyección).
# =============================================================================

# MSRP ilustrativo después de una subida de precio, para los ejemplos de la proyección
PROYECCION_MSRP_DESPUES_DE_SUBIDA = 64.99


def dibujar_caja_con_texto(diapositiva, x, y, ancho, alto, texto, color_relleno, tamaño=13, color_borde=None,
                           color_texto=None):
    """Caja redondeada con texto centrado. Atajo para los diagramas de la proyección."""
    caja = dibujar_rectangulo(diapositiva, x, y, ancho, alto, color_relleno, redondeado=True, color_borde=color_borde)
    escribir_dentro_de_forma(caja, texto, tamaño=tamaño, color=color_texto)
    return caja


def diapositiva_proyeccion_dos_momentos(presentacion):
    """
    La proyección necesita el MSRP en otro momento distinto de la transacción:
        - expiraciones pasadas: el MSRP vigente en la expiración (análisis)
        - expiraciones futuras: el MSRP vigente hoy (proyección)
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "The projection needs the MSRP at another moment",
                 "The model gives the price at the transaction. To project, we need the MSRP at expiration, or today")

    carriles = [
        {"titulo": "Past expirations", "subtitulo": "analysis", "momento": "Expiration",
         "precio": "MSRP valid at expiration", "y": 1.75,
         "resultados": [("Renewed: we know what they paid", COLOR_CORRECTO_SUAVE, COLOR_CORRECTO),
                        ("Churned: no transaction, no price", COLOR_ERROR_SUAVE, COLOR_ERROR)]},
        {"titulo": "Future expirations", "subtitulo": "projection", "momento": "Today",
         "precio": "MSRP valid today", "y": 4.0,
         "resultados": [("Projected amount at next renewal", COLOR_TARJETA, COLOR_MORADO),
                        ("Future price increases may not be planned yet", COLOR_TARJETA, COLOR_TEXTO_SECUNDARIO)]},
    ]

    for carril in carriles:
        y = carril["y"]
        dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, y, ANCHO_UTIL, 2.05, COLOR_TARJETA, redondeado=True,
                           color_borde=COLOR_MORADO)
        escribir_parrafos(diapositiva, MARGEN_IZQUIERDO + 0.3, y + 0.25, 2.6, 1.5, [
            [fragmento(carril["titulo"], 17, negrita=True, fuente=FUENTE_TITULOS)],
            [fragmento(carril["subtitulo"], 13, COLOR_MORADO, negrita=True)],
            [fragmento(carril["precio"], 13, COLOR_TEXTO_SECUNDARIO)],
        ], espacio_despues_de_parrafo=3)

        dibujar_caja_con_texto(diapositiva, 3.3, y + 0.6, 2.2, 0.85, "Last transaction\n(transaction model)",
                               COLOR_MSRP, tamaño=12)
        dibujar_flecha(diapositiva, 5.65, y + 0.85, 0.6, 0.38)
        dibujar_caja_con_texto(diapositiva, 6.4, y + 0.6, 2.2, 0.85, f"{carril['momento']}\n{carril['precio']}",
                               COLOR_MORADO, tamaño=12)
        dibujar_flecha(diapositiva, 8.75, y + 0.85, 0.6, 0.38)

        for posicion, (texto, relleno, color) in enumerate(carril["resultados"]):
            dibujar_caja_con_texto(diapositiva, 9.5, y + 0.3 + posicion * 0.78, 3.0, 0.62, texto, relleno, tamaño=12,
                                   color_borde=color, color_texto=color)

    banda = dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 6.3, ANCHO_UTIL, 0.8, COLOR_FONDO_PORTADA, redondeado=True)
    rellenar_marco_de_texto(banda.text_frame, [
        [fragmento("Amount at next renewal = MSRP at that moment × (1 − discount set by the plan)", 17,
                   COLOR_DESCUENTO_SECUNDARIO, negrita=True)],
        [fragmento("No discount plan: the discount is 0 and the customer pays the full MSRP.", 13,
                   COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO)],
    ], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, 2)

    poner_notas(diapositiva, "Two uses. For past expirations, the MSRP valid at expiration: it may include a price "
                             "increase. For projections, the MSRP valid today, because future increases may not be "
                             "planned yet.")


def diapositiva_renovado_frente_a_perdido(presentacion):
    """
    Si el cliente renueva, hay transacción y sabemos lo que pagó.
    Si se va, no hay transacción: hay que reconstruir el precio que le intentamos cobrar.
    Se usa el ejemplo real, suponiendo que no hay plan de descuentos ni subida de precio.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Renewed or churned: only one leaves a transaction",
                 f"Our example customer (order {MODELO_PEDIDO}) at expiration, with no discount plan")

    msrp_en_expiracion = MODELO_MSRP_CON_IMPUESTOS
    precio_intentado = msrp_en_expiracion
    price_delta = precio_intentado - MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS
    price_delta_porcentaje = 100 * price_delta / MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS

    columnas = [
        {"titulo": "Renewed", "color": COLOR_CORRECTO, "relleno": COLOR_CORRECTO_SUAVE, "x": MARGEN_IZQUIERDO,
         "parrafos": [
             [fragmento("A new transaction exists.", 15, negrita=True)],
             [fragmento("We know what the customer paid and, once the MSRP is in the transaction, the MSRP of that moment.", 14)],
             [fragmento("Price sensitivity: a positive case, with its price.", 14, COLOR_CORRECTO, negrita=True)],
         ]},
        {"titulo": "Churned", "color": COLOR_ERROR, "relleno": COLOR_ERROR_SUAVE, "x": 6.83,
         "parrafos": [
             [fragmento("No transaction.", 15, negrita=True)],
             [fragmento("What price did we try to charge? We must rebuild it:", 14)],
             [fragmento("MSRP at expiration × (1 − plan discount)", 14, COLOR_ERROR, negrita=True)],
             [fragmento(f"{formatear_numero(msrp_en_expiracion)} × (1 − 0) = {formatear_euros(precio_intentado)}",
                        14, negrita=True)],
             [fragmento(f"Last paid {formatear_euros(MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS)} → price delta "
                        f"{formatear_euros(price_delta)} (+{formatear_porcentaje(price_delta_porcentaje)})", 14,
                        COLOR_DESCUENTO, negrita=True)],
             [fragmento("Price sensitivity: a negative case, only if we can rebuild its price.", 14, COLOR_ERROR,
                        negrita=True)],
         ]},
    ]
    for columna in columnas:
        dibujar_rectangulo(diapositiva, columna["x"], 1.75, 5.9, 4.4, COLOR_TARJETA, redondeado=True,
                           color_borde=columna["color"])
        cabecera = dibujar_rectangulo(diapositiva, columna["x"] + 0.3, 1.95, 2.0, 0.5, columna["relleno"],
                                      redondeado=True, color_borde=columna["color"])
        escribir_dentro_de_forma(cabecera, columna["titulo"], tamaño=16, color=columna["color"])
        escribir_parrafos(diapositiva, columna["x"] + 0.3, 2.7, 5.3, 3.3, columna["parrafos"],
                          espacio_despues_de_parrafo=8)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 6.4, ANCHO_UTIL, 0.6,
                   "Price sensitivity needs both cases. Without the MSRP at expiration, churned customers have no price.",
                   tamaño=16, negrita=True)

    poner_notas(diapositiva, "When the customer renews, the renewal is a transaction and we know the price. When "
                             "the customer leaves, we do not know what we tried to charge. To compare renewed and "
                             "churned customers, we must rebuild that price with the MSRP at expiration.")


def diapositiva_price_cap(presentacion):
    """
    Price cap: clientes sin descuento que ya pagan el MSRP completo. Una subida de precio les devuelve
    margen de crecimiento. Tres clientes ilustrativos; el C es nuestro ejemplo real.
    """
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Price cap: who has room to grow, and who got it back",
                 f"Illustrative price increase: MSRP {formatear_euros(MODELO_MSRP_CON_IMPUESTOS)} → "
                 f"{formatear_euros(PROYECCION_MSRP_DESPUES_DE_SUBIDA)}. No discount plan at the next renewal")

    clientes = [
        {"cliente": "A", "pagado": MODELO_MSRP_CON_IMPUESTOS, "descuento": "0%",
         "msrp_antes": MODELO_MSRP_CON_IMPUESTOS, "msrp_ahora": MODELO_MSRP_CON_IMPUESTOS,
         "estado": "Price capped", "color": COLOR_ERROR},
        {"cliente": "B", "pagado": MODELO_MSRP_CON_IMPUESTOS, "descuento": "0%",
         "msrp_antes": MODELO_MSRP_CON_IMPUESTOS, "msrp_ahora": PROYECCION_MSRP_DESPUES_DE_SUBIDA,
         "estado": "Uncapped by the increase", "color": COLOR_CORRECTO},
        {"cliente": "C", "pagado": MODELO_IMPORTE_CLIENTE_CON_IMPUESTOS, "descuento": "60%",
         "msrp_antes": MODELO_MSRP_CON_IMPUESTOS, "msrp_ahora": PROYECCION_MSRP_DESPUES_DE_SUBIDA,
         "estado": "Room to grow: discount ends", "color": COLOR_DESCUENTO},
    ]

    cabeceras = ["Customer", "Last paid", "Discount", "MSRP then", "MSRP now", "Room to grow", "Status"]
    anchos = [1.9, 1.35, 1.2, 1.35, 1.35, 2.2, 2.78]

    def x_de_columna(numero_columna):
        return MARGEN_IZQUIERDO + sum(anchos[:numero_columna])

    dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 1.8, ANCHO_UTIL, 0.5, COLOR_FONDO_PORTADA, redondeado=True)
    for numero_columna, cabecera in enumerate(cabeceras):
        escribir_texto(diapositiva, x_de_columna(numero_columna) + 0.15, 1.8, anchos[numero_columna] - 0.2, 0.5,
                       cabecera, tamaño=13, color=COLOR_TEXTO_SOBRE_OSCURO, negrita=True, anclaje=MSO_ANCHOR.MIDDLE)

    for numero_fila, cliente in enumerate(clientes):
        y_fila = 2.45 + numero_fila * 0.85
        dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, y_fila, ANCHO_UTIL, 0.72, COLOR_TARJETA, redondeado=True)

        margen_de_crecimiento = cliente["msrp_ahora"] - cliente["pagado"]
        if margen_de_crecimiento > 0:
            texto_margen = (f"+{formatear_euros(margen_de_crecimiento)} "
                            f"(+{formatear_porcentaje(100 * margen_de_crecimiento / cliente['pagado'])})")
        else:
            texto_margen = formatear_euros(0)

        celdas = [
            cliente["cliente"] + (" · our example" if cliente["cliente"] == "C" else ""),
            formatear_euros(cliente["pagado"]),
            cliente["descuento"],
            formatear_euros(cliente["msrp_antes"]),
            formatear_euros(cliente["msrp_ahora"]),
            texto_margen,
        ]
        for numero_columna, texto in enumerate(celdas):
            es_el_margen = numero_columna == 5
            escribir_texto(diapositiva, x_de_columna(numero_columna) + 0.15, y_fila, anchos[numero_columna] - 0.2, 0.72,
                           texto, tamaño=14, negrita=(numero_columna == 0 or es_el_margen),
                           color=cliente["color"] if es_el_margen else COLOR_TEXTO, anclaje=MSO_ANCHOR.MIDDLE)
        dibujar_caja_con_texto(diapositiva, x_de_columna(6) + 0.1, y_fila + 0.13, anchos[6] - 0.3, 0.46,
                               cliente["estado"], COLOR_FONDO_CONTENIDO, tamaño=12, color_borde=cliente["color"],
                               color_texto=cliente["color"])

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 5.15, ANCHO_UTIL, 1.1, [
        [fragmento("Room to grow = MSRP now − last paid. ", 16, negrita=True),
         fragmento("A capped customer gets room back only with a price increase.", 16, COLOR_TEXTO_SECUNDARIO)],
        [fragmento("Questions for the whole pipeline: ", 16, negrita=True),
         fragmento("how many customers were capped and are now uncapped? How much incremental revenue?", 16,
                   COLOR_TEXTO_SECUNDARIO)],
    ], espacio_despues_de_parrafo=8)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 6.55, ANCHO_UTIL, 0.5,
                   "To answer them for every customer we need the MSRP now, and the last paid amount on the same basis.",
                   tamaño=15, color=COLOR_MORADO, negrita=True)

    poner_notas(diapositiva, "A pays the full MSRP and nothing changed: capped. B was capped but the price increase "
                             "gives room. C is our example: its 60% discount ends at renewal. Amounts with taxes; "
                             "the price increase is illustrative.")


def diapositiva_requisito_msrp_en_cualquier_fecha(presentacion):
    """El requisito de la proyección: poder consultar el MSRP en cualquier fecha."""
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Requirement: the MSRP at any date",
                 "On top of the transaction model, with correct and validated values")

    tarjetas = [
        {"titulo": "Minimum", "color": COLOR_MORADO,
         "texto": "Query the price list at a given moment and stamp the value (snapshot), by product × region "
                  "× currency. Enough to project with the MSRP of today."},
        {"titulo": "Ideal", "color": COLOR_CORRECTO,
         "texto": "A price history with validity from and to, with time. Then we rebuild the MSRP at each "
                  "expiration date, also for the past."},
    ]
    for posicion, tarjeta in enumerate(tarjetas):
        x_tarjeta = MARGEN_IZQUIERDO + posicion * 6.23
        dibujar_rectangulo(diapositiva, x_tarjeta, 1.75, 5.9, 2.2, COLOR_TARJETA, redondeado=True,
                           color_borde=tarjeta["color"])
        escribir_texto(diapositiva, x_tarjeta + 0.3, 1.95, 5.3, 0.45, tarjeta["titulo"], tamaño=18,
                       color=tarjeta["color"], negrita=True, fuente=FUENTE_TITULOS)
        escribir_texto(diapositiva, x_tarjeta + 0.3, 2.5, 5.3, 1.35, tarjeta["texto"], tamaño=15)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 4.2, ANCHO_UTIL, 0.4, "It depends on, in this order", tamaño=15,
                   color=COLOR_TEXTO_SECUNDARIO, negrita=True)
    dependencias = [
        ("1", "Transaction model clear", "The need", COLOR_PAGADO),
        ("2", "Values correct, checks C0 to C6 pass", "The gaps: cases 1 to 5", COLOR_ERROR),
        ("3", "MSRP at any date", "The gaps: projection", COLOR_MORADO),
    ]
    ancho_dependencia = 3.6
    for posicion, (numero, texto, parte, color) in enumerate(dependencias):
        x_dependencia = MARGEN_IZQUIERDO + posicion * (ancho_dependencia + 0.66)
        dibujar_rectangulo(diapositiva, x_dependencia, 4.65, ancho_dependencia, 1.0, COLOR_TARJETA, redondeado=True,
                           color_borde=color)
        dibujar_circulo_con_texto(diapositiva, x_dependencia + 0.2, 4.87, 0.55, numero, color, tamaño=15)
        escribir_parrafos(diapositiva, x_dependencia + 0.9, 4.65, ancho_dependencia - 1.0, 1.0, [
            [fragmento(texto, 14, negrita=True)],
            [fragmento(parte, 12, color)],
        ], anclaje=MSO_ANCHOR.MIDDLE, espacio_despues_de_parrafo=2)
        if posicion < len(dependencias) - 1:
            dibujar_flecha(diapositiva, x_dependencia + ancho_dependencia + 0.13, 4.95, 0.4, 0.4)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 5.95, ANCHO_UTIL, 0.5,
                   "With it: price cap, uncapped customers, incremental revenue and price sensitivity of churned "
                   "customers.", tamaño=15, color=COLOR_TEXTO_SECUNDARIO)

    poner_banda_de_acuerdo(diapositiva, 6)

    poner_notas(diapositiva, "Minimum: a snapshot of the price list at a moment, stamped for the calculation. Ideal: "
                             "the full history, so we can rebuild the MSRP at each expiration. All of this sits on "
                             "top of the transaction model and correct values.")


# =============================================================================
# PARTE 4: EL PLAN
# Tres retos. Cada uno sigue el mismo camino de cinco pasos:
#   1. explicarlo
#   2. acordar con Dmitry que se entiende y se comparte
#   3. llevarlo a los expertos dueños de la fuente
#   4. obtener de ellos un plan para corregirlo
#   5. con Dmitry, un plan de implementación
# El equipo de Dmitry no corrige las fuentes. Una vez corregidas, entrega los datos
# a Retención cada día, cumpliendo los criterios y restricciones que pone Retención.
# =============================================================================

# Los cinco pasos, iguales para los tres retos
PASOS_DE_CADA_RETO = [
    {"titulo": "Explain", "quien": "Retention",
     "texto": "What the challenge is, why it matters, and the evidence."},
    {"titulo": "Agree", "quien": "Retention + Dmitry",
     "texto": "We both understand the challenge and share it."},
    {"titulo": "Redirect", "quien": "Retention + Dmitry",
     "texto": "Take it to the experts who own the source."},
    {"titulo": "Expert plan", "quien": "Experts",
     "texto": "They give a plan and dates to fix it."},
    {"titulo": "Implement", "quien": "Dmitry's team",
     "texto": "Implementation plan. Then daily data for Retention."},
]

# Estado de cada paso hoy: el 1 está hecho (esta presentación), el 2 es lo que pedimos hoy,
# del 3 al 5 vienen después.
ESTADO_DE_LOS_PASOS = ["Done", "Today", "Next", "Next", "Next"]

# Los tres retos. "acuerdos_del_gap" dice qué acuerdos de la parte 2 se resuelven con cada reto.
# "detalles" son los límites de la técnica actual (reto 1) o lo que hace falta (retos 2 y 3).
RETOS = [
    {
        "numero": 1,
        "titulo": "The MSRP in the transaction",
        "resumen": "Every transaction carries the full MSRP valid at that moment.",
        "por_que": "The discount is always measured against the full MSRP valid at the transaction. "
                   "The transaction does not carry it, so today we rebuild it with a link to WebPriceList.",
        "titulo_de_los_detalles": "The current technique has three limits",
        "detalles": [
            ("Acquisition prices used for Retention",
             "WebPriceList is an acquisition price list. Renewal SKUs are not there: we map them to the "
             "purchase SKU and assume the same MSRP.", "Ex 3"),
            ("A link with limits",
             "We join by date, not by time. Prices change within the day, so the join can take the wrong "
             "price.", "Ex 6, 7, 8a"),
            ("Not every MSRP is available",
             "The API sends grossMsrp = 0, or without taxes. Some prices exist in no source.", "Ex 5, 8b"),
        ],
        "acuerdos_del_gap": [1, 3, 4, 5],
        "explicado_en": "The need and the gaps: the model, cases 1 to 5, the 9 examples.",
        "expertos": "Price list owners, eComm and Nexway API owners. Names to confirm.",
        "expertos_en_la_tabla": "Price lists, eComm, Nexway API",
        "plan_de_los_expertos": "How the MSRP valid at the transaction is sent or stored with every order, "
                                "renewals included.",
        "implementacion": "MSRP columns in the transaction table. Checks C0 to C6 on every row. Daily.",
    },
    {
        "numero": 2,
        "titulo": "Coverage close to 100%, at least for Nexway",
        "resumen": "We know how many transactions have a correct MSRP, and we close the gap.",
        "por_que": "A discount analysis on part of the transactions gives a biased view. "
                   "First we measure the coverage, then we close it, starting with Nexway.",
        "titulo_de_los_detalles": "What we need",
        "detalles": [
            ("Measure the coverage",
             "How many transactions have the MSRP and pass the checks? We do not have this number today.", ""),
            ("Nexway first",
             "Target: close to 100% for Nexway. Other providers after, each with its own target.", ""),
            ("Every gap has a reason",
             "Transactions without a correct MSRP are counted by case, not hidden.", "All 9 cases"),
        ],
        "acuerdos_del_gap": [2],
        "explicado_en": "Case 2 and the 9 examples. Coverage still to measure.",
        "expertos": "Dmitry's team measures, with the link we hand over. Source owners close the gaps.",
        "expertos_en_la_tabla": "Dmitry's team measures; source owners",
        "plan_de_los_expertos": "For each gap case: the fix, and when coverage reaches the target.",
        "implementacion": "Coverage by provider in the daily data, with an alert below the target.",
    },
    {
        "numero": 3,
        "titulo": "A reference price source",
        "resumen": "One agreed source for the MSRP at any date: room to grow, price cap, projections.",
        "por_que": "Room to grow, price cap and projections need the MSRP at any date. "
                   "Every analysis must take it from the same agreed source.",
        "titulo_de_los_detalles": "What the source must give",
        "detalles": [
            ("One reference",
             "One agreed source, so two analyses never use two different prices.", ""),
            ("The price at any date",
             "Minimum: a snapshot at a moment. Ideal: the full history, valid from and to, with time.", ""),
            ("Same basis as the transaction",
             "Same products, regions, currencies and tax basis as the transaction model.", ""),
        ],
        "acuerdos_del_gap": [6],
        "explicado_en": "The gaps: the projection and the price cap.",
        "expertos": "Price list owners. Names to confirm.",
        "expertos_en_la_tabla": "Price list owners",
        "plan_de_los_expertos": "Which source is the reference, and how its history is kept.",
        "implementacion": "The price list by date, refreshed daily, next to the transactions.",
    },
]


def texto_de_acuerdos(numeros_de_acuerdo):
    """[1, 3, 4, 5] -> 'agreements 1, 3, 4 and 5'; [2] -> 'agreement 2'."""
    if len(numeros_de_acuerdo) == 1:
        return f"agreement {numeros_de_acuerdo[0]}"
    todos_menos_el_ultimo = ", ".join(str(numero) for numero in numeros_de_acuerdo[:-1])
    return f"agreements {todos_menos_el_ultimo} and {numeros_de_acuerdo[-1]}"


def color_del_estado(estado):
    """Done en verde, Today en naranja, el resto en gris."""
    if estado == "Done":
        return COLOR_CORRECTO
    if estado == "Today":
        return COLOR_DESCUENTO
    return COLOR_TEXTO_SECUNDARIO


def dibujar_etiqueta_de_estado(diapositiva, x, y, ancho, estado):
    color = color_del_estado(estado)
    etiqueta = dibujar_rectangulo(diapositiva, x, y, ancho, 0.3, COLOR_FONDO_CONTENIDO, redondeado=True,
                                  color_borde=color)
    escribir_dentro_de_forma(etiqueta, estado, tamaño=11, color=color)


def diapositiva_tres_retos(presentacion):
    """Los tres retos que hay que entender, aceptar como bien definidos y planificar."""
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Three challenges to solve",
                 "We must understand them, agree that they are well defined, and plan the fix")

    separacion = 0.25
    ancho_tarjeta = (ANCHO_UTIL - separacion * (len(RETOS) - 1)) / len(RETOS)
    for posicion, reto in enumerate(RETOS):
        x_tarjeta = MARGEN_IZQUIERDO + posicion * (ancho_tarjeta + separacion)
        dibujar_rectangulo(diapositiva, x_tarjeta, 1.75, ancho_tarjeta, 4.45, COLOR_TARJETA, redondeado=True,
                           color_borde=COLOR_DESCUENTO)
        dibujar_circulo_con_texto(diapositiva, x_tarjeta + 0.25, 1.98, 0.6, str(reto["numero"]), COLOR_DESCUENTO,
                                  tamaño=18)
        escribir_texto(diapositiva, x_tarjeta + 1.0, 1.93, ancho_tarjeta - 1.2, 0.7, reto["titulo"], tamaño=17,
                       negrita=True, fuente=FUENTE_TITULOS, anclaje=MSO_ANCHOR.MIDDLE)
        escribir_texto(diapositiva, x_tarjeta + 0.25, 2.8, ancho_tarjeta - 0.5, 0.9, reto["resumen"], tamaño=14)

        escribir_texto(diapositiva, x_tarjeta + 0.25, 3.75, ancho_tarjeta - 0.5, 0.35,
                       reto["titulo_de_los_detalles"], tamaño=12, color=COLOR_TEXTO_SECUNDARIO, negrita=True)
        for numero_detalle, (titulo_detalle, _texto, _evidencia) in enumerate(reto["detalles"]):
            y_detalle = 4.15 + numero_detalle * 0.5
            escribir_parrafos(diapositiva, x_tarjeta + 0.25, y_detalle, ancho_tarjeta - 0.5, 0.45, [[
                fragmento(f"{'abc'[numero_detalle]}  ", 14, COLOR_DESCUENTO, negrita=True),
                fragmento(titulo_detalle, 14),
            ]])

        escribir_texto(diapositiva, x_tarjeta + 0.25, 5.75, ancho_tarjeta - 0.5, 0.35,
                       f"Closes {texto_de_acuerdos(reto['acuerdos_del_gap'])} of the gap", tamaño=12,
                       color=COLOR_CORRECTO, negrita=True)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 6.45, ANCHO_UTIL, 0.5,
                   "None of them can be fixed by Analytics or by Dmitry's team alone. They need the owners of "
                   "the sources.", tamaño=16, negrita=True)

    poner_notas(diapositiva, "The gap ended in six agreements. They group into three challenges. Today we want to "
                             "agree that these three are well defined. Then each one follows the same path.")


def diapositiva_camino_de_cada_reto(presentacion):
    """Los cinco pasos de cada reto y el papel de cada equipo."""
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "How each challenge is solved",
                 "The same five steps for the three challenges")

    separacion = 0.3
    ancho_paso = (ANCHO_UTIL - separacion * (len(PASOS_DE_CADA_RETO) - 1)) / len(PASOS_DE_CADA_RETO)
    for posicion, paso in enumerate(PASOS_DE_CADA_RETO):
        x_paso = MARGEN_IZQUIERDO + posicion * (ancho_paso + separacion)
        dibujar_rectangulo(diapositiva, x_paso, 1.75, ancho_paso, 2.35, COLOR_TARJETA, redondeado=True,
                           color_borde=COLOR_DESCUENTO)
        dibujar_circulo_con_texto(diapositiva, x_paso + 0.2, 1.95, 0.5, str(posicion + 1), COLOR_DESCUENTO,
                                  tamaño=15)
        escribir_texto(diapositiva, x_paso + 0.8, 1.95, ancho_paso - 0.9, 0.5, paso["titulo"], tamaño=16,
                       negrita=True, fuente=FUENTE_TITULOS, anclaje=MSO_ANCHOR.MIDDLE)
        caja_quien = dibujar_rectangulo(diapositiva, x_paso + 0.2, 2.6, ancho_paso - 0.4, 0.4, COLOR_MSRP,
                                        redondeado=True)
        escribir_dentro_de_forma(caja_quien, paso["quien"], tamaño=11)
        escribir_texto(diapositiva, x_paso + 0.2, 3.1, ancho_paso - 0.4, 0.95, paso["texto"], tamaño=13,
                       color=COLOR_TEXTO_SECUNDARIO)

        es_el_ultimo = posicion == len(PASOS_DE_CADA_RETO) - 1
        if not es_el_ultimo:
            dibujar_flecha(diapositiva, x_paso + ancho_paso + 0.04, 2.7, 0.22, 0.3)

    papeles = [
        {"quien": "Retention", "color": COLOR_PAGADO,
         "texto": "Explains the need. Sets the criteria and constraints of the data. Validates the result."},
        {"quien": "Experts", "color": COLOR_MSRP,
         "texto": "Own the sources. Fix them, with a plan and dates."},
        {"quien": "Dmitry's team", "color": COLOR_DESCUENTO,
         "texto": "Does not fix the sources. Once they are fixed, implements the solution and delivers the data "
                  "to Retention every day, meeting Retention's criteria."},
    ]
    separacion_papel = 0.25
    ancho_papel = (ANCHO_UTIL - separacion_papel * (len(papeles) - 1)) / len(papeles)
    for posicion, papel in enumerate(papeles):
        x_papel = MARGEN_IZQUIERDO + posicion * (ancho_papel + separacion_papel)
        dibujar_rectangulo(diapositiva, x_papel, 4.4, ancho_papel, 2.15, COLOR_TARJETA, redondeado=True,
                           color_borde=papel["color"])
        escribir_texto(diapositiva, x_papel + 0.25, 4.55, ancho_papel - 0.5, 0.45, papel["quien"], tamaño=17,
                       color=papel["color"], negrita=True, fuente=FUENTE_TITULOS)
        escribir_texto(diapositiva, x_papel + 0.25, 5.05, ancho_papel - 0.5, 1.45, papel["texto"], tamaño=14)

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 6.75, ANCHO_UTIL, 0.45,
                   "Each team commits only to what it owns.", tamaño=16, negrita=True)

    poner_notas(diapositiva, "Steps 1 and 2 are this meeting. Step 3 we do together. Step 4 belongs to the experts: "
                             "they own the sources and the dates. Step 5 is your team: you do not fix the source, "
                             "but once it is fixed you deliver the data every day, with our criteria.")


def diapositiva_reto(presentacion, reto):
    """Un reto: la explicación a la izquierda y los cinco pasos, con su estado, a la derecha."""
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, f"Challenge {reto['numero']}: {reto['titulo']}")

    # Izquierda: por qué, y los límites o lo que hace falta
    x_izquierda = MARGEN_IZQUIERDO
    ancho_izquierda = 6.3
    dibujar_rectangulo(diapositiva, x_izquierda, 1.45, ancho_izquierda, 1.2, COLOR_TARJETA, redondeado=True,
                       color_borde=COLOR_DESCUENTO)
    escribir_texto(diapositiva, x_izquierda + 0.25, 1.55, ancho_izquierda - 0.5, 1.0,
                   resaltar_entidades(reto["por_que"]), tamaño=15, anclaje=MSO_ANCHOR.MIDDLE)

    escribir_texto(diapositiva, x_izquierda, 2.85, ancho_izquierda, 0.35, reto["titulo_de_los_detalles"],
                   tamaño=14, color=COLOR_TEXTO_SECUNDARIO, negrita=True)

    alto_detalle = 1.05
    for numero_detalle, (titulo_detalle, texto_detalle, evidencia) in enumerate(reto["detalles"]):
        y_detalle = 3.25 + numero_detalle * (alto_detalle + 0.12)
        dibujar_rectangulo(diapositiva, x_izquierda, y_detalle, ancho_izquierda, alto_detalle, COLOR_TARJETA,
                           redondeado=True)
        dibujar_circulo_con_texto(diapositiva, x_izquierda + 0.18, y_detalle + 0.12, 0.38, "abc"[numero_detalle],
                                  COLOR_DESCUENTO, tamaño=12)
        escribir_texto(diapositiva, x_izquierda + 0.7, y_detalle + 0.1, 3.9, 0.4, titulo_detalle, tamaño=14,
                       negrita=True, anclaje=MSO_ANCHOR.MIDDLE)
        if evidencia != "":
            escribir_texto(diapositiva, x_izquierda + 4.6, y_detalle + 0.1, ancho_izquierda - 4.8, 0.4,
                           f"Seen in {evidencia}", tamaño=11, color=COLOR_TEXTO_SECUNDARIO, cursiva=True,
                           alineacion=PP_ALIGN.RIGHT, anclaje=MSO_ANCHOR.MIDDLE)
        escribir_texto(diapositiva, x_izquierda + 0.7, y_detalle + 0.5, ancho_izquierda - 0.9, 0.55,
                       resaltar_entidades(texto_detalle), tamaño=12, color=COLOR_TEXTO_SECUNDARIO)

    escribir_texto(diapositiva, x_izquierda, 6.85, ancho_izquierda, 0.35,
                   f"Closes {texto_de_acuerdos(reto['acuerdos_del_gap'])} of the gap.", tamaño=12,
                   color=COLOR_CORRECTO, negrita=True)

    # Derecha: los cinco pasos con lo concreto de este reto
    texto_de_cada_paso = [
        reto["explicado_en"],
        "We understand and share this challenge.",
        reto["expertos"],
        reto["plan_de_los_expertos"],
        reto["implementacion"],
    ]
    x_derecha = 7.15
    ancho_derecha = ANCHO_DIAPOSITIVA_PULGADAS - MARGEN_IZQUIERDO - x_derecha
    alto_paso = 0.95
    for posicion, paso in enumerate(PASOS_DE_CADA_RETO):
        y_paso = 1.45 + posicion * (alto_paso + 0.12)
        estado = ESTADO_DE_LOS_PASOS[posicion]
        es_el_acuerdo_de_hoy = estado == "Today"
        dibujar_rectangulo(diapositiva, x_derecha, y_paso, ancho_derecha, alto_paso, COLOR_TARJETA, redondeado=True,
                           color_borde=COLOR_DESCUENTO if es_el_acuerdo_de_hoy else None)
        dibujar_circulo_con_texto(diapositiva, x_derecha + 0.15, y_paso + 0.12, 0.4, str(posicion + 1),
                                  COLOR_DESCUENTO, tamaño=12)
        escribir_parrafos(diapositiva, x_derecha + 0.7, y_paso + 0.1, ancho_derecha - 1.85, 0.4, [[
            fragmento(paso["titulo"], 14, negrita=True),
            fragmento(f"  ·  {paso['quien']}", 12, COLOR_MSRP, negrita=True),
        ]], anclaje=MSO_ANCHOR.MIDDLE)
        dibujar_etiqueta_de_estado(diapositiva, x_derecha + ancho_derecha - 1.0, y_paso + 0.15, 0.85, estado)
        escribir_texto(diapositiva, x_derecha + 0.7, y_paso + 0.5, ancho_derecha - 0.85, 0.45,
                       resaltar_entidades(texto_de_cada_paso[posicion]), tamaño=12,
                       negrita=es_el_acuerdo_de_hoy, color=COLOR_TEXTO if es_el_acuerdo_de_hoy else
                       COLOR_TEXTO_SECUNDARIO)

    poner_notas(diapositiva, f"Challenge {reto['numero']}. Left: what it is and why. Right: the five steps for "
                             f"this challenge. Today we ask for step 2: that we both understand and share it.")


def diapositiva_reto_1(presentacion):
    diapositiva_reto(presentacion, RETOS[0])


def diapositiva_reto_2(presentacion):
    diapositiva_reto(presentacion, RETOS[1])


def diapositiva_reto_3(presentacion):
    diapositiva_reto(presentacion, RETOS[2])


def diapositiva_criterios_de_retencion(presentacion):
    """Los criterios y restricciones que pone Retención a los datos que entrega el equipo de Dmitry."""
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "Retention's criteria for the daily data",
                 "The data is ready for Retention only when all of them hold")

    grupos = [
        {"titulo": "Content", "color": COLOR_PAGADO, "criterios": [
            "Every transaction follows the transaction model (definitions in the need).",
            "The MSRP is the one valid at the transaction time, with taxes, on the same tax basis as the amount "
            "paid. The VAT rate is stored, also when it is 0.",
            "Checks C0 to C6 pass within ±0.01. Rows that fail are flagged, not removed.",
            "The price list can be queried at any date, for the same products, regions and currencies.",
        ]},
        {"titulo": "Delivery", "color": COLOR_DESCUENTO, "criterios": [
            "Updated every day.",
            "Coverage by provider, every day. Nexway close to 100%; a target agreed for the others.",
            "Any change in definitions or sources is agreed with Retention first.",
        ]},
    ]

    numero_de_criterio = 1
    separacion = 0.3
    ancho_grupo = (ANCHO_UTIL - separacion) / 2
    for posicion, grupo in enumerate(grupos):
        x_grupo = MARGEN_IZQUIERDO + posicion * (ancho_grupo + separacion)
        dibujar_rectangulo(diapositiva, x_grupo, 1.75, ancho_grupo, 4.75, COLOR_TARJETA, redondeado=True,
                           color_borde=grupo["color"])
        escribir_texto(diapositiva, x_grupo + 0.3, 1.92, ancho_grupo - 0.6, 0.45, grupo["titulo"], tamaño=18,
                       color=grupo["color"], negrita=True, fuente=FUENTE_TITULOS)
        for posicion_criterio, criterio in enumerate(grupo["criterios"]):
            y_criterio = 2.55 + posicion_criterio * 0.98
            dibujar_circulo_con_texto(diapositiva, x_grupo + 0.3, y_criterio + 0.03, 0.45, str(numero_de_criterio),
                                      grupo["color"], tamaño=13)
            escribir_texto(diapositiva, x_grupo + 0.95, y_criterio, ancho_grupo - 1.25, 0.9, criterio, tamaño=15)
            numero_de_criterio += 1

    escribir_texto(diapositiva, MARGEN_IZQUIERDO, 6.7, ANCHO_UTIL, 0.45,
                   "These are the acceptance criteria for the implementation plan with Dmitry's team.", tamaño=16,
                   negrita=True)

    poner_notas(diapositiva, "Your team does not fix the sources. What we ask from your team is the delivery: once "
                             "the sources are fixed, daily data that meets these criteria.")


def diapositiva_plan_y_estado(presentacion):
    """Tabla resumen: los tres retos por los cinco pasos, con lo concreto de cada celda."""
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_CONTENIDO)
    poner_titulo(diapositiva, "The plan and where we are",
                 "Three challenges, five steps. Dates for steps 4 and 5 come from the experts")

    ancho_columna_reto = 2.6
    ancho_columna_paso = (ANCHO_UTIL - ancho_columna_reto) / len(PASOS_DE_CADA_RETO)

    def x_de_columna(numero_columna):
        if numero_columna == 0:
            return MARGEN_IZQUIERDO
        return MARGEN_IZQUIERDO + ancho_columna_reto + (numero_columna - 1) * ancho_columna_paso

    dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, 1.75, ANCHO_UTIL, 0.5, COLOR_FONDO_PORTADA, redondeado=True)
    cabeceras = ["Challenge"] + [f"{posicion + 1} · {paso['titulo']}"
                                 for posicion, paso in enumerate(PASOS_DE_CADA_RETO)]
    for numero_columna, cabecera in enumerate(cabeceras):
        ancho = ancho_columna_reto if numero_columna == 0 else ancho_columna_paso
        escribir_texto(diapositiva, x_de_columna(numero_columna) + 0.15, 1.75, ancho - 0.2, 0.5, cabecera,
                       tamaño=13, color=COLOR_TEXTO_SOBRE_OSCURO, negrita=True, anclaje=MSO_ANCHOR.MIDDLE)

    alto_fila = 1.2
    for numero_fila, reto in enumerate(RETOS):
        y_fila = 2.4 + numero_fila * (alto_fila + 0.12)
        dibujar_rectangulo(diapositiva, MARGEN_IZQUIERDO, y_fila, ANCHO_UTIL, alto_fila, COLOR_TARJETA,
                           redondeado=True)
        escribir_parrafos(diapositiva, x_de_columna(0) + 0.15, y_fila, ancho_columna_reto - 0.25, alto_fila, [[
            fragmento(f"{reto['numero']}  ", 14, COLOR_DESCUENTO, negrita=True),
            fragmento(reto["titulo"], 14, negrita=True),
        ]], anclaje=MSO_ANCHOR.MIDDLE)

        # Lo concreto de cada paso en la tabla: corto, el detalle está en la diapositiva del reto
        textos_de_la_fila = [
            "This deck",
            "Today",
            reto["expertos_en_la_tabla"],
            "Pending",
            "Pending",
        ]
        for posicion, texto in enumerate(textos_de_la_fila):
            numero_columna = posicion + 1
            estado = ESTADO_DE_LOS_PASOS[posicion]
            dibujar_etiqueta_de_estado(diapositiva, x_de_columna(numero_columna) + 0.12, y_fila + 0.12, 0.85, estado)
            escribir_texto(diapositiva, x_de_columna(numero_columna) + 0.12, y_fila + 0.5,
                           ancho_columna_paso - 0.2, alto_fila - 0.55, texto, tamaño=11,
                           color=COLOR_TEXTO_SECUNDARIO)

    escribir_parrafos(diapositiva, MARGEN_IZQUIERDO, 6.4, ANCHO_UTIL, 0.9, [
        [fragmento("First step after today: ", 15, negrita=True),
         fragmento("we hand over how we link to WebPriceList, and your team measures its coverage (challenge 2).",
                   15, COLOR_TEXTO_SECUNDARIO)],
        [fragmento("We do not commit dates for work we do not own.", 15, COLOR_DESCUENTO, negrita=True)],
    ], espacio_despues_de_parrafo=4)

    poner_notas(diapositiva, "Steps 1 and 2 are today. The expert names are still to confirm. Steps 4 and 5 get "
                             "dates only when the experts give their plan.")


def diapositiva_lo_que_pedimos(presentacion):
    diapositiva = crear_diapositiva(presentacion, COLOR_FONDO_PORTADA)

    escribir_texto(diapositiva, 0.8, 0.9, 11.5, 0.9, "What we ask today", tamaño=40,
                   color=COLOR_TEXTO_SOBRE_OSCURO, negrita=True, fuente=FUENTE_TITULOS)

    peticiones = [
        "Agree that the three challenges are well defined, and that we share them.",
        "Agree the path: the experts fix the sources, your team implements and delivers the data daily.",
        "Help us name the experts for each challenge.",
        "Your team measures the coverage of our link to WebPriceList. We hand it over.",
    ]
    for posicion, peticion in enumerate(peticiones):
        y_peticion = 2.2 + posicion * 0.9
        dibujar_circulo_con_texto(diapositiva, 0.8, y_peticion, 0.6, str(posicion + 1), COLOR_DESCUENTO, tamaño=16)
        escribir_texto(diapositiva, 1.65, y_peticion - 0.05, 10.9, 0.7, peticion, tamaño=21,
                       color=COLOR_TEXTO_SOBRE_OSCURO, anclaje=MSO_ANCHOR.MIDDLE)

    escribir_parrafos(diapositiva, 0.8, 6.0, 11.5, 1.1, [
        [fragmento("Our part: ", 18, COLOR_DESCUENTO_SECUNDARIO, negrita=True),
         fragmento("the need, the transaction model, the checks, the criteria and the hand-over of the link.", 18,
                   COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO)],
        [fragmento("Then: ", 18, COLOR_DESCUENTO_SECUNDARIO, negrita=True),
         fragmento("with daily data that meets the criteria: discounts, price cap and projection.", 18,
                   COLOR_TEXTO_SECUNDARIO_SOBRE_OSCURO)],
    ], espacio_despues_de_parrafo=8)

    poner_notas(diapositiva, "We do not ask your team to fix the sources. We ask for agreement on the three "
                             "challenges, help to find the experts, the coverage measurement, and the daily "
                             "delivery once the sources are fixed.")


# =============================================================================
# 4. GENERACIÓN
# =============================================================================

def generar_presentacion():
    presentacion = Presentation()
    presentacion.slide_width = Inches(ANCHO_DIAPOSITIVA_PULGADAS)
    presentacion.slide_height = Inches(ALTO_DIAPOSITIVA_PULGADAS)

    # Cada diapositiva con la parte a la que pertenece ("need", "gap", "plan" o None) y su tema.
    # Criterio de tema:
    #   - claro:  diapositivas que explican conceptos (como tus diapositivas de contenido).
    #   - oscuro: diapositivas de orientación (resumen, agenda, modelo, cierre), como tu Outline,
    #             y diapositivas con datos reales, como tus tablas de Power BI.
    diapositivas_en_orden = [
        # Apertura
        (diapositiva_resumen, None, TEMA_OSCURO),
        (diapositiva_indice, None, TEMA_OSCURO),

        # 1 · La necesidad: lo que esperamos, con el ejemplo real y sin hablar del problema
        (diapositiva_portada, "need", TEMA_OSCURO),
        (diapositiva_modelo_en_porcentajes, "need", TEMA_OSCURO),
        (diapositiva_definiciones_del_modelo, "need", TEMA_OSCURO),
        (diapositiva_lo_que_necesita_retencion, "need", TEMA_CLARO),
        (diapositiva_por_que_por_transaccion, "need", TEMA_CLARO),
        (diapositiva_msrp_unico, "need", TEMA_CLARO),
        (diapositiva_descuento_de_la_transaccion, "need", TEMA_CLARO),
        (diapositiva_configuracion_no_es_descuento, "need", TEMA_CLARO),
        (diapositiva_la_regla, "need", TEMA_CLARO),
        (diapositiva_columnas, "need", TEMA_CLARO),
        (diapositiva_comprobaciones, "need", TEMA_CLARO),

        # 2 · Los gaps: cada caso termina en un acuerdo
        (diapositiva_mapping_de_conceptos, "gap", TEMA_OSCURO),
        (diapositiva_caso_1_el_ejemplo_en_los_datos, "gap", TEMA_OSCURO),
        (diapositiva_inferir_el_msrp, "gap", TEMA_CLARO),                  # acuerdo 1
        (diapositiva_realidad_hoy, "gap", TEMA_OSCURO),                     # caso 2
        (diapositiva_evidencia_cambios_de_precio, "gap", TEMA_OSCURO),      # acuerdo 2
        (diapositiva_ejemplo_cupon, "gap", TEMA_CLARO),                     # caso 3
        (diapositiva_evidencia_con_demand_level, "gap", TEMA_OSCURO),       # acuerdo 3
        (diapositiva_ejemplo_impuestos, "gap", TEMA_CLARO),                 # caso 4
        (diapositiva_evidencia_sin_demand_level, "gap", TEMA_OSCURO),       # acuerdo 4
        (diapositiva_evidencia_sin_explicacion, "gap", TEMA_OSCURO),        # caso 5 y acuerdo 5
        (diapositiva_evidencia_resumen, "gap", TEMA_OSCURO),

        # 2 · Los gaps (proyección): price cap y potencial, necesita el MSRP en cualquier fecha
        (diapositiva_proyeccion_dos_momentos, "gap", TEMA_CLARO),
        (diapositiva_ejemplo_cambio_de_msrp, "gap", TEMA_CLARO),
        (diapositiva_renovado_frente_a_perdido, "gap", TEMA_CLARO),
        (diapositiva_price_cap, "gap", TEMA_CLARO),
        (diapositiva_requisito_msrp_en_cualquier_fecha, "gap", TEMA_CLARO),   # acuerdo 6

        # 3 · El plan: tres retos, cada uno con el mismo camino de cinco pasos
        (diapositiva_tres_retos, "plan", TEMA_CLARO),
        (diapositiva_camino_de_cada_reto, "plan", TEMA_CLARO),
        (diapositiva_reto_1, "plan", TEMA_CLARO),
        (diapositiva_reto_2, "plan", TEMA_CLARO),
        (diapositiva_reto_3, "plan", TEMA_CLARO),
        (diapositiva_criterios_de_retencion, "plan", TEMA_CLARO),
        (diapositiva_plan_y_estado, "plan", TEMA_CLARO),
        (diapositiva_lo_que_pedimos, "plan", TEMA_OSCURO),
    ]

    for numero_de_diapositiva, (funcion_de_la_diapositiva, parte, tema) in enumerate(diapositivas_en_orden,
                                                                                       start=1):
        aplicar_tema(tema)
        funcion_de_la_diapositiva(presentacion)
        diapositiva_recien_creada = presentacion.slides[-1]
        poner_parte_y_numero(diapositiva_recien_creada, parte, numero_de_diapositiva)

    presentacion.save(NOMBRE_FICHERO_SALIDA)
    print(f"Presentación guardada en {NOMBRE_FICHERO_SALIDA} ({len(diapositivas_en_orden)} diapositivas)")


if __name__ == "__main__":
    generar_presentacion()
