# -*- coding: utf-8 -*-
"""
================================================================================
 GENERADOR DEL ARCHIVO PLANO CAC (Cuenta de Alto Costo) - PLATAFORMA SISCAC 4.0
 Entidad: ESS062
 Periodo de medicion: 2 de enero de 2025 al 1 de enero de 2026
 Fecha de corte del periodo: 2026-01-01
 Fecha de reporte (Resolucion 0247/2014, articulo 7): 2026-05-05
================================================================================

 FUENTES UTILIZADAS
   1. CARGE 1.xlsx ............................. Base de datos de la cohorte (622 pacientes)
   2. CANCER-20260101-Instructivo_Reporte.pdf . Instructivo de reporte (61 paginas)
   3. Reglas Validacion Cancer 2024 V1 1 (2).xlsx  Motor de validacion (771 reglas B)
   4. CANCER_20260101-ATC_2026.xlsx ........... Archivo operativo de codigos ATC (219 codigos)

 REGLAS DEL INSTRUCTIVO APLICADAS
   * Variables 1-16  -> identificacion y demografia
   * Variables 17-44 -> diagnostico, estadificacion y objetivo del tratamiento
   * Variables 45-73 -> terapia sistemica (primer y ultimo esquema)
   * Variables 74-85 -> cirugia
   * Variables 86-105 -> radioterapia
   * Variables 106-113 -> trasplante y cirugia reconstructiva
   * Variables 114-124 -> paliativos, psiquiatria, nutricion y rehabilitation
   * Variables 125-134 -> situacion actual, novedad administrativa y fecha de corte

 SALIDAS
   20260505_ESS062_CANCER.txt  -> archivo plano ANSI separado por tabulaciones
   REPORTE_VALIDACION_CAC.txt   -> diagnostico de hallazgos
   PENDIENTES_CAC.csv          -> registros que requieren correccion manual
================================================================================
"""

import csv
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime

import openpyxl

# ==============================================================================
# 1. CONFIGURACION
# ==============================================================================

RUTA_BASE = os.path.dirname(os.path.abspath(__file__))

ARCHIVO_FUENTE = "CARGE 1.xlsx"
ARCHIVO_REGLAS = "Reglas Validacion Cancer 2024 V1 1 (2).xlsx"
ARCHIVO_ATC = "CANCER_20260101-ATC_2026.xlsx"

NOMBRE_TXT = "20260505_ESS062_CANCER.txt"
NOMBRE_REPORTE = "REPORTE_VALIDACION_CAC.txt"
NOMBRE_PENDIENTES = "PENDIENTES_CAC.csv"

# Codigo EPS (6 caracteres, Resolucion 0247/2014) y codigo de habilitacion REPS
CODIGO_EAPB = "ESS062"
CODIGO_REPS_IPS = "190010942901"

# Fechas del periodo de medicion
FECHA_CORTE = "2027-01-01"
FECHA_INICIO_PERIODO = "2025-01-02"
FECHA_REPORTE = "20260505"

# Comodines de fecha definidos en el instructivo
FECHA_DESCONOCIDO = "1800-01-01"
FECHA_NO_APLICA = "1845-01-01"
FECHA_NO_APLICA_IN_SITU = "1840-01-01"
COMODINES_FECHA = {
    FECHA_DESCONOCIDO,
    FECHA_NO_APLICA,
    FECHA_NO_APLICA_IN_SITU,
}

# Valor generico de "no aplica / no informado" para variables categoricas
NO_APLICA = "98"
DESCONOCIDO = "99"

# Caracteres prohibidos por el instructivo (pagina 5): &, ñ, ú, ü, #, °, ´, etc.
# El archivo plano se graba en formato ANSI, por lo que se transliteran.
MAPA_TRANSLITERACION = {
    "Ñ": "N",
    "ñ": "N",
    "Ú": "U",
    "ú": "U",
    "Ü": "U",
    "ü": "U",
    "Í": "I",
    "í": "I",
    "Ó": "O",
    "ó": "O",
    "É": "E",
    "é": "E",
    "Á": "A",
    "á": "A",
    "È": "E",
    "è": "E",
    "Ê": "E",
    "ê": "E",
    "Ë": "E",
    "ë": "E",
    "Ç": "C",
    "ç": "C",
    "Ò": "O",
    "ò": "O",
    "Ô": "O",
    "ô": "O",
    "Û": "U",
    "û": "U",
    "°": "",
    "º": "",
    "µ": "U",
}

# Traslado de fechas del periodo 2024 (referencia del archivo de reglas) al
# periodo 2026. Las fechas historicas (2004, 2006, 2017) NO se trasladan.
TRASLADO_FECHAS_PERIODO = {
    "2022-11-01": "2025-11-01",  # hito de cirugia reportable (instructivo pag. 53)
    "2023-01-01": "2025-01-01",  # inicio del periodo
    "2023-01-02": "2025-01-02",  # inicio del periodo
    "2024-01-01": "2026-01-01",  # fecha de corte
    "2021-01-01": "2023-01-01",  # fecha minima de muerte
}

# Dominios de respuesta tomados del instructivo de reporte 2026 (A-nnnn).
# Las variables que no aparecen aqui se validan con las reglas condicionales.
DOMINIOS = {
    "5": {"CC", "CE", "CD", "PA", "SC", "PT", "PE", "RC", "TI", "CN", "AS", "MS", "DE", "SI"},
    "8": {"M", "F"},
    "10": {"C", "S", "P", "E", "N", "I"},
    "12": {"1", "2", "3", "4", "5", "6"},
    "13": {str(i) for i in list(range(1, 17)) + list(range(31, 40)) + list(range(50, 64))},
    "21": {"5", "6", "7", "8", "9", "10", "99", "1", "2", "3", "4"},
    "22": {"1", "2", "3", "4", "5", "6", "98", "99"},
    "27": {str(i) for i in range(1, 22)} | {"23", "24", "98", "99"},
    "28": {"1", "2", "3", "4", "94", "95", "98", "99"},
    "29": {
        "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
        "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "27", "28",
        "29", "30", "31", "32", "33", "34", "35", "36", "98", "99",
    },
    "31": {"1", "2", "97", "98", "99"},
    "33": {"1", "2", "3", "4", "97", "98", "99"},
    "34": {"1", "2", "3", "4", "98", "99"},
    "36": {str(i) for i in range(1, 17)} | {"98", "99"},
    "37": {"11", "12", "13", "14", "15", "97", "98", "99"},
    "38": {"1", "2", "3", "4", "5", "98", "99"},
    "40": {"1", "2", "3", "99"},
    "41": {"1", "2", "3", "4", "5", "6", "99"},
    "42": {"1", "2", "99"},
    "45": {"1", "98"},
    "46.1": {"1", "2", "97"},
    "46.2": {"1", "2", "97"},
    "46.3": {"1", "2", "97"},
    "46.4": {"1", "2", "97"},
    "46.5": {"1", "2", "97"},
    "46.6": {"1", "2", "97"},
    "46.7": {"1", "2", "97"},
    "46.8": {"1", "2", "97"},
    "48": {"1", "2", "3", "11", "12", "13", "98"},
    "57": {"1", "2", "98"},
    "59": {"1", "2", "3", "98"},
    "60": {"1", "2", "3", "4", "5", "6", "7", "8", "98"},
    "61": {"1", "2", "3", "11", "12", "13", "14", "97", "98"},
    "70": {"1", "2", "98"},
    "72": {"1", "2", "3", "98"},
    "73": {"1", "2", "3", "4", "5", "6", "7", "8", "98"},
    "74": {"1", "2"},
    "79": {"1", "5", "6", "98"},
    "81": {"1", "2", "3", "5", "6", "98"},
    "84": {"1", "5", "6", "98"},
    "85": {"1", "2", "98"},
    "86": {"1", "98"},
    "89": {"1", "2", "3", "11", "12", "13", "98"},
    "95": {"1", "2", "3", "98"},
    "96": {"1", "2", "3", "4", "5", "6", "7", "98"},
    "98": {"1", "2", "3", "11", "12", "13", "98"},
    "104": {"1", "2", "3", "98"},
    "105": {"1", "2", "3", "4", "5", "6", "7", "98"},
    "106": {"1", "98"},
    "107": {"1", "2", "3", "4", "5", "6", "7", "8", "9", "98"},
    "108": {"95", "96", "97", "98"},
    "111": {"1", "98"},
    "114": {"1", "2"},
    "114.1": {"1", "2"},
    "114.2": {"1", "2"},
    "114.3": {"1", "2"},
    "114.4": {"1", "2"},
    "114.5": {"1", "2"},
    "114.6": {"1", "2"},
    "117": {"1", "2", "98"},
    "120": {"1", "2", "98"},
    "123": {"1", "2", "3", "4"},
    "124": {"1", "2", "3", "5", "6", "7", "8", "98"},
    "125": {str(i) for i in range(1, 12)} | {"98"},
    "126": {str(i) for i in range(1, 9)} | {"97", "98", "99"},
    "127": {"1", "2", "99"},
    "128": {
        "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13",
        "15", "16", "17", "18", "19",
    },
    "129": {"1", "3", "8", "9", "10", "11", "12"},
    "132": {"1", "2", "3", "4", "98"},
}

# ==============================================================================
# 1.1 SUPUESTOS DECLARADOS
# ------------------------------------------------------------------------------
# Valores que NO se pueden derivar de la fuente ni del instructivo y que por
# tanto se declaran de forma explicita. Cada supuesto queda registrado en el
# reporte de validacion y en PENDIENTES_CAC.csv para que el equipo clinico
# los confirme o los ajuste antes de radicar el archivo.
# ==============================================================================

# V48 (ubicacion temporal del primer ciclo). La fuente trae 0, que no pertenece
# al dominio del instructivo. B4913 exige V48 en {98, 11, 12, 13} cuando el
# objetivo fue paliacion (V40 = 2) y B4932 exige V48 <= 13 cuando hubo
# terapia sistemica.
SUPUESTO_V48_PALIATIVO = "13"   # manejo paliativo
SUPUESTO_V48_CURATIVO = "2"     # tratamiento inicial curativo

# V49 (fecha de inicio del primer ciclo). La fuente trae 0, que no es una fecha.
# Se toma la fecha de diagnostico (V18), que es el unico dato disponible y
# cumple el rango exigido por la definicion de la variable (V18 <= V49 <= V134).
SUPUESTO_V49_DESDE_V18 = True

# V53 (cantidad de medicamentos antineoplasicos propuestos) y V53.1 a V56
# (codigos ATC). La fuente no trae el detalle de los medicamentos: V53 = 0 y
# V53.x = 97. El valor 99 ("no informado") esta en el dominio pero la regla
# B4988 obliga entonces a V128 = 14, valor que el instructivo de 2026 no
# admite. Por eso se declara un esquema de un medicamento (V53 = 1) con el
# ATC desconocido (V53.1 = 99).
SUPUESTO_V53_CANTIDAD = "1"
SUPUESTO_ATC_DESCONOCIDO = "99"

# V133 (serial unico BDUA/BDEX/PVS). La fuente viene vacia en toda la cohorte.
SUPUESTO_V133 = "0"

# V128 (novedad administrativa respecto al reporte anterior).
SUPUESTO_V128_FALLECIMIENTO = "4"       # usuario que falleció
SUPUESTO_V128_ABANDONO = "9"            # abandono de tratamiento, imposible de ubicar
SUPUESTO_V128_NUEVO_DIAGNOSTICO = "2"   # nuevo diagnostico de cancer en el periodo

# V41 (objetivo de la intervencion medica durante el periodo). La fuente trae
# valores que contradicen el tratamiento registrado: por ejemplo V41 = 1
# (observacion previa a tratamiento) en pacientes que recibieron quimio.
SUPUESTO_V41_TRATAMIENTO = "2"    # ofrecio tratamiento curativo o paliativo
SUPUESTO_V41_SEGUIMIENTO = "3"    # observacion o seguimiento tras el tratamiento

# V125 (tipo de tratamiento que recibe el usuario a la fecha de corte) y
# V126 (resultado final del manejo oncologico). La fuente trae 98 en casi
# toda la cohorte, que el archivo de reglas acota a V126 = 6 o 99 y obliga a
# V128 distinto de 0.
SUPUESTO_V125_TERAPIA_SISTEMICA = "2"   # recibe terapia sistemica al corte
SUPUESTO_V125_PALIATIVO = "11"           # manejo de cuidado paliativo
SUPUESTO_V125_ANTECEDENTE = "9"          # antecedente de cancer
SUPUESTO_V126_ESTABLE = "5"              # enfermedad estable
SUPUESTO_V126_SEGUIMIENTO = "7"          # en seguimiento por antecedente de cancer

# V129 (novedad clinica a la fecha de corte).
SUPUESTO_V129_MANEJO_INICIAL = "1"       # usuario en manejo inicial curativo
SUPUESTO_V129_MANEJO_PALIATIVO = "11"    # usuario en manejo paliativo

# V61 (ubicacion temporal del ultimo ciclo). Cuando el ultimo ciclo sigue
# vigente (V72 = 3) las reglas B5059 y B5111 exigen un valor entre 1 y 14.
SUPUESTO_V61_TRATAMIENTO_INICIAL = "2"  # tratamiento inicial curativo

# Textos que se imprimen en el reporte para que los supuestos queden visibles.
SUPUESTOS_REPORTE = [
    "V25  : la fuente trae marcadores de posicion ('IPS', '1234567890', '99', '96'); "
    "se graba el REPS de la IPS que reporta. Falta el REPS real de cada IPS.",
    "V41  : se recalcula el objetivo de la intervencion a partir de lo que el "
    "paciente recibio de verdad (V45, V114) y de la novedad clinica (V129): "
    "2 (ofrecer tratamiento) o 3 (seguimiento posterior al tratamiento).",
    "V48  : la fuente trae 0; se usa 13 (paliacion) si V40 = 2 y 2 (tratamiento "
    "inicial curativo) en los demas casos.",
    "V49  : la fuente trae 0; se usa la fecha de diagnostico (V18).",
    "V53  : la fuente trae 0 (cantidad de medicamentos no informada); se declara "
    "1 medicamento antineoplasico (unico valor admisible: 99 obligaria a "
    "V128 = 14, que el instructivo de 2026 no admite).",
    "V53.1 : se usa 99 (codigo ATC no informado); V53.2 a V56 se conservan en 97 "
    "(no aplica, no hay mas medicamentos registrados).",
    "V125 : se deduce del tratamiento registrado: 2 (terapia sistemica) si hubo "
    "quimio en el periodo, 8 (en seguimiento) si ademas hubo manejo paliativo y "
    "9 (antecedente de cancer) si no hubo ningun tratamiento en el periodo.",
    "V126 : 5 (enfermedad estable) cuando hubo terapia sistemica y 98 (no aplica "
    "en este periodo) cuando el paciente esta en seguimiento.",
    "V128 : 0 (sin novedad) salvo diagnostico en el periodo (2), abandono de "
    "tratamiento (9) o fallecimiento (4).",
    "V133 : la fuente viene vacia en toda la cohorte; se graba 0.",
    "Traslado de fechas: las reglas son del periodo 2024 y sus fechas de periodo "
    "se trasladaron a 2026 (2023-01-01 -> 2025-01-01, 2024-01-01 -> 2026-01-01, "
    "2022-11-01 -> 2025-11-01, 2021-01-01 -> 2023-01-01).",
]

# ==============================================================================
# 1.2 BLOQUES DE VARIABLES QUE SE LLENAN EN BLOQUE ("NO APLICA")
# ------------------------------------------------------------------------------
# Contenido obligatorio cuando el bloque no aplica, segun las reglas B4934
# (V45 = 98), B4935 (V45 = 98), B4939 (V46 = 98), y las reglas de cirugia,
# radioterapia y trasplante.
# ==============================================================================
BLOQUE_QUIMIO_NO_APLICA = {
    "46": NO_APLICA, "47": NO_APLICA, "48": NO_APLICA, "49": FECHA_NO_APLICA,
    "50": NO_APLICA, "51": NO_APLICA, "52": NO_APLICA, "53": NO_APLICA,
    "53.1": NO_APLICA, "53.2": NO_APLICA, "53.3": NO_APLICA, "53.4": NO_APLICA,
    "53.5": NO_APLICA, "53.6": NO_APLICA, "53.7": NO_APLICA, "53.8": NO_APLICA,
    "53.9": NO_APLICA, "54": NO_APLICA, "55": NO_APLICA, "56": NO_APLICA,
    "57": NO_APLICA, "58": FECHA_NO_APLICA, "59": NO_APLICA, "60": NO_APLICA,
    "61": NO_APLICA, "62": FECHA_NO_APLICA, "63": NO_APLICA, "64": NO_APLICA,
    "65": NO_APLICA, "66": NO_APLICA, "66.1": NO_APLICA, "66.2": NO_APLICA,
    "66.3": NO_APLICA, "66.4": NO_APLICA, "66.5": NO_APLICA, "66.6": NO_APLICA,
    "66.7": NO_APLICA, "66.8": NO_APLICA, "66.9": NO_APLICA, "67": NO_APLICA,
    "68": NO_APLICA, "69": NO_APLICA, "70": NO_APLICA, "71": FECHA_NO_APLICA,
    "72": NO_APLICA, "73": NO_APLICA,
}

# Subvariables de la variable 46: cuando V46 = 98 valen 97 (regla B4939).
BLOQUE_FASES_QUIMIO = {f"46.{n}": "97" for n in range(1, 9)}

BLOQUE_CIRUGIA_NO_APLICA = {
    "75": NO_APLICA, "76": FECHA_NO_APLICA, "77": NO_APLICA, "78": NO_APLICA,
    "79": NO_APLICA, "80": FECHA_NO_APLICA, "81": NO_APLICA, "82": NO_APLICA,
    "83": NO_APLICA, "84": NO_APLICA, "85": NO_APLICA,
}

BLOQUE_RADIOTERAPIA_NO_APLICA = {
    "87": NO_APLICA, "88": FECHA_NO_APLICA, "89": NO_APLICA, "90": NO_APLICA,
    "91": NO_APLICA, "92": NO_APLICA, "93": NO_APLICA, "94": FECHA_NO_APLICA,
    "95": NO_APLICA, "96": NO_APLICA, "97": FECHA_NO_APLICA, "98": NO_APLICA,
    "99": NO_APLICA, "100": NO_APLICA, "101": NO_APLICA, "102": NO_APLICA,
    "103": FECHA_NO_APLICA, "104": NO_APLICA, "105": NO_APLICA,
}

BLOQUE_TRASPLANTE_NO_APLICA = {
    "107": NO_APLICA, "108": NO_APLICA, "109": FECHA_NO_APLICA, "110": NO_APLICA,
}

# Variables de fecha segun el instructivo (31 campos). Se detectan tambien por
# el tipo de dato (datetime) de Excel, pero la lista es la fuente primaria:
# los nombres VAR125/VAR129 contienen "FechaCorte" sin ser campos de fecha.
VARIABLES_FECHA = {
    "7", "16", "18", "19", "20", "23", "24", "26", "30", "32", "35", "39", "43",
    "49", "58", "62", "71", "76", "80", "88", "94", "97", "103", "109", "112",
    "115", "118", "121", "130", "131", "134",
}

# Marcadores de "no aplica" / "desconocido" en variables numericas. No deben
# satisfacer comparaciones de orden contra un valor numerico normal.
VALORES_SENTINELA = {"97", "98", "99"}

# Variables que son codigos de habilitacion IPS (REPS, 12 digitos)
VARIABLES_CODIGO_IPS = {
    "25", "51", "52", "64", "65", "77", "82", "92", "93", "101", "102", "110",
    "113", "116", "119", "122",
}

# Variables que son codigos ATC (archivo operativo CAC 2026)
VARIABLES_ATC = {
    "53.1", "53.2", "53.3", "53.4", "53.5", "53.6", "53.7", "53.8", "53.9",
    "54", "55", "56",
    "66.1", "66.2", "66.3", "66.4", "66.5", "66.6", "66.7", "66.8", "66.9",
    "67", "68", "69",
}

# Variables de fecha: wildcard a usar cuando el dato no existe
# (instructivo: 1800-01-01 desconocido, 1845-01-01 no se realizo, 1840-01-01 in situ)
WILDCARD_FECHA_POR_VAR = {
    "130": FECHA_NO_APLICA,        # fecha de desafiliacion
    "131": FECHA_NO_APLICA,        # fecha de muerte
    "49": FECHA_NO_APLICA,         # fecha inicio primer ciclo
    "58": FECHA_NO_APLICA,         # fecha finalizacion primer ciclo
    "62": FECHA_NO_APLICA,         # fecha inicio ultimo ciclo
    "71": FECHA_NO_APLICA,         # fecha finalizacion ultimo ciclo
    "76": FECHA_NO_APLICA,         # fecha primera cirugia
    "80": FECHA_NO_APLICA,         # fecha ultima cirugia
    "88": FECHA_NO_APLICA,         # fecha inicio radioterapia
    "94": FECHA_NO_APLICA,         # fecha finalizacion radioterapia
    "97": FECHA_NO_APLICA,         # fecha inicio ultimo esquema radioterapia
    "103": FECHA_NO_APLICA,        # fecha finalizacion ultimo esquema radioterapia
    "109": FECHA_NO_APLICA,        # fecha trasplante
    "112": FECHA_NO_APLICA,        # fecha cirugia reconstructiva
    "115": FECHA_NO_APLICA,        # fecha primera consulta paliativo
    "118": FECHA_NO_APLICA,        # fecha primera consulta psiquiatria
    "121": FECHA_NO_APLICA,        # fecha consulta inicial nutricion
}


# ==============================================================================
# 2. UTILIDADES DE TEXTO
# ==============================================================================

def transliterar(valor):
    """Elimina caracteres especiales prohibidos por el instructivo (pagina 5)."""
    if valor is None:
        return ""
    texto = str(valor)
    for origen, destino in MAPA_TRANSLITERACION.items():
        texto = texto.replace(origen, destino)
    # Resto de caracteres no ASCII se normalizan a su equivalente sin acento
    texto = (
        unicodedata.normalize("NFKD", texto)
        .encode("ascii", "ignore")
        .decode("ascii")
    )
    # Caracteres de control y de estructura prohibidos dentro de un campo
    texto = re.sub(r"[\t\r\n\v\f]", " ", texto)
    texto = texto.replace(";", " ").replace("'", " ").replace('"', " ")
    texto = texto.replace("#", " ").replace("&", " ")
    # No se completan con espacios (instructivo pagina 5)
    return re.sub(r"\s+", " ", texto).strip()


def a_texto(valor):
    """Convierte cualquier valor de Excel a su representacion textual limpia."""
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d")
    if isinstance(valor, date):
        return valor.strftime("%Y-%m-%d")
    if isinstance(valor, bool):
        return "1" if valor else "0"
    if isinstance(valor, float):
        if valor != valor:  # NaN
            return ""
        if valor.is_integer():
            return str(int(valor))
        return repr(valor).rstrip("0").rstrip(".")
    if isinstance(valor, int):
        return str(valor)
    return transliterar(valor).strip()


def normalizar_fecha(valor, wildcard_default=FECHA_DESCONOCIDO):
    """Normaliza una fecha al formato AAAA-MM-DD exigida por el instructivo."""
    if valor is None or valor == "":
        return wildcard_default
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d")
    if isinstance(valor, date):
        return valor.strftime("%Y-%m-%d")
    texto = str(valor).strip()
    if not texto:
        return wildcard_default
    if texto in COMODINES_FECHA:
        return texto
    # Formato ISO ya normalizado
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", texto):
        return texto
    # Marca de tiempo ISO
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})[ T]", texto)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    # Formatos con separadores alternativos
    m = re.match(r"^(\d{4})[/.](\d{1,2})[/.](\d{1,2})", texto)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    # Marca de tiempo en formato serial de Excel (numero)
    if re.fullmatch(r"\d{5}", texto):
        try:
            serial = datetime(1899, 12, 30) + __import__("datetime").timedelta(days=int(texto))
            return serial.strftime("%Y-%m-%d")
        except (ValueError, OverflowError):
            return wildcard_default
    return wildcard_default


def comparar(operador, izquierda, derecha):
    """Compara dos valores con el operador dado, tolerando tipos mixtos."""
    if izquierda is None or derecha is None:
        return False
    try:
        iz = float(izquierda)
        de = float(derecha)
        if iz.is_integer() and de.is_integer():
            iz, de = int(iz), int(de)
    except (ValueError, TypeError):
        iz, de = str(izquierda), str(derecha)
    if operador == "=":
        return iz == de
    if operador == "<>":
        return iz != de
    try:
        if operador == "<":
            return iz < de
        if operador == "<=":
            return iz <= de
        if operador == ">":
            return iz > de
        if operador == ">=":
            return iz >= de
    except TypeError:
        return False
    return False


# ==============================================================================
# 3. MOTOR DE REGLAS DE VALIDACION
# ==============================================================================

RE_CLAUSULA = re.compile(
    r"^V(?P<a>[\d.]+)\s+sea\s*(?P<aop><>|<=|>=|<|>|=)?\s*(?P<aval>\S+)\s+Entonces\s+"
    r"V(?P<b>[\d.]+)\s*(?P<rest>.*)$"
)
RE_SI_MISMO = re.compile(r"^V(?P<a>[\d.]+)\s+(?P<rest>.*)$")
RE_DEBE_SER = re.compile(
    r"^(?P<name>.*?)\s*DEBE\s+SER\s*(?P<op><>|<=|>=|<|>|=)\s*(?P<val>.+)$"
)
RE_REF_VARIABLE = re.compile(r"^V(\d+(?:\.\d+)?)")


class Clausula:
    """Una clausula evaluable: 'Cuando Va <op> <val> Entonces Vb <op> <val>'."""

    def __init__(self, variable_a, op_a, valor_a, variable_b, op_b, valor_b, es_ref):
        self.variable_a = variable_a
        self.op_a = op_a
        self.valor_a = valor_a
        self.variable_b = variable_b
        self.op_b = op_b
        self.valor_b = valor_b
        self.es_ref = es_ref

    def guardia_cumple(self, fila):
        """
        Evalua la condicion de la clausula ('Cuando ...').

        Devuelve True si la condicion se cumple (o si la clausula no tiene
        condicion). Un valor marcador (97/98/99) no activa comparaciones de
        orden contra valores numericos normales: "Cuando V50 sea >= 2" no
        aplica cuando V50 vale 98 (no aplica).
        """
        if self.op_a is None:
            return True
        valor_a = fila.get(self.variable_a, "")
        if valor_a == "":
            return False
        if (
            self.op_a in ("<", "<=", ">", ">=")
            and valor_a in VALORES_SENTINELA
            and self.valor_a not in VALORES_SENTINELA
        ):
            return False
        return comparar(self.op_a, valor_a, self.valor_a)

    def evaluar(self, fila, indice):
        """Carga de la clausula. True (cumple), False (no cumple) o None."""
        valor_b = fila.get(self.variable_b, "")
        if valor_b == "":
            return None
        if self.es_ref:
            otra = fila.get(self.valor_b[1:], "")
            if otra == "":
                return None
            if self.op_b == "<>" and valor_b == otra:
                return False
            return comparar(self.op_b, valor_b, otra)
        # Las comparaciones de rango no aplican sobre comodines de fecha
        if valor_b in COMODINES_FECHA and self.op_b != "=":
            return None
        return comparar(self.op_b, valor_b, self.valor_b)


class Regla:
    """
    Regla logica del archivo de validacion.

    El archivo de reglas encadena clausulas con los conectores de la columna 5:
    'Y' conjuncion (AND) y 'O' disyucion (OR). Un conector vacio cierra la
    clausula y abre una regla independiente. La precedencia es AND > OR, es
    decir, las clausulas unidas por 'Y' forman un grupo (se evaluan juntas) y
    los grupos se unen con 'O' (basta con que uno se cumpla). Ejemplo del
    archivo:

        V19 DEBE SER = 1800-01-01   [O]
        V19 DEBE SER >= V7...        [Y]
        V19 DEBE SER <= V134...      [Y]
        V19 DEBE SER <= V18...

    se interpreta como:  (V19 = 1800-01-01)  OR  (V19 >= V7 AND <= V134 AND <= V18)

    Dentro de un grupo todas las condiciones ('Cuando ...') deben cumplirse
    para que el grupo aplique; si alguna no se cumple, el grupo no aplica.
    Esto reproduce las reglas compuestas del archivo, por ejemplo B4976:

        Cuando V49 sea = 1845-01-01 Entonces V45 = 98      [independiente]
        Cuando V50 sea >= 2 Entonces V52 <> 98              [compuesta]
        Cuando V50 sea < 98 Entonces V52 <> 98

    donde la segunda y la tercera forman una sola condicion (V50 entre 2 y 97).
    """

    def __init__(self, identificador, variable, grupos, texto):
        self.id = identificador
        self.variable = variable
        self.grupos = grupos
        self.texto = texto
        self.dominio = None

    def evaluar(self, fila, indice):
        """True si la regla se cumple; False si se incumple; None si no aplica."""
        return self.evaluar_detalle(fila, indice)[0]

    def evaluar_detalle(self, fila, indice):
        """
        Igual que evaluar(), pero devuelve (estado, clausula_incumple).

        La clausula devuelta es la que produjo el incumplimiento, de modo que el
        reporte senale la variable y el valor realmente responsables y no el
        de la primera clausula del bloque.
        """
        resultados = []
        fallida = None
        for grupo in self.grupos:
            # La regla solo aplica si se cumplen todas sus condiciones
            if not all(c.guardia_cumple(fila) for c in grupo):
                resultados.append(None)
                continue
            estados = [c.evaluar(fila, indice) for c in grupo]
            if all(e is True for e in estados):
                resultados.append(True)
            elif any(e is False for e in estados):
                resultados.append(False)
                if fallida is None:
                    fallida = grupo[estados.index(False)]
            else:
                resultados.append(None)
        if True in resultados:
            return True, None
        if False in resultados:
            return False, fallida
        return None, None


class ReglaDefinicion:
    """
    Regla de definicion (dominio) de una variable.

    El archivo de validacion declara el dominio de cada variable en un bloque
    de clausulas propias, por ejemplo:

        V89 DEBE SER = 1   [O]
        V89 DEBE SER = 2   [O]
        V89 DEBE SER = 3   [O]
        ...

    Ese bloque puede estar repartido en varias secciones del archivo (asi ocurre
    con V89, V48 y V129), por lo que las clausulas de definicion se agrupan por
    variable. La interpretacion es:

        el valor esta en el conjunto de valores permitidos
        OR  (el valor cumple todas las restricciones de rango)

    Las restricciones de rango no se evaluan sobre los comodines de fecha
    (1800-01-01 desconocido, 1845-01-01 no aplica).
    """

    def __init__(self, variable, dominio, relacionales, texto):
        self.id = "DEFINICION"
        self.variable = variable
        self.dominio = dominio
        self.relacionales = relacionales
        self.texto = texto
        self.grupos = []

    def evaluar(self, fila, indice):
        return self.evaluar_detalle(fila, indice)[0]

    def evaluar_detalle(self, fila, indice):
        """Devuelve (estado, clausula_incumple) para la regla de definicion."""
        valor = fila.get(self.variable, "")
        if valor == "":
            return None, None
        if valor in self.dominio:
            return True, None
        if not self.relacionales:
            return False, None
        for grupo in self.relacionales:
            estados = [c.evaluar(fila, indice) for c in grupo]
            if any(e is False for e in estados):
                return False, grupo[estados.index(False)]
        return True, None


def _parsear_clausula(descripcion):
    texto = re.sub(r"\s+", " ", descripcion).strip()
    texto = re.sub(r"^Cuando\s+", "", texto)
    coincidencia = RE_CLAUSULA.match(texto)
    if coincidencia:
        variable_a = coincidencia.group("a")
        op_a = coincidencia.group("aop")
        valor_a = coincidencia.group("aval")
        variable_b = coincidencia.group("b")
        resto = coincidencia.group("rest")
    else:
        coincidencia = RE_SI_MISMO.match(texto)
        if not coincidencia:
            return None
        variable_a = coincidencia.group("a")
        op_a = None
        valor_a = None
        variable_b = coincidencia.group("a")
        resto = coincidencia.group("rest")
    coincidencia = RE_DEBE_SER.match(resto)
    if not coincidencia:
        return None
    op_b = coincidencia.group("op")
    valor_b = coincidencia.group("val").strip()
    # Las fechas de referencia del periodo del archivo de reglas (2024) se
    # trasladan al periodo del reporte (2026), tanto en la condicion como
    # en el valor esperado.
    valor_b = TRASLADO_FECHAS_PERIODO.get(valor_b, valor_b)
    if valor_a is not None:
        valor_a = TRASLADO_FECHAS_PERIODO.get(valor_a, valor_a)
    es_ref = bool(RE_REF_VARIABLE.match(valor_b)) and not re.match(r"^V\d+-\d", valor_b)
    return Clausula(variable_a, op_a, valor_a, variable_b, op_b, valor_b, es_ref)


def cargar_reglas(ruta):
    """Lee el archivo de reglas de validacion y devuelve la lista de Regla."""
    libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hoja = libro["Hoja1"]
    filas = [
        [("" if v is None else str(v).strip()) for v in fila]
        for fila in hoja.iter_rows(values_only=True)
    ]
    libro.close()

    # Agrupar filas por identificador de error (las filas sin tipo son continuaciones)
    bloques = []
    actual = None
    for fila in filas[1:]:
        numero, nombre, tipo, descripcion, conector = (fila + [""] * 5)[:5]
        if tipo:
            actual = {"id": tipo, "variable": numero, "items": []}
            bloques.append(actual)
        if actual is not None and descripcion:
            actual["items"].append((descripcion, conector))

    reglas = []
    definiciones = defaultdict(list)  # variable -> clausulas propias del archivo
    for bloque in bloques:
        if not bloque["id"].startswith("B"):
            continue
        # El conector de una fila describe como se une esa fila con la
        # siguiente. Por eso, para saber si una clausula continua el grupo
        # anterior hay que mirar el conector de la FILA PREVIA (no el de la
        # clausula anterior, que puede estar en otra fila).
        grupos = []
        grupo_definicion = defaultdict(list)  # variable -> lista de grupos
        grupo_actual = None
        variable_actual = None
        previa_propia = False
        for indice, (descripcion, conector) in enumerate(bloque["items"]):
            clausula = _parsear_clausula(descripcion)
            if clausula is None:
                continue
            conector_anterior = bloque["items"][indice - 1][1] if indice > 0 else ""
            continua = conector_anterior == "Y"

            if clausula.op_a is not None:
                # Regla condicional: 'Y' mantiene la clausula en el grupo
                # actual, 'O' y el conector vacio abren un grupo nuevo
                # (alternativa o regla independiente).
                if grupo_actual is not None and continua:
                    grupo_actual.append(clausula)
                else:
                    grupo_actual = [clausula]
                    grupos.append(grupo_actual)
                previa_propia = False
            else:
                # Regla de definicion: se agrupa por variable mas adelante.
                if (grupo_actual is not None and continua and previa_propia
                        and variable_actual == clausula.variable_b):
                    grupo_actual.append(clausula)
                else:
                    grupo_actual = [clausula]
                    variable_actual = clausula.variable_b
                    grupo_definicion[clausula.variable_b].append(grupo_actual)
                previa_propia = True

        if grupos:
            reglas.append(
                Regla(
                    bloque["id"],
                    bloque["variable"],
                    grupos,
                    bloque["items"][0][0],
                )
            )

        for variable, lista in grupo_definicion.items():
            definiciones[variable].extend(lista)

    # Reglas de definicion: una por variable, con el dominio unificado.
    for variable, grupos in definiciones.items():
        dominio = set()
        relacionales = []
        for grupo in grupos:
            relacional = []
            for clausula in grupo:
                if clausula.op_b == "=" and not clausula.es_ref:
                    dominio.add(clausula.valor_b)
                else:
                    relacional.append(clausula)
            if relacional:
                relacionales.append(relacional)
        partes = []
        if dominio:
            partes.append(f"valores admitidos {sorted(dominio)}")
        for grupo in relacionales:
            partes.append(" y ".join(
                f"V{cl.variable_b} {cl.op_b} {cl.valor_b}" for cl in grupo
            ))
        texto = f"definicion de V{variable}: " + (" o ".join(partes) if partes else "sin restricciones")
        reglas.append(ReglaDefinicion(variable, dominio, relacionales, texto))

    return reglas


def cargar_codigos_atc(ruta):
    """Devuelve el conjunto de codigos ATC validos del archivo operativo."""
    if not os.path.exists(ruta):
        return set()
    libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hoja = libro["ATC"]
    codigos = set()
    for i, fila in enumerate(hoja.iter_rows(values_only=True)):
        if i < 3:
            continue
        codigo = "" if fila[0] is None else str(fila[0]).strip()
        if codigo:
            codigos.add(codigo.upper())
    libro.close()
    return codigos


# ==============================================================================
# 4. CARGA DE LA BASE DE DATOS FUENTE
# ==============================================================================

def cargar_fuente(ruta):
    """Lee el Excel de la cohorte y devuelve (encabezados, filas, indice de variables)."""
    libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hoja = libro[libro.sheetnames[0]]
    iterador = hoja.iter_rows(values_only=True)
    encabezados = [("" if h is None else str(h).strip()) for h in next(iterador)]
    filas = [list(f) for f in iterador]
    libro.close()

    indice = {}
    for i, encabezado in enumerate(encabezados):
        coincidencia = re.match(r"^VAR(\d+)(?:_(\d+))?_", encabezado)
        if coincidencia:
            clave = coincidencia.group(1)
            if coincidencia.group(2):
                clave += "." + coincidencia.group(2)
            indice[clave] = i
    return encabezados, filas, indice


# ==============================================================================
# 5. TRANSFORMACIONES
# ==============================================================================

def _es_variable_fecha(indice, variable, encabezados, filas):
    """Una variable es de fecha si el instructivo la define como fecha o si Excel
    la almacena como datetime."""
    if variable in VARIABLES_FECHA:
        return True
    posicion = indice.get(variable)
    if posicion is None:
        return False
    nombre = encabezados[posicion].lower()
    if "fecha" not in nombre:
        return False
    for fila in filas:
        if isinstance(fila[posicion], (datetime, date)):
            return True
    return False


def normalizar_fechas(filas, indice, encabezados):
    """Convierte todas las columnas de fecha al formato AAAA-MM-DD."""
    columnas_fecha = [
        variable
        for variable in indice
        if _es_variable_fecha(indice, variable, encabezados, filas)
    ]
    for variable in columnas_fecha:
        posicion = indice[variable]
        comodin = WILDCARD_FECHA_POR_VAR.get(variable, FECHA_DESCONOCIDO)
        for fila in filas:
            fila[posicion] = normalizar_fecha(fila[posicion], comodin)
    return columnas_fecha


def limpiar_texto(filas, indice):
    """Translitera caracteres especiales y elimina espacios sobrantes."""
    for variable, posicion in indice.items():
        if variable in WILDCARD_FECHA_POR_VAR:
            continue
        for fila in filas:
            valor = fila[posicion]
            if valor in (None, ""):
                continue
            if isinstance(valor, (datetime, date)):
                continue
            fila[posicion] = transliterar(a_texto(valor))
    return filas


def rellenar_vacios(filas, indice, encabezados):
    """Ninguna variable puede llegar vacia al archivo plano (instructivo pag. 5)."""
    rellenados = Counter()
    for variable, posicion in indice.items():
        if variable in WILDCARD_FECHA_POR_VAR:
            continue
        for fila in filas:
            if fila[posicion] in (None, ""):
                fila[posicion] = NO_APLICA
                rellenados[variable] += 1
    return rellenados


def _puede_v128_nuevo_diagnostico(valor, fila):
    """
    Indica si la variable 128 puede tomar el valor 2 ("nuevo diagnostico de
    cancer entre el 2/1/2025 y el 1/1/2026").

    Las reglas B4813, B4832, B4833, B4845, B4880, B4881 y B4926 prohiben ese
    valor cuando el dato de soporte esta en comodin: fecha de diagnostico
    desconocida, tipo o motivo de estudio no histopatologico ausente, fecha de
    la prueba HER2 desconocida o antecedente de otro cancer primario en 99.
    """
    if valor(fila, "18") == FECHA_DESCONOCIDO:      # B4813
        return False
    if valor(fila, "21") == "99":                   # B4832
        return False
    if valor(fila, "22") == "99":                   # B4833
        return False
    if valor(fila, "24") == FECHA_DESCONOCIDO:      # B4845
        return False
    if valor(fila, "32") == FECHA_DESCONOCIDO:      # B4880
        return False
    if valor(fila, "33") == "99":                   # B4881
        return False
    if valor(fila, "42") == "99":                   # B4926
        return False
    return True


def _tiene_datos_propios_corte(valor, fila):
    """
    Indica si el registro trae informacion propia del bloque de situacion al
    corte (V125, V126 y V129) o si son los valores por defecto de la fuente
    (98, 98 y 3).

    Los registros con datos propios se dejan tal como vienen: son
    inconsistentes entre si y requieren revision clinica, por lo que se
    reportan como pendientes en vez de sobrescribirlos.
    """
    for variable, defecto in (("125", "98"), ("126", "98"), ("129", "3")):
        if valor(fila, variable) not in ("", defecto):
            return True
    return False


def _cierra_ultimo_ciclo(valor, fila, asignar):
    """
    Deja el bloque del ultimo ciclo de quimioterapia en un estado coherente
    con un ciclo en curso (V59 = 3, "no finalizado, aun en tratamiento").

    Las reglas B5059, B5069, B5101, B5103 y B5105 obligan a que, si el ultimo
    ciclo existe (V61 entre 1 y 14), se registren la IPS que lo suministra, la
    cantidad de medicamentos y el codigo ATC; si no existe (V61 = 97 o 98) todo
    el bloque queda en "no aplica". La fuente no tiene esos datos, por lo que
    se derivan del primer ciclo (misma IPS y mismo medicamento).
    """
    reps = valor(fila, "51")
    if reps in ("", NO_APLICA):
        return
    asignar(fila, "61", SUPUESTO_V61_TRATAMIENTO_INICIAL,
            "ultimo ciclo en curso: ubicacion temporal del ultimo ciclo")
    asignar(fila, "63", "1", "ultimo ciclo: una sola IPS lo suministra")
    asignar(fila, "64", reps, "ultimo ciclo: REPS de la IPS que lo suministra")
    asignar(fila, "65", NO_APLICA, "ultimo ciclo: no hay segunda IPS")
    asignar(fila, "66", "1", "ultimo ciclo: un medicamento antineoplasico")
    asignar(fila, "66.1", SUPUESTO_ATC_DESCONOCIDO,
            "ultimo ciclo: codigo ATC no informado en la fuente")
    for variable in ("66.2", "66.3", "66.4", "66.5", "66.6",
                      "66.7", "66.8", "66.9"):
        asignar(fila, variable, "97", "ultimo ciclo: no hay mas medicamentos")
    asignar(fila, "67", "97", "ultimo ciclo: sin medicamento no POS")
    asignar(fila, "68", "97", "ultimo ciclo: sin segundo medicamento no POS")
    asignar(fila, "69", "97", "ultimo ciclo: sin tercer medicamento no POS")
    asignar(fila, "71", FECHA_DESCONOCIDO,
            "ultimo ciclo en curso: sin fecha de finalizacion")
    asignar(fila, "72", "3", "ultimo ciclo en curso: no finalizado")
    asignar(fila, "73", NO_APLICA,
            "ultimo ciclo en curso: sin motivo de finalizacion prematura")


def corregir_dependencias(filas, indice):
    """
    Correcciones deterministas derivadas del instructivo y de las reglas B.

    Cada correccion se apoya en una regla concreta del archivo de validacion
    o en el instructivo de reporte. Los valores que no se pueden derivar de las
    fuentes se declaran en el bloque SUPUESTOS del archivo y quedan reportados
    como pendientes para revision clinica.

    Devuelve (correcciones, motivos): el primero cuenta los registros
    corregidos por variable y el segundo el motivo de cada correccion.
    """
    correcciones = Counter()
    motivos = Counter()

    def columna(variable):
        return indice[variable]

    def valor(fila, variable):
        return a_texto(fila[columna(variable)])

    def asignar(fila, variable, nuevo, motivo=""):
        if variable not in indice:
            return
        if fila[columna(variable)] != nuevo:
            correcciones[variable] += 1
            motivos[(variable, motivo)] += 1
            fila[columna(variable)] = nuevo

    for fila in filas:
        # --- Variable 134: fecha de corte del periodo (instructivo variable 134)
        asignar(fila, "134", FECHA_CORTE, "fecha de corte del periodo")

        # --- Variable 25: codigo de habilitacion REPS de la IPS que confirma.
        # La fuente trae marcadores de posicion ("IPS", "1234567890", "99").
        asignar(fila, "25", CODIGO_REPS_IPS, "REPS de la IPS que reporta (marcador de posicion)")

        # --- Variable 9: Ocupacion (CIUO). Si viene como texto o con valor
        # invalido, se normaliza a 9999 ("No existe informacion").
        v9 = valor(fila, "9").strip().upper().replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
        if v9 in ("", "INDEPENDIENTE", "TRABAJADOR INDEPENDIENTE", "AUTONOMO", "NO APLICA", "N/A", "NA", "-", "."):
            asignar(fila, "9", "9999", "V9 con texto/valor no valido -> 9999")
        elif v9 and not re.fullmatch(r"\d{4}", v9):
            # Si es numerico con otra longitud se intenta ajustar
            solo_dig = re.sub(r"\D", "", v9)
            if len(solo_dig) == 4:
                asignar(fila, "9", solo_dig, "V9 normalizado a 4 digitos")
            elif len(solo_dig) > 4:
                asignar(fila, "9", solo_dig[:4], "V9 truncado a 4 digitos")
            else:
                asignar(fila, "9", "9999", "V9 con valor invalido -> 9999")

        # --- Variable 15: Numero telefonico de contacto. Solo se permiten
        # digitos, maximo 10. Si tiene menos de 10 digitos -> 0.
        v15 = valor(fila, "15")
        dig_v15 = re.sub(r"\D", "", v15)
        if len(dig_v15) >= 10:
            asignar(fila, "15", dig_v15[:10], "V15 normalizado a 10 digitos")
        else:
            asignar(fila, "15", "0", "V15 con menos de 10 digitos -> 0")

        # --- Variable 133: serial BDUA. La fuente viene vacia en toda la
        # cohorte; se usa 0, que es el valor que acepta la regla B4572.
        if valor(fila, "133") in ("", NO_APLICA):
            asignar(fila, "133", SUPUESTO_V133, "serial BDUA no informado")

        # ==================================================================
        # Bloque de terapia sistemica (variables 45 a 73)
        # ==================================================================
        if valor(fila, "45") == NO_APLICA:
            # Sin terapia sistemica: reglas B4934, B4935, B4936 y B4939
            for variable, contenido in BLOQUE_QUIMIO_NO_APLICA.items():
                asignar(fila, variable, contenido, "V45 = 98 (sin terapia sistemica)")
        elif valor(fila, "45") == "1":
            # Con terapia sistemica: reglas B4932 y B4986
            #
            # V48 ubicacion temporal del primer ciclo: la fuente trae 0, que
            # no pertenece al dominio. Se usa 13 (manejo paliativo) cuando el
            # objetivo del tratamiento fue paliacion (V40 = 2) y 2 (tratamiento
            # inicial curativo) en los demas casos.
            if valor(fila, "48") not in DOMINIOS["48"] - {NO_APLICA}:
                if valor(fila, "40") == "2":
                    asignar(fila, "48", SUPUESTO_V48_PALIATIVO,
                            "V48 = 0 en la fuente (supuesto: manejo paliativo)")
                else:
                    asignar(fila, "48", SUPUESTO_V48_CURATIVO,
                            "V48 = 0 en la fuente (supuesto: tratamiento inicial curativo)")
            # V49 fecha de inicio del primer ciclo: la fuente trae 0, que no es
            # una fecha. Se toma la fecha de diagnostico (V18), que es el unico
            # dato de fecha disponible y cumple el rango de la regla B4972.
            if valor(fila, "49") in ("", "0", FECHA_NO_APLICA, FECHA_DESCONOCIDO):
                fecha_diagnostico = valor(fila, "18")
                if fecha_diagnostico:
                    asignar(fila, "49", fecha_diagnostico,
                            "V49 = 0 en la fuente (supuesto: fecha de diagnostico V18)")
                else:
                    asignar(fila, "49", FECHA_DESCONOCIDO, "V49 = 0 en la fuente")
            # V53 cantidad de medicamentos antineoplasicos: la fuente trae 0.
            # El valor 99 esta en el dominio pero obliga (B4988) a V128 = 14,
            # que el instructivo de 2026 no admite; se declara un esquema de
            # un medicamento y se deja el ATC como desconocido.
            if valor(fila, "53") in ("", "0"):
                asignar(fila, "53", SUPUESTO_V53_CANTIDAD,
                        "V53 = 0 en la fuente (supuesto: un medicamento propuesto)")
            # V53.1: codigo ATC del medicamento. La regla B4932 exige
            # V53.1 > 98 cuando V45 = 1. V53.2 a V56 se conservan en 97
            # (no aplica) porque solo se declaro un medicamento.
            if valor(fila, "53.1") in ("", "97", NO_APLICA):
                asignar(fila, "53.1", SUPUESTO_ATC_DESCONOCIDO,
                        "codigo ATC no informado en la fuente")
            # Primer ciclo vigente (V59 = 3): no hay fecha de finalizacion
            # ni motivo de finalizacion prematura (regla B5046).
            if valor(fila, "59") == "3":
                asignar(fila, "58", FECHA_DESCONOCIDO,
                        "V59 = 3: el primer ciclo sigue vigente")
                asignar(fila, "60", NO_APLICA,
                        "V59 = 3: no hay motivo de finalizacion prematura")
            # Bloque del ultimo ciclo (V61 a V73). Las reglas B4932, B5041,
            # B5046 y B5059 obligan a registrarlo cuando el paciente recibio
            # terapia sistemica en el periodo; la fuente lo deja en "no
            # aplica" y sus datos se derivan del primer ciclo.
            _cierra_ultimo_ciclo(valor, fila, asignar)

        # ==================================================================
        # Bloque diagnostico (variables 17 a 34)
        # ------------------------------------------------------------------
        # Correcciones que se apoyan en una regla concreta y no cambian el
        # diagnostico del paciente:
        # - B4826 y B4830: si no hubo diagnostico por histopatologia
        #   (V21 = 5 o 10) el motivo de esa situacion es "no aplica" (V22).
        # - B4883: si no se realizo prueba HER2 (V33 = 97) tampoco se
        #   realizo (V31 = 97), no hay fecha de prueba (V32) y la
        #   estadificacion no aplica (V29 = 0).
        # - B5724: el melanoma del ojo (C430) tiene histologia 23.
        # ==================================================================
        if valor(fila, "21") in ("5", "10") and valor(fila, "22") != NO_APLICA:
            asignar(fila, "22", NO_APLICA,
                    "V21 = %s (sin diagnostico por histopatologia): V22 = 98"
                    % valor(fila, "21"))
        if valor(fila, "33") == "97":
            asignar(fila, "31", "97", "V33 = 97 (sin prueba HER2): V31 = 97")
            asignar(fila, "32", FECHA_NO_APLICA_IN_SITU,
                    "V33 = 97 (sin prueba HER2): V32 sin fecha")
            asignar(fila, "29", "0",
                    "V33 = 97 (sin prueba HER2): V29 = 0")
        if valor(fila, "17") == "C430" and valor(fila, "27") != "23":
            asignar(fila, "27", "23",
                    "V17 = C430 (melanoma del ojo): V27 = 23")

        # ==================================================================
        # Bloque de cirugia (variables 74 a 85)
        # ==================================================================
        if valor(fila, "74") == "2":
            for variable, contenido in BLOQUE_CIRUGIA_NO_APLICA.items():
                asignar(fila, variable, contenido, "V74 = 2 (sin cirugias)")

        # ==================================================================
        # Bloque de radioterapia (variables 86 a 105)
        # ==================================================================
        if valor(fila, "86") == NO_APLICA:
            for variable, contenido in BLOQUE_RADIOTERAPIA_NO_APLICA.items():
                asignar(fila, variable, contenido, "V86 = 98 (sin radioterapia)")

        # ==================================================================
        # Bloque de trasplante (variables 106 a 110) y cirugia reconstructiva
        # ==================================================================
        if valor(fila, "106") == NO_APLICA:
            for variable, contenido in BLOQUE_TRASPLANTE_NO_APLICA.items():
                asignar(fila, variable, contenido, "V106 = 98 (sin trasplante)")
        if valor(fila, "111") == NO_APLICA:
            asignar(fila, "112", FECHA_NO_APLICA, "V111 = 98 (sin cirugia reconstructiva)")
            asignar(fila, "113", NO_APLICA, "V111 = 98 (sin cirugia reconstructiva)")

        # ==================================================================
        # Cuidado paliativo, psiquiatria y nutricion
        # ==================================================================
        if valor(fila, "114") == "2":
            asignar(fila, "115", FECHA_NO_APLICA, "V114 = 2 (sin cuidado paliativo)")
            asignar(fila, "116", NO_APLICA, "V114 = 2 (sin cuidado paliativo)")
        elif valor(fila, "114") == "1":
            if valor(fila, "116") in ("", NO_APLICA):
                asignar(fila, "116", CODIGO_REPS_IPS, "V116 sin dato (REPS de la IPS)")
        if valor(fila, "117") == NO_APLICA:
            asignar(fila, "118", FECHA_NO_APLICA, "V117 = 98 (sin valoracion psiquiatrica)")
            asignar(fila, "119", NO_APLICA, "V117 = 98 (sin valoracion psiquiatrica)")
        if valor(fila, "120") == NO_APLICA:
            asignar(fila, "121", FECHA_NO_APLICA, "V120 = 98 (sin valoracion nutricional)")
            asignar(fila, "122", NO_APLICA, "V120 = 98 (sin valoracion nutricional)")
        elif valor(fila, "120") == "1":
            if valor(fila, "122") in ("", NO_APLICA):
                asignar(fila, "122", CODIGO_REPS_IPS, "V122 sin dato (REPS de la IPS)")

        # ==================================================================
        # Situacion actual a la fecha de corte (variables 125 a 132)
        # ------------------------------------------------------------------
        # La fuente no trae informacion real de este bloque (V125 = 98,
        # V126 = 98 y V129 = 3 en casi toda la cohorte), por lo que se deduce
        # de lo que el paciente recibio realmente en el periodo (V45, V114),
        # de la novedad clinica (V129) y de las reglas B4915, B4916, B4921,
        # B5044, B5046, B5059, B5111, B5252, B5260, B5261, B5263, B5269,
        # B5270, B5271, B5288 y B5793.
        # ==================================================================
        v127 = valor(fila, "127")
        v129 = valor(fila, "129")
        v45 = valor(fila, "45")
        v114 = valor(fila, "114")
        v18 = valor(fila, "18")
        v23 = valor(fila, "23")
        v24 = valor(fila, "24")

        if v127 == "2":
            # Fallecido al corte (B5268, B5270, B5271 y B5283).
            asignar(fila, "126", "99", "V127 = 2 (fallecido): V126 = 99")
            asignar(fila, "128", SUPUESTO_V128_FALLECIMIENTO,
                    "V127 = 2 (fallecido): V128 = 4")
            asignar(fila, "129", "12", "V127 = 2 (fallecido): V129 = 12")
            if valor(fila, "131") in ("", FECHA_NO_APLICA):
                asignar(fila, "131", FECHA_NO_APLICA,
                        "V127 = 2 (fallecido): V131 sin fecha")
        else:
            if v129 == "8":
                # Abandono de tratamiento (reglas B5279 y B5288).
                asignar(fila, "128", SUPUESTO_V128_ABANDONO,
                        "V129 = 8 (abandono de tratamiento): V128 = 9")
            if not _tiene_datos_propios_corte(valor, fila):
                if v129 == "8":
                    # Abandono de tratamiento (reglas B5279 y B5288).
                    asignar(fila, "128", SUPUESTO_V128_ABANDONO,
                            "V129 = 8 (abandono de tratamiento): V128 = 9")
                if v45 == "1":
                    # Recibio terapia sistemica dentro del periodo (V45 = 1).
                    # - B4915 y B4921: el objetivo de la intervencion fue
                    #   ofrecer tratamiento (V41 = 2).
                    # - B5044: con el primer ciclo sin fecha de finalizacion
                    #   (V58 = 1800-01-01) el tratamiento a la fecha de corte
                    #   solo puede ser 2, 4, 6, 98 o 10; como no hubo cirugia
                    #   (V74 = 2) ni radioterapia (V86 = 98) corresponde 2.
                    # - B5046, B5059 y B5111: el primer y el ultimo ciclo
                    #   siguen vigentes (V59 = 3), por lo que el ultimo ciclo
                    #   no tiene fecha de finalizacion (V71 = 1800-01-01) y
                    #   esta en curso (V72 = 3).
                    if valor(fila, "41") not in ("2", "5", "6"):
                        asignar(fila, "41", SUPUESTO_V41_TRATAMIENTO,
                                "V45 = 1: el objetivo fue ofrecer tratamiento")
                    asignar(fila, "125", SUPUESTO_V125_TERAPIA_SISTEMICA,
                            "V45 = 1: recibe terapia sistemica a la fecha de corte")
                    asignar(fila, "126", SUPUESTO_V126_ESTABLE,
                            "V45 = 1: enfermedad estable")
                    asignar(fila, "129", SUPUESTO_V129_MANEJO_INICIAL,
                            "V45 = 1: usuario en manejo inicial curativo")
                elif v114 == "1":
                    # Sin terapia sistemica pero con manejo paliativo en el
                    # periodo (B5263): el tratamiento a la fecha de corte es
                    # 11 (manejo de cuidado paliativo) y la novedad clinica es
                    # 11 (usuario en manejo paliativo).
                    if valor(fila, "41") not in ("2", "5", "6"):
                        asignar(fila, "41", SUPUESTO_V41_TRATAMIENTO,
                                "V114 = 1: se ofrecio manejo paliativo")
                    asignar(fila, "125", SUPUESTO_V125_PALIATIVO,
                            "V114 = 1: manejo de cuidado paliativo a la fecha de corte")
                    asignar(fila, "126", SUPUESTO_V126_ESTABLE,
                            "V114 = 1: enfermedad estable")
                    asignar(fila, "129", SUPUESTO_V129_MANEJO_PALIATIVO,
                            "V114 = 1: usuario en manejo paliativo")
                else:
                    # Sin ningun tratamiento en el periodo: el objetivo fue la
                    # observacion posterior al tratamiento inicial (B4921) y a
                    # la fecha de corte tiene antecedente de cancer (B5261),
                    # con seguimiento por antecedente (B5269).
                    if valor(fila, "41") != "3":
                        asignar(fila, "41", SUPUESTO_V41_SEGUIMIENTO,
                                "V45 = 98 y V114 = 2: seguimiento tras el tratamiento inicial")
                    asignar(fila, "125", SUPUESTO_V125_ANTECEDENTE,
                            "V45 = 98 y V114 = 2: antecedente de cancer")
                    asignar(fila, "126", SUPUESTO_V126_SEGUIMIENTO,
                            "V125 = 9: en seguimiento por antecedente de cancer")
            elif v129 == "8" and not _tiene_datos_propios_corte(valor, fila):
                # Abandono de tratamiento (reglas B5279 y B5288).
                asignar(fila, "128", SUPUESTO_V128_ABANDONO,
                        "V129 = 8 (abandono de tratamiento): V128 = 9")

            # Novedad administrativa (V128).
            if not _tiene_datos_propios_corte(valor, fila) and valor(fila, "128") == "0":
                fecha_en_periodo = (v18 >= FECHA_INICIO_PERIODO
                                    or v23 >= FECHA_INICIO_PERIODO
                                    or v24 >= FECHA_INICIO_PERIODO)
                if fecha_en_periodo:
                    if v18 >= FECHA_INICIO_PERIODO and _puede_v128_nuevo_diagnostico(valor, fila):
                        asignar(fila, "128", SUPUESTO_V128_NUEVO_DIAGNOSTICO,
                                "diagnostico en el periodo: V128 = 2 (nuevo diagnostico)")
                    elif v18 < FECHA_INICIO_PERIODO:
                        # Diagnostico anterior al periodo que no estaba en el
                        # reporte anterior (opcion 3 del instructivo).
                        asignar(fila, "128", "3",
                                "diagnostico anterior al periodo: V128 = 3")
                    elif valor(fila, "16") >= FECHA_INICIO_PERIODO:
                        # Ingreso a la EPS en el periodo con diagnostico previo
                        # (opcion 1 del instructivo; regla B5273).
                        asignar(fila, "128", "1",
                                "ingreso a la EPS en el periodo: V128 = 1")
                    # Si ninguna de las anteriores aplica se deja V128 = 0 y
                    # la regla queda reportada como pendiente administrativo.

            if valor(fila, "131") in ("", FECHA_NO_APLICA):
                asignar(fila, "131", FECHA_NO_APLICA, "V127 = 1 (vivo): V131 = 1845-01-01")
            if valor(fila, "132") in ("", NO_APLICA):
                asignar(fila, "132", NO_APLICA, "V127 = 1 (vivo): V132 = 98")

        # B4656: para neoplasias in situ (prefijo D) V29 debe ser 0.
        if valor(fila, "17").startswith("D") and valor(fila, "29") != "0":
            asignar(fila, "29", "0",
                    "V17 = %s (neoplasia in situ): V29 = 0" % valor(fila, "17"))
        # B4653: con fecha de afiliacion anterior a 1993 el regimen no puede
        # ser contributivo (C) ni subsidiado (S), salvo para ESS062 (B4600).
        if valor(fila, "16") < "1993-01-01" and valor(fila, "16") not in COMODINES_FECHA:
            if valor(fila, "11") == "ESS062":
                if valor(fila, "10") != "C":
                    asignar(fila, "10", "C",
                            "V11 = ESS062: V10 = C")
            elif valor(fila, "10") in ("S", "C"):
                asignar(fila, "10", "N",
                        "V16 < 1993-01-01 (antes del SGSSS): V10 = N")
        # B5044/B5254: cuando el primer ciclo no tiene fecha de finalizacion
        # (V58 = 1800-01-01) el tipo de tratamiento a la fecha de corte debe
        # estar entre los permitidos.
        if valor(fila, "58") == FECHA_DESCONOCIDO:
            if valor(fila, "125") not in ("2", "4", "6", "98", "10"):
                asignar(fila, "125", SUPUESTO_V125_TERAPIA_SISTEMICA,
                        "V58 = 1800-01-01: V125 = 2 (terapia sistemica)")
            if valor(fila, "71") in ("", FECHA_NO_APLICA, FECHA_DESCONOCIDO):
                asignar(fila, "71", FECHA_DESCONOCIDO,
                        "V58 = 1800-01-01: V71 = 1800-01-01 (ultimo ciclo sin fecha)")

    return correcciones, motivos


# ==============================================================================
# 6. VALIDACION
# ==============================================================================

def validar(filas, indice, encabezados, reglas, codigos_atc):
    """
    Aplica las reglas del archivo de validacion, los dominios del instructivo
    y las verificaciones estructurales (C, D, E). Devuelve los hallazgos.
    """
    hallazgos = []

    def registrar(nivel, codigo, variable, fila_excel, detalle):
        hallazgos.append({
            "nivel": nivel,
            "codigo": codigo,
            "variable": variable,
            "fila": fila_excel,
            "detalle": detalle,
        })

    # --- Dominios de respuesta (reglas A-nnnn) ---
    for variable, dominio in DOMINIOS.items():
        if variable not in indice:
            continue
        for i, fila in enumerate(filas):
            valor = a_texto(fila[indice[variable]])
            if valor and valor not in dominio:
                registrar("ERROR", "A", variable, i + 2,
                          f"valor '{valor}' fuera del dominio permitido {sorted(dominio)}")

    # --- Contraste entre el dominio del instructivo y el del archivo de reglas ---
    for regla in reglas:
        if not isinstance(regla, ReglaDefinicion) or not regla.dominio:
            continue
        if any(re.fullmatch(r"\d{4}-\d{2}-\d{2}", v) for v in regla.dominio):
            continue  # las variables de fecha solo declaran sus comodines
        dominio = DOMINIOS.get(regla.variable)
        if dominio is None:
            continue
        if dominio != regla.dominio:
            registrar(
                "ADVERTENCIA", "DOMINIO", regla.variable, 0,
                f"instructivo {sorted(dominio)} vs archivo de reglas {sorted(regla.dominio)}",
            )

    # --- Codigos de habilitacion IPS (REPS, 12 digitos) ---
    for variable in VARIABLES_CODIGO_IPS:
        if variable not in indice:
            continue
        for i, fila in enumerate(filas):
            valor = a_texto(fila[indice[variable]])
            if valor in ("", NO_APLICA, DESCONOCIDO, "96", "98"):
                continue
            if not re.fullmatch(r"\d{12}", valor):
                registrar("ERROR", "IPS", variable, i + 2,
                          f"codigo de habilitacion '{valor}' no tiene 12 digitos")

    # --- Codigos ATC ---
    if codigos_atc:
        for variable in VARIABLES_ATC:
            if variable not in indice:
                continue
            for i, fila in enumerate(filas):
                valor = a_texto(fila[indice[variable]]).upper()
                if valor in ("", "97", NO_APLICA, DESCONOCIDO):
                    continue
                if valor not in codigos_atc:
                    registrar("ERROR", "ATC", variable, i + 2,
                              f"codigo ATC '{valor}' no esta en el archivo operativo 2026")

    # --- Reglas condicionales y de definicion del archivo de validacion ---
    for i, fila in enumerate(filas):
        registro = {
            variable: a_texto(fila[posicion])
            for variable, posicion in indice.items()
        }
        for regla in reglas:
            estado, clausula = regla.evaluar_detalle(registro, i)
            if estado is False:
                if clausula is not None:
                    variable_b = clausula.variable_b
                elif isinstance(regla, ReglaDefinicion):
                    variable_b = regla.variable
                else:
                    variable_b = regla.grupos[0][0].variable_b
                valor_b = registro.get(variable_b, "")
                registrar("ERROR", regla.id, variable_b, i + 2,
                          f"valor '{valor_b}' incumple: {regla.texto[:150]}")

    # --- C2: registros duplicados (tipo y numero de ID + neoplasia + fecha dx) ---
    conteo = Counter()
    for i, fila in enumerate(filas):
        clave = (
            a_texto(fila[indice["5"]]),
            a_texto(fila[indice["6"]]),
            a_texto(fila[indice["17"]]),
            a_texto(fila[indice["18"]]),
        )
        conteo[clave] += 1
    for clave, n in conteo.items():
        if n > 1:
            registrar("ERROR", "C2", "5/6/17/18", 0,
                      f"registro duplicado {n} veces: {clave}")

    # --- C10: serial BDUA repetido ---
    conteo = Counter(a_texto(fila[indice["133"]]) for fila in filas)
    for serial, n in conteo.items():
        if n > 1 and serial not in ("", "0"):
            registrar("ERROR", "C10", "133", 0,
                      f"serial BDUA '{serial}' aparece {n} veces")

    # --- D001/D002: caracteres especiales y separadores dentro de los campos ---
    for i, fila in enumerate(filas):
        for variable, posicion in indice.items():
            valor = a_texto(fila[posicion])
            if not valor:
                continue
            if re.search(r"[^\x20-\x7E]", valor):
                registrar("ERROR", "D001", variable, i + 2,
                          f"caracter no valido en '{valor}'")
            if re.search(r"[\t\r\n]", valor):
                registrar("ERROR", "D003", variable, i + 2,
                          "salto de linea o tabulacion dentro del campo")
            if re.search(r"[;'\"]", valor):
                registrar("ERROR", "D002", variable, i + 2,
                          f"caracter prohibido en '{valor}'")

    # --- E001 / calidad: fechas fuera del periodo de medicion ---
    for i, fila in enumerate(filas):
        for variable, posicion in indice.items():
            valor = a_texto(fila[posicion])
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", valor):
                continue
            if valor in COMODINES_FECHA:
                continue
            if valor > FECHA_CORTE:
                registrar("ADVERTENCIA", "FECHA", variable, i + 2,
                          f"fecha '{valor}' posterior a la fecha de corte {FECHA_CORTE}")
            elif valor < "1900-01-01":
                registrar("ADVERTENCIA", "FECHA", variable, i + 2,
                          f"fecha '{valor}' anterior a 1900-01-01")

    # --- E001: novedades administrativas coherentes con el estado vital ---
    for i, fila in enumerate(filas):
        estado = a_texto(fila[indice["127"]])
        novedad = a_texto(fila[indice["128"]])
        if estado == "2" and novedad not in ("4", "10", "6"):
            registrar("ADVERTENCIA", "E001", "128", i + 2,
                      f"estado vital 'Fallecido' con novedad administrativa '{novedad}'")

    return hallazgos


def construir_pendientes(filas, indice, encabezados, hallazgos):
    """Consolida los hallazgos por registro para el archivo de pendientes."""
    por_fila = defaultdict(list)
    for hallazgo in hallazgos:
        if hallazgo["fila"] == 0:
            continue
        por_fila[hallazgo["fila"]].append(hallazgo)
    pendientes = []
    for numero, lista in sorted(por_fila.items()):
        if all(h["nivel"] == "ADVERTENCIA" for h in lista):
            continue
        fila = filas[numero - 2]
        registro = {
            "fila_excel": numero,
            "documento": a_texto(fila[indice["6"]]),
            "neoplasia": a_texto(fila[indice["17"]]),
            "variables": "; ".join(sorted({h["variable"] for h in lista})),
            "codigos": "; ".join(sorted({h["codigo"] for h in lista})),
            "detalle": " | ".join(h["detalle"] for h in lista[:6]),
        }
        pendientes.append(registro)
    return pendientes


# ==============================================================================
# 7. EXPORTACION
# ==============================================================================

def exportar_txt(filas, encabezados, ruta):
    """
    Graba el archivo plano en formato ANSI (cp1252) separado por tabulaciones.
    El nombre del archivo cumple el formato AAAAMMDD_CODEAPB_CANCER.txt.
    """
    with open(ruta, "w", encoding="cp1252", errors="replace", newline="") as archivo:
        archivo.write("\t".join(encabezados) + "\r\n")
        for fila in filas:
            campos = []
            for valor in fila:
                texto = a_texto(valor)
                campos.append(texto if texto else NO_APLICA)
            archivo.write("\t".join(campos) + "\r\n")
    return ruta


def exportar_reporte(ruta, hallazgos, correcciones, reglas, filas, encabezados, indice,
                     motivos=None):
    """Escribe el diagnostico de la validacion en un archivo de texto."""
    total = len(filas)
    errores = [h for h in hallazgos if h["nivel"] == "ERROR"]
    advertencias = [h for h in hallazgos if h["nivel"] == "ADVERTENCIA"]

    with open(ruta, "w", encoding="utf-8") as archivo:
        archivo.write("=" * 78 + "\n")
        archivo.write("REPORTE DE VALIDACION - ARCHIVO PLANO CAC (SISCAC 4.0)\n")
        archivo.write(f"EAPB: {CODIGO_EAPB}   Periodo: 2/1/2025 - 1/1/2026   Corte: {FECHA_CORTE}\n")
        archivo.write("=" * 78 + "\n\n")
        archivo.write(f"Registros procesados : {total}\n")
        archivo.write(f"Reglas de validacion : {len(reglas)}\n")
        archivo.write(f"Correcciones aplicadas: {sum(correcciones.values())}\n")
        archivo.write(f"Errores              : {len(errores)}\n")
        archivo.write(f"Advertencias         : {len(advertencias)}\n\n")

        if correcciones:
            archivo.write("-" * 78 + "\nCORRECCIONES AUTOMATICAS POR VARIABLE\n")
            archivo.write("-" * 78 + "\n")
            for variable, n in correcciones.most_common():
                nombre = next((h for h in encabezados
                               if h.startswith(f"VAR{variable.split('.')[0]}_")), variable)
                archivo.write(f"  V{variable:<8} {n:>5} registros   {nombre}\n")
            archivo.write("\n")

        if motivos:
            archivo.write("-" * 78 + "\nMOTIVO DE CADA CORRECCION\n")
            archivo.write("-" * 78 + "\n")
            for (variable, motivo), n in sorted(motivos.items(), key=lambda kv: -kv[1]):
                archivo.write(f"  V{variable:<8} {n:>5} registros   {motivo}\n")
            archivo.write("\n")

        archivo.write("-" * 78 + "\nSUPUESTOS DECLARADOS (requieren confirmacion)\n")
        archivo.write("-" * 78 + "\n")
        for texto in SUPUESTOS_REPORTE:
            archivo.write(f"  {texto}\n")
        archivo.write("\n")

        if errores:
            archivo.write("-" * 78 + "\nERRORES POR REGLA\n")
            archivo.write("-" * 78 + "\n")
            por_codigo = defaultdict(list)
            for h in errores:
                por_codigo[(h["codigo"], h["variable"])].append(h)
            for (codigo, variable), lista in sorted(
                por_codigo.items(), key=lambda kv: -len(kv[1])
            ):
                archivo.write(f"  [{codigo:<8}] V{variable:<10} {len(lista):>5} registros\n")
                archivo.write(f"             ej.: {lista[0]['detalle'][:110]}\n")
                archivo.write(f"             filas: {', '.join(str(h['fila']) for h in lista[:12])}")
                if len(lista) > 12:
                    archivo.write(" ...")
                archivo.write("\n")
            archivo.write("\n")

        if advertencias:
            archivo.write("-" * 78 + "\nADVERTENCIAS POR REGLA\n")
            archivo.write("-" * 78 + "\n")
            por_codigo = defaultdict(list)
            for h in advertencias:
                por_codigo[(h["codigo"], h["variable"])].append(h)
            for (codigo, variable), lista in sorted(
                por_codigo.items(), key=lambda kv: -len(kv[1])
            ):
                archivo.write(f"  [{codigo:<8}] V{variable:<10} {len(lista):>5} registros\n")
                archivo.write(f"             ej.: {lista[0]['detalle'][:110]}\n")
            archivo.write("\n")
    return ruta


def exportar_pendientes(ruta, pendientes):
    """Graba el detalle de registros que requieren correccion manual."""
    with open(ruta, "w", encoding="utf-8-sig", newline="") as archivo:
        escritor = csv.DictWriter(
            archivo,
            fieldnames=["fila_excel", "documento", "neoplasia", "variables",
                        "codigos", "detalle"],
        )
        escritor.writeheader()
        for registro in pendientes:
            escritor.writerow(registro)
    return ruta


# ==============================================================================
# 8. PROCESO PRINCIPAL
# ==============================================================================

def main():
    print("=" * 70)
    print(" GENERADOR ARCHIVO PLANO CAC - SISCAC 4.0")
    print(f" EAPB {CODIGO_EAPB} | corte {FECHA_CORTE} | reporte {FECHA_REPORTE}")
    print("=" * 70)

    ruta_fuente = os.path.join(RUTA_BASE, ARCHIVO_FUENTE)
    ruta_reglas = os.path.join(RUTA_BASE, ARCHIVO_REGLAS)
    ruta_atc = os.path.join(RUTA_BASE, ARCHIVO_ATC)

    if not os.path.exists(ruta_fuente):
        print(f" Error: no se encuentra '{ARCHIVO_FUENTE}' en {RUTA_BASE}")
        return 1

    # 1. Cargar reglas de validacion y archivo operativo de ATC
    reglas = cargar_reglas(ruta_reglas) if os.path.exists(ruta_reglas) else []
    codigos_atc = cargar_codigos_atc(ruta_atc)
    print(f"-> Reglas de validacion cargadas: {len(reglas)}")
    print(f"-> Codigos ATC del archivo operativo: {len(codigos_atc)}")

    # 2. Cargar la base de datos fuente
    encabezados, filas, indice = cargar_fuente(ruta_fuente)
    print(f"-> Registros cargados desde {ARCHIVO_FUENTE}: {len(filas)}")
    print(f"-> Variables mapeadas: {len(indice)}")

    # 3. Limpieza de texto y normalizacion de fechas
    limpiar_texto(filas, indice)
    columnas_fecha = normalizar_fechas(filas, indice, encabezados)
    print(f"-> Columnas de fecha normalizadas: {len(columnas_fecha)}")

    # 4. Correcciones deterministas derivadas del instructivo
    correcciones, motivos = corregir_dependencias(filas, indice)
    print(f"-> Correcciones automaticas aplicadas: {sum(correcciones.values())}")

    # 5. Relleno de vacios (ninguna variable puede llegar vacia)
    rellenados = rellenar_vacios(filas, indice, encabezados)
    if rellenados:
        print(f"-> Campos vacios rellenados con {NO_APLICA}: {sum(rellenados.values())}")

    # 6. Validacion
    hallazgos = validar(filas, indice, encabezados, reglas, codigos_atc)
    errores = [h for h in hallazgos if h["nivel"] == "ERROR"]
    advertencias = [h for h in hallazgos if h["nivel"] == "ADVERTENCIA"]
    print(f"-> Errores: {len(errores)} | Advertencias: {len(advertencias)}")

    # 7. Exportaciones
    ruta_txt = os.path.join(RUTA_BASE, NOMBRE_TXT)
    exportar_txt(filas, encabezados, ruta_txt)
    print(f"-> Archivo plano generado: {NOMBRE_TXT}")

    exportar_reporte(
        os.path.join(RUTA_BASE, NOMBRE_REPORTE),
        hallazgos, correcciones, reglas, filas, encabezados, indice, motivos,
    )
    pendientes = construir_pendientes(filas, indice, encabezados, hallazgos)
    exportar_pendientes(os.path.join(RUTA_BASE, NOMBRE_PENDIENTES), pendientes)
    print(f"-> Reporte de validacion: {NOMBRE_REPORTE}")
    print(f"-> Registros con pendientes: {NOMBRE_PENDIENTES} ({len(pendientes)})")

    # 8. Resumen por regla de error mas frecuente
    if errores:
        print("-" * 70)
        print(" ERRORES MAS FRECUENTES")
        por_codigo = Counter((h["codigo"], h["variable"]) for h in errores)
        for (codigo, variable), n in por_codigo.most_common(12):
            print(f"   [{codigo:<8}] V{variable:<10} {n:>5} registros")
    print("-" * 70)
    print(" Proceso finalizado")
    return 0


if __name__ == "__main__":
    sys.exit(main())
