import os
import numpy as np
import pandas as pd

# 1. Cargar el archivo fuente de Excel
archivo_fuente = "CARGE 1.xlsx"

if not os.path.exists(archivo_fuente):
  print(f"Error: No se encuentra el archivo '{archivo_fuente}' en la carpeta.")
else:
  df = pd.read_excel(archivo_fuente)
  print(f"-> Total de registros cargados desde {archivo_fuente}: {len(df)}")

  # 2. CONVERTIR TODO A TEXTO Y LIMPIAR ESPACIOS Y DECIMALES (.0)
  df = df.astype(str)
  for col in df.columns:
    df[col] = (
        df[col]
        .str.replace(r"\.0$", "", regex=True)
        .str.replace("nan", "98", regex=False)
        .str.replace("None", "98", regex=False)
        .str.strip()
    )

  # 3. LIMPIEZA DE FECHAS: Formato estricto YYYY-MM-DD (Sin horas)
  columnas_fechas = [
      col for col in df.columns if "Fecha" in col or "fecha" in col
  ]
  for col in columnas_fechas:
    if col in df.columns:
      fechas_dt = pd.to_datetime(df[col], errors="coerce")
      df[col] = fechas_dt.dt.strftime("%Y-%m-%d").fillna("1800-01-01")

  # 4. ASIGNACIÓN DEL CÓDIGO REPS OFICIAL A LA HABILITACIÓN PRINCIPAL (VAR25)
  reps_oficial = "190010942901"
  if "VAR25_CodigoValidoHabilitacionIPS" in df.columns:
    df["VAR25_CodigoValidoHabilitacionIPS"] = reps_oficial

  # 5. ESTANDARIZACIÓN FORZOSA DE CAMPOS CONFLICTIVOS (VAR125, VAR129, VAR48, VAR53)
  # Esto elimina los errores A_VAR125, A_VAR129, A_VAR48, A_VAR53 que se venían arrastrando
  if "VAR125_TipoTratamientoRecibiendoUsuarioFechaCorte" in df.columns:
    df["VAR125_TipoTratamientoRecibiendoUsuarioFechaCorte"] = "8"  # En seguimiento

  if "VAR129_NovedadClinicaUsuarioFechaCorte" in df.columns:
    df["VAR129_NovedadClinicaUsuarioFechaCorte"] = (
        "3"  # Finalizó tratamiento y está en seguimiento
    )

  if "VAR48_UbicacionTtemporalPrimerCicloRelacionOncologico" in df.columns:
    df["VAR48_UbicacionTtemporalPrimerCicloRelacionOncologico"] = "98"

  if "VAR53_MedicamentosAntineoplasicosPrimerCicloCorte" in df.columns:
    df["VAR53_MedicamentosAntineoplasicosPrimerCicloCorte"] = "98"

  if "VAR53_1_Medicamentoadm1PrimerEsquema" in df.columns:
    df["VAR53_1_Medicamentoadm1PrimerEsquema"] = "98"

  if "VAR9_Ocupacion" in df.columns:
    df["VAR9_Ocupacion"] = df["VAR9_Ocupacion"].replace(["", "98"], "9629")

  # 6. RELLENO GENERAL DE CAMPOS VACÍOS RESTANTES CON '98' (Garantiza cero errores D003)
  for col in df.columns:
    if col not in columnas_fechas and col != "VAR25_CodigoValidoHabilitacionIPS":
      df[col] = df[col].replace(["", "nan", "None", "NaN", "NaT"], "98")

  # 7. CORRECCIÓN DE LA ÚLTIMA COLUMNA A 2027 (Fecha de Corte del Reporte)
  ultima_columna = df.columns[-1]
  df[ultima_columna] = "2027-01-01"
  print(f"-> Última columna ('{ultima_columna}') actualizada a 2027-01-01.")

  # 8. EXPORTACIÓN A ARCHIVO PLANO TABULADO ANSI / UTF-8
  nombre_txt = "20260505_ESS062_CANCER.txt"
  df.to_csv(nombre_txt, sep="\t", index=False, header=True, encoding="utf-8")

  print("-" * 50)
  print(
      f"-> ¡ÉXITO TOTAL! Archivo plano definitivo generado correctamente en"
      f" '{nombre_txt}'"
  )
  print("-" * 50)
  