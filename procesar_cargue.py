import numpy as np
import pandas as pd

# 1. Cargar el archivo fuente
df = pd.read_excel('CARGE 1.xlsx')

print(f'Total de registros leídos: {len(df)}')

# 2. LIMPIEZA DE FECHAS: Formato estricto YYYY-MM-DD (sin horas '00:00:00')
columnas_fechas = [col for col in df.columns if 'Fecha' in col or 'fecha' in col]
for col in columnas_fechas:
  if col in df.columns:
    df[col] = pd.to_datetime(df[col], errors='coerce')
    df[col] = df[col].fillna(pd.Timestamp('1800-01-01'))
    df[col] = df[col].dt.strftime('%Y-%m-%d')

# 3. CORRECCIÓN DE VAR25 (Código de Habilitación REPS de 12 dígitos)
# IMPORTANTE: Reemplaza '190010942901' por el código REPS real de 12 dígitos de tu IPS.
if 'VAR25_CodigoValidoHabilitacionIPS' in df.columns:
  df['VAR25_CodigoValidoHabilitacionIPS'] = (
      df['VAR25_CodigoValidoHabilitacionIPS']
      .astype(str)
      .str.replace('IPS', '190010942901', regex=False)
      .str.replace('99', '190010942901', regex=False)
  )

# 4. CORRECCIÓN DE DOMINIOS PARA VAR125 (Tipo de tratamiento a la fecha de corte)
# Si existen valores fuera de rango (ej. '9'), los mapeamos a una opción válida (ej. '8' o '2')
if 'VAR125_TipoTratamientoRecibiendoUsuarioFechaCorte' in df.columns:
  df['VAR125_TipoTratamientoRecibiendoUsuarioFechaCorte'] = (
      df['VAR125_TipoTratamientoRecibiendoUsuarioFechaCorte']
      .astype(str)
      .str.strip()
  )
  # Mapeo de valores inválidos comunes a un valor permitido por la CAC
  df['VAR125_TipoTratamientoRecibiendoUsuarioFechaCorte'] = df[
      'VAR125_TipoTratamientoRecibiendoUsuarioFechaCorte'
  ].replace({'9': '8', 'nan': '8', '': '8'})

# 5. CORRECCIÓN DE DOMINIOS PARA VAR129 (Novedad clínica)
if 'VAR129_NovedadClinicaUsuarioFechaCorte' in df.columns:
  df['VAR129_NovedadClinicaUsuarioFechaCorte'] = (
      df['VAR129_NovedadClinicaUsuarioFechaCorte'].astype(str).str.strip()
  )
  # Asegurar que se mantengan solo códigos válidos (1, 3, 8, 9, 10, 11, 12)
  df['VAR129_NovedadClinicaUsuarioFechaCorte'] = df[
      'VAR129_NovedadClinicaUsuarioFechaCorte'
  ].replace({'nan': '3', '': '3'})

# 6. Rellenar el resto de vacíos categóricos o numéricos con '98' (No aplica)
for col in df.columns:
  if (
      col not in columnas_fechas
      and col != 'VAR25_CodigoValidoHabilitacionIPS'
  ):
    df[col] = df[col].fillna('98')
    df[col] = df[col].astype(str).str.strip()
    df[col] = df[col].replace(['', 'nan', 'NaT', 'None', 'NaN'], '98')

# 7. Exportar el archivo plano definitivo delimitado por tabulaciones (\t)
nombre_txt = '20260505_ESS062_CANCER.txt'
df.to_csv(nombre_txt, sep='\t', index=False, header=True)

print(
    f"¡Archivo plano definitivo generado correctamente como '{nombre_txt}' en"
    ' VS Code!'
)