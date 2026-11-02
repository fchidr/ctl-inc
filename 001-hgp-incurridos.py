#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
  @Name: 01-hgp-incurridos.py
  @Author: fcarbajo
  @Site: 
  @Descripción: Script de extracción de incurridos por actividad utilizando la API usando el formato CSV para la descarga.
  @Usage:
    python 01-hgp-incurridos.py {AAAA-MM}
    Para la ejecución del script es necesario ejecutar 'redmine:hgp-conf'
"""

__author__ = 'Fernando J. Carbajo'
__version__ = '1.0.9'
__date__ = '2026-08-29'

import os
import sys
from pathlib import Path

script_path = Path(os.path.dirname(os.path.abspath(__file__)))
script_base = script_path.parent
script_libs = f"{script_path}\\libs"
# La librería común está en el directorio Scripts de la raiz del proyecto
common_libs = f"{script_base}\\.pylibs"
path_in = script_path / "data"
path_out = script_path / "data"
PYTHON_PROFILE = os.environ['PYTHON_PROFILE']
PY_NAME = (os.environ['PY_NAME'])

py_full_version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
py_version = sys.version_info.major

sys.path.insert(0, script_libs)
sys.path.insert(1, common_libs)

import helper as hlp
import helper_pd as hpd
import coloriza as cl

import requests
import pandas as pd
import warnings
import csv
import re
import shutil
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter
#import json
#import yaml


def download_file(enlace, fichero='fichero', params=None, cookies=None):

  requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecurePlatformWarning)
  requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)

  cl.cout_pair("  Fichero: ", fichero)
  if _debug:
    cl.print_dbg("    URL:", enlace)

  if cookies:
    for cookie in cookies:
      s.cookies.set(cookie['name'], cookie['value'])
    s = requests.Session()
    respuesta = s.get(enlace, params, headers=headers, verify=False).content
  else:
    respuesta = requests.get(enlace, params, headers=headers, verify=False).content

  if respuesta:
    with open(fichero, 'wb') as f:
      f.write(respuesta)
      f.close()

  print("  {}".format(cl.ctext("-- Descarga completa --", cl.Fore.CYAN)))

  return fichero

def get_issue_time_entries_lm_project(proj_id, tipo='csv'):

  print(cl.ctext("\n  Obteniendo incurridos."))
  cl.cout_pair("  Proyecto: ", proj_id)

  # Parámetros último mes
  params = {
    "utf8": "✓",
    "set_filter": "1",
    "sort": "spent_on:desc",
    "f[]": ["spent_on", "issue.tracker_id", ""],
    "op[spent_on]": "lm",
    "op[issue.tracker_id]": "!",
    "v[issue.tracker_id][]": "21",
    "c[]": ["issue.id", "issue.category", "issue", "hours", "user", "spent_on"],
    "t[]": "hours"
  }

  if tipo == "xml":
    url_params = f"limit={query_response_limit}"

  query_url = f"{redmine_url}/projects/{proj_id}/time_entries.{tipo}"
  nfichero = f"{proj_id}-issues-time_entries-lm"

  if _debug:
    cl.print_dbg("URL query: ",  query_url)

  if fecha == None:
    nfichero = hlp.gen_file_name_timestamp(nfichero)
  else:
    nfichero = "%s-%s" % (nfichero, fecha)

  nfichero = path_out / f"{nfichero}.{tipo}"
  fichero = download_file(query_url, nfichero, params)

  return fichero

def copy_excel_from_template(archivo_origen, nombre_destino, carpeta_destino="."):
  """
  Copia una plantilla Excel y evita sobrescribir archivos existentes.

  Parámetros:
    archivo_origen (str): Ruta de la plantilla Excel.
    nombre_destino (str): Nombre base del nuevo archivo (sin extensión opcional).
    carpeta_destino (str): Carpeta donde se guardará la copia.

  Retorna:
    str: Ruta final del archivo creado.
  """

  origen = Path(archivo_origen)
  destino_dir = Path(carpeta_destino)

  # Crear la carpeta destino si no existe
  destino_dir.mkdir(parents=True, exist_ok=True)

  # Asegurar extensión .xlsx
  if not nombre_destino.lower().endswith(".xlsx"):
    nombre_destino += ".xlsx"

  destino = destino_dir / nombre_destino

  # Si existe, agregar numeración
  contador = 1
  while destino.exists():
    destino = destino_dir / f"{Path(nombre_destino).stem} ({contador}){Path(nombre_destino).suffix}"
    contador += 1

  # Copiar archivo
  shutil.copy2(origen, destino)

  return str(destino)

def guardar_csv(dataframe, ruta_archivo, delimitador=";", decimal=",", incluir_indice=False, codificacion="utf-8"):
  """
  Guarda un DataFrame en un archivo CSV.

  :param dataframe: DataFrame de pandas a guardar.
  :param ruta_archivo: Ruta completa del archivo CSV de salida.
  :param separador: Separador de columnas (por defecto ',').
  :param incluir_indice: Si True, incluye el índice del DataFrame.
  :param codificacion: Codificación del archivo (por defecto 'utf-8').
  """
  try:
    # Validar que el argumento es un DataFrame
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError("El objeto proporcionado no es un DataFrame de pandas.")

    base, ext = os.path.splitext(ruta_archivo)
    #if ext.lower() != ".csv":
    p = Path(ruta_archivo)
    if p.stem.lower() != "csv":
      ruta_archivo = f"{base}.csv"

    # Crear carpeta si no existe
    carpeta = os.path.dirname(ruta_archivo)
    if carpeta and not os.path.exists(carpeta):
      os.makedirs(carpeta)

    # Guardar el CSV
    dataframe.to_csv(
        ruta_archivo,
        sep=delimitador,
        decimal=decimal,
        quotechar='"',
        quoting=1,
        index=incluir_indice,
        encoding=codificacion
    )
    #print(f"Archivo CSV guardado en: {ruta_archivo}")
    return ruta_archivo

  except Exception as e:
    print(f"Error al guardar el CSV: {e}")
    return None


def leer_csv(ruta_csv, codificacion = 'utf-8', delimiter = '\t', tiene_encabezados = True):
  """
  Lee el CSV y devuelve:
    encabezados_csv: lista de nombres de columnas
    filas_csv: lista de filas, cada una como lista de valores
  """

  with ruta_csv.open(mode = "r", encoding = codificacion, newline = "") as archivo:
    lector = csv.reader(archivo, delimiter = delimiter)
    filas = list(lector)

  if not filas:
    raise ValueError("El archivo CSV está vacío.")

  if tiene_encabezados:
    encabezados = [celda.strip() for celda in filas[0]]
    datos = filas[1:]
  else:
    encabezados = []
    datos = filas

  # Elimina filas completamente vacías
  datos = [fila for fila in datos if any(str(valor).strip() for valor in fila)]

  return encabezados, datos

def get_struct_excel_table(ws, tabla):
  """
  Obtiene los encabezados de la tabla Excel a partir de su primera fila.
  """

  min_col, min_row, max_col, max_row = range_boundaries(tabla.ref)

  encabezados = []

  for columna in range(min_col, max_col + 1):
    valor = ws.cell(row=min_row, column=columna).value
    encabezados.append("" if valor is None else str(valor).strip())

  return encabezados, min_row, min_col, max_col, max_row

def copy_format_row(ws, fila_origen, fila_destino, min_col, max_col):
  """
  Copia el formato de una fila existente hacia una nueva fila.
  """

  for columna in range(min_col, max_col + 1):
    origen = ws.cell(row=fila_origen, column=columna)
    destino = ws.cell(row=fila_destino, column=columna)

    if origen.has_style:
      destino._style = origen._style

    if origen.number_format:
      destino.number_format = origen.number_format

    if origen.alignment:
      destino.alignment = origen.alignment

    if origen.protection:
      destino.protection = origen.protection

def procesa_issue_time_entries(fichero, nombre_salida, tipo_salida='csv'):

  if _debug:
    cl.print_dbg("\n  Fichero de entrada: ", fichero)
    cl.print_dbg("  Fichero de salidad: ", nombre_salida)

  df = pd.read_csv(fichero, sep=";", decimal=",", encoding='latin-1')

  #print(df.info())
  """
  <class 'pandas.core.frame.DataFrame'>
  RangeIndex: 106 entries, 0 to 105
  Data columns (total 6 columns):
     #   Column     Non-Null Count  Dtype
    ---  ------     --------------  -----
     0   ID Ticket  106 non-null    int64
     1   Categoría  106 non-null    object
     2   Ticket     106 non-null    object
     3   Horas      106 non-null    float64
     4   Usuario    106 non-null    object
     5   Fecha      106 non-null    object
  dtypes: float64(1), int64(1), object(4)
  """

  df = df.rename(columns={"ID Ticket": "Petición", "Ticket": "Asunto"})
  df['Usuario'] = df['Usuario'].apply(hlp.normalize_user_name)
  df['Usuario'] = df['Usuario'].str.replace('Sanchez', 'Sánchez', regex=False)
  df['Asunto'] = df['Asunto'].str.replace(r'Tarea\s+#\d+:\s*', '', regex=True)

  #if _debug:
  #  print(df)

  #if tipo_salida == 'csv':
  if Path(nombre_salida).suffix.lower() == '.csv':
    guardar_csv(dataframe=df, ruta_archivo=nombre_salida, delimitador='\t')

  #return nombre_salida

def procesa_excel(fichero_excel, nombre_hoja, nombre_tabla, fichero_csv, proyecto, codificacion_csv = 'utf-8', delimitador_csv = '\t', con_cabecera_csv = True):
  if not Path(fichero_excel).exists():
    raise FileNotFoundError(f"No existe el archivo Excel: {fichero_excel}"
  )

  if not Path(fichero_csv).exists():
    raise FileNotFoundError(f"No existe el archivo CSV: {fichero_csv}"
  )

  if _debug:
    cl.print_dbg("  Fichero Excel: ", fichero_excel)

  cl.cout_pair("\n  Completando Excel con datos del proyecto ", proyecto.upper())
  cl.cout_pair("    Hoja: ", nombre_hoja)
  cl.cout_pair("    Tabla: ", nombre_tabla)

  # Listas de columnas
  columnas_numero = hlp.config_get(cfg, 'excel', f"{proyecto}.columnas.numero")
  columnas_decimal = hlp.config_get(cfg, 'excel', f"{proyecto}.columnas.decimal")
  columnas_fecha = hlp.config_get(cfg, 'excel', f"{proyecto}.columnas.fecha")

  #if _debug:
  #  cl.cout_pair("  Procesando CSV: ", fichero_csv)
  #  cl.print_dbg("    Columnas numero: ", [columnas_numero])
  #  cl.print_dbg("    Columnas decimal: ", [columnas_decimal])
  #  cl.print_dbg("    Columnas fecha: ", [columnas_fecha])

  # Formatos de columnas
  dia_primero = hlp.str_to_bool(hlp.config_get(cfg, 'formatos', 'dia_primero'))
  formato_numero = hlp.config_get(cfg, 'formatos', 'numero', default = '#,##0')
  formato_decimal = hlp.config_get(cfg, 'formatos', 'decimal', default = '#,##0.00')
  formato_fecha = hlp.config_get(cfg, 'formatos', 'fecha', default = 'dd/mm/yyyy')
  separador_miles = hlp.config_get(cfg, 'formatos', 'separador.miles', default = '.')
  separador_decimal = hlp.config_get(cfg, 'formatos', 'separador.decimal', default = ',')

  # Procesado del libro
  # Si el archivo es un Excel con macros '.xlsm', usamos keep_vba=True para conservar el contenido VBA.
  es_xlsm = Path(fichero_excel).suffix.lower() == ".xlsm"
  wb = load_workbook(filename = fichero_excel, keep_vba = es_xlsm, data_only = False)

  # Validar hoja
  if nombre_hoja not in wb.sheetnames:
    raise KeyError(f"No existe la hoja '{nombre_hoja}'. Hojas disponibles: {wb.sheetnames}")

  ws = wb[nombre_hoja]

  # Validar tabla
  if nombre_tabla not in ws.tables:
    tablas_disponibles = list(ws.tables.keys())
    raise KeyError(f"No existe la tabla '{nombre_tabla}' en la hoja '{nombre_hoja}'. Tablas disponibles: {tablas_disponibles}")

  tabla = ws.tables[nombre_tabla]

  # Obtener estructura de la tabla Excel
  (encabezados_excel, fila_encabezado, primera_columna, ultima_columna, ultima_fila_actual) = get_struct_excel_table(ws, tabla)
  numero_columnas_excel = len(encabezados_excel)

  encabezados_csv, datos_csv = leer_csv(fichero_csv, codificacion = codificacion_csv, delimiter = delimitador_csv)

  #if con_cabecera_csv:
  #  if len(encabezados_csv) > numero_columnas_excel:

  # Relación entre columnas CSV y columnas Excel por nombre
  posiciones_csv = {hlp.normalizar_cadena(nombre): posicion for posicion, nombre in enumerate(encabezados_csv)}

  asignacion_columnas = []
  for nombre_excel in encabezados_excel:
    clave = hlp.normalizar_cadena(nombre_excel)
    asignacion_columnas.append(posiciones_csv.get(clave))

  columnas_no_encontradas = [nombre for nombre, posicion in zip(encabezados_excel, asignacion_columnas) if posicion is None]
  if columnas_no_encontradas:
    print("  Aviso: estas columnas Excel no están en el CSV y quedarán vacías:")
    for columna in columnas_no_encontradas:
      print(f"    - {columna}")
  else:
    # Si el CSV no tiene encabezados, se copian las columnas por posición, desde la primera columna de la tabla.
    if any(len(fila) > numero_columnas_excel for fila in datos_csv):
      raise ValueError("Alguna fila del CSV tiene más columnas que la tabla Excel.")

    asignacion_columnas = list(range(numero_columnas_excel))

  # Diccionario con Nombre de columna y posición de la columna.
  mapa_columnas = {hlp.normalizar_cadena(nombre): posicion for nombre, posicion in zip(encabezados_excel, asignacion_columnas) if posicion is not None}
  #print(f"Mapa Columnas {mapa_columnas}")
  #print(f"Asignación Columnas {asignacion_columnas}")
  #print(f"Columnas CSV {posiciones_csv}")

  aux = [c.lower() for c in [columnas_numero]]
  cols_columnas_numero = [valor for clave, valor in mapa_columnas.items() if clave.lower() in aux]
  aux = [c.lower() for c in [columnas_decimal]]
  cols_columnas_decimal = [valor for clave, valor in mapa_columnas.items() if clave.lower() in aux]
  aux = [c.lower() for c in [columnas_fecha]]
  cols_columnas_fecha = [valor for clave, valor in mapa_columnas.items() if clave.lower() in aux]

  # La primera fila de datos de la tabla está después del encabezado
  primera_fila_datos = fila_encabezado + 1

  # Escribir los datos del CSV
  for indice_fila, fila_csv in enumerate(datos_csv):
    fila_excel = primera_fila_datos + indice_fila

    """
    mapa_columnas {'petición': 0, 'categoría': 1, 'asunto': 2, 'horas': 3, 'usuario': 4, 'fecha': 5}
    posiciones_csv {'petición': 0, 'categoría': 1, 'asunto': 2, 'horas': 3, 'usuario': 4, 'fecha': 5}
    asignacion_columnas [0, 1, 2, 3, 4, 5, None, None]
    """

    # Copiar cada valor del CSV a la columna correspondiente
    for indice_columna_excel, posicion_csv in enumerate(asignacion_columnas):
      columna_excel = primera_columna + indice_columna_excel

      if posicion_csv is not None and posicion_csv < len(fila_csv):
        valor = fila_csv[posicion_csv]
        formato = ''
        if posicion_csv in cols_columnas_numero:
          valor = int(valor)
          formato = formato_numero
        elif posicion_csv in cols_columnas_decimal:
          valor = float(valor.replace(',', '.'))
          formato = formato_decimal
        elif posicion_csv in cols_columnas_fecha:
          #d, m, a = valor.strip('-')
          valor = datetime.strptime(valor, '%d-%m-%Y')
          formato = formato_fecha
      else:
        # Columnas que no existen en el CSV
        valor = None

      cell = ws.cell(row = fila_excel, column=columna_excel)
      cell.value = valor
      cell.number_format = formato

  # Calcular la última fila que debe ocupar la tabla
  # La tabla conserva siempre su fila de encabezados.
  if datos_csv:
    ultima_fila_tabla = primera_fila_datos + len(datos_csv) - 1
  else:
    # Una tabla Excel debe conservar al menos una fila de datos
    ultima_fila_tabla = fila_encabezado + 1

  # Redimensionar la tabla
  nueva_referencia = (
    f"{get_column_letter(primera_columna)}{fila_encabezado}:"
    f"{get_column_letter(ultima_columna)}{ultima_fila_tabla}"
  )

  tabla.ref = nueva_referencia

  # Guardar el resultado
  wb.save(fichero_excel)

if __name__ == '__main__':
  global redmine_url
  global query_response_limit, fecha

  # Carga de configuración.
  cfg = hlp.config('hgp-incurridos.ini')
  _debug = hlp.config_dbg(cfg, section='aplicacion', option='debug')

  if py_version < 3:
    version = "%s.%s.%s" % (sys.version_info.major, sys.version_info.minor, sys.version_info.micro)
    print("Se requiere Python 3, la versión actual es %s" % (version))
    sys.exit(1)

  print(f"Python: {py_full_version}")
  clave_api_key = cfg.get("aplicacion", "env_api_key")
  if _debug:
    cl.print_dbg("  Clave API: ", clave_api_key)

  API_KEY = os.getenv(clave_api_key)

  if not API_KEY:
    print(f"\n  No se ha declarado valor para {clave_api_key}.")
    sys.exit(1)

#  headers = {
#  "X-Redmine-API-Key": API_KEY,
#  "Accept": "text/csv",
#  "Accept-Charset": "utf-8"
#  }
  headers = {
    "X-Redmine-API-Key": API_KEY, 
  }

  # Marca temporal del nombre de ficheros
  fecha = hlp.get_fecha_hm()

  redmine_url = cfg.get("aplicacion", "redmine_url_base")
  query_response_limit = (int(cfg.get("aplicacion", "query_limit")) if cfg.get("aplicacion", "query_limit").isnumeric() else 100)

  cl.cout_pair("  URL: ", redmine_url)
  cl.cout_pair("  Limit: ", query_response_limit)
  cl.cout_pair("  Marca temporal: ", fecha)

  # Proyectos a descargar
  lst_proyectos_scrum = hlp.str_to_list(cfg.get("proyectos", "scrum.lst_proyectos"))
  #cl.cout_pair("  Lista proyectos: ", lst_proyectos_scrum)
  #for entradas in lst_proyectos_scrum:
  #  nom_proyecto, proyecto = entradas.split(':')
  #  cl.cout_pair("  Nombre proyecto: ", nom_proyecto)
  #  cl.cout_pair("  Código proyecto: ", proyecto)

  # Plantilla Excel.
  plantilla_xlsx = cfg.get("excel", "excel.template")
  xlsx_salida = hlp.gen_file_name_timestamp(plantilla_xlsx)
  cl.cout_pair("  Plantilla Excel: ", plantilla_xlsx)
  cl.cout_pair("  Excel salida: ", xlsx_salida)

  #ficheros = []
  ficheros = {}
  # Copia de plantilla.
  #xlsx_salida = copy_excel_from_template(plantilla_xlsx, xlsx_salida, path_out)
  for entradas in lst_proyectos_scrum:
    nom_proyecto, proyecto = entradas.split(':')
    fichero = get_issue_time_entries_lm_project(proyecto)
    if Path(fichero).exists():
      #ficheros.append(fichero)
      ficheros[nom_proyecto] = fichero

  # Copia de plantilla.
  xlsx_salida = copy_excel_from_template(plantilla_xlsx, xlsx_salida, path_out)
  #ficheros = {'retech': f'{path_in}\\sds-hgp-01936-8-issues-time_entries-lm-20260828-2031.csv', 'pae': f'{path_in}\\sds-hgp-01960-1-issues-time_entries-lm-20260828-2031.csv'}
  for nom_proyecto, fichero in ficheros.items():
    fsalida = path_out / f"{Path(fichero).stem}-normalizado{Path(fichero).suffix}"
    procesa_issue_time_entries(fichero, fsalida)
    hoja_incurridos = cfg.get("excel", f"{nom_proyecto}.incurridos")
    tabla_incurridos = cfg.get("excel", f"{nom_proyecto}.tabla")
    procesa_excel(xlsx_salida, hoja_incurridos, tabla_incurridos, fsalida, nom_proyecto)


"""
  Hojas:
    0: Resumen
      tbl_RETECH_ResumenGeneral
      tbl_RETECH_Portal_Fases
      tbl_RETECH_PIM_Fases
      tbl_RETECH_HGN_Fases
      tbl_RETECH_Actividades
      tbl_PAE_ResumenGeneral
      tbl_PAE_Fases
      tbl_PAE_Actividades
    1: Balance Proyecto
      tbl_Balance_RETECH
      tbl_Balance_PAE
    2: Balance Perfil
      tblBalanceRETECH
      tblPerfilRETECH
      tblBalancePAE
      tblPerfilPAE
    3: R. Usuarios Proyectos
      tblRUsuarioProyecto
      tblProyectosSumatorios
    4: Inc. RETECH
      tblIncurridosRETECH
    5: Inc. PAE
      tblIncurridosPAE
    6: R. RETECH
      tblReuniones_RETECH
      tblGestion_RETECH
      tblDocumentacion_RETECH
      tblDesarrollo_RETECH
    7: R. PAE
      tblReuniones_PAE
      tblGestion_PAE
      tblDocumentacion_PAE
      tblDesarrollo_PAE

  Tabla incurridos RETECH: tblIncurridosRETECH
  Tabla incurridos PAE: tblIncurridosPAE

"""

"""
  Funciones:
    tablas: tblIncurridosRETECH, tblIncurridosPAE
    =SI(ESNUMERO(COINCIDIR([@Petición];$M$2:$M$428;0));"Sí";"No")
"""