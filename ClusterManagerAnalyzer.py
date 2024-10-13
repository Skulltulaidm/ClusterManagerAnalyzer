# --- propósito ---
# Autor: Luis Alberto Rodríguez Solis
# Fecha: 12/10/24
# Versión: 1.0
# Descripción: Este script analiza un archivo Excel con información de clusters y gerentes. 
# Identifica filas con celdas vacías y clusters asignados a múltiples gerentes.
# -------------------------------------------

import pandas as pd
import unicodedata
from collections import defaultdict

# --- Configuración del usuario ---
ruta_archivo = 'ruta_a_tu_archivo.xlsx'  # Ruta Completa del archivo Excel
hoja = 'nombre_de_la_hoja'  # Nombre de la hoja en el archivo Excel
fila_encabezado = 1  # Número de fila donde están los encabezados (1-indexado)
# ---------------------------------


def normalizar_texto(texto):
    """
    Normaliza el texto eliminando acentos, espacios extra y haciendo que sea insensible a mayúsculas/minúsculas.

    :param texto: Cadena de texto a normalizar
    :return: Texto normalizado
    """
    if isinstance(texto, str):
        # Eliminar acentos
        texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
        # Eliminar espacios extras y convertir a minúsculas
        texto = texto.strip().lower()
        # Reemplazar múltiples espacios por uno solo
        texto = ' '.join(texto.split())
    return texto


def leer_archivo_excel(ruta_archivo, hoja, fila_encabezado):
    """
    Lee el archivo Excel y selecciona las columnas 'Gerente' y 'Cluster'.

    :param ruta_archivo: Ruta del archivo Excel
    :param hoja: Nombre de la hoja en el archivo Excel
    :param fila_encabezado: Fila donde se encuentran los encabezados
    :return: DataFrame con las columnas 'Gerente' y 'Cluster'
    """
    try:
        df = pd.read_excel(ruta_archivo, sheet_name=hoja, header=fila_encabezado-1)
        df.columns = [normalizar_texto(col) for col in df.columns]
        if 'gerente' not in df.columns or 'cluster' not in df.columns:
            raise ValueError("No se encontraron las columnas 'Gerente' o 'Cluster' en el archivo.")
        return df[['gerente', 'cluster']]
    
    except Exception as e:
        print(f"Ocurrió un error: {e}")
        return None


def analizar_datos(df):
    """
    Analiza los datos del DataFrame para identificar celdas vacías y clusters con gerentes duplicados.

    :param df: DataFrame con las columnas 'Gerente' y 'Cluster'
    :return: Tupla con dos DataFrames:
             - vacios: Filas con celdas vacías en 'Gerente' o 'Cluster'
             - clusters_duplicados: Clusters asignados a múltiples gerentes
    """
    try:
        df['gerente'] = df['gerente'].apply(normalizar_texto)
        df['cluster'] = df['cluster'].apply(normalizar_texto)
        df['cluster'] = pd.to_numeric(df['cluster'], errors='coerce').fillna(0).astype(int)
        
        # Identificar celdas vacías
        vacios = df[df['gerente'].isna() | (df['cluster'] == 0)]
        
        # Identificar clusters asignados a múltiples gerentes
        cluster_gerentes = defaultdict(set)
        for _, row in df.iterrows():
            if pd.notna(row['cluster']) and row['cluster'] != 0:
                cluster = str(row['cluster'])
                gerente = str(row['gerente'])
                if gerente != 'nan':
                    cluster_gerentes[cluster].add(gerente)
        
        # Retornar clusters con múltiples gerentes
        clusters_duplicados = {cluster: gerentes for cluster, gerentes in cluster_gerentes.items() if len(gerentes) > 1}
        
        return vacios, clusters_duplicados
    
    except Exception as e:
        print(f"Ocurrió un error al analizar los datos: {e}")
        return None, None


# --- Ejecución del análisis ---
df = leer_archivo_excel(ruta_archivo, hoja, fila_encabezado)
if df is not None:
    vacios, clusters_duplicados = analizar_datos(df)
    
    if vacios is not None and clusters_duplicados is not None:
        # Mostrar los resultados con pandas
        
        # Mostrar filas con celdas vacías
        if not vacios.empty:
            display(vacios)
        
        # Mostrar clusters duplicados (una fila por gerente)
        if clusters_duplicados:
            rows = []
            for cluster, gerentes in clusters_duplicados.items():
                for gerente in gerentes:
                    rows.append((cluster, gerente))
            df_clusters_duplicados = pd.DataFrame(rows, columns=['Cluster', 'Gerente'])
            display(df_clusters_duplicados)
