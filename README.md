# Análisis de Clusters y Gerentes en Excel

## Descripción

Este script te permite analizar un archivo Excel que contiene información sobre **clusters** y **gerentes**. El análisis te ayudará a:
1. Encontrar filas con datos faltantes en las columnas de **Cluster** y **Gerente**.
2. Ver si algún **Cluster** ha sido asignado a más de un gerente.

Es una herramienta útil si trabajas con datos en Excel y necesitas verificar la organización de tus gerentes y clusters de manera rápida.

---

## ¿Cómo usar este proyecto?

### Paso 1: Descargar el proyecto

1. En la parte superior de esta página, haz clic en el botón verde que dice **Code**.
2. Luego selecciona **Download ZIP** para descargar todos los archivos en tu computadora.
3. Extrae el archivo `.zip` en tu computadora (haz clic derecho y selecciona "Extraer aquí" o "Extract").

### Paso 2: Instalar Python

1. Asegúrate de tener **Python** instalado en tu computadora.
2. Si no lo tienes, puedes descargar Python desde aquí: [https://www.python.org/downloads/](https://www.python.org/downloads/)
   - Durante la instalación, asegúrate de marcar la opción **Add Python to PATH**.

### Paso 3: Instalar los paquetes necesarios

1. Abre la carpeta donde extrajiste el archivo `.zip`.
2. Haz clic derecho dentro de la carpeta mientras mantienes la tecla **Shift** presionada y selecciona **Abrir ventana de PowerShell aquí** o **Abrir terminal** (en Mac o Linux).
3. Escribe el siguiente comando para instalar las herramientas necesarias:

    ```bash
    pip install -r requirements.txt
    ```

4. Espera a que se instalen los paquetes.

### Paso 4: Ejecutar el script

1. Abre el archivo llamado `analisis_clusters.py` en cualquier editor de texto (por ejemplo, **Notepad** o **VSCode**).
2. Cambia las siguientes tres líneas para que apunten a tu archivo Excel:
   
    ```python
    ruta_archivo = 'ruta_a_tu_archivo.xlsx'  # Cambia esto con la ruta a tu archivo Excel
    hoja = 'nombre_de_la_hoja'  # Cambia esto por el nombre de la hoja en tu archivo Excel
    fila_encabezado = 1  # Número de fila donde están los encabezados (1-indexado)
    ```

3. Guarda los cambios.
4. Regresa a la ventana de **PowerShell** o **Terminal**.
5. Escribe el siguiente comando para ejecutar el script:

    ```bash
    python ClusterManagerAnalyzer.py
    ```

### Paso 5: Ver los resultados

1. El script analizará tu archivo Excel y mostrará:
   - Las filas con celdas vacías en las columnas de **Gerente** y **Cluster**.
   - Los **clusters** que tienen más de un gerente asignado, organizados en filas separadas.

---

## ¿Qué necesitas para usar este proyecto?

- **Python**: Asegúrate de tener Python instalado. Puedes descargarlo desde: [https://www.python.org/downloads/](https://www.python.org/downloads/).
- **Un archivo Excel**: El archivo Excel debe tener las columnas de **Gerente** y **Cluster** en una de las hojas.

---

## Preguntas frecuentes

### 1. ¿Qué pasa si no tengo Python instalado?
No te preocupes, puedes descargar Python fácilmente desde [este enlace](https://www.python.org/downloads/). Asegúrate de seguir las instrucciones en el **Paso 2**.

### 2. No sé cómo encontrar la ruta de mi archivo Excel. ¿Qué hago?
Puedes copiar la ruta completa del archivo Excel haciendo lo siguiente:
- **Windows**: Haz clic derecho en el archivo, selecciona **Propiedades**, y copia la "Ubicación". Luego, añade el nombre del archivo al final de la ruta, por ejemplo: `C:/Usuarios/TuNombre/Documentos/mi_archivo.xlsx`.
- **Mac**: Haz clic derecho en el archivo, selecciona **Obtener información**, y copia la "Ruta".

### 3. Mi archivo Excel tiene varias hojas. ¿Cuál es el nombre de la hoja?
Puedes ver el nombre de cada hoja en la parte inferior de tu archivo Excel. Usa ese nombre en el script.

---

## Autor

- Creado por Luis Alberto Rodriguez Solis.
- Contacto: luisa.rodriguezs@ext.cemex.com
