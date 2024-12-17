# English:

# JSON File Indexer and Text Extractor

This script recursively scans a directory for specific file types (`.pdf`, `.docx`, `.xls`, `.xlsx`), extracts text content from these files, and converts the results into a JSON file.

## Features

- Converts `.xls` files to `.xlsx` format for compatibility.
- Extracts text from the following file types:
  - PDF (`.pdf`)
  - Word documents (`.docx`)
  - Excel spreadsheets (`.xls`, `.xlsx`)
- Outputs a JSON file with the extracted content, file paths, and filenames.

## Requirements

Before running the script, ensure you have the following dependencies installed:

- Python 3.8+
- `PyPDF2` for PDF text extraction
- `python-docx` for Word document text extraction
- `openpyxl` for Excel file handling
- `xlrd` for reading `.xls` files

You can install these dependencies using the following command:

```bash
pip install PyPDF2 python-docx openpyxl xlrd
```

## File Structure

- Input Directory: `/Users/luisalbertocerelli/Desktop/00-Todo/14_Hackaton_Real/Assets_empresa`
- Output File: `/Users/luisalbertocerelli/Desktop/00-Todo/14_Hackaton_Real/PRINCIPAL/Json_conversion/documentos_indexados.json`

## How to Use

1. **Setup Input Directory**:
   - Place all the files you want to process in the directory specified as `root_dir`.

2. **Run the Script**:
   - Execute the script using Python:
     ```bash
     python script.py
     ```

3. **Output JSON**:
   - After running the script, a JSON file containing the extracted content will be saved in the location specified by `output_json`.

## Code Breakdown

### Convert `.xls` to `.xlsx`

The `convertir_xls_a_xlsx` function converts `.xls` files to `.xlsx` using `xlrd` and `openpyxl`:

```python
def convertir_xls_a_xlsx(ruta_xls, ruta_xlsx):
    libro_xls = xlrd.open_workbook(ruta_xls)
    libro_xlsx = Workbook()
    hoja_xlsx = libro_xlsx.active

    hoja_xls = libro_xls.sheet_by_index(0)
    for row in range(hoja_xls.nrows):
        for col in range(hoja_xls.ncols):
            hoja_xlsx.cell(row=row+1, col=col+1).value = hoja_xls.cell_value(row, col)

    libro_xlsx.save(ruta_xlsx)
```

### Extract Text from Files

The `extraer_texto_archivo` function processes files based on their type:

- PDF: Extracts text from all pages using `PyPDF2`.
- DOCX: Extracts text from paragraphs using `python-docx`.
- XLSX/XLS: Extracts cell values using `openpyxl`.

```python
def extraer_texto_archivo(file_path, file_type):
    if file_type == "pdf":
        reader = PdfReader(file_path)
        text = " ".join([page.extract_text() for page in reader.pages])
    elif file_type == "docx":
        doc = Document(file_path)
        text = "\n".join([p.text for p in doc.paragraphs])
    elif file_type in ["xlsx", "xls"]:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        text = ""
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows(values_only=True):
                text += " ".join([str(cell) for cell in row if cell is not None]) + "\n"
        wb.close()
    return text
```

### Recursively Process Files

The `convertir_a_json_recursivamente` function traverses the directory and processes supported file types:

```python
def convertir_a_json_recursivamente(directorio):
    documentos = []
    for root, _, files in os.walk(directorio):
        for filename in files:
            if filename.endswith((".pdf", ".docx", ".xlsx", ".xls")):
                file_path = os.path.join(root, filename)
                text = extraer_texto_archivo(file_path, file_type)
                if text:
                    documentos.append({
                        "id": filename,
                        "path": file_path,
                        "content": text
                    })

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(documentos, f, ensure_ascii=False, indent=4)
```

### Run the Script

Call the main function to process files:

```python
convertir_a_json_recursivamente(root_dir)
```

## Notes

- Ensure file paths are correctly specified for your environment.
- Unsupported file types are skipped with a warning message.
- The script handles exceptions and logs errors for problematic files.

## Output Example

The output JSON structure:

```json
[
    {
        "id": "example.pdf",
        "path": "/path/to/example.pdf",
        "content": "Extracted text content..."
    },
    {
        "id": "example.xlsx",
        "path": "/path/to/example.xlsx",
        "content": "Extracted text content..."
    }
] 
```



# Español: 

# Indexador y Extractor de Texto a JSON

Este script analiza recursivamente un directorio en busca de archivos específicos (`.pdf`, `.docx`, `.xls`, `.xlsx`), extrae el contenido de texto de dichos archivos y lo convierte en un archivo JSON.

## Características

- Convierte archivos `.xls` a formato `.xlsx` para garantizar la compatibilidad.
- Extrae texto de los siguientes tipos de archivo:
  - PDF (`.pdf`)
  - Documentos Word (`.docx`)
  - Hojas de cálculo Excel (`.xls`, `.xlsx`)
- Genera un archivo JSON con el contenido extraído, las rutas de los archivos y los nombres de archivo.

## Requisitos

Antes de ejecutar el script, asegúrate de tener instaladas las siguientes dependencias:

- Python 3.8+
- `PyPDF2` para la extracción de texto de archivos PDF
- `python-docx` para la extracción de texto de documentos Word
- `openpyxl` para el manejo de archivos Excel
- `xlrd` para leer archivos `.xls`

Puedes instalar estas dependencias con el siguiente comando:

```bash
pip install PyPDF2 python-docx openpyxl xlrd
```

## Estructura de Archivos

- Directorio de Entrada: `/Users/luisalbertocerelli/Desktop/00-Todo/14_Hackaton_Real/Assets_empresa`
- Archivo de Salida: `/Users/luisalbertocerelli/Desktop/00-Todo/14_Hackaton_Real/PRINCIPAL/Json_conversion/documentos_indexados.json`

## Cómo Usar

1. **Configura el Directorio de Entrada**:
   - Coloca todos los archivos que deseas procesar en el directorio especificado como `root_dir`.

2. **Ejecuta el Script**:
   - Ejecuta el script con Python:
     ```bash
     python script.py
     ```

3. **Archivo JSON Generado**:
   - Tras ejecutar el script, se guardará un archivo JSON con el contenido extraído en la ubicación especificada como `output_json`.

## Desglose del Código

### Convertir `.xls` a `.xlsx`

La función `convertir_xls_a_xlsx` convierte archivos `.xls` a `.xlsx` usando `xlrd` y `openpyxl`:

```python
def convertir_xls_a_xlsx(ruta_xls, ruta_xlsx):
    libro_xls = xlrd.open_workbook(ruta_xls)
    libro_xlsx = Workbook()
    hoja_xlsx = libro_xlsx.active

    hoja_xls = libro_xls.sheet_by_index(0)
    for row in range(hoja_xls.nrows):
        for col in range(hoja_xls.ncols):
            hoja_xlsx.cell(row=row+1, col=col+1).value = hoja_xls.cell_value(row, col)

    libro_xlsx.save(ruta_xlsx)
```

### Extraer Texto de Archivos

La función `extraer_texto_archivo` procesa los archivos según su tipo:

- PDF: Extrae texto de todas las páginas usando `PyPDF2`.
- DOCX: Extrae texto de los párrafos usando `python-docx`.
- XLSX/XLS: Extrae valores de las celdas usando `openpyxl`.

```python
def extraer_texto_archivo(file_path, file_type):
    if file_type == "pdf":
        reader = PdfReader(file_path)
        text = " ".join([page.extract_text() for page in reader.pages])
    elif file_type == "docx":
        doc = Document(file_path)
        text = "\n".join([p.text for p in doc.paragraphs])
    elif file_type in ["xlsx", "xls"]:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        text = ""
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows(values_only=True):
                text += " ".join([str(cell) for cell in row if cell is not None]) + "\n"
        wb.close()
    return text
```

### Procesamiento Recursivo de Archivos

La función `convertir_a_json_recursivamente` recorre el directorio y procesa los tipos de archivo soportados:

```python
def convertir_a_json_recursivamente(directorio):
    documentos = []
    for root, _, files in os.walk(directorio):
        for filename in files:
            if filename.endswith((".pdf", ".docx", ".xlsx", ".xls")):
                file_path = os.path.join(root, filename)
                text = extraer_texto_archivo(file_path, file_type)
                if text:
                    documentos.append({
                        "id": filename,
                        "path": file_path,
                        "content": text
                    })

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(documentos, f, ensure_ascii=False, indent=4)
```

### Ejecutar el Script

Llama a la función principal para procesar los archivos:

```python
convertir_a_json_recursivamente(root_dir)
```

## Notas

- Asegúrate de especificar correctamente las rutas de los archivos según tu entorno.
- Los tipos de archivo no soportados se omiten con un mensaje de advertencia.
- El script maneja excepciones y registra errores para archivos problemáticos.

## Ejemplo de Salida

La estructura del archivo JSON generado:

```json
[
    {
        "id": "example.pdf",
        "path": "/path/to/example.pdf",
        "content": "Contenido de texto extraído..."
    },
    {
        "id": "example.xlsx",
        "path": "/path/to/example.xlsx",
        "content": "Contenido de texto extraído..."
    }
]
```
