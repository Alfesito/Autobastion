# Autobastion

Herramienta para automatizar el proceso de copia y pega de los controles de las guías de bastionado del CIS desde documentos PDF y Word a un archivo Excel. El uso de PDF y Word es debido a que cada uno de estos archivos permite optener el texto con los comandos, de otra forma la copia de los comandos sería más dificil.

## Requirements

- Instalar Python y las siguientes librerias:

```shell
pip install python-docx openpyxl tqdm PyPDF2 mtranslate
```
- Descargar el PDF de la guía de bastionado del CIS y convertirlo a Word utilizando la herramienta: https://www.adobe.com/es/acrobat/online/pdf-to-word.html

- Dentro de la función main, se deben especificar las rutas absolutas de los archivos de entrada (Word y PDF) y del archivo de salida (Excel).

## Pasos a segurir

1. Ejecutar el comando para extraer la información del Word y el PDF. Es necesario especificar las rutas de estos archivos. Adicionalmente, una vez generado el documento Word, es importante revisar cómo se organizan los encabezados, ya que pueden variar según el documento.

```shell
python AutoBastion.py --word_path "ruta/del/archivo.docx" --pdf_path "ruta/del/archivo.pdf" --excel_path "ruta/del/salida.xlsx" --heading_level 3
```
**--word_path**:
    Descripción: Especifica la ruta del archivo Word de entrada.
    Tipo: Cadena de texto (str).
    Requerido: Sí.
    Ejemplo: --word_path "C:/Documentos/CIS_Guia.docx"
**--pdf_path**:
    Descripción: Especifica la ruta del archivo PDF de entrada.
    Tipo: Cadena de texto (str).
    Requerido: Sí.
    Ejemplo: --pdf_path "C:/Documentos/Guia.pdf"
**--excel_path**:
    Descripción: Especifica la ruta del archivo Excel de salida.
    Tipo: Cadena de texto (str).
    Requerido: Sí.
    Ejemplo: --excel_path "C:/Documentos/Output.xlsx"
**--heading_level**:
    Descripción: Define el nivel de encabezado que se utilizará para extraer los controles del documento Word. Los niveles de encabezado en Word se utilizan para organizar el documento en jerarquías, donde 1 es el nivel más alto (encabezado principal), 2 es un subencabezado de nivel 1, 3 es un subencabezado de nivel 2, y así sucesivamente. El uso del encabezado debe de ser único en todo el documento, de otra forma puede haber problemas a la hora de seleccionar los controles.
    Tipo: Entero (int).
    Requerido: No (tiene un valor predeterminado).
    Valor Predeterminado: 3
    Ejemplo: --heading_level 3

2. Después, se puede traducir el contenido de las columnas con el siguiente comando:

```shell
python TraduceColumnas.py
```

¡Cuidado! También se traducirán los comandos. No se recomienda en general su uso.

## A tener en cuenta

- Se recomienda el filtrado de ciertas palabras clave como: **CIS** y **P a g e**. Para cambiarlas o eliminarlas.

- **¡¡IMPORTANTE!!** Es obligatorio revisar todas las filas y columnas para detectar posibles errores. Esta herramienta facilita el proceso de copia y pega de los controles, además de la traducción. Sin embargo, puede haber ocasiones en las que no identifique correctamente algunos campos debido a problemas de formato.