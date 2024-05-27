# Autobastion

Herramienta para automatizar el copia y pega de los controles de las guias de bastionado del CIS de un pdf y word a excel.

## Requirements

- Instalar python y las siguientes librerias:

```shell
pip install python-docx openpyxl tqdm PyPDF2 mtranslate
```
- Descargar el pdf de la guia de bastionado del CIS y convertirlo a word con la herramienta: https://www.adobe.com/es/acrobat/online/pdf-to-word.html

- Dentro de la función main, hay que indicar las rutas absolutas para el input (word y pdf), como del output (excel).

## Pasos a segurir

1. Ejecutamos el comnado para sacar la información del word y el pdf. Es necesario poner la ruta de estos archivos. Adicionalmente, una vez generado el word, hay que ver como se organizan los headings, porque puede variar según el documento.

```shell
python AutoBastion.py
```
Una vez generado el nuevo documento excel la columna 2 (subdominio no suele ser precisa), por lo que se recomienda hacerla a mano.

2. Después es posible traducir las columnas con el siguiente comando:

```shell
python TraduceColumnas.py
```
**¡¡IMPORTANTE!!** Se recomienda ver todas las filas y columnas para encontrar fallos, esta herramienta facilita a la hora de copiar y pegar contoles, además de traducir. Sin embargo, hay ocasiones en las que no identifica bien algunos campos por temas de formato.