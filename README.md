# Autobastion

Herramienta para automatizar el copia y pega de los controles de las guias de bastionado del CIS de un pdf y word a excel.

## Requirements

- Instalar python y las siguientes librerias:

```bash
$ pip install python-docx openpyxl tqdm PyPDF2
```
- Descargar el pdf de la guia de bastionado del CIS y convertirlo a word con la herramienta: https://www.adobe.com/es/acrobat/online/pdf-to-word.html

- Dentro de la función main, hay que indicar las rutas absolutas para el input (word y pdf) como del output (excel)