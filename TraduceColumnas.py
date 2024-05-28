import openpyxl
from mtranslate import translate
from tqdm import tqdm

# Cargar el archivo de Excel
archivo_excel = r'D:\Usuarios\Ralfamole\Documentos\Cosas de Andres\VS\Autobastion\output.xlsx'
libro = openpyxl.load_workbook(archivo_excel)
hoja = libro.active

# Especificar las columnas que deseas traducir (A, B, C, E, F, G y H)
columnas_a_traducir = ['F', 'G', 'H']

# Calcular el total de celdas a traducir
total_celdas = (hoja.max_row - 1) * len(columnas_a_traducir)

# Configurar la barra de progreso
barra_progreso = tqdm(total=total_celdas, desc="Progreso")

# Iterar a través de las filas comenzando desde la fila 2
for fila in range(2, hoja.max_row + 1):
    # Iterar sobre las columnas especificadas
    for columna in columnas_a_traducir:
        # Leer el valor de la celda en la columna actual
        celda = hoja[columna + str(fila)]
        texto_original = celda.value

        # Traducir el texto si no está vacío
        if texto_original:
            try:
                # Separar el texto antes y después de "#!/usr/bin/env bash"
                partes = texto_original.split("#!/usr/bin/env bash", 1)
                texto_a_traducir = partes[0]
                
                # Traducir solo el texto antes de "#!/usr/bin/env bash"
                texto_traducido = translate(texto_a_traducir, 'es')

                # Recombinar el texto traducido con la parte que no debe ser traducida
                if len(partes) > 1:
                    texto_final = texto_traducido + '\n\n' + "#!/usr/bin/env bash" + partes[1]
                else:
                    texto_final = texto_traducido

                # Escribir el texto final en la misma celda
                celda.value = texto_final
            except Exception as e:
                print(f"Error al traducir la fila {fila}, columna {columna}: {e}")
        
        # Actualizar la barra de progreso
        barra_progreso.update(1)

# Cerrar la barra de progreso
barra_progreso.close()

# Guardar los cambios en el archivo de Excel
libro.save(r'D:\Usuarios\Ralfamole\Documentos\Cosas de Andres\VS\Autobastion\output_es.xlsx')
