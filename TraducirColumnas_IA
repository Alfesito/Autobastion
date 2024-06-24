import openpyxl
from mtranslate import translate
from tqdm import tqdm
import requests
import json

#Tu clave de API
api_key = "AIzaSyAkO8l32JaDyHcwxVSE7JutO5-NrYRTLCc"

#URL de la API
url = f"https://generativelanguage.googleapis.com/v1/models/gemini-pro:generateContent?key={api_key}"

# Cargar el archivo de Excel
archivo_excel = r'C:\Users\aalfarofernandez\OneDrive - Deloitte (O365D)\Documents\Scripts\AutoBast\output.xlsx'
libro = openpyxl.load_workbook(archivo_excel)
hoja = libro.active

# Especificar las columnas que deseas traducir (A, B, C, E, F, G y H)
columnas_a_traducir = ['F', 'H', 'I']

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
        if texto_original and texto_original != "-":
            try:
                #Datos de la solicitud
                data = {
                    "contents": [
                        {
                            "role": "user",
                            "parts": [{"text": f"Traduceme al ingles el texto natural pero no los comandos o scripts: {texto_original}"}]
                        }
                    ]
                }

                #Encabezados de la solicitud
                headers = {
                    "Content-Type": "application/json"
                }

                #Envío de la solicitud POST a la API
                response = requests.post(url, headers=headers, data=json.dumps(data))
                response.raise_for_status()  # Lanza un error para códigos de estado HTTP malos
                result = response.json()

                #Imprimir la respuesta completa para ver su estructura
                print(json.dumps(result, indent=4))

                #Verificar si hay candidatos y mostrar el contenido
                if "candidates" in result and len(result["candidates"]) > 0:
                    content = result["candidates"][0]["content"]["parts"][0]["text"]
                    #print("Response from AI:", content)
                    celda.value = content
                else:
                    print("No valid response received from the API.")
                
            except Exception as e:
                print(f"Error al traducir la fila {fila}, columna {columna}: {e}")
        
        # Actualizar la barra de progreso
        barra_progreso.update(1)

# Cerrar la barra de progreso
barra_progreso.close()

# Guardar los cambios en el archivo de Excel
libro.save(r'C:\Users\aalfarofernandez\OneDrive - Deloitte (O365D)\Documents\Scripts\AutoBast\output_es.xlsx')
