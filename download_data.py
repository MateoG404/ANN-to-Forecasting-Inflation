import os
import pandas as pd
import requests
import openpyxl 

#Declaramos las variables necesarias para generar la URL e iterar para descargar los exceles

url = "https://www.dane.gov.co/files/investigaciones/boletines/ipc/anexo_ipc_jun22.xlsx"
url = "https://www.dane.gov.co/files/investigaciones/boletines/ipc/anexo_ipc_"

meses = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
años = ["20", "21", "22"]

#Extraemos directorio actual y en caso de no tener una carpeta donde alojar el excel la creamos

current_directory = os.getcwd()
current_directory = current_directory.replace( "\\" , "/")
excel_inflation_folder = current_directory + "/excel_inflation"

if not os.path.isdir(excel_inflation_folder):
    os.makedirs(excel_inflation_folder)
    print("La carpeta fue creada exitosamente")
else:
    print("la carpeta ya existe")

#Verificamos si existe algún archivo en la carpeta de destino, en caso negativo descargamos, acá queda pendiente 2023

files_in_folder = os.listdir(excel_inflation_folder)

if not files_in_folder:
    for año in años:
        for mes in meses:
            url = "https://www.dane.gov.co/files/investigaciones/boletines/ipc/anexo_ipc_jun22.xlsx"
            url = "https://www.dane.gov.co/files/investigaciones/boletines/ipc/anexo_ipc_"

            url = url + mes + año + ".xlsx"
            response = requests.get(url)

            with open(excel_inflation_folder + "/inflacion" + año + mes + ".xlsx" ,"wb") as file:
                file.write(response.content)
else:
    print("La carpeta contiene archivos:")

# descargados todos los exceles debemos abrirlos, extraer los datos y pegarlos en un nuevo dataset añadiendo la columna data
 
def extract_data(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)

            try:
                workbook = openpyxl.load_workbook(file_path)
                sheet_names = workbook.sheetnames

                if "8" in sheet_names:
                    sheet8 = workbook["8"]
                    
                    # Assuming you want to read data from all cells in Sheet 8
                    data = []
                    for row in sheet8.iter_rows(min_row = 9, max_row=62, values_only=True):
                        data.append(row)

                    # Process data as per your requirements here
                    print(f"Data from {filename}, Sheet 8:")
                    for row_data in data:
                        print(row_data)

                workbook.close()

            except Exception as e:
                print(f"Error processing {filename}: {e}")

# Specify the folder path where the .xlsx files are located
extract_data(excel_inflation_folder)



# c) osea que al correr este .py me descargue los archivos del DANE y despues los abra y me genere el xlsx que estoy necesitando, que ese sea el output
# d) y luego que eliminé los archivos descargados, es decir una descarga temporal, porque no me interesa tenerlos en mi repositorio



