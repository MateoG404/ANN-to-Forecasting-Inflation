from download_data import download_data


if __name__ == '__main__':

    descarga_datos = download_data()
    descarga_datos.download_sheets()
    
    descarga_datos.extract_data("2022",10,197,22,"8")    
    descarga_datos.extract_data("2021",10,197,21,"8")
    descarga_datos.extract_data("2020",10,197,20,"8")
    descarga_datos.extract_data("2019",10,197,19,"7")
    
