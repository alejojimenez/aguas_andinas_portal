import os
import re
import time
import shutil
import pandas as pd

def rename_file(folder_path_input, folder_path_output, folder_path_config):
    print('Entrando en la funcion upload...')
    print('----------------------------------------------------------------------')
    
    # Especifica la ruta de tu archivo Excel
    excel_file = folder_path_config + "clientes.xlsx"

    # Especifica el nombre de la hoja en la que se encuentran los datos
    hoja_excel = "Hoja1"

    # Carga los datos de Excel en un DataFrame
    df = pd.read_excel(excel_file, sheet_name=hoja_excel)
    print('Dataframe ', df)
    print('----------------------------------------------------------------------')
    
    # Variable array
    file_name_list = []
    
    # Bucle para obtener lista de nombre de archivos
    for add_file_list in os.listdir(folder_path_input):
        if add_file_list.endswith(".pdf"):
            file_name_list.append(add_file_list)
    
    print('Cantidad Elem. file_name_list: ', len(file_name_list))
    print('----------------------------------------------------------------------')
    
    # Ordenar lista de archivos por nombre
    new_file_name_list_sort = sorted(file_name_list)
    print('file_name_list_sort: ', new_file_name_list_sort, len(new_file_name_list_sort))
    print('----------------------------------------------------------------------')
    
    # Contador de archivos
    file_count = 0
    
    # Recorrer lista con cada archivo, abrir y extraer numero factura
    for x in range(0, len(new_file_name_list_sort)):
        file_count += 1
        input_file = folder_path_input + new_file_name_list_sort[x]
        print('Archivo PDF', input_file, file_count)
        print('----------------------------------------------------------------------')
        time.sleep(2)

        # Separacion de elementos en nombre
        file_name_split = re.split(pattern = r"[_/ / ]", string = str(new_file_name_list_sort[x]))
        print('Separacion de elementos en nombre: ', file_name_split)
        print('--------------------------------------------------------------------------')

        # Cruce datos faltantes para ontener
        for index, row in df.iterrows():
            
            df_nro_cliente = df.loc[index, 'nro_cliente']
            df_sucursal = df.loc[index, 'sucursal']
            print('Nro. Cliente: ', df_nro_cliente, 'Sucursal: ', df_sucursal, 'Split_Capitalize: ', file_name_split[1].capitalize())
            print('--------------------------------------------------------------------------')
            
            if str(file_name_split[1].capitalize()) in str(df_sucursal):

                # Componer nuevo nombre
                new_file_name_combined = str(df_nro_cliente)+'_'+str(file_name_split[0])+'_'+str(file_name_split[-1])
                print('Nuevo nombre compuesto: ', new_file_name_combined)
                print('--------------------------------------------------------------------------')
                
                # Mover a la carpeta output con el nuevo nombre
                source = input_file
                dest = folder_path_output + new_file_name_combined
                shutil.copy(source, dest)
                print('Copiando archivo a nuevo destino: ', source, dest)
                print('--------------------------------------------------------------------------')        
                
                break
                
                
if __name__ == '__main__':
    
    # Obtener en una lista todos los archivos 
    FOLDER_PATH_INPUT = '../input/'
    FOLDER_PATH_OUTPUT = '../output/'
    FOLDER_PATH_CONFIG = '../config/'
    
    rename_file(folder_path_input=FOLDER_PATH_INPUT, 
                folder_path_output=FOLDER_PATH_OUTPUT,
                folder_path_config=FOLDER_PATH_CONFIG
                )



