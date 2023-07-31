import os
import time
import shutil
import requests
import re

#from domain.chrome_node import ChromeNode

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from openpyxl import load_workbook
from datetime import datetime,timedelta

class Scraper_Aguas():

    def __init__(self,url, email, password, driver_path):
        print(url, email, password, driver_path)
        self.url = url
        self.email = email
        self.password = password
        self.driver_path = driver_path

    def wait(self, seconds):
        return WebDriverWait(self.driver, seconds)

    def close(self):
        self.driver.close()
        self.driver = None

    def quit(self):
        self.driver.quit()
        self.driver = None    

    def login(self):
        
        driver_exe = '.domain\\chromedriver.exe'
        credencials = '.\\config\\credenciales.xlsx'

        print('Entrando en la funcion login...')
        print('----------------------------------------------------------------------')
        
        #Seteo variables
        email = self.email
        url = self.url
        driver_path = self.driver_path
        password  = self.password
        
        options = webdriver.ChromeOptions()
            
        self.driver = webdriver.Chrome(driver_path, options=options)
        self.driver.get(url)
        self.driver.maximize_window()
        self.driver.implicitly_wait(40)

        # Controlar evento alert() Notificacion
        try:
            alert = WebDriverWait(self.driver, 35).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.accept()
                    
        except:
            print("No se encontró ninguna alerta.")  

        #Botones ID que utilizaremos para logear
        selector_ingreso_cuenta = 'wlogin_login'
        selector_rut_input = 'rut2'
        selector_password_input = 'clave'
        selector_login_button = 'b_login_1'

        # Seleccionar campo cuenta
        intentos = 0
        reintentar = True
        while (reintentar):
            try:
                print('Try en la funcion login para el campo cuenta...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                element_cuenta= self.driver.find_element(By.ID, selector_ingreso_cuenta)
                element_cuenta.click()
                reintentar = False
            except:    
                print('Exception en la funcion click campo cuenta')
                print('----------------------------------------------------------------------')
                reintentar = intentos <= 3                
       
       
        time.sleep(10) 
        # Seleccionar campo rut y setear usuario
        intentos_rut = 0
        reintentar = True
        while (reintentar):
            try:
                print('Try en la funcion click para el campo rut...', intentos_rut)
                print('----------------------------------------------------------------------')
                intentos += 1
                element_username = self.driver.find_element(By.ID, selector_rut_input)
                element_username.click()
                element_username.send_keys(email)
                reintentar = False
            except:    
                print('Exception en la funcion click rut cliente')
                print('----------------------------------------------------------------------')
                reintentar = intentos_rut <= 3
                
        # Seleccionar campo clave y setear clave
        intentos = 0
        reintentar = True
        while (reintentar):
            try:
                print('Try en la funcion click para el campo rut...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                element_password = self.driver.find_element(By.ID, selector_password_input)
                element_password.click()
                element_password.send_keys(password)
                reintentar = False
            except:    
                print('Exception en la funcion click rut cliente')
                print('----------------------------------------------------------------------')
                reintentar = intentos <= 3

        # Seleccionar boton y hacer en boton ingresar
        intentos = 0
        reintentar = True
        while (reintentar):
            try:
                print('Try en la funcion click_element_xpath para el boton ingresar...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                element_button_ingresar = self.driver.find_element(By.ID, selector_login_button)
                element_button_ingresar.click()
                reintentar = False
            except:    
                print('Exception en la funcion click_element_xpath')
                print('----------------------------------------------------------------------')
                reintentar = intentos <= 3
    
    
    def scrapping_aguas(self,posicion):

        #Buscamos la tabla que contiene el boton de boletas
        WebDriverWait(self.driver, 30).until(
        EC.presence_of_element_located((By.ID,'scta')))
        boton_sucursal= self.driver.find_element(By.ID,'scta')
        boton_sucursal.click()

        #Buscamos las sucursales y generamos lista con los nombres
        lista_sucursales = []
        boton_sucursal_abierto = self.driver.find_elements(By.ID,'selCu')
        for index,boton in enumerate(boton_sucursal_abierto):
            #print(index,'-',boton.get_attribute('innerText'))
            sucursal = boton.get_attribute('innerText')
            resultado = re.sub(r"[^a-zA-Z ]", "", sucursal).strip()
            print(resultado)
            lista_sucursales.append(resultado)
        
        cantidad_sociedades = len(lista_sucursales)
        print(f'el largo de los locales es {cantidad_sociedades}')
        
        while posicion < cantidad_sociedades+1:
            
            #Boton de las sociedades
            intento_sociedades = 0
            reintentar_sociedades = True
            while (reintentar_sociedades):
                try:
                    time.sleep(5)
                    boton_sucursal_of = self.driver.find_elements(By.ID,'selCu')
                    time.sleep(5)    
                    menu_sucursal = boton_sucursal_of[posicion]
                    menu_sucursal.click()
                    reintentar_sociedades = False
                    time.sleep(5)
                except:    
                    print('menu sociedades no esta clicleable')
                    print('----------------------------------------------------------------------')
                    self.driver.execute_script("window.scrollTo(0, 0)")
                    reintentar_sociedades = intento_sociedades <= 5

            #Boton resumen boletas
            intento_resumen= 0
            reintentar_resumen = True
            while (reintentar_resumen):
                try:
                    boton_sucursal_of = self.driver.find_element(By.ID,'pestanaResumenBoletas')
                    time.sleep(2)    
                    boton_sucursal_of.click()
                    reintentar_resumen = False
                    time.sleep(5)
                except:    
                    print('menu resumen boletas no esta clicleable')
                    print('----------------------------------------------------------------------')
                    reintentar_resumen = intento_resumen <= 3

            #Boton ver mas boletas           
            intento_ver_mas= 0
            reintentar_ver_mas = True
            while (reintentar_ver_mas):
                try:
                    boton_sucursal_of = self.driver.find_element(By.LINK_TEXT,'Ver más documentos')
                    time.sleep(2)    
                    boton_sucursal_of.click()
                    reintentar_ver_mas = False
                    time.sleep(5)
                except:    
                    print('link ver mas boletas no esta clicleable')
                    print('----------------------------------------------------------------------')
                    self.driver.execute_script("window.scrollBy(0, 200)")
                    reintentar_ver_mas = intento_ver_mas <= 5
                    
            #Encontrar la tabla y sus elementos          
            intento_ver_tabla= 0
            reintentar_ver_tabla = True
            while (reintentar_ver_tabla):
                try:
                    tabla_element = self.driver.find_element(By.XPATH,'//*[@id="myTablePer"]')
                    time.sleep(2)
                    celdas = tabla_element.find_elements(By.TAG_NAME,'td')
                    reintentar_ver_tabla = False
                    time.sleep(5)
                except:    
                    print('Buscando tabla con datos')
                    print('----------------------------------------------------------------------')
                    reintentar_ver_tabla = intento_ver_tabla <= 5

            fila = 1
            while fila < 14:            
                
                if fila == 13:
                    #Siguiente pagina
                    intentos_cambio_hoja = 0
                    reintentar_cambio = True
                    while (reintentar_cambio):
                        try:
                            tabla_element_cambio = self.driver.find_element(By.XPATH,f'//*[@id="tabs-1"]/div[1]/ul/li[3]')
                            time.sleep(2)
                            tabla_element_cambio.click()
                            reintentar_cambio = False
                        except:
                            print('no hay numero de factura')
                            print('----------------------------------------------------------------------')
                            reintentar_cambio = intentos_cambio_hoja <= 3
                    
                #Numero de factura
                intentos_factura = 0
                reintentar_factura = True
                while (reintentar_factura):
                    try:
                        tabla_element_nfact = self.driver.find_element(By.XPATH,f'//*[@id="myTablePer"]/tbody/tr[{fila}]/td[1]')
                        factura = tabla_element_nfact.text
                        time.sleep(2)
                        reintentar_factura = False
                    except:
                        print('no hay numero de factura')
                        print('----------------------------------------------------------------------')
                        reintentar_factura = intentos_factura <= 3
                
                #Mes y año
                intentos_fecha= 0
                reintentar_fecha = True
                while (reintentar_fecha):
                    try: 
                        tabla_element_fecha = self.driver.find_element(By.XPATH,f'//*[@id="myTablePer"]/tbody/tr[{fila}]/td[2]')
                        fecha = tabla_element_fecha.text
                        partes = fecha.split("/")
                        mes = str(partes[0])
                        año = str(partes[1])
                        time.sleep(2)
                        reintentar_fecha = False
                    except:
                        print('no hay fecha')
                        print('----------------------------------------------------------------------')
                        reintentar_fecha = intentos_fecha <= 3
                
                #Boton descarga
                intentos_down= 0
                reintentar_down= True                
                while (reintentar_down):
                    try:
                        
                        tabla_element_descarga = self.driver.find_element(By.XPATH,f'//*[@id="myTablePer"]/tbody/tr[{fila}]/td[7]')
                        time.sleep(2)
                        tabla_element_descarga.click()
                        reintentar_down = False
                    except:
                        print('no hay boton de descarga')
                        print('----------------------------------------------------------------------')
                        reintentar_down = intentos_down <= 3    
        

                time.sleep(5)

                # Esperar a que se abra la ventana emergente
                intentos = 0
                reintentar_ventana = True
                while (reintentar_ventana):
                    try:
                        # Obtiene el identificador de la ventana actual
                        current_window = self.driver.current_window_handle
                        print('Ventana principal: ', current_window)
                        print('----------------------------------------------------------------------')
                        print('Try en la funcion manejo de ventanas abiertas...', intentos)
                        print('----------------------------------------------------------------------')
                        intentos += 1
                        window_handles_all = self.driver.window_handles
                        reintentar_ventana = False
                            
                    except:    
                        print('Exception en la funcion manejo de ventanas abiertas')
                        print('----------------------------------------------------------------------')
                        print('----------------------------------------------------------------------')
                        time.sleep(60) #espera para que cargue ventana emergente
                                    
                        reintentar_ventana = intentos <= 3
                                    
                # Obtiene los identificadores de las ventanas abiertas
                self.driver.implicitly_wait(35)
                print('Ventanas abiertas: ', window_handles_all, len(window_handles_all))
                print('----------------------------------------------------------------------')            
                                    
                # Cambiar al manejo de la ventana emergente
                for window_handle in window_handles_all:
                    if window_handle != current_window:
                        self.driver.switch_to.window(window_handle)
                        print('Ventana emergente: ', window_handle)
                        print('----------------------------------------------------------------------')
                        break

                # Esperar hasta que el elemento esté presente en la página
                self.driver.implicitly_wait(35)
                            
                try:
                    # Obtener la URL de la ventana emergente
                    ventana_emergente_url = self.driver.current_url
                    print("URL de la ventana emergente:", ventana_emergente_url)
                    print('----------------------------------------------------------------------')

                    # Realizar una solicitud GET para obtener la data binaria del documento
                    response = requests.get(ventana_emergente_url, stream=True)

                    # Obtener el nombre del archivo a partir de los datos del proceso de descarga
                    folder_path = './input/'
                    file_name = folder_path + str(factura) +"_"+ str(lista_sucursales[posicion])+"_"+ str(mes)+"_"+str(año)+".pdf" 

                    # Guardar la data binaria en un archivo PDF
                    with open(file_name, 'wb') as file:
                        response.raw.decode_content = True
                        shutil.copyfileobj(response.raw, file)
                        print("Guardando archivo:", file_name)
                        print('----------------------------------------------------------------------')
                                    
                except:    
                    print("No se encontró el elemento con el id especificado...")
                    print('----------------------------------------------------------------------')
                                
                    count += 1
                    print('Conteo de documentos: ', count)
                    print('----------------------------------------------------------------------')                

                # Cerrar la ventana emergente
                time.sleep(5)
                self.driver.close()
                time.sleep(15)                            

                # Cambiar de nuevo al manejo de ventana principal
                self.driver.switch_to.window(current_window)
                print('Cual ventana es: ', current_window)
                print('----------------------------------------------------------------------')
                        
                time.sleep(10)
                
                print('pasamos al siguiente archivo')
                fila +=1 
            break
        print('Pasamos a la siguiente sociedad')
