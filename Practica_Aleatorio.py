from time import sleep
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By 
from webdriver_manager.chrome import ChromeDriverManager
import pandas
    
def lectorPrint():

    x = pandas.read_excel(r'C:\Users\migue\Documents\Progamación&Hacking\Proyectos\Web_Scraping\datosRUT_CC.xlsx')
    idsDIAN = x.get('CC-RUT')
    return tuple(idsDIAN)

ser = Service(r'C:\Users\migue\Documents\Progamación&Hacking\Proyectos\Web_Scraping\chromedriver.exe')
op = webdriver.ChromeOptions()
op.add_experimental_option('excludeSwitches', ['enable-logging']) #Deshabilita validador de mouse
# op.add_argument('--headless')                                     #Activación de segundo plano
            
enlace = webdriver.Chrome(service=ser, options=op)

enlace.get("https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces")

datosSea = lectorPrint()

for ccRUT in datosSea:
    inData = enlace.find_element(value='vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit')
    inData.clear()
    inData.send_keys(ccRUT)
    enterData = enlace.find_element(value="vistaConsultaEstadoRUT:formConsultaEstadoRUT:btnBuscar")
    enterData.click()
    sleep(2)
    try:
        enterData = enlace.find_element(value="divMensajeShadow")
        enterData.find_element(By.XPATH,"//img[@src='imagenes/es/botones/botcerrarrerror.gif']").click()
        print(ccRUT,"No existe el usuario")
    except Exception:
        print(ccRUT,"Existe la cuenta")

enlace.close()
enlace.stop_client()
ser.stop()
