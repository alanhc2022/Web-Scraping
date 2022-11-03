#El archivo es ejecutado desde Anaconda - Spyder.
#Para poder sacar el XPATH puedes ver un vídeo tutorial en Youtube.
#Requiere descargar el chromedriver_win32
#https://chromedriver.chromium.org/downloads
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import pandas as pd
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

#!pip install selenium
#!pip install webdriver_manager.chrome
#!pip install pytesseract
try:
    import Image    
except ImportError:
    from PIL import Image    

#pytesseract una librería de procesamiento de imágenes, leer en internet para más detalle.
import pytesseract

#Leemos nuestro archivo excel, con una sola columna de "DNI"

df = pd.read_excel("C:/Users/Nombre/Pictures/Prueba_dnis3.xlsx")
a=df['DNI'].tolist()
al=[]

#convertimos los DNI, en formato 8 dígitos

for num in a:
    al.append(str(num).rjust(8, '0'))

lis=[]
lis2=[]
lis3=[]
import time

inicio = time.time() #para calcular el tiempo de ejecución de nuestro for

#Usamos el try/except: Try cuando cumpla la condición y no tenga errores, en caso tenga algún error
#ejecutará el Excetp, es decir, si no hay información de la persona, hará un comparativo para ver
#si no hay resultados, o el captcha fue mal puesto.

for i in range(len(df)):
    try:
        website = 'https://constancias.sunedu.gob.pe/verificainscrito'
        path = 'C:/Users/Nombre/Pictures/Saved Pictures/chromedriver_win32' 
        
        #Acceder a la web 
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.get(website)
        #Luego de cargar la web, espera un tiempo para recién poner el número del documento
        time.sleep(1) 
        documento = driver.find_element(By.XPATH,'//*[@id="doc"]')
        #Ingresa el número del documento a la web
        documento.send_keys(al[i])
        #Verifica si hay captcha o una imagen / no es necesario el código.
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="captchaImg"]/img')))
        
        #Hacemos la captura de pantalla correspondiente
        time.sleep(1)     
        driver.save_screenshot("screenshot.jpg")
        img=Image.open('screenshot.jpg')
        #img.size -> Podemos ver la dimensión de la imagen para hacer el recorte que se requiere.
        img_recortada = img.crop((815,55,940,95))
        #Guardamos el recorte
        gray = img_recortada.convert('L')
        gray.save("recorte.jpg")
        #Es necesario descargar los accesos que estarán en la descripción.
        #https://github.com/UB-Mannheim/tesseract/wiki
        #tesseract-ocr-w64-setup-v5.2.0.20220712.exe (64 bit) resp.
        pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Nombre\AppData\Local\Tesseract-OCR\tesseract.exe'
        #Obtenemos el texto del captcha
        captcha = pytesseract.image_to_string(gray)
        #Obtenemos la caja de texto donde se escribe el texto del captcha
        codigo = driver.find_element(By.XPATH,'//*[@id="captcha"]')
         #Escribimos el texto en la web
        codigo.send_keys(captcha) 
        
        #Luego de haber ingresado el captcha, buscamos la información.
        WebDriverWait(driver, 5)\
            .until(EC.element_to_be_clickable((By.XPATH,'//*[@id="buscar"]')))\
            .click()
        #Le damos un tiempo antes de extraer la información, dado que se satura la web de sunedu
        #y también depende de la velocidad de la red, en caso 5 no sea suficiente, se agrega más.
        time.sleep(5)        
        #Extraemos el nombre del graduado, grado académico y la universidad de la tabla.
        #Resalto que aún falta mejorar, para extraer más información de la tabla, ya que extraigo
        #la primera fila, bajo ciertas condiciones se podría modificar y extraer cuando sólo tenga 1
        # fila, 2 filas, 3 filas, etc.
        nombre_graduado= driver.find_element(By.XPATH,'//*[@id="finalData"]/tr/td[1]')
        grado_academico= driver.find_element(By.XPATH,'//*[@id="finalData"]/tr/td[2]')
        universidad= driver.find_element(By.XPATH,'//*[@id="finalData"]/tr/td[3]')
        
        #Guardamos el nombre en forma de texto, y lo agregadmos a nuestra lista "lis"            
        a=nombre_graduado.text
        b=grado_academico.text
        c=universidad.text
        datos=[a,b,c]
        lis.append(datos)
        #Si cumplió las condiciones, es decir, que haya información, luego de extraer se cierra la web.
        driver.close()
    except:
        #Al salir error, se ejecuta lo siguiente: Hacemos el comparativo si el error es porque 
        #"No se encontraron resultados", guardamos el DNI, con una observación "No se encontró resultados"
        #y lo agregamos a la lista "lis". Si no cumple esa condicón quiere decir que el captcha fue mal puesto
        # y lo guardamos en la lis3 los números de "DNI" y la observación "Captcha no reconocido"
        comparativo = driver.find_element(By.XPATH,'//*[@id="frmError_Body"]').text
        if( comparativo == "No se encontraron resultados."):
            datos1=[al[i],'No se encontró resultados']
            lis2.append(datos1)
        else:
            datos2=[al[i],'Captcha no reconocido']
            lis3.append(datos2)
        #Al finalizar las acciones, cerramos la web.
        driver.close()
        
fin = time.time()

(fin-inicio)/60 #Tiempo de ejecución del for

#len(lis) Lista registrados / bachiller
#len(lis2) Lista no registrados / sin título alguno
#len(lis3) Lista captcha mal puesto

#Aquí podemos agrupar todas las listas en un sólo archivo Excel, agregando un campo a lis2 y lis3.
#Juntarlos en un DataFrame, o simplemente guardar en diferentes Excel, para diferenciar de los cuales
# se obtuvo o no información alguna.
column_names = ["NOMBRE", "GRADO","UNIVERSIDAD"]
datos=pd.DataFrame(lis, columns=column_names)
datos.to_excel('SUNEDU_1.xlsx', header=True, index=False)

column_names = ["NUMERDOC", "OBSERVACION"]
datos2=pd.DataFrame(lis2, columns=column_names)
datos2.to_excel('SUNEDU_2.xlsx', header=True, index=False)

column_names = ["NUMERDOC", "OBSERVACION"]
datos3=pd.DataFrame(lis3, columns=column_names)
datos3.to_excel('SUNEDU_3.xlsx', header=True, index=False)