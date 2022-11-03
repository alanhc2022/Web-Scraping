#Ejecutado desde Anaconda - Spyder
#Para poder sacar el XPATH puedes ver un vídeo tutorial en Youtube.
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import pandas as pd
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
#!pip install selenium
#Lista de DNI que requiere buscar información.
#Una sola columna con número de DNI.
df = pd.read_excel("C:/Users/HP/Documents/Prueba_dnis2.xlsx")
a=df['DNI'].tolist()
al=[]
#Convertimos los DNI a 8 dígitos
for num in a:
    al.append(str(num).rjust(8, '0'))

lis=[]
    for i in range(len(df)):
    #try:
        website = 'https://e-consultaruc.sunat.gob.pe/'
        path = 'C:/Users/HP/Documents/Comunidata/chromedriver_win32'         
        #Ingresar a la Web
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.get(website)
        #Damos click al boton "Documento", para colocar el número.
        click_botones = driver.find_element(By.XPATH,'//*[@id="btnPorDocumento"]')
        click_botones.click() #Realiza el click
        #Obtenemos la caja de texto donde se ingresa el DNI
        documento = driver.find_element(By.NAME,"search2")
        #Ingresamos el número de DNI
        documento.send_keys(al[i])
        #Esperamos hasta que el texto esté escrito en la caja de texto del DNI
        time.sleep(2)
        
        #Obtenemos el botón de búsqueda
        boton_busqueda = driver.find_element(By.CLASS_NAME,"btn-primary")
        #Damos click al botón de búsqueda
        boton_busqueda.click()
        time.sleep(3)
        #Acá hacemos un comparativo cuando haya o no haya información de la persona.
        comparativo = driver.find_element(By.XPATH,'//div[@class="panel-heading"]').text
        if( comparativo == "Relación de contribuyentes"):
            #Si la cabecera sale Relación de contribuyentes, hacemos click.
            WebDriverWait(driver, 5)\
                .until(EC.element_to_be_clickable((By.XPATH,
                                                  '/html/body/div/div[2]/div/div[3]/div[2]/a')))\
                .click()
            #Extraemos la información de la tabla que se desea.
            ruc_nombre= driver.find_element(By.XPATH,'/html/body/div/div[2]/div/div[3]/div[2]/div[1]/div/div[2]/h4[@class="list-group-item-heading"]')
            Tipo_contribuyente= driver.find_element(By.XPATH,'/html/body/div/div[2]/div/div[3]/div[2]/div[2]/div/div[2]/p[@class="list-group-item-text"]')
            DNI_nombre= driver.find_element(By.XPATH,'/html/body/div/div[2]/div/div[3]/div[2]/div[3]/div/div[2]/p[@class="list-group-item-text"]')
            fecha_inscripción= driver.find_element(By.XPATH,'/html/body/div/div[2]/div/div[3]/div[2]/div[5]/div/div[2]/p[@class="list-group-item-text"]')
            fecha_inicio_act= driver.find_element(By.XPATH,'/html/body/div/div[2]/div/div[3]/div[2]/div[5]/div/div[4]/p[@class="list-group-item-text"]')
            estado_contribuye= driver.find_element(By.XPATH,'/html/body/div/div[2]/div/div[3]/div[2]/div[6]/div/div[2]/p[@class="list-group-item-text"]')
            condicion_contribuye= driver.find_element(By.XPATH,'/html/body/div/div[2]/div/div[3]/div[2]/div[7]/div/div[2]/p[@class="list-group-item-text"]')
            #Almacenamos en una lista = 'lis'
            #Se pueden agregar más campos.
            a=ruc_nombre.text
            b=Tipo_contribuyente.text
            c=estado_contribuye.text
            d=fecha_inscripción.text
            e=condicion_contribuye.text
            datos=[a,b,c,d]
            lis.append(datos)
            driver.close()
        else:
            #Si no encuentra información, cierra la página.
            #También podías agregar aquí, una tabla con el número de documento y la observación de que no hay info.
            driver.close()
    return []
    #Vemos la lista
    lis
    #Acá puedes agregar más columnas, depende de la información que deseas extraer.
    column_names = ["RUC", "TIPO_CONTRIBUYENTE","ESTADO_CONTRIBUYENTE","FECHA_INSCRIPCIÓN"]
    datos=pd.DataFrame(lis, columns=column_names)
    #Almacenamos en un CSV, también puede ser en Excel...
    datos.to_csv('numeros.csv', header=True, index=False)


