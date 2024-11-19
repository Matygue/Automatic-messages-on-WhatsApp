import pandas as pd
import webbrowser as web
import pyautogui as pg
import time
from urllib.parse import quote

# Cargar el archivo Excel
data = pd.read_excel("C:/Users/MATY/Desktop/Ayudas a mis papas/Clientes.xlsx", sheet_name="Ventas")


for i in range(len(data)):
    celular = str(data.loc[i, 'Celular']).strip()  

    # Asegurarse de que cada número tiene el prefijo "549"
    if not celular.startswith("549"):
        celular = "549" + celular

    # Crear el mensaje
    mensaje = (
        "Hola buenas tardes, soy Daniel de G.Y.G Computación.\n\n"
        "Quería comentarte que tenemos precios especiales a comercios del pueblo en insumos para impresoras, como tóner y cartuchos.\n\n"
        "¿Te gustaría que te envíe la lista para que veas nuestros precios?"
    )

    # Codificar el mensaje para la URL
    mensaje_cifrado = quote(mensaje)

    # Crear la URL completa para enviar en WhatsApp
    url = f"https://web.whatsapp.com/send?phone={celular}&text={mensaje_cifrado}"


    web.open(url)
    
    # Espera para que la página cargue y envíe el mensaje
    time.sleep(12)  

    pg.click(1600,984)      
    time.sleep(5)           
    pg.press('enter')       
    time.sleep(3)          
    pg.hotkey('ctrl', 'w')  
    time.sleep(2)
  
