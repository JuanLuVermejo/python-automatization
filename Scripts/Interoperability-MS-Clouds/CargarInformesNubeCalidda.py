from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
# Configura las opciones del navegador
chrome_options = Options()
#chrome_options.add_argument("--headless")  # Opcional: ejecuta en modo headless (sin interfaz gráfica)

# Inicializa el controlador de Chrome
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    # Abre la página de login de OneDrive
    driver.get("https://onedrive.live.com/login/")

    time.sleep(20)

    # Espera a que la página se cargue y obtén el contenido del <h1>
    h1_element = driver.find_element(By.XPATH, '//h1[@class="row text-title margin-bottom-16"]')
    print(h1_element.text)

finally:
    # Cierra el navegador
    driver.quit()
