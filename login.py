import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def iniciar_sesion(url, tipo_id, usuario, password):
    options = Options()
    options.add_experimental_option("detach", True)

    # Descarga e instala automáticamente el ChromeDriver correcto
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 20)

    driver.get(url)
    time.sleep(5)

    iframe_login = driver.find_elements(By.TAG_NAME, "iframe")[0]
    driver.switch_to.frame(iframe_login)

    Select(wait.until(EC.element_to_be_clickable((By.ID, "tipoId")))).select_by_value(tipo_id)

    wait.until(EC.element_to_be_clickable((By.ID, "username"))).send_keys(usuario)
    wait.until(EC.element_to_be_clickable((By.NAME, "josso_password"))).send_keys(password)

    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Ingresar']"))).click()

    time.sleep(8)
    driver.switch_to.default_content()

    rol = Select(wait.until(EC.element_to_be_clickable((By.XPATH, "//select"))))
    rol.select_by_visible_text("Usuario SENA")

    time.sleep(5)

    return driver, wait