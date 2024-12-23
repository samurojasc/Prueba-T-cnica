import os
import plotly.graph_objects as go
import zipfile
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from datetime import datetime


def configurar_navegador():
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    firefox_options = Options()
    firefox_options.add_argument('--headless')
    driver = webdriver.Firefox(options=firefox_options)
    return driver, download_dir


def descargar_archivo(driver, download_dir):
    url = "https://guiadevalores.fasecolda.com/ConsultaExplorador/Default.aspx?url=C:\inetpub\wwwroot\Fasecolda\ConsultaExplorador\Guias\GuiaValores_Unificada%20-%20337%20en%20adelante"
    driver.get(url)

    # Esperar y localizar filas de la tabla
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'gv_data'))
    )
    rows = driver.find_elements(By.CSS_SELECTOR, '#gv_data tr.gridViewRow, #gv_data tr.gridViewAlternateRow')

    last_row = None
    last_date = None

    for row in rows:
        cols = row.find_elements(By.TAG_NAME, 'td')
        date_text = cols[-1].text.strip()
        date_object = datetime.strptime(date_text, '%m/%d/%Y %I:%M:%S %p')
        if not last_date or date_object > last_date:
            last_date = date_object
            last_row = cols[1].find_element(By.TAG_NAME, 'a')

    if last_row:
        print(f"Última carpeta encontrada: {last_row.text}")
        last_row.click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'gv_data'))
        )

        file_link = driver.find_element(By.CSS_SELECTOR, '#gv_data tr.gridViewRow a')
        file_name = file_link.text
        file_link.click()

        print("Esperando la descarga del archivo ZIP...")
        zip_file_path = os.path.join(download_dir, file_name)
        return zip_file_path
    else:
        print("No se encontró ninguna carpeta.")
        return None


def procesar_archivo(zip_file_path):
    script_dir = os.getcwd()
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(script_dir)
        print("Archivo ZIP extraído.")

    for file in sorted(os.listdir(script_dir), key=lambda f: os.path.getmtime(os.path.join(script_dir, f)), reverse=True):
        if file.endswith('.xlsx') or file.endswith('.xls'):
            return os.path.join(script_dir, file)
    print("No se encontró ningún archivo Excel en el ZIP.")
    return None


def graficar_datos(codigos_data):
    años = list(range(1970, 2027))
    codigos_data.iloc[:, codigos_data.columns.get_loc('1970'):].replace(0, pd.NA, inplace=True)
    promedios_por_marca = codigos_data.groupby('Marca')[list(map(str, años))].mean()
    promedios_por_marca_t = promedios_por_marca.T

    fig = go.Figure()
    for marca in promedios_por_marca_t.columns:
        fig.add_trace(go.Scatter(
            x=promedios_por_marca_t.index,
            y=promedios_por_marca_t[marca],
            mode='lines',
            name=marca,
            hoverinfo='name+y'
        ))

    fig.update_layout(
        title='Curvas de Depreciación por Marca (1970-2026)',
        xaxis_title='Año',
        yaxis_title='Valor Promedio (en unidades)',
        showlegend=True,
        margin=dict(l=40, r=40, t=40, b=40),
        height=600,
        width=1000,
        template='plotly_white'
    )
    fig.show()


if __name__ == '__main__':
    driver, download_dir = configurar_navegador()

    try:
        zip_file_path = descargar_archivo(driver, download_dir)
        if zip_file_path:
            excel_file_path = procesar_archivo(zip_file_path)
            if excel_file_path:
                print(f"Archivo Excel encontrado: {excel_file_path}")
                codigos_data = pd.read_excel(excel_file_path, sheet_name="Codigos")
                graficar_datos(codigos_data)
    finally:
        driver.quit()


