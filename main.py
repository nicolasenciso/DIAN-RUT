import pandas as pd
import xlsxwriter
import os
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

URL_DIAN = 'https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces'

def check_RUT(cedula):

    driver = webdriver.Chrome()
    driver.get(URL_DIAN)
    msg = ''

    try:
        search = driver.find_element(By.NAME, "vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit")
        search.send_keys(cedula)
        search.send_keys(Keys.RETURN)
    except NoSuchElementException:
        msg = "ERROR EN MUISCA"

    try:
        elements = driver.find_element(By.ID, "vistaConsultaEstadoRUT:formConsultaEstadoRUT:estado")
        msg = elements.text
    except NoSuchElementException:
        msg = "NO EXISTE"

    driver.quit()
    return msg

def read_Excel(input_file, output_file):

    if os.path.exists(input_file):
        input_excel = pd.read_excel(input_file)
        data = pd.DataFrame(input_excel)

        output_excel = xlsxwriter.Workbook(output_file)
        output_sheet = output_excel.add_worksheet()
        output_sheet.write('A1', 'CASO')
        output_sheet.write('B1', 'RESULTADO')
        index = 2

        for caso in data['casos']:
            output_sheet.write('A'+str(index), str(caso))
            output_sheet.write('B'+str(index), check_RUT(str(caso)))
            index += 1

        output_excel.close()
    else:
        print("No existe el archivo excel a leer")

    print('FINALIZADO')

read_Excel('/home/execlenovo/Desktop/NITS.xlsx', 'Resultados_RUT.xlsx')

