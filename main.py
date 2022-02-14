# Author: Nicolas Enciso:: twitter @NicolasEncisz
# GNU license. You can use this program, but it belongs the author to the original (Nicolas Enciso)
# Github: nicolasenciso :: nricardoe@unal.edu.co for more info

import pandas as pd
import xlsxwriter
import os
from selenium import webdriver
import tkinter
from tkinter import messagebox
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

def read_Excel(input_file, output_file, frame):

    exists = True
    if not os.path.exists(input_file):
        exists = False
        messagebox.showinfo("ERROR", "La ruta del Excel con los casos no existe")

    if exists:
        input_excel = pd.read_excel(input_file)
        data = pd.DataFrame(input_excel)

        output_excel = xlsxwriter.Workbook(output_file)
        output_sheet = output_excel.add_worksheet()
        output_sheet.write('A1', 'CASO')
        output_sheet.write('B1', 'RESULTADO')
        index = 2

        num_cases = 0
        for i in data['casos']:
            num_cases += 1

        case_number = 1
        for caso in data['casos']:

            #last_nit.pack_forget()
            #last_nit = tkinter.Label(frame, text="Revisando NIT: "+str(caso), font="Helvetica 13 bold", pady=10)
            #last_nit.pack()

            #status.pack_forget()
            #status = tkinter.Label(frame, text="Caso "+str(case_number)+" de "+str(num_cases), font="Helvetica 13 bold", pady=5)
            #status.pack(side=tkinter.BOTTOM)
            #case_number += 1

            output_sheet.write('A'+str(index), str(caso))
            output_sheet.write('B'+str(index), check_RUT(str(caso)))
            index += 1

        output_excel.close()


    finished = tkinter.Label(frame, text="FINALIZADO", font="Helvetica 13 bold", bg='#738f65', width="600")
    finished.pack()

def GUI():
    window = tkinter.Tk()
    window.title("CONSULTA NIT en DIAN")
    window.geometry("600x350")

    frame = tkinter.Frame(window)

    title = tkinter.Label(frame, text="CONSULTA AUTOMATICA DE NIT EN DIAN", bg="#738f65", width="600",
                          font="Helvetica 13 bold", pady=10)

    title.pack(fill=tkinter.X)

    title_input_excel = tkinter.Label(frame, text="Ingrese la ruta completa del Excel con los casos:",
                                      font="Helvetica 13 bold", pady=10)
    title_input_excel.pack()

    input_excel = tkinter.Entry(frame)
    input_excel.pack(fill=tkinter.X)

    label_blank = tkinter.Label(frame, pady=7)
    label_blank.pack()

    title_output_excel = tkinter.Label(frame, text="Ingrese la ruta completa del Excel a guardar resultados:",
                                       font="Helvetica 13 bold", pady=10)
    title_output_excel.pack()

    output_excel = tkinter.Entry(frame)
    output_excel.pack(fill=tkinter.X)

    label_blank = tkinter.Label(frame, pady=7)
    label_blank.pack()

    #status = tkinter.Label(frame, text="Estado de la revision:", font="Helvetica 13 bold", pady=5)
    #status.pack()

    #last_nit = tkinter.Label(frame, text="Revisando NIT: ", font="Helvetica 13 bold", pady=10)
    #last_nit.pack()

    execute_button = tkinter.Button(frame, text="EJECUTAR", activeforeground="green",
                                    command=lambda: read_Excel(input_excel.get(), output_excel.get(), frame),
                                    font="Helvetica 13 bold")
    execute_button.pack()
    frame.pack()
    window.mainloop()

GUI()



#'/home/execlenovo/Desktop/NITS.xlsx'
#'/home/execlenovo/Desktop/res.xlsx'