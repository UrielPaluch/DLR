from os import stat
import tkinter
import xlsxwriter
from tkinter import *
from tkinter.ttk import Combobox
import tkinter.messagebox
from tkinter import filedialog
from pathlib import Path
from xlsxwriter import *
from functools import partial
import tkinter.font as font
from tkinter import ttk
 
# PANTALLA DE INICIO VACIA
root = Tk()
root.geometry("900x300")
root.title("main")

# DROPDOWN OPTIONS PARA ELEGIR LA PESTANIA DE CONFIGURACION
def abrir():
    global pngGuardar
    # PANTALLA DE CONFIGURAICON
    ConfHoja = Toplevel()
    ConfHoja.title('Configuacion de hoja')
    ConfHoja.geometry('700x300')
    ConfHoja.configure(background='#F08080') #F08080 COLOR SALMON
    ConfHoja.iconbitmap('C:/Users/Pedri/Desktop/Zurich/DLR/Iconos/icono.ico')
    ConfHoja.resizable(False,False)
    myFont = font.Font(family='Nunito')

    # CREO LOS LABELS QUE VA A TENER A LA IZQUIERDA DEL DROPBOX
    labelHoja = Label(ConfHoja,text='Tamaño de hoja: ', font=(myFont,11), bg='#F08080')
    labelHoja.grid(column = 2, row= 2,padx= 5, pady = 5, sticky='w')
    labelLetra = Label(ConfHoja, text = 'Tamaño de letra: ',font=(myFont,11),bg='#F08080' )
    labelLetra.grid(column = 2, row= 4,padx= 5, pady = 5, sticky='w')
    labelVori = Label(ConfHoja, text='Orientacion vertical:',font=(myFont,11), bg='#F08080')
    labelVori.grid(column = 2, row= 6,padx= 5, pady = 5, sticky='w')
    labelHori = Label(ConfHoja, text='Orientacion horizontal:',font=(myFont,11), bg='#F08080')
    labelHori.grid(column = 2, row= 8,padx= 5, pady = 5, sticky='w')

    # CREO EL DROPBOX PARA LA HOJA
    tipoHoja = ['A4', 'A5']
    boxtipoHoja = Combobox(ConfHoja,values = tipoHoja, width=3)
    boxtipoHoja.grid(column = 3, row= 2,padx= 5, pady = 5, sticky='w')

    # TEXT FIELD PARA LA LETRA 
    tmnLetra = Entry(ConfHoja, text = '14' ,width=4) 
    tmnLetra.grid(column = 3, row= 4,padx= 5, pady = 5, sticky='w')

    # DROPBOX PARA ORIENTACIONES VERTICAL Y HORIZONTAL DE TEXTO
    ori_vert = ['Arriba', 'Centro', 'Abajo']
    boxOri_vert = Combobox(ConfHoja, values=ori_vert, width = 8)
    boxOri_vert.grid(column = 3, row= 6,padx= 5, pady = 5, sticky='w')

    ori_hori = ['Izquierda', 'Centro', 'Derecha']
    boxOri_hori = Combobox(ConfHoja, values=ori_hori, width = 8)
    boxOri_hori.grid(column = 3, row= 8,padx= 5, pady = 5, sticky='w')

    # ACA EMPIEZO A USAR XLSXWRITER
    workbook = xlsxwriter.Workbook('C:/Users/Pedri/Desktop/Prueba.xlsx') 
    worksheet = workbook.add_worksheet()
    # BOTON DE GUARDADO  
    def guardar (): # guarda en la lista y cierra la ventana
        boxtipoHoja.get() # Tipo de hoja
        tmnLetra.get() # tamanio de letra
        boxOri_vert.get() # orientacion vertical
        boxOri_hori.get() # orientacion horizontal

        # Valida el campo de numero
        try:
            float(tmnLetra.get())   
        except ValueError:    
            tkinter.messagebox.showerror('Error', 'Ingrese un numero')

        formatocelda = workbook.add_format()
        #Tamanio de letra 
        formatocelda.set_font_size(tmnLetra.get())

        #Orientaciones verticales
        if boxOri_vert.get() == 'Arriba':
            formatocelda.set_align('top')
        elif  boxOri_vert.get() == 'Centro':
            formatocelda.set_align('vcenter')
        else:
            formatocelda.set_align('bottom')

        #Orientaciones horizontales
        if boxOri_hori.get() == 'Derecha':
            formatocelda.set_align('right')
        elif  boxOri_hori.get() == 'Centro':
            formatocelda.set_align('center')
        else:
            formatocelda.set_align('left')

        # Formato de Hoja
        if boxtipoHoja.get() == 'A5':
            worksheet.set_landscape()
            worksheet.set_paper(11)
        else:
            worksheet.set_paper(9)

        #Margenes siempre 0
        worksheet.set_margins(left=0.0, right=0.0, top=0.0, bottom=0.0)
        #centrado de hoja para impresion
        worksheet.center_horizontally()
        worksheet.center_vertically()

        # cierra la ventana y guarda
        if float(tmnLetra.get()):
            tkinter.messagebox.showinfo('Guarado', 'La configuracion ha sido guardada correctamente')
            ConfHoja.destroy() # cierra la ventana y guarda


    # Boton de guardado
    pngGuardar = PhotoImage(file='C:/Users/Pedri/Desktop/Zurich/DLR/Iconos/discket.png') # Esto permite poner la imagen como boton
    savebtn = Button(ConfHoja, image = pngGuardar, command = guardar, bg='#F08080',activebackground= '#F08080' , borderwidth=0)
    #state=DISABLED Esto bloquea el boton, estaria bueno agregarlo cuando no hay nada tipeado
    savebtn.grid(column = 70, row = 20,padx = 25, pady = 15)

btnOptions = Button(root, text = 'Configuracion de hoja', command = abrir)
btnOptions.grid(column = 0, row = 0)
#boxOptions = Combobox(root, values = btnOptions).pack()

root.mainloop()





