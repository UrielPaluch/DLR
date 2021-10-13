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
    ConfHoja.geometry('800x300')
    ConfHoja.configure(background='#F08080') #F08080 COLOR SALMON
    ConfHoja.iconbitmap('C:/Users/Pedri/Desktop/Zurich/DLR/Iconos/icono.ico')
    ConfHoja.resizable(False,False)
    myFont = font.Font(family='Nunito')

    # CREO EL LABEL QUE VA A TENER A LA IZQUIERDA DEL DROPBOX PARA LA HOJA
    # GENERAL
    labelHoja = Label(ConfHoja,text='Tamaño de hoja: ', font=(myFont,11), bg='#F08080')
    labelHoja.grid(column = 2, row= 2,padx= 5, pady = 5, sticky='w')
    labelTituloGeneral = Label(ConfHoja, text = 'Configuracion de campos', font= (myFont,16), bg ='#F08080')
    labelTituloGeneral.grid(column = 4 , row = 3, padx = 10, pady = 5)
    # BOX PARA LA HOJA
    tipoHoja = ['A4', 'A5']
    boxtipoHoja = Combobox(ConfHoja,values = tipoHoja, width=3)
    boxtipoHoja.grid(column = 3, row= 2,padx= 5, pady = 5, sticky='w')
    #Canvas.create_line(15, 25, 200, 25) # para hacer una linea abajo del titulo

    # LABELS NUMERO DE RECETA 
    labelTitulo_receta = Label(ConfHoja,text = 'Numero de receta', font=(myFont,14), bg='#F08080')
    labelTitulo_receta.grid(column = 2,row = 4,padx= 5, pady = 5, sticky='w')

    labelLetra_receta = Label(ConfHoja, text = 'Tamaño de letra: ',font=(myFont,11),bg='#F08080' )
    labelLetra_receta.grid(column = 2, row= 6,padx= 5, pady = 5, sticky='w')

    labelVori_receta = Label(ConfHoja, text='Orientacion vertical:',font=(myFont,11), bg='#F08080')
    labelVori_receta.grid(column = 2, row= 8,padx= 5, pady = 5, sticky='w')

    labelHori_receta = Label(ConfHoja, text='Orientacion horizontal:',font=(myFont,11), bg='#F08080')
    labelHori_receta.grid(column = 2, row= 10,padx= 5, pady = 5, sticky='w')

    labelA_texto_receta = Label(ConfHoja, text='Ajustar Texto:',font=(myFont,11), bg='#F08080')
    labelA_texto_receta.grid(column = 2, row= 12,padx= 5, pady = 5, sticky='w')

    # BOXS NUMERO DE RECETA
        # TEXT FIELD TAMANIO LETRA 
    tmnLetra_receta = Entry(ConfHoja, text = '14' ,width=4) 
    tmnLetra_receta.grid(column = 3, row= 6,padx= 5, pady = 5, sticky='w')
        # BOX ORIENTACION VERTICAL DE TEXTO
    ori_vert_receta = ['Arriba', 'Centro', 'Abajo']
    boxOri_vert_receta = Combobox(ConfHoja, values=ori_vert_receta, width = 8)
    boxOri_vert_receta.grid(column = 3, row= 8,padx= 5, pady = 5, sticky='w')
        # BOX ORIENTACION HORIZONTAL  DE TEXTO
    ori_hori_receta = ['Izquierda', 'Centro', 'Derecha']
    boxOri_hori_receta = Combobox(ConfHoja, values=ori_hori_receta, width = 8)
    boxOri_hori_receta.grid(column = 3, row= 10,padx= 5, pady = 5, sticky='w')
        # BOX AJUSTAR TEXTO 
    boxA_texto_receta = Checkbutton(ConfHoja, bg ='#F08080' , activebackground='#F08080', activeforeground="white")
    boxA_texto_receta.grid(column = 3, row= 12,padx= 5, pady = 5, sticky='w')

    # ACA EMPIEZO A USAR XLSXWRITER
    workbook = xlsxwriter.Workbook('C:/Users/Pedri/Desktop/Prueba.xlsx') 
    worksheet = workbook.add_worksheet()
    # BOTON DE GUARDADO  
    def guardar (): # guarda en la lista y cierra la ventana
        boxtipoHoja.get() # Tipo de hoja
        tmnLetra_receta.get() # tamanio de letra
        boxOri_vert_receta.get() # orientacion vertical
        boxOri_hori_receta.get() # orientacion horizontal

        # Valida el campo de numero
        try:
            float(tmnLetra_receta.get())   
        except ValueError:    
            tkinter.messagebox.showerror('Error', 'Ingrese un numero')

        formatocelda = workbook.add_format()
        #Tamanio de letra 
        formatocelda.set_font_size(tmnLetra_receta.get())

        #Orientaciones verticales
        if boxOri_vert_receta.get() == 'Arriba':
            formatocelda.set_align('top')
        elif  boxOri_vert_receta.get() == 'Centro':
            formatocelda.set_align('vcenter')
        else:
            formatocelda.set_align('bottom')

        #Orientaciones horizontales
        if boxOri_hori_receta.get() == 'Derecha':
            formatocelda.set_align('right')
        elif  boxOri_hori_receta.get() == 'Centro':
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
        if float(tmnLetra_receta.get()):
            tkinter.messagebox.showinfo('Guarado', 'La configuracion ha sido guardada correctamente')
            ConfHoja.destroy() # cierra la ventana y guarda


    # Boton de guardado
    pngGuardar = PhotoImage(file='C:/Users/Pedri/Desktop/Zurich/DLR/Iconos/discket.png') # Esto permite poner la imagen como boton
    savebtn = Button(ConfHoja, image = pngGuardar, command = guardar, bg='#F08080',activebackground= '#F08080' , borderwidth=0)
    #state=DISABLED Esto bloquea el boton, estaria bueno agregarlo cuando no hay nada tipeado
    #savebtn.grid(column = 20, row = 20,padx = 25, pady = 15) # Chequear lo de como se muestra

btnOptions = Button(root, text = 'Configuracion de hoja', command = abrir)
btnOptions.grid(column = 0, row = 0)
#boxOptions = Combobox(root, values = btnOptions).pack()

root.mainloop()





