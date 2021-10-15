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
    ConfHoja.geometry('450x250')
    ConfHoja.configure(background='#F1948A') #F1948A COLOR SALMON
    ConfHoja.iconbitmap('C:/Users/Pedri/Desktop/Zurich/DLR/Iconos/icono.ico')
    ConfHoja.resizable(False,False)
    myFont = font.Font(family='Nunito')

    # CREO EL LABEL QUE VA A TENER A LA IZQUIERDA DEL DROPBOX PARA LA HOJA
    # GENERAL
    frameTituloGeneral = Frame(ConfHoja, bg ='#EC7063')
    frameTituloGeneral.grid(column = 2, row = 0,columnspan= 6, sticky='ew')
    labelTituloGeneral = Label(frameTituloGeneral, text = 'Configuracion General', font= (myFont,18), bg ='#EC7063')
    labelTituloGeneral.grid(column = 2, row = 0, sticky = 'w')
    labelHoja = Label(ConfHoja,text='Tipo de hoja: ', font=(myFont,11), bg='#F1948A')
    labelHoja.grid(column = 2, row= 1,padx= 5, pady = 5, sticky='w')
    labelTituloCampos = Label(ConfHoja, text = 'Campos: ', font= (myFont,16), bg ='#EC7063')
    labelTituloCampos.grid(column = 2 , row = 4, sticky='w')

    # BOX PARA LA HOJA
    tipoHoja = ['A4', 'A5']
    boxtipoHoja = Combobox(ConfHoja,values = tipoHoja, width=3)
    boxtipoHoja.grid(column = 3, row= 1,padx= 5, pady = 5, sticky='w')
    #Canvas.create_line(15, 25, 200, 25) # para hacer una linea abajo del titulo

    # LABELS DE TITULO
        #RECETA
    labelTitulo_receta = Label(ConfHoja,text = 'Receta', font=(myFont,14), bg='#F1948A')
    labelTitulo_receta.grid(column = 3,row = 4,padx= 5, pady = 5, sticky='w')
        #MEDICO
    labelTitulo_medico = Label(ConfHoja,text = 'Medico', font=(myFont,14), bg='#F1948A')
    labelTitulo_medico.grid(column = 4,row = 4,padx= 5, pady = 5, sticky='w')
        #FORMULA
    labelTitulo_formula = Label(ConfHoja,text = 'Formula', font=(myFont,14), bg='#F1948A')
    labelTitulo_formula.grid(column = 5,row = 4,padx= 5, pady = 5, sticky='w')

    # LABELS GENERALES PARA LAS TRES CASILLAS
    # Tamanio letra
    labelLetra_receta = Label(ConfHoja, text = 'Tamaño de letra: ',font=(myFont,11),bg='#F1948A' )
    labelLetra_receta.grid(column = 2, row= 6,padx= 5, pady = 5, sticky='w')
    # Orientacion vertical
    labelVori_receta = Label(ConfHoja, text='Orientacion vertical:',font=(myFont,11), bg='#F1948A')
    labelVori_receta.grid(column = 2, row= 8,padx= 5, pady = 5, sticky='w')
    # Orientacion horizontal
    labelHori_receta = Label(ConfHoja, text='Orientacion horizontal:',font=(myFont,11), bg='#F1948A')
    labelHori_receta.grid(column = 2, row= 10,padx= 5, pady = 5, sticky='w')
    # Ajustar texto
    labelA_texto_receta = Label(ConfHoja, text='Ajustar Texto:',font=(myFont,11), bg='#F1948A')
    labelA_texto_receta.grid(column = 2, row= 12,padx= 5, pady = 5, sticky='w')

    # BOXS NUMERO DE RECETA
        # TEXT FIELD TAMANIO LETRA 
    tmnLetra_receta = Entry(ConfHoja, width=4) 
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
    boxA_texto_receta = Checkbutton(ConfHoja, bg ='#F1948A' , activebackground='#F1948A', activeforeground="white")
    boxA_texto_receta.grid(column = 3, row= 12,padx= 5, pady = 5, sticky='w')

    # BOXS MEDICO
        # TEXT FIELD TAMANIO LETRA 
    tmnLetra_medico = Entry(ConfHoja, width=4) 
    tmnLetra_medico.grid(column = 4, row= 6,padx= 5, pady = 5, sticky='w')
        # BOX ORIENTACION VERTICAL DE TEXTO
    ori_vert_medico = ['Arriba', 'Centro', 'Abajo']
    boxOri_vert_medico = Combobox(ConfHoja, values=ori_vert_receta, width = 8)
    boxOri_vert_medico.grid(column = 4, row= 8,padx= 5, pady = 5, sticky='w')
        # BOX ORIENTACION HORIZONTAL  DE TEXTO
    ori_hori_medico = ['Izquierda', 'Centro', 'Derecha']
    boxOri_hori_medico = Combobox(ConfHoja, values=ori_hori_receta, width = 8)
    boxOri_hori_medico.grid(column = 4, row= 10,padx= 5, pady = 5, sticky='w')
        # BOX AJUSTAR TEXTO 
    boxA_texto_medico = Checkbutton(ConfHoja, bg ='#F1948A' , activebackground='#F1948A', activeforeground="white")
    boxA_texto_medico.grid(column = 4, row= 12,padx= 5, pady = 5, sticky='w')

    # BOXS FORMULAS
        # TEXT FIELD TAMANIO LETRA 
    tmnLetra_formula = Entry(ConfHoja ,width=4) 
    tmnLetra_formula.grid(column = 5, row= 6,padx= 5, pady = 5, sticky='w')
        # BOX ORIENTACION VERTICAL DE TEXTO
    ori_vert_formula = ['Arriba', 'Centro', 'Abajo']
    boxOri_vert_formula = Combobox(ConfHoja, values=ori_vert_receta, width = 8)
    boxOri_vert_formula.grid(column = 5, row= 8,padx= 5, pady = 5, sticky='w')
        # BOX ORIENTACION HORIZONTAL  DE TEXTO
    ori_hori_formula = ['Izquierda', 'Centro', 'Derecha']
    boxOri_hori_formula = Combobox(ConfHoja, values=ori_hori_receta, width = 8)
    boxOri_hori_formula.grid(column = 5, row= 10,padx= 5, pady = 5, sticky='w')
        # BOX AJUSTAR TEXTO 
    boxA_texto_formula = Checkbutton(ConfHoja, bg ='#F1948A' , activebackground='#F1948A', activeforeground="white")
    boxA_texto_formula.grid(column = 5, row= 12,padx= 5, pady = 5, sticky='w')

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
            float(tmnLetra_medico.get())  
            float(tmnLetra_formula.get())  
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
    savebtn = Button(ConfHoja, image = pngGuardar, command = guardar, bg='#F1948A',activebackground= '#F1948A' , borderwidth=0)
    #state=DISABLED Esto bloquea el boton, estaria bueno agregarlo cuando no hay nada tipeado
    savebtn.grid(column = 6, row = 12) # Chequear lo de como se muestra

btnOptions = Button(root, text = 'Configuracion de hoja', command = abrir)
btnOptions.grid(column = 0, row = 0)
#boxOptions = Combobox(root, values = btnOptions).pack()

root.mainloop()





