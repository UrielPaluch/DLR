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
 
workbook = xlsxwriter.Workbook('C:/Users/Pedri/Desktop/Prueba.xlsx') 
worksheet = workbook.add_worksheet()
window = Tk()
window.title('Configuacion de hoja')
window.geometry('300x200+10+20')
window.configure(background='#ADD8E6')
window.iconbitmap('C:/Users/Pedri/Desktop/Zurich/DLR/Iconos/icono.ico')
myFont = font.Font(family='Nunito')

# CREO LOS LABELS QUE VA A TENER A LA IZQUIERDA DEL DROPBOX
Hoja = Label(window,text='Tamaño de hoja: ', font=(myFont,11), bg='#ADD8E6')
Hoja.place(x = 20, y = 20)
Letra = Label(window, text = 'Tamaño de letra: ',font=(myFont,11),bg='#ADD8E6' )
Letra.place(x = 20, y = 50)
Vorientacion = Label(window, text='Orientacion vertical:',font=(myFont,11), bg='#ADD8E6')
Vorientacion.place(x=20,y=80)
Horientacion = Label(window, text='Orientacion horizontal:',font=(myFont,11), bg='#ADD8E6')
Horientacion.place(x=20,y=110)

# CREO EL DROPBOX PARA LA HOJA
formatoHoja = ['A4', 'A5']
lbformato = Combobox(window,values = formatoHoja, width=3)
lbformato.place(x=180,y=23)

# TEXT FIELD PARA LA LETRA 
NumLetra = Entry(window, text = '14' ,width=4) 
NumLetra.place(x=180, y=53)


# DROPBOX PARA ORIENTACIONES DE TEXTO
Overt = ['Arriba', 'Centro', 'Abajo']
lbOvert = Combobox(window, values=Overt, width = 8)
lbOvert.place(x=180,y=83)

Ohori = ['Izquierda', 'Centro', 'Derecha']
lbOhori = Combobox(window, values=Ohori, width = 8)
lbOhori.place(x=180,y=113)

# BOTON DE GUARDADO  
def guardar (): # guarda en la lista y cierra la ventana

    lbformato.get() # Tipo de hoja
    #NumLetra.get() # tamanio de letra
    lbOvert.get() # orientacion vertical
    lbOhori.get() # orientacion horizontal
# que vuelva a mostar la entrada y no corte
    try:
        float(NumLetra.get())   
    except ValueError:    
        tkinter.messagebox.showerror('Error', 'Ingrese un numero')


    if lbformato.get() == 'A4':
        worksheet.set_paper(9)
    elif lbformato.get() == 'A5':
        worksheet.set_paper(11)

    tkinter.messagebox.showinfo('Guarado', 'La configuracion ha sido guardada correctamente')
    window.destroy() # cierra la ventana y guarda


pngGuardar = PhotoImage(file='C:/Users/Pedri/Desktop/Zurich/DLR/Iconos/discket.png')
# Esto permite poner la imagen como boton
savebtn = Button(window, image = pngGuardar, command=guardar, bg='#ADD8E6',activebackground= '#ADD8E6' , borderwidth=0)
'''state=DISABLED Esto bloquea el boton, estaria bueno agregarlo cuando no hay nada tipeado''' 
savebtn.place(x= 240,y=155)

# MARGENES PARA LA HOJA
worksheet.set_margins(left=0.0, right=0.0, top=0.0, bottom=0.0)
worksheet.center_horizontally()
worksheet.center_vertically()

window.mainloop()





