import xlsxwriter
from tkinter import *
from tkinter.ttk import Combobox
from pathlib import Path
from xlsxwriter import *
from functools import partial
 
workbook = xlsxwriter.Workbook('Prueba.xlsx') #agregar path
worksheet = workbook.add_worksheet()
window = Tk()
window.title('Configuacion de hoja')
window.geometry('300x200+10+20')


# CREO LOS LABELS QUE VA A TENER A LA IZQUIERDA DEL DROPBOX
Hoja = Label(window,text='Tamaño de hoja: ', font=('Calibri',12))
Hoja.place(x = 20, y = 20)
Letra = Label(window, text = 'Tamaño de letra: ',font=('Calibri',12) )
Letra.place(x = 20, y = 50)
Vorientacion = Label(window, text='Orientacion vertical:',font=('Calibri',12))
Vorientacion.place(x=20,y=80)
Horientacion = Label(window, text='Orientacion horizontal:',font=('Calibri',12))
Horientacion.place(x=20,y=110)

# CREO EL DROPBOX PARA LA HOJA
formatoHoja = ['A4', 'A5']
lbformato = Combobox(window,values = formatoHoja, width=3)
lbformato.place(x=180,y=23)

# TEXT FIELD PARA LA LETRA 
NumLetra = Entry(window, text="14",width=3) 
NumLetra.place(x=180, y=53)

# DROPBOX PARA ORIENTACIONES
Overt = ['Arriba', 'Centro', 'Abajo']
lbOvert = Combobox(window, values=Overt, width = 8)
lbOvert.place(x=180,y=83)

Ohori = ['Izquierda', 'Centro', 'Derecha']
lbOhori = Combobox(window, values=Ohori, width = 8)
lbOhori.place(x=180,y=113)

# BOTON DE GUARDADO  
def guardar (): # guarda en la lista y cierra la ventana
    thoja = lbformato.get()
    tletra = NumLetra.get()
    orivert = lbOvert.get()
    orihor = lbOhori.get()

    window.destroy() # agregar pop up guardado exitosamente mbox.show 

    if thoja == 'A4':
        worksheet.set_paper(9)
    elif thoja == 'A5':
        worksheet.set_paper(11)
    
savebtn = Button(window, text = 'Guardar', command=guardar)
savebtn.place(x= 220,y=160)

# MARGENES PARA LA HOJA
worksheet.set_margins(left=0.0, right=0.0, top=0.0, bottom=0.0)
worksheet.center_horizontally()
worksheet.center_vertically()

window.mainloop()





