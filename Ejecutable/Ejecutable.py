from tkinter import *
from datetime import timedelta, date, datetime
from tkcalendar import *
from tkinter import messagebox as m_box
import pymysql
import re
import pandas as pd
from tkinter import filedialog
import tkinter.font as font
#Comento los unused imports, si anda bien sin esto los borro
#from tkinter import ttk
#from tkinter.font import Font, nametofont
#import os
#import sys
import pandas.io.formats.excel

#Inicializo las variables para que puedan ser globales
pantalla = ''
idDCI = []
#--------------------------------------------------------------------------------------------------#
#Autocomplete
class AutocompleteEntry(Entry):
    global myFont
    global pantalla

    def __init__(self, lista, *args, **kwargs):
        Entry.__init__(self, *args, **kwargs)
        self.lista = lista        
        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = StringVar()
        #self.bind("<KeyRelease>", self.caps)
        self.var.trace('w', self.changed)
        self.bind("<Right>", self.selection)
        self.bind("<Up>", self.up)
        self.bind("<Down>", self.down)
        #Tecla enter
        self.bind("<Return>", self.selection)
        #Tecla tab
        self.bind("<Tab>", self.end)
        self.lb_up = False
        
    def changed(self, name, index, mode):  
        
        if self.var.get() == '':
            self.lb.destroy()
            self.lb_up = False
        else:
            words = self.comparison()
            if words:            
                if not self.lb_up:
                    #El width 40 es para que me ocupe el ancho del entry
                    self.lb = Listbox(width = 40)
                    if pantalla == 'root':
                        self.lb = Listbox(root, width = 40)
                    if pantalla == 'abmMedicosFrame':
                        self.lb = Listbox(abmMedicosFrame, width = 40)
                    if pantalla == 'altaFormulaFrame':
                        self.lb = Listbox(frameFormula, width = 40)
                    self.lb['font'] = myFont
                    self.lb.bind("<Double-Button-1>", self.selection)
                    self.lb.bind("<Right>", self.selection)
                    self.lb.place(x=self.winfo_x(), y=self.winfo_y()+self.winfo_height())
                    self.lb_up = True
                self.lb.delete(0, END)
                for w in words:
                    self.lb.insert(END,w)
            else:
                if self.lb_up:
                    self.lb.destroy()
                    self.lb_up = False
        
    def selection(self, event):
        if self.lista != lista_medicos:
            #Con esto consigo el ID que corresponde al DCI. 
            global idDCI 
            idDCI.append(self.lista.index(self.lb.get(ACTIVE)))
            print(idDCI)

        if self.lb_up:
            self.var.set(self.lb.get(ACTIVE))
            self.lb.destroy()
            self.lb_up = False
            self.icursor(END)
        
        
    
    def end(self, event):
        if self.lb_up:
            self.lb.destroy()
            self.lb_up = False
            self.icursor(END)
    
    def up(self, event):
        global primerIndex
        global ultimoIndex
        
        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != '0':
                self.lb.selection_clear(first=index)
                index = str(int(index)-1)                
                self.lb.selection_set(first=index)
                self.lb.activate(index)
                if int(index) < primerIndex:
                    primerIndex = primerIndex - 1
                    self.lb.yview_scroll(-1, "units")
                    ultimoIndex -= 1
                if int(index) == -1:
                    primerIndex = 0
                    ultimoIndex = 9



    def down(self, event):
        global ultimoIndex
        global primerIndex

        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != END:                        
                self.lb.selection_clear(first=index)
                index = str(int(index)+1)        
                self.lb.selection_set(first=index)
                self.lb.activate(index)
                if int(index) > ultimoIndex:
                    ultimoIndex = ultimoIndex + 1
                    self.lb.yview_scroll(1, "units")
                    primerIndex += 1
                if int(self.lb.size()) == int(index):
                    self.lb.yview_scroll(-primerIndex, "units")
                    primerIndex = 0
                    ultimoIndex = 9

    def comparison(self):
        global ultimoIndex
        global primerIndex
        ultimoIndex = 9
        primerIndex = 0
        pattern = re.compile('.*' + self.var.get() + '.*')
        return [w for w in self.lista if re.match(pattern, w)]

    #Es para las mayusculas, pero no se porque no anda
    '''
    def caps(self, event):
        self.var.set(self.var.get().upper())
    '''
    def endPublico(self):
        self.lb.destroy()
        self.icursor(END)
    
#--------------------------------------------------------------------------------------------------#
#Base de datos
class DataBase():
    def __init__(self):
        try:
            self.connection = pymysql.connect(host='localhost', user = 'root', password='saN#jOy9*8zT', db='psicotropicos')
            self.cursor = self.connection.cursor()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo establecer la conexion con la base de datos ' + str(e))
    
    def traigoFormulas(self):
        sql = 'SELECT formula FROM tabla_formula'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def traigoDCIformulas(self):
        sql = 'SELECT DCI FROM tabla_formula'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

    def traigoDCIformulas3(self):
        sql = 'SELECT DCI FROM tabla_formula3'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

    def traigoFormulas3(self):
        sql = 'SELECT formula3 FROM tabla_formula3'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def traigoMedicos(self):
        sql = 'SELECT nombre FROM tabla_medicos'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

    def traigoPsico(self, desde, hasta):
        sql = 'SELECT * FROM tabla_psico WHERE (control <= (%s)  AND control >= (%s))'

        try:
            self.cursor.execute(sql, (hasta, desde))
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

    def traigoPsico3(self, desde, hasta):
        sql = 'SELECT * FROM tabla_psico3  WHERE id_receta <= (%s) AND id_receta >= (%s) GROUP BY control'

        try:
            self.cursor.execute(sql, (hasta, desde))
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def traigoUnaFormula(self, idreceta):
        sql = 'SELECT * FROM tabla_psico WHERE id_receta = (%s)'

        try:
            self.cursor.execute(sql, (idreceta))
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

    def grabarFormula(self, formula, dci = NONE):
        global dfFormula
        lista_formulas.append(formula)
        lista_DCIformulas.append(dci)
        dfFormula = pd.DataFrame(lista_formulas, columns=['Formula'])

        sql = 'INSERT INTO tabla_formula (formula, DCI) VALUES (%s, %s)'

        try:
            self.cursor.execute(sql, (formula, dci))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def grabarFormula3(self, formula):
        global dfFormula3
        lista_formulas3.append(formula)
        dfFormula3 = pd.DataFrame(lista_formulas3, columns=['Formula3'])

        sql = 'INSERT INTO tabla_formula3 (formula3) VALUES (%s)'

        try:
            self.cursor.execute(sql, (formula))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def ultimoID(self):
        sql = 'SELECT id_receta FROM tabla_psico  ORDER BY id_receta DESC LIMIT 1'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def ultimaReceta(self):
        sql = 'SELECT * FROM tabla_psico ORDER BY ignorar DESC LIMIT 1'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def graboMedicamento(self, id, nombre, fecha, formula, cant, control):
        sql = 'INSERT INTO tabla_psico (id_receta, nombre_medico, fecha, formula, cant, control) VALUES (%s, %s, %s, %s, %s, %s)'

        try:
            self.cursor.execute(sql, (id, nombre, fecha, formula, cant, control))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

    def graboMedicamento3(self, id, nombre, fecha, formula, cant, control):
        sql = 'INSERT INTO tabla_psico3 (id_receta, nombre_medico, fecha, formula, cant, control) VALUES (%s, %s, %s, %s, %s, %s)'

        try:
            self.cursor.execute(sql, (id, nombre, fecha, formula, cant, control))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

    def graboMedico(self, nombre):
        sql = 'INSERT INTO tabla_medicos (nombre) VALUES (%s)'

        try:
            self.cursor.execute(sql, (nombre))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def sinMovimiento(self, fecha, n_control):
        sql = 'INSERT INTO tabla_psico (id_receta, nombre_medico, fecha, formula, cant, control) VALUES (%s, %s, %s, %s, %s, %s)'

        try:
            self.cursor.execute(sql, (None, 'SIN MOVIMIENTO', fecha, None, None, n_control))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def borroMedico(self, nombre):
        global dfMedicos
        lista_medicos.remove(nombre)
        dfMedicos = pd.DataFrame(lista_medicos, columns=['Medicos'])

        sql = 'DELETE FROM tabla_medicos WHERE nombre = (%s)'

        try:
            self.cursor.execute(sql, (nombre))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def borroPsico3(self, control):
        sql = 'DELETE FROM tabla_psico3 WHERE control = (%s)'

        try:
            self.cursor.execute(sql, (control))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def borroFormula(self, formula):
        global dfFormula
        lista_formulas.remove(formula)
        dfFormula = pd.DataFrame(lista_formulas, columns=['Formula'])

        sql = 'DELETE FROM tabla_formula WHERE formula = (%s)'

        try:
            self.cursor.execute(sql, (formula))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def borroFormula3(self, formula):
        global dfFormula3
        lista_formulas3.remove(formula)
        dfFormula3 = pd.DataFrame(lista_formulas3, columns=['Formula'])

        sql = 'DELETE FROM tabla_formula3 WHERE formula3 = (%s)'

        try:
            self.cursor.execute(sql, (formula))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def borrarReceta(self, id_borrar):
        sql = 'DELETE FROM tabla_psico WHERE control = (%s)'

        try:
            self.cursor.execute(sql, (id_borrar))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

    def borrarReceta3(self, id_borrar):
        sql = 'DELETE FROM tabla_psico3 WHERE control = (%s)'

        try:
            self.cursor.execute(sql, (id_borrar))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

    def borrarUltimaReceta(self):
        ultima = database.ultimaReceta()
        if ultima[0][2] == 'SIN MOVIMIENTO':
            sql = 'DELETE FROM tabla_psico ORDER BY ignorar DESC LIMIT 1'
            
            try:
                self.cursor.execute(sql)
                self.connection.commit()
            except Exception as e:
                m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
            
            return(True)

        sql = 'DELETE FROM tabla_psico WHERE id_receta = (%s)'
        sql2 = 'DELETE FROM tabla_psico3 WHERE id_receta = (%s)'
        try:
            self.cursor.execute(sql, (ultimoID))
            self.cursor.execute(sql2, (ultimoID))
            self.connection.commit()
            return(False)
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def actualizoPsico(self, nombre, formula, cant, control):
        sql = 'UPDATE tabla_psico set formula = (%s), nombre_medico = (%s), cant = (%s) WHERE control = (%s)'

        try:
            self.cursor.execute(sql, (formula, nombre, cant, control))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def updatePsico3(self, nombre, formula, cant, control):
        sql = 'UPDATE tabla_psico3 set formula = (%s), nombre_medico = (%s), cant = (%s) WHERE control = (%s)'

        try:
            self.cursor.execute(sql, (formula, nombre, cant, control))
            self.connection.commit()
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def backupFormula(self):
        sql = 'SELECT * FROM tabla_formula'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def backupFormula3(self):
        sql = 'SELECT * FROM tabla_formula3'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def backupMedicos(self):
        sql = 'SELECT * FROM tabla_medicos'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def backupTablaPsico(self):
        sql = 'SELECT * FROM tabla_psico'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))
    
    def backupTablaPsico3(self):
        sql = 'SELECT * FROM tabla_psico3'

        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror('Error', 'No se pudo ejecutar la consulta ' + str(e))

#--------------------------------------------------------------------------------------------------#
#Base de datos web
class DataBaseWeb():
    def __init__(self):
        try:
            #DE ACA NO HAY QUE CAMBIAR NADA.
            self.connection = pymysql.connect(host='190.106.131.222', user = 'c17790_uriel', password='saN#jOy9*8zT', db='c17790_usuarios')
            self.cursor = self.connection.cursor()
        except Exception as e:
            m_box.showerror("Error al iniciar", "Fallo la conexión: " + str(e))
    
    def traigoUsuarios(self):
        sql = "SELECT * FROM usuarios WHERE Usuario = 'Agnese'"
        try:
            self.cursor.execute(sql)
            return(self.cursor.fetchall())
        except Exception as e:
            m_box.showerror("Error al iniciar", "Comuniquese con su distribuidor: " + str(e))

#--------------------------------------------------------------------------------------------------#
#Boton
class HoverButton(Button):
    def __init__(self, master, **kw):
        Button.__init__(self,master=master,**kw)
        self.defaultBackground = self["background"]
        self["borderwidth"] = 0
        self.bind("<FocusIn>", self.on_enter)
        self.bind("<FocusOut>", self.on_leave)

    def on_enter(self, e):
        self['borderwidth'] = 1

    def on_leave(self, e):
        self['borderwidth'] = 0
#--------------------------------------------------------------------------------------------------#
#On hover 
class ToolTip(object):

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 57
        y = y + cy + self.widget.winfo_rooty() +27
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(tw, text=self.text, justify=LEFT,
                      background="#ffffe0", relief=SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)
    
    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)
#--------------------------------------------------------------------------------------------------#
database = DataBase()
ultimoID1 = database.ultimoID()

#Conexion a la data base web
databaseweb = DataBaseWeb()
def chequeo_pago():
    usuarioWeb = databaseweb.traigoUsuarios()
    try:
        pago = usuarioWeb[0][2]
    except Exception as e:
        m_box.showerror("Error al iniciar", "Error DBW1: " + str(e))
    #Si no pago
    if (pago != "Si"):
        #Se cierra
        m_box.showerror("Error al iniciar", "Falta de pago, comuniquese con su distribuidor")
        #La idea es qeu tire error y se cierre
        root.destroy()
chequeo_pago()

formulas = database.traigoFormulas()
#Creo una lista bien con todas las formulas
lista_formulas = []
for i in range(0, len(formulas),1):
    lista_formulas.append(formulas[i][0])
#Creo un dataframe
dfFormula = pd.DataFrame(lista_formulas, columns=['Formula'])

DCIformulas = database.traigoDCIformulas()
#Creo una lista bien con los DCI
lista_DCIformulas = []
for i in range(0, len(DCIformulas),1):
    lista_DCIformulas.append(DCIformulas[i][0])
#Creo un dataframe
dfDCIFormula = pd.DataFrame(lista_DCIformulas, columns=['DCIFormula'])

DCIformulas3 = database.traigoDCIformulas3()
#Creo una lista bien con los DCI
lista_DCIformulas3 = []
for i in range(0, len(DCIformulas3),1):
    lista_DCIformulas3.append(DCIformulas3[i][0])
#Creo un dataframe
dfDCIFormula3 = pd.DataFrame(lista_DCIformulas3, columns=['DCIFormula'])

formulas3 = database.traigoFormulas3()
#Creo una lista bien con todas las formulas
lista_formulas3 = []
for i in range(0, len(formulas3),1):
    lista_formulas3.append(formulas3[i][0])
#Creo un dataframe
dfFormula3 = pd.DataFrame(lista_formulas3, columns=['Formula3'])

medicos = database.traigoMedicos()
#Creo una lista bien con todas los medicos
lista_medicos = []
for i in range(0, len(medicos),1):
    lista_medicos.append(medicos[i][0])
#lista_medicos_aux = lista_medicos
dfMedicos = pd.DataFrame(lista_medicos, columns=['Medicos'])

#Lo devuelve como si fuera un diccionario, por eso el 0, 0
ultimoID = ultimoID1[0][0]
def funUltimaFecha():
    global ultimaFecha
    #Pido la última fecha de la última receta
    ultimaFecha = database.ultimaReceta()
    ultimaFecha = ultimaFecha[0][3]
    ultimaFecha = datetime.strptime(ultimaFecha, '%d-%m-%Y')

#Genero el cuadro para que el usuario pueda pedir la información que desea
#Creo la raiz
root = Tk() 

#Cambia el título
root.title("Libro Recetario Digital - Zurich Software")

#Para que se abra en el centro
ws = root.winfo_screenwidth()
hs = root.winfo_screenheight()
x = ws/18
y = hs/10
root.geometry('+%d+%d' % (x, y))

#Cambio el icono
root.iconbitmap("Iconos/icono.ico")

#Navigation bar
menubar = Menu(root)
root.config(menu=menubar)

#Creo el objeto font
myFont = font.Font(family='Nunito')

#tearoff = 0 es para que no aparezca el de ejemplo
#Instancio la barra
filemenu = Menu(menubar, tearoff=0)

#No se puede agregar imagenes al menubar
menubar.add_cascade(label="Opciones", menu=filemenu)

filemenu.add_command(label="Fecha", command = lambda: seleccionoFecha(), font=("Calibri", 11))
filemenu.add_command(label="Alta fórmula", command = lambda: pantallaFormula(), font=("Calibri", 11))
filemenu.add_command(label="Alta médico", command = lambda: pantallaMedicos(), font=("Calibri", 11))
filemenu.add_command(label="Reporte L. Recetario", command = lambda: pantallaPsico(), font=("Calibri", 11))
filemenu.add_command(label="Reporte para L. Contralor", command = lambda: pantallaPsico3(), font=("Calibri", 11))
filemenu.add_command(label="Backup", command = lambda: backup(), font=("Calibri", 11))

#--------------------------------------------------------------------------------------------------#
#Pantalla selecciono fecha

def guardoFecha(fecha):
    varFecha.set(fecha)

    #Cada vez que cambie de fecha chequeo que haya pagado.
    #De esta forma me aseguro que este conectado a internet y no hace lenta la carga.
    chequeo_pago()

def seleccionoFecha():
    global Calendario
    global myFont
    global ultimaFecha

    #Creo el frame fecha
    fechaFrame = Toplevel()

    #Para fijar el tamaño
    fechaFrame.geometry("360x300")

    #Cambio el icono
    fechaFrame.iconbitmap("Iconos/icono.ico")

    #Cambio el colo de fondo
    fechaFrame.configure(background='#cfcfcf')

    #Para que no se pueda cambiar el tamaño de la pantalla
    fechaFrame.resizable(False, False)    

    #Texto seleccione fecha
    textoLabel = Label(fechaFrame, text="Seleccione la fecha de carga", font=18, bg ='#9e9e9e')
    #Selecciono el tipo de letra
    textoLabel['font'] = myFont
    textoLabel.pack(fill = X)

    diaHoy = ultimaFecha
    #Calendario
    Calendario = Calendar(fechaFrame, font=(11), selectmode="day", year = diaHoy.year,  month = diaHoy.month, day = diaHoy.day, date_pattern = 'dd-mm-y')
    Calendario.pack(pady = 10)
    
    #Para que se muestre la imagen
    img_discket = PhotoImage(file='Iconos/discket.png')

    #Boton grabar
    btnGrabarFecha = HoverButton(fechaFrame, activebackground='#cfcfcf', bg = '#cfcfcf', image = img_discket, text = "Grabar", font=(12), command = lambda:guardoFecha(Calendario.get_date()))
    btnGrabarFecha.pack(side = RIGHT, padx = 5)
    #Si no pongo esto no se muestra
    btnGrabar.image = img_discket
    CreateToolTip(btnGrabarFecha, text = "Grabar")

#--------------------------------------------------------------------------------------------------#
#Pantalla ABM medicos
def pantallaMedicos():
    global abmMedicosFrame
    global pantalla

    #Creo el frame
    abmMedicosFrame = Toplevel()

    #Cambio el colo de fondo
    abmMedicosFrame.configure(background='#e087b0')

    #Selecciono el icono
    abmMedicosFrame.iconbitmap("Iconos/icono.ico")

    #Selecciono el tamaño de la pantalla
    abmMedicosFrame.geometry("460x340")

    #Para que no se pueda cambiar el tamaño de la pantalla
    abmMedicosFrame.resizable(False, False)  

    #Texto seleccione fecha
    textoLabel = Label(abmMedicosFrame, text="Ingrese el nombre del médico", font=18, bg ='#d65891')
    textoLabel.pack(fill = X)
    textoLabel['font'] = myFont
    
    def cambioabmMedicosFrame(event):
        global pantalla
        pantalla = 'abmMedicosFrame'
    
    #Entry formula
    entryNombreMedico = AutocompleteEntry(lista_medicos, abmMedicosFrame, width = 40)
    entryNombreMedico.pack(pady = 15)
    entryNombreMedico['font'] = myFont
    entryNombreMedico.bind('<KeyRelease>', cambioabmMedicosFrame)
    entryNombreMedico.bind('<Button-1>', cambioabmMedicosFrame)

    #Para que se muestre la imagen
    img_discket = PhotoImage(file='Iconos/discket.png')

    #Boton grabar
    btnGrabar = HoverButton(abmMedicosFrame, activebackground='#e087b0', bg = '#e087b0', font=12, image = img_discket, text="Grabar")
    btnGrabar.pack(side = RIGHT, padx = 10, anchor = SW)
    btnGrabar['borderwidth'] = 0
    #Si no pongo esto la imagen no se muestra
    btnGrabar.image = img_discket
    #Onhover
    CreateToolTip(btnGrabar, text = "Grabar")

    def clickerGrabar(event):
        filtro2 = existeMedico(str(entryNombreMedico.get()))
        if filtro2 == False:
            m_box.showinfo('Aviso!', 'Archivo guardado exitosamente')
        else:
            m_box.showerror('Error', 'Ya existe un médico con ese nombre')
        
        #Pongo los campos vacios
        entryNombreMedico.delete(0, 'end')
        
        #Para que el cursor se ponga en el Nombre del médico
        entryNombreMedico.focus_set()

    btnGrabar.bind("<Right>", clickerGrabar)
    #Click izquierdo
    btnGrabar.bind("<Button-1>", clickerGrabar)
    #Tecla enter
    btnGrabar.bind("<Return>", clickerGrabar)

    #Boton borrar
    #Para que se muestre la imagen
    img_delete = PhotoImage(file='Iconos/delete.png')

    #Boton grabar
    btnBorrar = HoverButton(abmMedicosFrame, activebackground='#e087b0', bg = '#e087b0', font=12, image = img_delete, text="Delete")
    btnBorrar.pack(side = LEFT, padx = 10, anchor = SW)
    btnBorrar['borderwidth'] = 0
    #Si no pongo esto la imagen no se muestra
    btnBorrar.image = img_delete
    #Onhover
    CreateToolTip(btnBorrar, text = "Borrar")

    def clickerBorrar(event):
        global dfMedicos
        filtro = dfMedicos[dfMedicos['Medicos'].str.contains(str(entryNombreMedico.get()).upper())]
        filtro2 = valildarFiltro2(filtro, str(entryNombreMedico.get()).upper(), 'Medicos')
        if filtro2 == True:
            try:
                database.borroMedico(str(entryNombreMedico.get()).upper())
                m_box.showinfo('Aviso!', 'Archivo borrado exitosamente')
            except Exception as e:
                m_box.showerror('Error', 'El siguiente error ocurrio al borrar el archivo ' + str(e))
        else:
            m_box.showerror('Error', 'No existe un médico con ese nombre')
        #Pongo los campos vacios
        entryNombreMedico.delete(0, 'end')

        #Para que el cursor se ponga en el Nombre del médico
        entryNombreMedico.focus_set()

    btnBorrar.bind("<Right>", clickerBorrar)
    #Click izquierdo
    btnBorrar.bind("<Button-1>", clickerBorrar)
    #Tecla enter
    btnBorrar.bind("<Return>", clickerBorrar)

#--------------------------------------------------------------------------------------------------#
#Pantalla cargo medicamento
def grabarMedicamento():
    global ultimoID
    global dfFormula3
    global ultimaFecha
    global idDCI

    try:
        int(cant1.get())
        int(cant2.get())
        int(cant3.get())
    except Exception:
        m_box.showerror('Error', 'Ingrese un número válido')
        return()

    #Para que no me pueda ingresar fechas salteadas
    filtroFecha = validoFecha()
    if filtroFecha == False:
        m_box.showerror('Error', 'No es posible ingresar una receta con esa fecha')
        return()

    if entryNombre.get() == '':
        m_box.showerror('Error', 'Ingrese un nombre')
        return()

    ultimoID += 1

    if cant1.get() != '0':
        
        control = str(ultimoID) + '1'
        control = int(control)

        if str(lista_DCIformulas[idDCI[0]]) != "":
            database.graboMedicamento(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), str(lista_DCIformulas[idDCI[0]]+ ", " + entryFormula1.get()), cant1.get(), control)
        else:
            database.graboMedicamento(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), entryFormula1.get(), cant1.get(), control)

        filtro3 = dfFormula3[dfFormula3['Formula3'].str.contains(entryFormula1.get())]
        
        filtro4 = valildarFiltro2(filtro3, entryFormula1.get(), 'Formula3')
        if filtro4 == True:
            if str(lista_DCIformulas[idDCI[0]]) != "":
                database.graboMedicamento3(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), str(lista_DCIformulas[idDCI[0]]+ ", " + entryFormula1.get()), cant1.get(), control)
            else:
                database.graboMedicamento3(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), entryFormula1.get(), cant1.get(), control)
    else:
        m_box.showerror('Error', "Ingrese una cantidad")
        return()
        
    if cant2.get() != '0':

        control = str(ultimoID) + '2'
        control = int(control)

        if str(lista_DCIformulas[idDCI[1]]) != "":
            database.graboMedicamento(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), str(lista_DCIformulas[idDCI[1]]+ ", " + entryFormula2.get()), cant2.get(), control)
        else:
            database.graboMedicamento(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), entryFormula2.get(), cant2.get(), control)

        filtro3 = dfFormula3[dfFormula3['Formula3'].str.contains(entryFormula2.get())]
        
        filtro4 = valildarFiltro2(filtro3, entryFormula2.get(), 'Formula3')
        if filtro4 == True:
            if str(lista_DCIformulas[idDCI[1]]) != "":
                database.graboMedicamento3(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), str(lista_DCIformulas[idDCI[1]]+ ", " + entryFormula2.get()), cant2.get(), control)
            else:
                database.graboMedicamento3(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), entryFormula2.get(), cant2.get(), control)

    if cant3.get() != '0':

        control = str(ultimoID) + '3'
        control = int(control)

        if str(lista_DCIformulas[idDCI[2]]) != "":
            database.graboMedicamento(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), str(lista_DCIformulas[idDCI[2]]+ ", " + entryFormula3.get()), cant3.get(), control)
        else:
            database.graboMedicamento(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), entryFormula3.get(), cant3.get(), control)
        
        filtro3 = dfFormula3[dfFormula3['Formula3'].str.contains(entryFormula3.get())]
        
        filtro4 = valildarFiltro2(filtro3, entryFormula3.get(), 'Formula3')
        if filtro4 == True:
            if str(lista_DCIformulas[idDCI[2]]) != "":
                database.graboMedicamento3(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), str(lista_DCIformulas[idDCI[2]]+ ", " + entryFormula3.get()), cant3.get(), control)
            else:
                database.graboMedicamento3(ultimoID, str(entryNombre.get()).upper(), Calendario.get_date(), entryFormula3.get(), cant3.get(), control)
    
    varIDreceta.set(ultimoID+1)

    ultimaFecha = datetime.strptime(Calendario.get_date(), '%d-%m-%Y')
    
    #Limpio la lista
    idDCI = []
    limpioLosCampos()

def existeMedico(nombre):
    global dfMedicos
    filtro = dfMedicos[dfMedicos['Medicos'].str.contains(nombre)]
    filtro2 = valildarFiltro2(filtro, nombre, 'Medicos')
    if filtro2 == False:
        lista_medicos.append(nombre.upper())
        dfMedicos = pd.DataFrame(lista_medicos, columns=['Medicos'])
        database.graboMedico(nombre.upper())
    return(filtro2)

def existeMedicamento():
    global dfFormula

    if cant1.get() != '0':
        filtro = dfFormula[dfFormula['Formula'].str.contains(entryFormula1.get())]
        
        if len(filtro) != 1:
            filtro2 = valildarFiltro2(filtro, str(entryFormula1.get()), 'Formula')
            
            if filtro2 == False:
                m_box.showerror('Error', "No existe la fórmula " + str(entryFormula1.get()) + ", ingrese otra")
                return(False)
    
    if cant2.get() != '0':
        filtro = dfFormula[dfFormula['Formula'].str.contains(entryFormula2.get())]

        if len(filtro) != 1:
            filtro2 = valildarFiltro2(filtro, str(entryFormula2.get()), 'Formula')
            
            if filtro2 == False:
                m_box.showerror('Error', "No existe la fórmula " + str(entryFormula2.get()) + ", ingrese otra")
                return(False)
    
    if cant3.get() != '0':
        filtro = dfFormula[dfFormula['Formula'].str.contains(entryFormula3.get())]

        if len(filtro) != 1:
            filtro2 = valildarFiltro2(filtro, str(entryFormula3.get()), 'Formula')
            
            if filtro2 == False:
                m_box.showerror('Error', "No existe la fórmula " + str(entryFormula3.get()) + ", ingrese otra")
                return(False)

def valildarFiltro2(filtro, entry, nombreTabla):
    lista_filtro = list(filtro[nombreTabla])
    esta = False

    for i in range(0, len(lista_filtro), 1):
        if lista_filtro[i] == entry:
            esta = True
            return(esta)
    
    return(esta)

def limpioLosCampos():
    #Pongo los campos vacios
    entryNombre.delete(0, 'end')
    entryFormula1.delete(0, 'end')
    entryFormula2.delete(0, 'end')
    entryFormula3.delete(0, 'end')
    entryCant1.delete(0, 'end')
    entryCant2.delete(0, 'end')
    entryCant3.delete(0, 'end')
    entryCant1.insert(0, '1')
    entryCant2.insert(0, '0')
    entryCant3.insert(0, '0')

def cambioPantalla(event):
    global pantalla
    pantalla = 'root'

def validoFecha():
    if ultimaFecha  == datetime.strptime(Calendario.get_date(), '%d-%m-%Y') or ultimaFecha + timedelta(days=1)  == datetime.strptime(Calendario.get_date(), '%d-%m-%Y'):
        return(True)
    else:
        return(False)

def esLista3(formula):
    #Para saber si es de lista 3
    filtro3 = dfFormula3[dfFormula3['Formula3'].str.contains(formula)]
    filtro4 = valildarFiltro2(filtro3, formula, 'Formula3')
    if filtro4 == True:
        return(True)
    else:
        return(False)

def ABMpsico3(formulaLista3, formula, nombreMedico, entryFormula, entryCant, Ncontrol, Nreceta):

    filtro = esLista3(entryFormula)

    #Si la cantidad ingresada es cero, directamente lo borra
    if entryCant == '0':
        #Borro la receta de lista 3
        database.borroPsico3(int(Ncontrol))
        return()

    #Si las dos formulas son iguales no hay que hacer nada
    if entryFormula.upper() == formula:
        return()

    #Si la receta de la base de datos es de lista 3
    if formulaLista3 == True:
        
        #Si la formula ingresada es de lista 3 tambien
        if filtro == True:
            
            database.updatePsico3(nombreMedico, str(entryFormula).upper(), entryCant, int(Ncontrol))
        else:
            #Borro la receta de lista 3
            database.borroPsico3(int(Ncontrol))
    #Si la receta de la base de datos NO es de lista 3
    else:

        #Si la formula ingresada es de lista 3
        if filtro == True:
            database.graboMedicamento3(Nreceta, nombreMedico, fecha, str(entryFormula).upper(), entryCant, int(Ncontrol))

#Para que no se pueda cambiar el tamaño de la pantalla
root.resizable(False, False)

#Tamaño de la pantalla
root.geometry("1252x447")

#Cambio el colo de fondo
root.configure(background='#98abca')

#Creo un frame para el título
frameTitulo = Frame()
frameTitulo.grid(row = 0, column = 0, columnspan = 5, sticky="nsew")
frameTitulo.configure(background='#5071A5')

#Texto titulo
labelTitulo = Label(frameTitulo, text="Ingrese el medicamento", font=(18), bg ='#5071A5')
labelTitulo.pack(fill = X)
labelTitulo['font'] = myFont

#Creo un frame para ubicar el texto id de la receta y la fecha
frameTexto = Frame()
frameTexto.grid(row = 3, column = 2)
frameTexto.configure(background='#98abca')

#Creo un frame para ubicar las variables id de la receta y la fecha
frameVariables = Frame()
frameVariables.grid(row = 4, column = 2)
frameVariables.configure(background='#98abca')

#Texto id_receta
lalbelID_receta = Label(frameTexto, text = "Número de receta", font=(12), bg ='#98abca')
lalbelID_receta.grid(row = 0, column = 0, padx = 60)
lalbelID_receta['font'] = myFont

#Variable id_receta
varIDreceta = IntVar()
varIDreceta.set(ultimoID + 1)

entryIDreceta = Entry(frameVariables, textvariable = varIDreceta, font=12, justify='center')
entryIDreceta.grid(row = 0, column = 0, pady = 10, padx = 10)
entryIDreceta['font'] = myFont

def clickerBuscar(event):
    global control1
    global control2
    global control3
    global formula1
    global formula2
    global formula3
    global formula1Lista3
    global formula2Lista3
    global formula3Lista3
    global fecha
    
    formula2Lista3 = False
    formula3Lista3 = False

    formula2 = ''
    formula3 = ''

    #Valido que el número sea un int
    try:
        int(entryIDreceta.get())
    except:
        m_box.showerror('Error', 'Ingrese un número válido')
        return()

    #Inicializo las variables globales
    control1 = 0
    control2 = 0
    control3 = 0

    #Primero limpio los campos por las dudas
    limpioLosCampos()

    #Traigo los datos
    datos = database.traigoUnaFormula(entryIDreceta.get())
    #Creo un dataframe
    dfDatos = pd.DataFrame(datos, columns = ['ignorar', 'id_receta', 'nombre', 'fecha', 'formula', 'cant', 'control'])
    
    #Saco la columna que no me sirve
    dfDatos.drop('ignorar', axis = 1, inplace = True)

    #Los datos que si o si pido
    #Si no encuentra nombre del médico, significa que no no existe la receta
    try:
        nombre = dfDatos['nombre'][0]
    except:
        m_box.showinfo('Aviso', 'No se encontro la receta con el número: ' + str(entryIDreceta.get()))
        return()

    entryNombre.insert(0, nombre)

    #Para que no se abra el listbox
    AutocompleteEntry.endPublico(entryNombre)

    fecha = dfDatos['fecha'][0]
    varFecha.set(fecha)

    #Depende de cuantos datos tenga el frame, cuanto asigno
    if len(dfDatos) == 1:
        #Pongo manualmente en cero las cantidades
        entryCant1.delete(0, 'end')

        formula1 = dfDatos['formula'][0]
        entryFormula1.insert(0, formula1)

        formula1Lista3 = esLista3(formula1)

        try:
            #Para que no se abra el listbox
            AutocompleteEntry.endPublico(entryFormula1)
        except Exception as e:
            print(e)

        cant1_aux = dfDatos['cant'][0]
        entryCant1.insert(0, cant1_aux)

        control1 = dfDatos['control'][0]

    if len(dfDatos) == 2:
        #Pongo manualmente las cantidades en cero
        entryCant1.delete(0, 'end')

        formula1 = dfDatos['formula'][0]
        entryFormula1.insert(0, formula1)

        #Para saber si es de lista 3
        formula1Lista3 = esLista3(formula1)

        cant1_aux = dfDatos['cant'][0]
        entryCant1.insert(0, cant1_aux)

        control = dfDatos['control'][1]
        control = str(control)

        formula2 = dfDatos['formula'][1]
        cant2_aux = dfDatos['cant'][1]

        control1 = dfDatos['control'][0]

        #Si el número de control termina en 3
        if (control[-1] == '3'):
            control3 = control
            #Pongo manualmente en cero las cantidades
            entryCant3.delete(0, 'end')
            
            entryFormula3.insert(0, formula2)
            entryCant3.insert(0, cant2_aux)
            
            try:
            #Para que no se abra el listbox
                AutocompleteEntry.endPublico(entryFormula3)
                AutocompleteEntry.endPublico(entryFormula1)
            except e:
                print(e)
        #Si el número de control termina en 2
        else:
            control2 = control
            #Pongo manualmente en cero las cantidades
            entryCant2.delete(0, 'end')

            entryFormula2.insert(0, formula2)
            entryCant2.insert(0, cant2_aux)
            
            try:
            #Para que no se abra el listbox
                AutocompleteEntry.endPublico(entryFormula2)
                AutocompleteEntry.endPublico(entryFormula1)
            except e:
                print(e)


        formula2Lista3 = esLista3(formula2)

        
    if len(dfDatos) == 3:
        #Pongo manualmente en cero las cantidades
        entryCant1.delete(0, 'end')
        entryCant2.delete(0, 'end')
        entryCant3.delete(0, 'end')


        formula1 = dfDatos['formula'][0]
        entryFormula1.insert(0, formula1)

        formula1Lista3 = esLista3(formula1)

        cant1_aux = dfDatos['cant'][0]
        entryCant1.insert(0, cant1_aux)

        formula2 = dfDatos['formula'][1]
        entryFormula2.insert(0, formula2)

        formula2Lista3 = esLista3(formula2)

        cant2_aux = dfDatos['cant'][1]
        entryCant2.insert(0, cant2_aux)

        formula3 = dfDatos['formula'][2]
        entryFormula3.insert(0, formula3)

        formula3Lista3 = esLista3(formula3)

        try:
            #Para que no se abra el listbox
            AutocompleteEntry.endPublico(entryFormula1)
            AutocompleteEntry.endPublico(entryFormula2)
            AutocompleteEntry.endPublico(entryFormula3)
        except Exception as e:
            print(e)

        cant3_aux = dfDatos['cant'][2]
        entryCant3.insert(0, cant3_aux)

        control1 = dfDatos['control'][0]
        control2 = dfDatos['control'][1]
        control3 = dfDatos['control'][2]

#Para que se muestre la imagen
img_buscar = PhotoImage(file='Iconos/buscar.png')

#Boton buscar
btnBuscar = HoverButton(frameVariables, text="buscar", font=(12), image=img_buscar, activebackground='#98abca', bg = '#98abca')
btnBuscar.grid(row = 0, column = 1)
btnBuscar.bind("<Right>", clickerBuscar)
#Click izquierdo
btnBuscar.bind("<Button-1>", clickerBuscar)
#Tecla enter
btnBuscar.bind("<Return>", clickerBuscar)
#Onhover
CreateToolTip(btnBuscar, text = "Buscar")

#Texto fecha
labelFecha = Label(frameTexto, text = "Fecha", font=(12), bg ='#98abca')
labelFecha.grid(row = 0, column = 2, padx = 60)
labelFecha['font'] = myFont

#Hice una función porque sino, no me acepta Calendario.get()
def func_varFecha():
    global varFecha
    global myFont
    #Variable fecha
    varFecha = StringVar()
    varFecha.set(Calendario.get_date())
    labelMedicamentoFecha = Label(frameVariables, textvariable = varFecha, font=12, bg ='#98abca')
    labelMedicamentoFecha.grid(row = 0, column = 2, padx = 20)
    labelMedicamentoFecha['font'] = myFont

#Creo un frame para los nombres
frameNombres = Frame()
frameNombres.grid(row = 1, column = 2, columnspan = 5, sticky="nsew")
frameNombres.configure(background='#708ab4')

#Texto nombre y apellido del médico
labelNombre = Label(frameNombres, text = "Nombre y apellido", font=(12), bg ='#708ab4')
labelNombre.grid(row = 0, column = 2, pady = 3, padx = 160)
labelNombre['font'] = myFont

#Entrada del nombre del médico
entryNombre = AutocompleteEntry(lista_medicos, root, width = 40)
entryNombre.grid(row = 2, column = 2, padx = 10)
entryNombre['font'] = myFont
#Para que se seleccione bien el predictivo
entryNombre.bind('<Button-1>', cambioPantalla)
entryNombre.bind('<KeyRelease>', cambioPantalla)

#Texto fórmula
labelFormula = Label(frameNombres, text = "Fórmula", font=12, bg ='#708ab4')
labelFormula.grid(row = 0, column = 3, padx = 160)
labelFormula['font'] = myFont

#Entrada de las fórmulas
#Fórmula 1
entryFormula1 = AutocompleteEntry(lista_formulas, root, width = 40)
entryFormula1.grid(row = 2, column = 3, padx = 5, pady = 5)
entryFormula1['font'] = myFont
#Para que se seleccione bien el predictivo
entryFormula1.bind('<Button-1>', cambioPantalla)
entryFormula1.bind('<KeyRelease>', cambioPantalla)

#Fórmula 2
entryFormula2 = AutocompleteEntry(lista_formulas, root, width = 40)
entryFormula2.grid(row = 3, column = 3, padx = 5, pady = 15)
entryFormula2['font'] = myFont
#Para que se seleccione bien el predictivo
entryFormula2.bind('<Button-1>', cambioPantalla)
entryFormula2.bind('<KeyRelease>', cambioPantalla)

#Fórmula 3
entryFormula3 = AutocompleteEntry(lista_formulas, root, width = 40)
entryFormula3.grid(row = 4, column = 3, padx = 5)
entryFormula3['font'] = myFont
#Para que se seleccione bien el predictivo
entryFormula2.bind('<Button-1>', cambioPantalla)
entryFormula3.bind('<KeyRelease>', cambioPantalla)

#Texto cantidad
labelCantidad = Label(frameNombres, text = "Cantidad", font=12, bg ='#708ab4')
labelCantidad.grid(row = 0, column = 4, padx = 160)
labelCantidad['font'] = myFont

#Entrada de las cantidades
cant1 = StringVar(root, value = 1)
#Cantidad 1
entryCant1 = Entry(root, textvariable=cant1, width = 40)
entryCant1.grid(row = 2, column = 4, padx = 5)
entryCant1['font'] = myFont

cant2 = StringVar(root, value = 0)
#Cantidad 2
entryCant2 = Entry(root, textvariable=cant2, width = 40)
entryCant2.grid(row = 3, column = 4, padx = 5)
entryCant2['font'] = myFont


cant3 = StringVar(root, value = 0)
#Cantidad 3
entryCant3 = Entry(root, textvariable=cant3, width = 40)
entryCant3.grid(row = 4, column = 4, padx = 5)
entryCant3['font'] = myFont

def clickerGrabar(event):
    global ultimaFecha
    global control1
    global control2
    global control3
    global formula1Lista3
    global formula2Lista3
    global formula3Lista3
    global fecha
    global formula1
    global formula2
    global formula3

    try:
        idReceta = int(entryIDreceta.get())
    except Exception:
        m_box.showerror('Error', 'Ingrese un número válido')
        return()

    nombreMedico = str(entryNombre.get()).upper()

    #Grabar
    if idReceta == ultimoID + 1:
        if nombreMedico == 'SIN MOVIMIENTO':
            #Tomo el último ID y le concateno un 4
            control = str(ultimoID) + '4'
            #Lo convierto a int
            control = int(control)

            #Para que no me pueda ingresar fechas salteadas
            filtroFecha = validoFecha()
            if filtroFecha == False:
                m_box.showerror('Error', 'No es posible ingresar una receta con esa fecha')
                return()

            database.sinMovimiento(Calendario.get_date(), control)
            ultimaFecha  = datetime.strptime(Calendario.get_date(), '%d-%m-%Y')

            #NO ANDA SI EL CALENDARIO NO ESTA ABIERTO. QUEDA PARA MAS ADELANTE
            #Para que pase a la fecha siguiente
            #Fijo la fecha de la view calendario
            '''Calendario.selection_set(ultimaFecha + timedelta(days = 1))'''
            #Fijo la vista de la view donde se guardan las recetas
            '''guardoFecha((ultimaFecha + timedelta(days = 1)).strftime('%d-%m-%Y'))'''
            #Para guardar la fecha uso Calendario.get_date() entonces con eso basta
            
            limpioLosCampos()
        else:
            existeMedico(nombreMedico)
            existe = existeMedicamento()
            #Si no existe
            if existe == False:
                #Para que el cursor se ponga en el Nombre del médico
                entryNombre.focus_set()
                return()

            grabarMedicamento()
            entryNombre.focus_set()
        return()
    #Actualiza
    elif idReceta <= ultimoID:
        try:
            existeMedico(str(entryNombre.get()).upper())

            if str(entryNombre.get()).upper() == 'SIN MOVIMIENTO':
                controlSM = str(idReceta) + '4'
                database.sinMovimiento(fecha, controlSM)
                database.borrarReceta(int(control1))
                database.borrarReceta(int(control2))
                database.borrarReceta(int(control3))
                database.borrarReceta3(int(control1))
                database.borrarReceta3(int(control2))
                database.borrarReceta3(int(control3))
                return()

            #Es para que el programa tire error en el caso que toque grabar sin antes la lupita
            print(control1)

            #Actualiza
            existe = existeMedicamento()

            if control1 != 0:
                #No se puede borrar la primera
                if nombreMedico == '' or int(entryCant1.get()) <= 0 or str(entryFormula1.get()).upper() == '':
                    m_box.showerror('Error', 'No puede ser vacio el nombre del médico ni la primer fórmula')
                    return()
                
                if existe == False:
                    m_box.showinfo('Aviso','O ponga la cantidad en cero')
                    return()

                database.actualizoPsico(nombreMedico, str(entryFormula1.get()).upper(), entryCant1.get(), int(control1))
                
                ABMpsico3(formula1Lista3, formula1, nombreMedico, entryFormula1.get(), entryCant1.get(), control1, idReceta)

            if control2 != 0:
                #Borrar
                if str(entryFormula2.get()).upper() == '' or int(entryCant2.get()) <= 0:
                    database.borrarReceta(int(control2))
                    ABMpsico3(formula2Lista3, formula2, nombreMedico, entryFormula2.get(), entryCant2.get(), control2, idReceta)
                    m_box.showinfo('Aviso', 'Receta eliminada con éxito')
                    return()

                if existe == False:
                    m_box.showinfo('Aviso', 'Ponga la cantidad en cero')
                    return()

                database.actualizoPsico(nombreMedico, str(entryFormula2.get()).upper(), entryCant2.get(), int(control2))
                ABMpsico3(formula2Lista3, formula2, nombreMedico, entryFormula2.get(), entryCant2.get(), control2, idReceta)
            else:
                #Creo nuevo
                if entryFormula2.get() != '' and int(entryCant2.get()) > 0:
                    control2 = str(idReceta) + '2'
                    control2 = int(control2) 
                    database.graboMedicamento(idReceta, nombreMedico, fecha, str(entryFormula2.get()).upper(), entryCant2.get(), control2)
                    ABMpsico3(formula2Lista3, formula2, nombreMedico, entryFormula2.get(), entryCant2.get(), control2, idReceta)
                
                if entryFormula2.get() != '' and int(entryCant2.get()) <= 0:
                    m_box.showwarning('AVISO', 'Ingrese una cantidad válida')
                    return()

            

            if control3 != 0:
                #Borrar
                if str(entryFormula3.get()).upper() == '' or int(entryCant3.get()) <= 0:
                    ABMpsico3(formula3Lista3, formula3, nombreMedico, entryFormula3.get(), entryCant3.get(), control3, idReceta)
                    database.borrarReceta(int(control3))
                    m_box.showinfo('Aviso', 'Receta borrada con éxito')
                    return()

                #Actualiza
                if existe == False:
                    m_box.showinfo('Aviso', 'Ponga la cantidad en cero')
                    return()

                ABMpsico3(formula3Lista3, formula3, nombreMedico, entryFormula3.get(), entryCant3.get(), control3, idReceta)
                database.actualizoPsico(nombreMedico, str(entryFormula3.get()).upper(), entryCant3.get(), int(control3))
            else:
                #Creo nuevo
                if entryFormula3.get() != '' and int(entryCant3.get()) > 0:
                    control3 = str(idReceta) + '3'
                    control3 = int(control3) 
                    ABMpsico3(formula3Lista3, formula3, nombreMedico, entryFormula3.get(), entryCant3.get(), control3, idReceta)
                    database.graboMedicamento(idReceta, nombreMedico, fecha, str(entryFormula3.get()).upper(), entryCant3.get(), control3)
                
                if entryFormula3.get() != '' and int(entryCant3.get()) <= 0:
                    m_box.showwarning('AVISO', 'Ingrese una cantidad válida')
                    return()
            
            limpioLosCampos()
            entryIDreceta.delete(0, END)
            entryIDreceta.insert(0, str(ultimoID + 1))
            m_box.showinfo('Aviso', 'Actualización realizada con exito')
            return()
        except Exception:
            limpioLosCampos()
            m_box.showerror('Error', 'No se pudo realizar la actualización')
    #Nada
    else:
        m_box.showinfo('Aviso', 'El número de receta que corresponde es: ' + str(ultimoID + 1))
        entryIDreceta.delete(0, END)
        entryIDreceta.insert(0, str(ultimoID + 1))
        return()

#Frame botones
frameBotones = Frame()
frameBotones.grid(row = 7, column = 2, columnspan = 1)
frameBotones.configure(background='#98abca')

def clickerSaltear(event):
    
    #Para que el cursor se ponga en el Nombre del médico
    entryNombre.focus_force()

    #Con esto el Tab se frena, entonces no se va al boton de borrar
    return 'break'

#Para que se muestre la imagen
img_discket = PhotoImage(file='Iconos/discket.png')

#Boton grabar
btnGrabar = HoverButton(frameBotones, text="Grabar", font=(12), image=img_discket, activebackground='#98abca', bg = '#98abca')
btnGrabar.grid(row = 0, column = 2, padx = 62)
btnGrabar.bind("<Right>", clickerGrabar)
#Click izquierdo
btnGrabar.bind("<Button-1>", clickerGrabar)
#Tecla enter
btnGrabar.bind("<Return>", clickerGrabar)
#Para que grabe con ctrl+s
root.bind("<Control-s>", clickerGrabar)
root.bind("<Control-S>", clickerGrabar)
#Para que saltee borrar con el tab
btnGrabar.bind("<Tab>", clickerSaltear)
#Onhover
CreateToolTip(btnGrabar, text = "Grabar")

def clickerBorrar(event):
    global myFont
    

    frameLogin = Toplevel(bg = '#e0bdfa')

    #Cambio el icono
    frameLogin.iconbitmap("Iconos/icono.ico")

    labelPermiso = Label(frameLogin, text = 'Se requieren permisos de administrador', font = 12, bg = '#cb91f7')
    labelPermiso['font'] = myFont
    labelPermiso.pack()

    labelUsuario = Label(frameLogin, text = 'Usuario', font = 12, bg = '#e0bdfa')
    labelUsuario['font'] = myFont
    labelUsuario.pack()

    entryUsuario = Entry(frameLogin, font = 12)
    entryUsuario['font'] = myFont
    entryUsuario.pack()

    labelContraseña = Label(frameLogin, text = 'Contraseña', font = 12, bg = '#e0bdfa')
    labelContraseña['font'] = myFont
    labelContraseña.pack()

    entryContraseña = Entry(frameLogin, font = 12, show = '*')
    entryContraseña['font'] = myFont
    entryContraseña.pack(padx = 10)

    #Para que se muestre la imagen
    img_loguin = PhotoImage(file='Iconos/loguin.png')

    btnLoguin = HoverButton(frameLogin, text = 'Borrar', font=(12), image=img_loguin, activebackground='#e0bdfa', bg = '#e0bdfa', command = lambda: borrar())
    btnLoguin.pack(pady = 10)
    #Si no pongo esto la imagen no se muestra
    btnLoguin.image = img_loguin

    def borrar(event):
        global ultimoID
        if entryUsuario.get() == 'NELSON' and entryContraseña.get() == '1978':
            try:
                sinMovimiento = database.borrarUltimaReceta()
                if sinMovimiento == False:
                    varIDreceta.set(ultimoID)
                    ultimoID -= 1
                    funUltimaFecha()
                m_box.showinfo('Aviso', 'Última receta borrada con exito')
            except Exception as e:
                m_box.showerror('Error', 'Se produjo un error: ' + str(e))
        else:
            m_box.showerror('Error', 'Usuario y/o contraseña incorrectos')
        frameLogin.withdraw()

    btnLoguin.bind("<Right>", borrar)
    #Click izquierdo
    btnLoguin.bind("<Button-1>", borrar)
    #Tecla enter
    btnLoguin.bind("<Return>", borrar)
    #Onhover
    CreateToolTip(btnLoguin, text = "Ingresar")

#Para que se muestre la imagen
img_delete = PhotoImage(file='Iconos/delete.png')

btnBorrar = HoverButton(frameBotones, text = 'Borrar', font=(12), image=img_delete, activebackground='#98abca', bg = '#98abca')
btnBorrar.grid(row = 0, column = 0, padx = 105)
btnBorrar.bind("<Right>", clickerBorrar)
#Click izquierdo
btnBorrar.bind("<Button-1>", clickerBorrar)
#Tecla enter
btnBorrar.bind("<Return>", clickerBorrar)
#Onhover
CreateToolTip(btnBorrar, text = "Borrar la última receta")

frameFooter = Frame()
frameFooter.place(x = 0, y = 400, width = 1252)
frameFooter.configure(background='#98abca')

labelFooter = Label(frameFooter, text = 'Zurich Software', font = 9, fg ='#c0cce0', background='#98abca')
labelFooter.pack(fill = X)
labelFooter['font'] = myFont


#--------------------------------------------------------------------------------------------------#
#Alta fórmula
def pantallaFormula():
    global frameFormula

    #Creo el frame fecha
    altaFormulaFrame = Toplevel()

    #Cambio el colo de fondo
    altaFormulaFrame.configure(background='#aad9af')

    #Selecciono el icono
    altaFormulaFrame.iconbitmap("Iconos/icono.ico")

    #Selecciono el tamaño de la pantalla
    altaFormulaFrame.geometry("460x340")

    #Para que no se pueda cambiar el tamaño de la pantalla
    altaFormulaFrame.resizable(False, False)  

    #Texto seleccione fecha
    textoLabel = Label(altaFormulaFrame, text="Ingrese la fórmula", font=18, bg ='#89cb8f')
    textoLabel.pack(fill = X)
    textoLabel['font'] = myFont

    def cambioAltaFormulaFrame(event):
        global pantalla
        pantalla = 'altaFormulaFrame'

    #Creo un frame para el entry y para el checkbox
    frameFormula = Frame(altaFormulaFrame, bg = '#aad9af')
    #Si no le pongo width y height al frame no me sale el listbox para el predictivo
    frameFormula['height'] = 280
    frameFormula['width'] = 450
    frameFormula.grid_propagate(False)
    frameFormula.pack()

    #Texto nombre comercial
    textoNombreComercial = Label(frameFormula, text = "Ingrese el nombre comercial", font = "14", bg = '#aad9af')
    textoNombreComercial.grid(row = 0, column = 0, padx = 5, sticky=SW, pady = (10,0))
    textoNombreComercial['font'] = myFont

    #Entry formula
    entryFormula = AutocompleteEntry(lista_formulas, frameFormula, width = 40, font=12)
    entryFormula.grid(row = 1, column = 0, padx = 5)
    entryFormula['font'] = myFont
    entryFormula.bind('<KeyRelease>', cambioAltaFormulaFrame)
    entryFormula.bind('<Button-1>', cambioAltaFormulaFrame)

    #Texto nombre DCI
    textoNombreDCI = Label(frameFormula, text = "Ingrese el nombre DCI", font = "14", bg = '#aad9af')
    textoNombreDCI.grid(row = 2, column = 0, padx = 5, sticky=SW, pady = (10,0))
    textoNombreDCI['font'] = myFont

    #Entry formula
    entryFormulaDCI = AutocompleteEntry(lista_DCIformulas, frameFormula, width = 40, font=12)
    entryFormulaDCI.grid(row = 3, column = 0, padx = 5)
    entryFormulaDCI['font'] = myFont
    entryFormulaDCI.bind('<KeyRelease>', cambioAltaFormulaFrame)
    entryFormulaDCI.bind('<Button-1>', cambioAltaFormulaFrame)


    esPsico3 = IntVar()
    checkPsico3 = Checkbutton(frameFormula, variable = esPsico3, bg = '#aad9af', activebackground='#aad9af', onvalue = 1, offvalue = 0)
    checkPsico3.grid(row = 1, column = 1, padx = 5)
    #Onhover
    CreateToolTip(checkPsico3, text = "¿Es Psico3?")

    #Para que se muestre la imagen
    img_discket = PhotoImage(file='Iconos/discket.png')

    #Boton grabar
    btnGrabar = HoverButton(altaFormulaFrame, activebackground='#aad9af', bg = '#aad9af', font=(12), image = img_discket, text="Grabar")
    btnGrabar.pack(side = RIGHT, padx = 10, anchor = SW)
    btnGrabar['borderwidth'] = 0
    #Si no pongo esto la imagen no se muestra
    btnGrabar.image = img_discket
    #Onhover
    CreateToolTip(btnGrabar, text = "Grabar")

    def existeFormula():
        global idDCI

        formula = str(entryFormula.get()).upper()
        dci = str(entryFormulaDCI.get()).upper()
        filtro = dfFormula[dfFormula['Formula'].str.contains(formula)]

        if formula == '':
            m_box.showerror('Error', "La formula no puede ser vacia")
            altaFormulaFrame.focus_force()
            return()

        if len(filtro) != 0:
            filtro2 = valildarFiltro2(filtro, formula, 'Formula')
            
            if filtro2 == True:
                m_box.showerror('Error', "Ya existe la fórmula " + formula)
                altaFormulaFrame.focus_force()
                return()
        
        if esPsico3.get() == 1:
            database.grabarFormula3(formula)

        database.grabarFormula(formula, dci)
        m_box.showinfo('Aviso!', 'Archivo guardado exitosamente')

        #Pongo los campos vacios
        entryFormula.delete(0, 'end')
        checkPsico3.deselect()

        #Vacío la lista idDCI
        idDCI = []

        altaFormulaFrame.focus_force()


    

    def clickerGrabar(event):
        existeFormula()

    btnGrabar.bind("<Right>", clickerGrabar)
    #Click izquierdo
    btnGrabar.bind("<Button-1>", clickerGrabar)
    #Tecla enter
    btnGrabar.bind("<Return>", clickerGrabar)

    #Para que se muestre la imagen
    img_delete = PhotoImage(file='Iconos/delete.png')

    #Boton borrar
    btnBorrar = HoverButton(altaFormulaFrame, activebackground='#aad9af', bg = '#aad9af', font=12, image = img_delete, text="Delete")
    btnBorrar.pack(side = LEFT, padx = 10, anchor = SW)
    btnBorrar['borderwidth'] = 0
    #Si no pongo esto la imagen no se muestra
    btnBorrar.image = img_delete
    #Onhover
    CreateToolTip(btnBorrar, text = "Borrar")

    def clickerBorrar(event):
        global dfFormula
        formula = str(entryFormula.get().upper())
        filtro = dfFormula[dfFormula['Formula'].str.contains(formula)]
        filtro2 = valildarFiltro2(filtro, formula, 'Formula')
        if filtro2 == True:
            filtro3 = dfFormula3[dfFormula3['Formula3'].str.contains(formula)]
            filtro4 = valildarFiltro2(filtro3, formula, 'Formula3')
            try:
                if filtro4 == True:
                    database.borroFormula3(formula)
                database.borroFormula(formula)
                m_box.showinfo('Aviso!', 'Fórmula borrada exitosamente')
            except Exception as e:
                m_box.showerror('Error', 'El siguiente error ocurrio al borrar el archivo ' + str(e))
        else:
            m_box.showerror('Error', 'No existe una fórmula con ese nombre')

        #Pongo los campos vacios
        entryFormula.delete(0, 'end')

        #Para que el cursor se ponga en el Nombre del médico
        entryFormula.focus_set()
        
    btnBorrar.bind("<Right>", clickerBorrar)
    #Click izquierdo
    btnBorrar.bind("<Button-1>", clickerBorrar)
    #Tecla enter
    btnBorrar.bind("<Return>", clickerBorrar)

#--------------------------------------------------------------------------------------------------#
#Pantalla imprimir psicotrópicos
def pantallaPsico():

    #Creo la pantalla
    framePantallaPsico = Toplevel()

    #Cambio el colo de fondo
    framePantallaPsico.configure(background='#f9c97f')

    #Selecciono el icono
    framePantallaPsico.iconbitmap("Iconos/icono.ico")

    #Selecciono el tamaño de la pantalla
    framePantallaPsico.geometry("300x200")

    #Para que no se pueda cambiar el tamaño de la pantalla
    framePantallaPsico.resizable(False, False)  

    #Texto titulo
    labelTituloPsico = Label(framePantallaPsico, text = 'Reporte L. Recetario', font=(18), bg ='#f5a326')
    labelTituloPsico.pack(fill = X)
    labelTituloPsico['font'] = myFont

    #Texto numero de receta
    labelNumeroReceta = Label(framePantallaPsico, text = 'Ingrese el numero de receta', font=(12), bg ='#f6b34c')
    labelNumeroReceta.pack(fill = X)
    labelNumeroReceta['font'] = myFont

    #Texto desde
    labelDesde = Label(framePantallaPsico, text = 'Desde', font=(12), bg ='#f9c97f')
    labelDesde.pack()
    labelDesde['font'] = myFont

    #Entry desde
    entryDesde = Entry(framePantallaPsico)
    entryDesde.pack()
    entryDesde['font'] = myFont

    #Texto hasta
    labelHasta = Label(framePantallaPsico, text = 'Hasta', font=(12), bg ='#f9c97f')
    labelHasta.pack()
    labelHasta['font'] = myFont

    #Entry hasta
    entryHasta = Entry(framePantallaPsico)
    entryHasta.pack()
    entryHasta['font'] = myFont

    #Para que se muestre la imagen
    img_discket = PhotoImage(file='Iconos/discket.png')

    #Boton para consultar
    botonConsultar = HoverButton(framePantallaPsico, activebackground='#f9c97f', bg = '#f9c97f', image = img_discket, font=(12), text = "Consultar")
    botonConsultar.pack(side = RIGHT, padx = 5)
    #Si no pongo esto la imagen no se muestra
    botonConsultar.image = img_discket
    #Onhover
    CreateToolTip(botonConsultar, text = "Grabar")

    def clickerBuscar(event):
        browsefunc()
        try:
            guardarDatos(entryDesde.get(), entryHasta.get())
            framePantallaPsico.focus_force()
        except Exception:
            m_box.showerror('Error', 'Error al guardar el archivo')
            framePantallaPsico.focus_force()

    botonConsultar.bind("<Right>", clickerBuscar)
    #Click izquierdo
    botonConsultar.bind("<Button-1>", clickerBuscar)
    #Tecla enter
    botonConsultar.bind("<Return>", clickerBuscar)

#Función para elegir a donde se va a descargar el archivo
def browsefunc(): 
    #Definimos una variable global para poder llamarla en la funcion excel, sin llamar a la funcion
    global directory

    directory = filedialog.askdirectory(initialdir='.')

def guardarDatos(desde, hasta):
    try:
        desde = int(desde)
        hasta = int(hasta)
    except Exception as e:
        m_box.showerror('Error', 'Ingrese un número válido')
        return()

    lista_datos = database.traigoPsico(str(desde)+'1', str(hasta)+'4')
    dfDatos = pd.DataFrame(lista_datos, columns = ['ignorar', 'Numero', 'Nombre del medico', 'Fecha', 'Formula', 'Cajas', 'control'])
    dfDatos.drop(['ignorar', 'control'], axis = 1, inplace = True)
    
    #Esto es lo que quiere Agnese
    dfDatos['Fecha'].replace({'-2020':''}, regex = True, inplace = True)
    dfDatos['Fecha'].replace({'-2021':''}, regex = True, inplace = True)
    dfDatos['Fecha'].replace({'-2022':''}, regex = True, inplace = True)
    dfDatos['Fecha'].replace({'-2023':''}, regex = True, inplace = True)
    dfDatos['Fecha'].replace({'-2024':''}, regex = True, inplace = True)


    lista_numeros = list(dfDatos.Numero)
    lista_numeros_aux = []
    lista_numeros_aux.append(lista_numeros[0])
    for i in range(0,len(lista_numeros)-1,1):
        if(lista_numeros[i] == lista_numeros[i+1]):
            lista_numeros_aux.append('')
        else:
            lista_numeros_aux.append(lista_numeros[i+1])

    dfDatos.drop('Numero', axis = 1, inplace = True)
    dfDatos['Numero'] = lista_numeros_aux
    dfDatos.set_index('Numero', inplace = True)

    try:
        dfDatos.to_excel(directory + "/" + str(desde) + "-" + str(hasta) + " Libro recetario.xlsx")

        #Para que se le pueda cambiar el formato
        pandas.io.formats.excel.ExcelFormatter.header_style = None

        #Para que se formatee bien la impresion
        #Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(directory + "/" + str(desde) + "-" + str(hasta) + " Libro recetario.xlsx", engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        dfDatos.to_excel(writer, sheet_name='Sheet1')

        # Get the xlsxwriter workbook and worksheet objects.
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        #Para darle formato a la celda
        cell_format = workbook.add_format()
        cell_format.set_font_size(14)
        cell_format.set_align('center')
        cell_format.set_align('bottom')

        #Para darle formato a la celda de la forma que pidio Agnese
        cell_format_2 = workbook.add_format()
        cell_format_2.set_font_size(14)
        cell_format_2.set_align('left')
        cell_format_2.set_align('bottom')

        #Para que se ponga la hoja A5
        worksheet.set_paper(11)
        #Margenes
        '''worksheet.set_margins(left = 0, right = 0, top = 0.75, bottom = 0)'''
        worksheet.set_margins(left = 0, right = 0, top = 0, bottom = 0)

        #Lo centra
        worksheet.center_horizontally()
        worksheet.center_vertically()

        #Para que se ponga impresión horizontal
        worksheet.set_landscape()

        #Esta en puntos (es una unidad de medida)
        #Numero
        worksheet.set_column('A:A', 9.57, cell_format)
        #Nombre del medico
        worksheet.set_column('B:B', 28.86, cell_format_2)
        #Fecha
        worksheet.set_column('C:C', 7.14, cell_format)
        #Fórmula
        worksheet.set_column('D:D', 48.86, cell_format_2)
        #Cantidad
        worksheet.set_column('E:E', 4.86, cell_format)

        #Para cambiar las filas
        worksheet.set_default_row(29.15)

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        


    except Exception as e:
        m_box.showerror('Error', 'No se pudo guardar el archivo, seleccione una carpeta')
        print(e)
        return()

    m_box.showinfo('Aviso!', 'Archivo guardado exitosamente')

#--------------------------------------------------------------------------------------------------#
#Pantalla para imprimir psicotropicos de lista 3
def pantallaPsico3():
    #Creo la pantalla
    framePantallaPsico3 = Toplevel()

    #Cambio el colo de fondo
    framePantallaPsico3.configure(background='#d8a5aa')

    #Selecciono el icono
    framePantallaPsico3.iconbitmap("Iconos/icono.ico")

    #Selecciono el tamaño de la pantalla
    framePantallaPsico3.geometry("300x200")

    #Para que no se pueda cambiar el tamaño de la pantalla
    framePantallaPsico3.resizable(False, False) 

    #Texto titulo
    labelTituloPsico = Label(framePantallaPsico3, bg ='#d16e72', text = 'Reporte L. Contralor', font=(18))
    labelTituloPsico.pack(fill = X)
    labelTituloPsico['font'] = myFont

    #Texto numero de receta
    labelNumeroReceta = Label(framePantallaPsico3, bg ='#c8868a', text = 'Ingrese el numero de receta', font=(12))
    labelNumeroReceta.pack(fill = X)
    labelNumeroReceta['font'] = myFont

    #Texto desde
    labelDesde = Label(framePantallaPsico3, bg ='#d8a5aa', text = 'Desde', font=(12))
    labelDesde.pack()
    labelDesde['font'] = myFont

    #Entry desde
    entryDesde = Entry(framePantallaPsico3)
    entryDesde.pack()
    entryDesde['font'] = myFont

    #Texto hasta
    labelHasta = Label(framePantallaPsico3, bg ='#d8a5aa', text = 'Hasta', font=( 12))
    labelHasta.pack()
    labelHasta['font'] = myFont

    #Entry hasta
    entryHasta = Entry(framePantallaPsico3)
    entryHasta.pack()
    entryHasta['font'] = myFont

    #Para que se muestre la imagen
    img_discket = PhotoImage(file='Iconos/discket.png')

    #Boton para consultar
    botonConsultar = HoverButton(framePantallaPsico3, activebackground='#d8a5aa', bg = '#d8a5aa', image = img_discket, font=("Calibri", 12), text = "Consultar")
    botonConsultar.pack(side = RIGHT, padx = 5)
    #Si no pongo esto la imagen no se muestra
    botonConsultar.image = img_discket
    #Onhover
    CreateToolTip(botonConsultar, text = "Grabar")

    def clickerBuscar(event):
        try:
            guardarDatos3(entryDesde.get(), entryHasta.get())
            framePantallaPsico3.focus_force()
        except Exception as e:
            m_box.showerror('Error', 'Error al guardar el archivo ' + str(e))
            framePantallaPsico3.focus_force()

    botonConsultar.bind("<Right>", clickerBuscar)
    #Click izquierdo
    botonConsultar.bind("<Button-1>", clickerBuscar)
    #Tecla enter
    botonConsultar.bind("<Return>", clickerBuscar)

def guardarDatos3(desde, hasta):
    browsefunc()
    try:
        desde = int(desde)
        hasta = int(hasta)
    except Exception as e:
        m_box.showerror('Error', 'Ingrese un número válido ' + str(e))
        return()

    lista_datos = database.traigoPsico3(desde, hasta)
    dfDatos = pd.DataFrame(lista_datos, columns = ['ignorar', 'Numero', 'Nombre del medico', 'Fecha', 'Formula', 'Cajas', 'Control'])
    dfDatos.drop('ignorar', axis = 1, inplace = True)
    dfDatos.drop('Control', axis = 1, inplace = True)
    
    #Esto es lo que quiere Agnese
    dfDatos['Fecha'].replace({'-2020':''}, regex = True, inplace = True)
    dfDatos['Fecha'].replace({'-2021':''}, regex = True, inplace = True)
    dfDatos['Fecha'].replace({'-2022':''}, regex = True, inplace = True)
    dfDatos['Fecha'].replace({'-2023':''}, regex = True, inplace = True)
    dfDatos['Fecha'].replace({'-2024':''}, regex = True, inplace = True)

    lista_numeros = list(dfDatos.Numero)
    lista_numeros_aux = []
    lista_numeros_aux.append(lista_numeros[0])
    for i in range(0,len(lista_numeros)-1,1):
        if(lista_numeros[i] == lista_numeros[i+1]):
            lista_numeros_aux.append('')
        else:
            lista_numeros_aux.append(lista_numeros[i+1])

    dfDatos.drop('Numero', axis = 1, inplace = True)
    dfDatos['Numero'] = lista_numeros_aux
    dfDatos.set_index('Numero', inplace = True)

    try:
        dfDatos.to_excel(directory + "/" + str(desde) + "-" + str(hasta) + " Reporte libro contralor.xlsx")

    except Exception as e:
        m_box.showerror('Error', 'No se pudo guardar el archivo, seleccione una carpeta')
        return()
    
    m_box.showinfo('Aviso!', 'Archivo guardado exitosamente')
#--------------------------------------------------------------------------------------------------#
#Backup

def browseBackup():
    #Definimos una variable global para poder llamarla en la funcion excel, sin llamar a la funcion
    global pathBackup

    pathBackup = filedialog.askdirectory(initialdir='.')

def backup():
    global pathBackup

    try:
        #Selecciono el path
        browseBackup()

        #Pido los datos y creo los DataFrame

        lista_formula = database.backupFormula()
        dfBackUpFormula = pd.DataFrame(lista_formula, columns = ['id_formula', 'formula'])
        dfBackUpFormula.set_index('id_formula', inplace = True)


        lista_formula3 = database.backupFormula3()
        dfBackUpFormula3 = pd.DataFrame(lista_formula3, columns = ['idtabla_formula3', 'formula3'])
        dfBackUpFormula3.set_index('idtabla_formula3', inplace = True)
        
        lista_medicosBackUp = database.backupMedicos()
        dfBackUpMedicos = pd.DataFrame(lista_medicosBackUp, columns = ['id_medicos', 'nombre'])
        dfBackUpMedicos.set_index('id_medicos', inplace = True)

        lista_tablapsico = database.backupTablaPsico()
        dfTablaPsico = pd.DataFrame(lista_tablapsico, columns = ['ignorar', 'id_receta', 'nombre_medico', 'fecha', 'formula', 'cant', 'control'])
        dfTablaPsico.set_index('ignorar', inplace = True)

        lista_tablapsico3 = database.backupTablaPsico3()
        dfTablaPsico3 = pd.DataFrame(lista_tablapsico3, columns = ['ignorar', 'id_receta', 'nombre_medico', 'fecha', 'formula', 'cant', 'control'])
        dfTablaPsico3.set_index('ignorar', inplace = True)

        #Guardo los datos
        dfBackUpFormula.to_excel(pathBackup + "/" + " backup-formulas libro recetario.xlsx")
        dfBackUpFormula3.to_excel(pathBackup + "/" + " backup-formulas libro contralor.xlsx")
        dfBackUpMedicos.to_excel(pathBackup + "/" + " backup-medicos.xlsx")
        dfTablaPsico.to_excel(pathBackup + "/" + " backup-libro recetario.xlsx")
        dfTablaPsico3.to_excel(pathBackup + "/" + " backup-reporte libro contralor.xlsx")


        m_box.showinfo('Aviso', 'Archivo guardado exitosamente')
    except Exception:
        m_box.showerror('Error', 'Algo salió mal')
    
        
#--------------------------------------------------------------------------------------------------#
#Es para que se abra primero
funUltimaFecha()
seleccionoFecha()
func_varFecha()

root.mainloop()