# Load libraries
from pandas import read_csv
from pandas import to_datetime
import openpyxl
import tkinter as tk
from tkinter import filedialog
from datetime import datetime, date, timedelta
from PIL import Image


asl_path = " "
global_path = " "

#Genero pop up para obtener la ruta de los archivos de ASL y global
root= tk.Tk()
root.wm_title("Cruce ASL vs Globalcollect!")

mainwindow = tk.Canvas(root, width = 450, height = 400, bg = 'white', relief = 'raised')
mainwindow.pack()

def okASL():
    label_asl.config(text='OK', fg="green")

def okGlobal():
    label_global.config(text='OK', fg="green")

def popupmsg(msg):
    popup = tk.Tk()
    popup.wm_title("!")
    label = tk.Label(popup, text=msg, font=('helvetica', 12, 'bold'))
    label.pack(side="top", fill="x", pady=10)
    B1 = tk.Button(popup, text="Aceptar", command = popup.destroy)
    B1.pack()
    popup.mainloop()

#Obtiene la ruta del archivo de ASL
def getCSVASL ():
  global asl_path
  asl_path = filedialog.askopenfilename()
  if(asl_path != " "):
    okASL()
  print(asl_path)

#Obtiene la ruta del archivo de Globalcollect
def getCSVGlobal ():
  global global_path
  global_path = filedialog.askopenfilename()
  if(global_path != " "):
    okGlobal()
  print(global_path)



#Ejecuta el cruce
def ejecutarCruce():
  print (asl_path)
  print (global_path)
  # Load dataset
  #Defino las columnas
  names1 = ['MerchantID', 'ContractID',  'OrderID', 'EffortID',  'MerchantReference', 'PaymentReference',  'CustomerID', 'StatusID', 'StatusDescription', 'PaymentProduct ID', 'PaymentProductDescription',  'OrderCountryCode', 'OrderCurrencyCode',  'OrderAmount', 'RequestCurrencyCode',  'RequestAmount', 'PaidCurrency',  'PaidAmount',  'ReceivedDate',  'StatusDate',  'RejectionCode', 'Remarks']
  #Cargo el archivo
  globalcollect = read_csv(global_path,header=0 , names=names1, encoding='latin-1', index_col=False, low_memory=False)

  #Defino las columnas
  names2 = ['Merchant',  'Orden',  'Suborden', 'F12',  'Tipo Factura', 'Terminal', 'Secuencia',  'Numero De Boleta', 'Canal',  'Fuente Abastecimiento',  'RUT',  'Nombre Cliente', 'Sexo', 'Direccion Cliente',  'Comuna', 'Ciudad', 'Region', 'Pais', 'Codigo Postal',  'Fono Comprador', 'Direccion Despacho', 'Comuna Receptor',  'Ciudad Receptor',  'Region Receptor',  'Pais Receptor',  'Cod. Postal Recep.', 'Fono Receptor',  'Metodo Envio', 'Codigo Linea', 'Descripcion Linea',  'Codigo Sublinea',  'Descripcion Sublinea', 'Codigo Clase', 'Descripcion Clase',  'Codigo Subclase',  'Descripcion Subclase', 'SKU',  'Descripcion SKU',  'Tamano', 'Regalo', 'Medio De Pago',  'Banco',  'Numero De Tarjeta',  'Numero Cuotas',  'Cantidad', 'Codigo Subclase2',  'Descripcion Subclase2', 'Fecha Coloc Orden',  'Hora Coloc Orden', 'Fecha Entrega Orden',  'Fecha Reparto',  'Fecha Pactada',  'Fecha Picking',  'Fecha Ruta', 'Fecha Validacion', 'Fecha Boleta', 'Estado ASL', 'Fec. Cambio Estado', 'Retencion Novios', 'Codigo Vendedor',  'Nombre Vendedor',  'Email',  'Precio Con Rebaja',  'Precio', 'Valor Flete',  'Total',  'Bulto',  'CodigoAutorizacion',  'OC Agrupada PF', 'Hora Validacion',  'Hora Boleta']
  #Cargo el archivo
  asl = read_csv(asl_path, delimiter=';',header=0, names=names2, encoding='latin-1', low_memory=False)
  asl.to_csv("asl.csv", index=False)

  print("Globalcollect:")
  # shape
  print(globalcollect.shape)
  # head
  #print(globalcollect.head(20))

  print("ASL: ")
  # shape
  print(asl.shape)

  #filtro que el orderid de global no esté en el archivo de ASL y que el effortID sea 1
  cruce2 = globalcollect[~globalcollect.OrderID.isin(asl.CodigoAutorizacion) & globalcollect.EffortID.isin(['1'])]

  #Ordeno por fecha
  cruceOrdenado = cruce2.sort_values(by='ReceivedDate')
  #agrego columna de fecha con hora argentina (ya que la hora de global es de holanda)
  cruceOrdenado['fechaArg'] = cruceOrdenado['ReceivedDate']
  #convierto la fecha de global a la fecha argentina
  cruceOrdenadoFecha = cruceOrdenado.apply(lambda x: to_datetime(x, dayfirst=True, yearfirst=False) if x.name == 'fechaArg' else x)
  cruceFechaArg = cruceOrdenadoFecha.apply(lambda x: x - timedelta(hours=5) if x.name == 'fechaArg' else x)

  #imprimo una muestra del cruce para validar en la consola
  print("Cruce: ")
  print(cruceFechaArg.shape)
  print(cruceFechaArg.head(35))
  #Genero export en excel
  cruceFechaArg.to_excel("cruce_soar.xlsx","cruce")
  popupmsg("Cruce se generó con exito!!")


#Defino los botones y labels de la ventana

#labels de estado de carga del archivo (- si no lo cargó y OK si ya lo cargó)
label_asl = tk.Label(text='-',bg='white', fg='red', font=('helvetica', 12, 'bold'))
label_global = tk.Label(text='-',bg='white', fg='red', font=('helvetica', 12, 'bold'))

#Boton ruta asl
imagenASL = tk.PhotoImage(file="images/buttonASL.png")
print(imagenASL)
buttonAsl_csv = tk.Button(text="Importar archivo de ASL", command=getCSVASL, bg='white', fg='white', font=('helvetica', 12, 'bold'),image=imagenASL, border=0)
#Boton ruta global
imagenGlobal = tk.PhotoImage(file="images/buttonGlobal.png")
buttonGlobal_csv = tk.Button(text="Importar archivo de GLOBALCOLLECT", command=getCSVGlobal, bg='white', fg='white', font=('helvetica', 12, 'bold'),image=imagenGlobal, border=0)
#Boton cruce
imagenCruce = tk.PhotoImage(file="images/buttonCruce.png")
buttonCruce = tk.Button(text="Generar cruce", command=ejecutarCruce, bg='white', fg='white', font=('helvetica', 12, 'bold'),image=imagenCruce, border=0)


#Mi logo :)
imagenBypipeok = tk.PhotoImage(file="images/bypipeoksize.png")
byPipe = tk.Label(text='Compiled by Pipe',bg='white', fg='black', font=('helvetica', 12, 'bold'), image=imagenBypipeok)

muerteok = tk.PhotoImage(file="images/muerteoksize.png")
muerte_label = tk.Label(bg='white', fg='black', image=muerteok)

mainwindow.create_window(155, 60, window=buttonAsl_csv)
mainwindow.create_window(200, 130, window=buttonGlobal_csv)
mainwindow.create_window(225, 250, window=buttonCruce)

mainwindow.create_window(400, 60, window=label_asl)
mainwindow.create_window(400, 130, window=label_global)

mainwindow.create_window(360,380, window=byPipe)
mainwindow.create_window(50,380, window=muerte_label)

root.mainloop()




