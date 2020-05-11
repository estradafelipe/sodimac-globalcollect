# Load libraries
from pandas import read_csv
from pandas import to_datetime
import openpyxl
import tkinter as tk
from tkinter import filedialog
from datetime import datetime, date, time, timedelta
import calendar
from PIL import Image


asl_path = " "
global_path = " "

#Genero pop up para obtener la ruta de los archivos de ASL y global
root= tk.Tk()

popup = tk.Canvas(root, width = 450, height = 400, bg = 'white', relief = 'raised')
popup.pack()

#Obtiene la ruta del archivo de ASL
def getCSVASL ():
  global asl_path
  asl_path = filedialog.askopenfilename()
  #asl_path = import_file_path
  okASL()
  print(asl_path)

#Obtiene la ruta del archivo de Globalcollect
def getCSVGlobal ():
  global global_path
  global_path = filedialog.askopenfilename()
  #global_path = import_file_path
  okGlobal()
  print(global_path)

def okASL():
    label_asl.config(text='OK', fg="green")

def okGlobal():
    label_global.config(text='OK', fg="green")

#Ejecuta el cruce
def ejecutarCruce():
  print (asl_path)
  print (global_path)
  # Load dataset
  #Defino las columnas
  names1 = ['MerchantID', 'ContractID',  'OrderID', 'EffortID',  'MerchantReference', 'PaymentReference',  'CustomerID', 'StatusID', 'StatusDescription', 'PaymentProduct ID', 'PaymentProductDescription',  'OrderCountryCode', 'OrderCurrencyCode',  'OrderAmount', 'RequestCurrencyCode',  'RequestAmount', 'PaidCurrency',  'PaidAmount',  'ReceivedDate',  'StatusDate',  'RejectionCode', 'Remarks']
  #Cargo el archivo
  globalcollect = read_csv(global_path,header=0 , names=names1, encoding='latin-1', index_col=False)

  #Defino las columnas
  names2 = ['Merchant',  'Orden',  'Suborden', 'F12',  'Tipo Factura', 'Terminal', 'Secuencia',  'Numero De Boleta', 'Canal',  'Fuente Abastecimiento',  'RUT',  'Nombre Cliente', 'Sexo', 'Direccion Cliente',  'Comuna', 'Ciudad', 'Region', 'Pais', 'Codigo Postal',  'Fono Comprador', 'Direccion Despacho', 'Comuna Receptor',  'Ciudad Receptor',  'Region Receptor',  'Pais Receptor',  'Cod. Postal Recep.', 'Fono Receptor',  'Metodo Envio', 'Codigo Linea', 'Descripcion Linea',  'Codigo Sublinea',  'Descripcion Sublinea', 'Codigo Clase', 'Descripcion Clase',  'Codigo Subclase',  'Descripcion Subclase', 'SKU',  'Descripcion SKU',  'Tamano', 'Regalo', 'Medio De Pago',  'Banco',  'Numero De Tarjeta',  'Numero Cuotas',  'Cantidad', 'Codigo Subclase2',  'Descripcion Subclase2', 'Fecha Coloc Orden',  'Hora Coloc Orden', 'Fecha Entrega Orden',  'Fecha Reparto',  'Fecha Pactada',  'Fecha Picking',  'Fecha Ruta', 'Fecha Validacion', 'Fecha Boleta', 'Estado ASL', 'Fec. Cambio Estado', 'Retencion Novios', 'Codigo Vendedor',  'Nombre Vendedor',  'Email',  'Precio Con Rebaja',  'Precio', 'Valor Flete',  'Total',  'Bulto',  'CodigoAutorizacion',  'OC Agrupada PF', 'Hora Validacion',  'Hora Boleta']
  #Cargo el archivo
  asl = read_csv(asl_path, delimiter=';',header=0, names=names2, encoding='latin-1')

  print("Globalcollect:")
  # shape
  print(globalcollect.shape)
  # head
  #print(globalcollect.head(20))

  print("ASL: ")
  # shape
  print(asl.shape)

  #asl.to_excel("asl.xlsx", "asl")
  # head
  #print(asl.head(20))

  #filtro que el orderid de global no esté en el archivo de ASL y que el effortID sea 1
  cruce2 = globalcollect[~globalcollect.OrderID.isin(asl.CodigoAutorizacion) & globalcollect.EffortID.isin(['1'])]

  #Ordeno por fecha
  cruceOrdenado = cruce2.sort_values(by='ReceivedDate')
  #agrego columna de fecha con hora argentina (ya que la hora de global es de holanda)
  cruceOrdenado['fechaArg'] = cruceOrdenado['ReceivedDate']
  #convierto la fecha de global a la fecha argentina
  cruceOrdenadoFecha = cruceOrdenado.apply(lambda x: to_datetime(x) if x.name == 'fechaArg' else x)
  cruceFechaArg = cruceOrdenadoFecha.apply(lambda x: x - timedelta(hours=5) if x.name == 'fechaArg' else x)

  #imprimo una muestra del cruce para validar en la consola
  print("Cruce: ")
  print(cruceFechaArg.shape)
  print(cruceFechaArg.head(35))
  #Genero export en excel
  cruceFechaArg.to_excel("cruce_soar.xlsx","cruce")

  return(0)

#Defino los botones y labels de la ventana

#labels de estado de carga del archivo (- si no lo cargó y OK si ya lo cargó)
label_asl = tk.Label(text='-',bg='white', fg='red', font=('helvetica', 12, 'bold'))
label_global = tk.Label(text='-',bg='white', fg='red', font=('helvetica', 12, 'bold'))

#Boton ruta asl
imagenASL = tk.PhotoImage(file="buttonASL.png")
buttonAsl_csv = tk.Button(text="Importar archivo de ASL", command=getCSVASL, bg='white', fg='white', font=('helvetica', 12, 'bold'),image=imagenASL, border=0)
#Boton ruta global
imagenGlobal = tk.PhotoImage(file="buttonGlobal.png")
buttonGlobal_csv = tk.Button(text="Importar archivo de GLOBALCOLLECT", command=getCSVGlobal, bg='white', fg='white', font=('helvetica', 12, 'bold'),image=imagenGlobal, border=0)
#Boton cruce
imagenCruce = tk.PhotoImage(file="buttonCruce.png")
buttonCruce = tk.Button(text="Generar cruce", command=ejecutarCruce, bg='white', fg='white', font=('helvetica', 12, 'bold'),image=imagenCruce, border=0)


#Mi logo :)
imagenBypipeok = tk.PhotoImage(file="bypipeoksize.png")
byPipe = tk.Label(text='Compiled by Pipe',bg='white', fg='black', font=('helvetica', 12, 'bold'), image=imagenBypipeok)

muerteok = tk.PhotoImage(file="muerteoksize.png")
muerte_label = tk.Label(bg='white', fg='black', image=muerteok)

popup.create_window(200, 60, window=buttonAsl_csv)
popup.create_window(200, 130, window=buttonGlobal_csv)
popup.create_window(200, 250, window=buttonCruce)

popup.create_window(400, 60, window=label_asl)
popup.create_window(400, 130, window=label_global)

popup.create_window(350,380, window=byPipe)
popup.create_window(100,380, window=muerte_label)

root.mainloop()




