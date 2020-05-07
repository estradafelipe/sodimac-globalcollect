# Load libraries
from pandas import read_csv
import openpyxl
import tkinter as tk
from tkinter import filedialog


asl_path = " "
global_path = " "

#Genero pop up para obtener la ruta de los archivos de ASL y global
root= tk.Tk()

popup = tk.Canvas(root, width = 300, height = 300, bg = 'white', relief = 'raised')
popup.pack()

#Obtiene la ruta del archivo de ASL
def getCSVASL ():
  global asl_path
  asl_path = filedialog.askopenfilename()
  #asl_path = import_file_path
  print(asl_path)

#Obtiene la ruta del archivo de Globalcollect
def getCSVGlobal ():
  global global_path
  global_path = filedialog.askopenfilename()
  #global_path = import_file_path
  print(global_path)

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
  print("Cruce: ")
  print(cruceOrdenado.shape)
  print(cruceOrdenado.head(35))
  #Genero export en excel
  cruceOrdenado.to_excel("cruce_soar.xlsx","cruce")

  exit()

#Boton ruta asl
buttonAsl_csv = tk.Button(text="Import ASL CSV File", command=getCSVASL, bg='green', fg='white', font=('helvetica', 12, 'bold'))
#Boton ruta global
buttonGlobal_csv = tk.Button(text="Import GLOBAL CSV File", command=getCSVGlobal, bg='green', fg='white', font=('helvetica', 12, 'bold'))
#Boton cruce
buttonCruce = tk.Button(text="Generar cruce", command=ejecutarCruce, bg='red', fg='white', font=('helvetica', 12, 'bold'))

popup.create_window(150, 60, window=buttonAsl_csv)
popup.create_window(150, 100, window=buttonGlobal_csv)
popup.create_window(150, 160, window=buttonCruce)

root.mainloop()




