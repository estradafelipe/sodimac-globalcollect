# Load libraries
from pandas import read_csv
from pandas import merge
from pandas.plotting import scatter_matrix
from matplotlib import pyplot
from sklearn.model_selection import train_test_split
from sklearn.model_selection import cross_val_score
from sklearn.model_selection import StratifiedKFold
from sklearn.metrics import classification_report
from sklearn.metrics import confusion_matrix
from sklearn.metrics import accuracy_score
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
from sklearn.naive_bayes import GaussianNB
from sklearn.svm import SVC
import openpyxl

# Load dataset
#url = "C:/Users/feliestrada/Desktop/report.csv"
names1 = ['MerchantID', 'ContractID',  'OrderID', 'EffortID',  'MerchantReference', 'PaymentReference',  'CustomerID', 'StatusID', 'StatusDescription', 'PaymentProduct ID', 'PaymentProductDescription',  'OrderCountryCode', 'OrderCurrencyCode',  'OrderAmount', 'RequestCurrencyCode',  'RequestAmount', 'PaidCurrency',  'PaidAmount',  'ReceivedDate',  'StatusDate',  'RejectionCode', 'Remarks']
globalcollect = read_csv(r"C:/Users/feliestrada/Desktop/report.csv",header=0 , names=names1, encoding='latin-1', index_col=False)

names2 = ['Merchant',  'Orden',  'Suborden', 'F12',  'Tipo Factura', 'Terminal', 'Secuencia',  'Numero De Boleta', 'Canal',  'Fuente Abastecimiento',  'RUT',  'Nombre Cliente', 'Sexo', 'Direccion Cliente',  'Comuna', 'Ciudad', 'Region', 'Pais', 'Codigo Postal',  'Fono Comprador', 'Direccion Despacho', 'Comuna Receptor',  'Ciudad Receptor',  'Region Receptor',  'Pais Receptor',  'Cod. Postal Recep.', 'Fono Receptor',  'Metodo Envio', 'Codigo Linea', 'Descripcion Linea',  'Codigo Sublinea',  'Descripcion Sublinea', 'Codigo Clase', 'Descripcion Clase',  'Codigo Subclase',  'Descripcion Subclase', 'SKU',  'Descripcion SKU',  'Tamano', 'Regalo', 'Medio De Pago',  'Banco',  'Numero De Tarjeta',  'Numero Cuotas',  'Cantidad', 'Codigo Subclase2',  'Descripcion Subclase2', 'Fecha Coloc Orden',  'Hora Coloc Orden', 'Fecha Entrega Orden',  'Fecha Reparto',  'Fecha Pactada',  'Fecha Picking',  'Fecha Ruta', 'Fecha Validacion', 'Fecha Boleta', 'Estado ASL', 'Fec. Cambio Estado', 'Retencion Novios', 'Codigo Vendedor',  'Nombre Vendedor',  'Email',  'Precio Con Rebaja',  'Precio', 'Valor Flete',  'Total',  'Bulto',  'CodigoAutorizacion',  'OC Agrupada PF', 'Hora Validacion',  'Hora Boleta']
asl = read_csv(r"C:/Users/feliestrada/Desktop/InfVtasAcum_WEB.csv", delimiter=';',header=0, names=names2, encoding='latin-1')

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

cruce2 = globalcollect[~globalcollect.OrderID.isin(asl.CodigoAutorizacion) & globalcollect.EffortID.isin(['1'])]

cruceOrdenado = cruce2.sort_values(by='ReceivedDate')
print("Cruce: ")
print(cruceOrdenado.shape)
print(cruceOrdenado.head(35))


cruceOrdenado.to_excel("cruce_soar2.xlsx","cruce")



