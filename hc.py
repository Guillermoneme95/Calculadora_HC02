import pandas as pd
from pandas import ExcelWriter
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.dates as dates
from datetime import datetime

def eliminar_lista(lista):
    longitud= len(lista)
    while longitud > 0:
        lista.pop(indice)
        longitud =  longitud - 1

def control_int(i):
    try:
        a = int(i)
        return a
    except ValueError:
        print("Ingrese un numero entre 0-6")
         
            
fechayhora = datetime.now()

#Ingreso de datos por teclado

print("Hola! Esta es una herramienta para el cálculo de la huella de carbono de una organización\n")

print("La huella de carbono identifica la cantidad de emisiones de GEI que son liberadas a la atmósfera como consecuencia del desarrollo de cualquier actividad.Permite identificar todas las fuentes de emisiones de GEI y establecer a partir de este conocimiento, medidas de reducción efectivas.\n ")

print("Para realizar el calculo de la huella de carbono vamos a especificar los alcances que tendremos en cuenta:\n")

print("Alcance 1:\t Emisiones directas de GEI\n")

print("Alcance 2:\t Emisiones indirectas de GEI asociadas a la generación de electricidad adquirida y consumida por la organización\n")

#Parámetros iniciales

empresa_name = 'No asignado'

actividad = 'No asignado'

act1= []

act2=[]

n_act1=[]

n_act2=[]

act_total=[]

n_act_total=[]

menuprincipal = 0

menu= 0

consumo= 0

litros = 0

string= ''
#Menu principal

#Verificación del valor de entrada

menu = input("Menu de operaciones: \n 1-Agregar Nombre de la empresa \n 2- Especificar la actividad a la que se dedica \n 3-Agregar Alcance 1 \n 4-Agregar Alcance 2 \n 5-Eliminar Tablas de Alcance \n 6- Ver planilla completa\n 0- Finalizar e imprimir informe\nSELECCIÓN:")

#Control del ingreso de datos

menuprincipal = control_int(menu)

print(".................................................................\n")

#Main loop
while menuprincipal !=0:

    if menuprincipal == 1:

        empresa_name=input("Ingese el nombre de la empresa: ")

    elif menuprincipal == 2:

        actividad= input("Especifique la actividad a la que se dedica la empresa: ")


    elif menuprincipal == 3:

        print("Agregar item de Alcance 1")
        try:

            nuevo_nombre1= str(input("Ingrese nombre de elemento: "))

            cantidad_hs =float(input("Ingrese la cantidad de horas diarias de dicha actividad: "))
            
            cantidad_semanas =float(input("Ingrese la cantidad de semanas de dicha actividad: "))

            uso_mensual= float(input("Ingrese frecuencia de uso mensual: "))

            cantidad_elem= float(input("Ingrese la cantidad de elementos: "))

            cantidad_litros= float(input("Ingrese la cantidad de litros de combustible empleados por dicha actividad: "))
                                    
            factor_em= float(input("Ingrese el factor de emision de dicha actividad en [KgC02eq/l] : "))

            n_act1.append(nuevo_nombre1)

            n_act_total.append(nuevo_nombre1)

            nuevo_valor1= cantidad_hs * factor_em * uso_mensual * cantidad_elem * cantidad_litros * cantidad_semanas

            litros= litros + (cantidad_hs  * uso_mensual * cantidad_elem * cantidad_litros * cantidad_semanas) 
            
            act1.append(nuevo_valor1)

            act_total.append(nuevo_valor1)

            bandera = 0

        except:
            print("El dato ingresado no corresponde con la caracteristica del dato solicitado")
            bandera = 1

        finally:
            if (bandera == 0):
                print("Elemento {} fue agregado exitosamente a la tabla de Alcance 1".format(nuevo_nombre1))



    elif menuprincipal == 4:

        print("Agregar item de Alcance 2")

        try:
            nuevo_nombre2= str(input("Ingrese nombre de elemento: "))

            cantidad_horas= float(input("Ingrese la cantidad de horas de uso semanales: "))

            cantidad_dias= float(input("Ingrese la cantidad de dias de uso semanales: "))

            cantidad_semanas= float(input("Ingrese la cantidad de semanas al año que se utiliza: "))

            consumo_prom= float(input("Ingrese el consumo eléctrico promedio: "))

            cantidad_elementos= float(input("Ingrese la cantidad de elementos: "))

            factor_emision = float(input("Ingrese el factor de emision en [KgC02eq/kwh]"))

            n_act2.append(nuevo_nombre2)

            n_act_total.append(nuevo_nombre2)

            nuevo_valor2= consumo_prom * cantidad_elementos * cantidad_horas * cantidad_dias * factor_emision * cantidad_semanas

            consumo= consumo + (consumo_prom * cantidad_elementos * cantidad_horas * cantidad_dias * cantidad_semanas)
            
            act2.append(nuevo_valor2)

            act_total.append(nuevo_valor1)

            bandera = 0

        except:

            print("El dato ingresado no corresponde con la caracteristica del dato solicitado")
            bandera = 1

        finally:
            if( bandera == 0):
                print("Elemento {} fue agregado exitosamente a la tabla de Alcance 2".format(nuevo_nombre2))



    elif menuprincipal == 5:

        try:
            eliminar_lista(n_act1)
            eliminar_lista(act1)
            eliminar_lista(n_act2)
            eliminar_lista(act2)
            eliminar_lista(n_act_total)
            eliminar_lista(act_total)
            consumo=0 
            litros=0 
            
        except:

            print("Error al eliminar la tabla")

        else:

            print("La Tabla de Alcance 1 Fue eliminada exitosamente")



    elif menuprincipal == 6:

        print("Mostrando planilla completa")

        print("Nombre de la empresa: {}\n".format(empresa_name))

        print("Actividad:{}\n".format(actividad))

        dataframe_act1 = pd.DataFrame(act1)

        dataframe_n1=pd.DataFrame(n_act1)

        dataframe1= pd.concat([dataframe_n1,dataframe_act1],axis=1)

        dataframe_act2 = pd.DataFrame(act2)

        dataframe_n2= pd.DataFrame(n_act2)

        dataframe2= pd.concat([dataframe_n2,dataframe_act2],axis=1)

        print("Alcance 1\n")

        print(dataframe1)

        print("\n")

        print("Alcance 2\n")

        print(dataframe2)


    elif menuprincipal == 0:

        print("Informe de actividad")


    else:

        print("Por favor ingrese una opción correcta")

#Comprobación de ingreso de datos interna
    
    menu = input("Menu de operaciones: \n 1-Agregar Nombre de la empresa \n 2- Especificar la actividad a la que se dedica \n 3-Agregar Alcance 1 \n 4-Agregar Alcance 2 \n 5-Eliminar Tablas de Alcance \n 6- Ver planilla completa\n 0- Finalizar e imprimir informe\nSELECCIÓN:")      
    menuprincipal=control_int(menu)
    print("\n\n")
   
            

Titulo = 'Empresa: {}\nActividad: {}'.format(empresa_name, actividad)

#Graficos

fig, (ax1,ax2,ax3)= plt.subplots(3, figsize=(10, 8))

plt.title(Titulo)

#Grafico Alcance 1
ax1.pie(act1, labels=n_act1, autopct='%1.1f%%',

        shadow=True, startangle=90)

ax1.axis('equal') 

ax1.set_title('Alcance 1') 

#Grafico Alcance 2
ax2.pie(act2,labels=n_act2, autopct='%1.1f%%',

        shadow=True, startangle=90)

ax2.axis('equal')  

ax2.set_title('Alcance 2')

#Grafico Total
ax3.pie(act_total ,labels=n_act_total, autopct='%1.1f%%',

        shadow=True, startangle=90)

ax3.axis('equal')  

ax3.set_title('Total')

#Guardado de imagen como .png
plt.savefig('Informe.png')

#Calculo de HC
sumatoria1 = 0

sumatoria2 = 0


for x in range(len(act1)):

    sumatoria1= sumatoria1 + act1[x]
    

for x in range(len(act2)):

    sumatoria2= sumatoria2 + act2[x]


sumatoria_total= sumatoria1 + sumatoria2


#Gestion de Data Frames

dataframe_act1 = pd.DataFrame(act1)

dataframe_n1=pd.DataFrame(n_act1)

dataframe1= pd.concat([dataframe_n1,dataframe_act1],axis=1)

dataframe_act2 = pd.DataFrame(act2)

dataframe_n2= pd.DataFrame(n_act2)

dataframe2= pd.concat([dataframe_n2,dataframe_act2],axis=1)

df=pd.concat([dataframe1,dataframe2],axis=0)

#Muestra de Data Frames
print(Titulo)

print("\n")
print(df)

#Conversion de Data Frame a .xlsx
df.to_excel("informe.xlsx")


#Creacion y escritura de informe.txt

hora = fechayhora.hour
minuto = fechayhora.minute
segundo = fechayhora.second
microsegundo = fechayhora.microsecond
año= fechayhora.year
mes= fechayhora.month
dia = fechayhora.day

file=open("Informe_escrito.txt","w")

file.write("Empresa: ")

file.write(empresa_name)

file.write("\n")

file.write("Actividad que realiza: ")

file.write(actividad)

file.write("\n")

file.write("Calculo de huella de carbono con alcance 1+2: ")

file.write(str(sumatoria_total) + "[KgC02eq]")

file.write("El consumo en litros de combustible fosil es: {} [L]\n".format(litros))

file.write("El consumo de potencia eléctrica en [KW] es: {} [KW]\n".format(consumo))

file.write("Fecha de emisión: {}/{}/{}\n".format(dia, mes , año))

file.write("Hora de emisión: {}:{}:{} \n".format(hora, minuto , segundo))

file.write("\n\n")

file.write("---------------------------------------------------------------------------------------------------------------\n")

file.write("Este informe es el primer paso para la obtención de datos para iniciar acciones, el mismo se complementa con los gráficos previstos por la misma herramienta \n")

file.write("El calculo de la Huella de carbono para esta organizacion fue realizado por el grupo '13' con el objetivo de completar el trabajo práctico numero 2 \nAsignatura: Gestión Ambiental, Salud Ocupacional y Seguridad.\n\n")

file.write("Integrantes:\n\n")

file.write("Neme Simonetti, Guillermo Sebastian \n Coca, Luis Rogelio \n Lopez, Hector Ramon \n Gómez Morales Sergio Ismael\n\n\n")

file.write("¡Muchas gracias!")

#Muestra de graficas 
plt.show()