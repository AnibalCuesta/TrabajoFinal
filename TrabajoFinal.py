import pandas as pd
import smtplib as sm
from email.mime.text import MIMEText
from smtplib import SMTP
from tkinter import *
from tkinter import filedialog
#Coneccion
servidor = SMTP("smtp.gmail.com", 587)
servidor.ehlo()
servidor.starttls()

remitente = "oacuesta.csscpces@gmail.com" #Correo de la cuenta que envia el mail
#Inicio de sesion
servidor.login(remitente, "kejx cfii gjtu rgkm")

Ventana=Tk()
Ventana.title("Informe de Débitos Rechazados")
Ventana.iconbitmap(r"C:\Users\oacue\OneDrive\Documentos\Cursos\Itta Pynthon\Trabajo Final\correo.ico")
Ventana.geometry("800x400")
Ventana.resizable(False,False)

# Ventana.minsize(700,300)
# Ventana.maxsize(800,400)
#Ventana.iconbitmap(r"C:\Users\oacue\OneDrive\Documentos\Cursos\Itta Pynthon\Trabajo Final\ico_Ventana.ico")
Ventana.configure(bg="skyblue")
frame1=Frame(Ventana)
frame1.configure(height=400,width=200,background="red",bd=5)
frame1.pack()
label_busca=Label(frame1,text="Selecciona Archivo")
label_busca.config(font=("arial",12,"italic"))
label_busca.pack()

Buscar=Button(frame1,text="Buscar")
Buscar.pack()
# Variable que deseas mostrar en el campo Entry 
valor = StringVar() 
valor.set("")

entrada=Entry(frame1, textvariable=valor, state='disabled',width=80)
entrada.pack()

Enviar=Button(frame1,text="Enviar Correos")
Enviar.pack()

frame2=Frame()
frame2.configure(height=400,width=200,background="blue",bd=5)
frame2.pack(fill=X)

label_formato=Label(frame2,text="El Archivo Excel debe respetar el siguiente formato \n Columna1: Matricula \n Columna2: Nombre \n Columna3: Error \n Columna4: Mail")
label_formato.config(font=("arial",12,"bold"))
label_formato.pack()

label_informe=Label(frame2,text="")
label_informe.config(font=("arial",14,"bold"),bg="red")
label_informe.pack()

def boton_buscar():
    global Archivo
    Archivo = filedialog.askopenfilename( filetypes=(("Archivos de Excel", "*.xlsx"),) )
    if Archivo:
        try:
            valor.set(Archivo)
            entrada.config(state="normal")
            entrada.config(state="disabled")
            df = pd.read_excel(Archivo) # Mostrar las primeras filas del DataFrame en la etiqueta
            entrada.config(text=df.head().to_string())
        except Exception as e:
            entrada.config(text=f"Error al leer el archivo: {e}")    


def enviar_correos():

    #dataframe = pd.read_excel(r"C:\Users\oacue\OneDrive\Documentos\Cursos\Itta Pynthon\Trabajo Final\Prueba.xlsx")
    dataframe = pd.read_excel(Archivo)

    Matricula = dataframe['Matricula']
    Nombre = dataframe['Nombre']
    Rechazo = dataframe['Error']
    Mail = dataframe['Mail']
    ind=0

    for i in Matricula:
        # construir mensaje para enviar
        mensaje = (f"Buenos Días Cr.{i} {Nombre[ind]}\n"
                f"Informamos que su Debíto Automatico vino rechazado por :{Rechazo[ind]} Su pago mensual lo puede hacer por caja \n"
                f"Saludos Cordiales")
        #print(mensaje)

            # Configurar cuentas
        destinatario = Mail[ind] #Correo al que mandar el mail
        #Configuarar Mensaje
        mensaje_correo = MIMEText(mensaje)
        mensaje_correo["To"] = destinatario
        mensaje_correo["Subject"] = "Aviso débito Rechazado" # asunto del mail
        mensaje_correo["From"] = remitente

        # #Enviar mail ----------------comento hasta q termine si no manda correos con a¡cada pruebaç
        servidor.sendmail(remitente, destinatario, mensaje_correo.as_string())

        ind=ind+1
    label_informe.config(text="Correos Enviados")


Buscar.config(command=boton_buscar)
Enviar.config(command=enviar_correos)
Ventana.mainloop()