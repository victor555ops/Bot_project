import speech_recognition as sr 
import webbrowser
import tkinter as tk
import pyttsx3
import datetime 
import requests #Para los URL
import sys 
import pywhatkit
import os
from PIL import ImageTk, Image

r = sr.Recognizer()
engine = pyttsx3.init()

engine.setProperty("language", "es-ES")
engine.setProperty("rate", 145)


def obtener_hora():
  hora = datetime.datetime.now().time()

  hora_str = hora.strftime("%I:%M %p")
  return hora_str

print(obtener_hora())


def abrir_segunda():
  ventana.withdraw() 
  
  ventana_segunda = tk.Toplevel() #ayuda a manejar mas ventanas despues de la primera
  ventana_segunda.title("Segunda ventana")
  ventana_segunda.geometry("550x300")
  imagen2 = Image.open("Sabana.jpg")

  foto2 = ImageTk.PhotoImage(imagen2)
  
  marco = tk.Frame(ventana)
  marco.pack()
  
  etiqueta = tk.Label(marco, image=foto2)
  etiqueta.pack()
   
  etiqueta_2 = tk.Label(ventana_segunda, text="Bienvenido a la seccion de ayuda.\nEste programa puede hacer lo siguiente con solo picar el boton de hablar:\n- Abrir la paqueteria de office. Por ejemplo: Word.\n- Te dira la hora, fecha y estado del clima.\n- Reproduce videos en youtube.\n- Abrir google o hacer la busqueda por ti, solo tienes que pedirle tu busqueda.\nhacer preguntas interactivas.\n-por ultimo, puedes cerrar el programa diciendo cerrar programa.\n si quieres volver a la primera ventana, preciona el boton de abajo")
  etiqueta_2.pack()

  boton_2 = tk.Button(ventana_segunda, text="Volver a la primera ventana", command=lambda: volver_primera(ventana, ventana_segunda))
  boton_2.pack()


def volver_primera(ventana, ventana_segunda): 
  ventana.deiconify()
  ventana_segunda.destroy()



def obtener_fecha():
  fecha = datetime.datetime.now().date()
  fecha_str = fecha.strftime("%d/%m/%Y")
  return fecha_str


def abrir_explorador():
  os.startfile("explorer.exe")

  
def crear_carpeta(nombre):
  escritorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'desktop')
  carpeta = os.path.join(escritorio, nombre)
   
  os.makedirs(carpeta, exist_ok=True)
  
  
def abrir_carpeta(nombre):
  carpeta = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', nombre)
  
  if os.path.exists(carpeta):
    os.startfile(carpeta)
  else:
    print(f"La carpeta {nombre} no existe en el escritorio.")
    engine.say(f"La carpeta {nombre} no existe en el escritorio")
    engine.runAndWait()



def saludar_usuario():
  hora = datetime.datetime.now().time()

  if hora.hour < 12:
    saludo = "Hasta Pronto, que tenga un excelente dia, si necesita ayuda, no dude en contactarme"
  elif hora.hour < 18:
    saludo = "Hasta pronto, que tenga una excelente tarde, si necesita ayuda, no dude en contactarme"
  else:
    saludo = "Hasta pronto, que tenga una linda noche, si necesita ayuda no dude en contactarme"
  
  engine.say(saludo)
  engine.runAndWait()
 
  print(saludo)


def escuchar_y_procesar():
  
  with sr.Microphone() as source:
  
    r.adjust_for_ambient_noise(source)
 
    print("Escuchando...")
    audio = r.listen(source,) 
 
    try:
    
      texto = r.recognize_google(audio, language="es-ES")

      print("Has dicho: " + texto)
     
      if "youtube" in texto.lower():
      
        musica = texto.lower().split("youtube")[1].strip()
        
        engine.say("Reproduciendo " + musica)
        engine.runAndWait()
      
        print("Reproduciendo " + musica, "en youtube")
        pywhatkit.playonyt(musica)
   
      elif "abrir word" in texto.lower() or "Cerrar word" in texto.lower():
      
        webbrowser.open("C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE")

        engine.say("Abriendo Word")
        engine.runAndWait()
      
        print("Abriendo Word")
        engine.runAndWait()
      
      
      elif "excel" in texto.lower():
       
        webbrowser.open("C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE")
       
        engine.say("Abriendo Excel")
        engine.runAndWait()
      
        print("Abriendo Excel")
      
      elif "powerpoint" in texto.lower():
       
        webbrowser.open("C:/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE")
        
        engine.say("Abriendo PowerPoint")
        engine.runAndWait()
   
        print("Abriendo PowerPoint")
      
      elif "google" in texto.lower():
     
        if len(texto.lower().split("google")) > 1:
         
          busqueda = texto.lower().split("google")[1].strip()
      
          url = "https://www.google.com/search?q=" + busqueda
        else:
         
          url = "https://www.google.com"
       
        webbrowser.open(url)
        
        engine.say("abriendo Google con la busqueda " + busqueda)
        engine.runAndWait()
      
        print("Abriendo Google con la búsqueda: " + busqueda)
      
      elif "facebook" in texto.lower():
        
        webbrowser.open("https://www.facebook.com")
        engine.say("Abriendo facebook")
        engine.runAndWait()
        
      elif "whatsapp" in texto.lower():
        
        webbrowser.open("https://web.whatsapp.com")
        engine.say("Abriendo whatsApp")
        engine.runAndWait()
        
      elif "instagram" in texto.lower():
        
        webbrowser.open("https://www.instagram.com")
        engine.say("Abriendo Instagram")
        engine.runAndWait()
      
      elif "twitter" in texto.lower():
       
        webbrowser.open("https://twitter.com")
        engine.say("Abriendo twitter")
        engine.runAndWait()
        
      elif "tiktok" in texto.lower():
        
        webbrowser.open("https://www.tiktok.com/login?lang=es&redirect_url=https%3A%2F%2Fwww.tiktok.com%2Fupload%3Flang%3Des")
        engine.say("Abriendo tiktok")
        engine.runAndWait()
          
      elif "hora" in texto.lower():
   
        hora = obtener_hora()
        
        engine.say("La hora actual es " + hora)
        engine.runAndWait()
       
        print("La hora actual es " + hora)
      
      elif "fecha" in texto.lower():
   
        fecha = obtener_fecha()
    
        engine.say("La fecha actual es " + fecha)
        engine.runAndWait()
       
        print("La fecha actual es " + fecha)
      
      elif "cerrar programa" in texto.lower():
     
        saludar_usuario()
      
        sys.exit()
        
      elif "edad" in texto.lower() or "cuántos años tienes" in texto.lower():
      
        engine.say("Soy un modelo de asistente o bot creado desde visual studio, por lo tanto no tengo una edad definida, simplemente la fecha de mi creacion")
        engine.runAndWait()
      
        print(f"Soy un modelo de asistente o bot creado desde visual studio, por lo tanto no tengo una edad defininida, simplemente la fecha de mi creacion")
      
      elif "gracias" in texto.lower():
        
        engine.say("Claro, es un placer atenderte")
        engine.runAndWait()
      
      elif "clima" in texto.lower() or "tiempo" in texto.lower():

      
        api_key = "ab613795bd25da5f26bb66bfcc72d475"

        lugar = "veracruz"

        url = f"http://api.openweathermap.org/data/2.5/weather?q={lugar}&appid={api_key}&units=metric&lang=es"

        respuesta = requests.get(url)

        datos = respuesta.json()

        print(f"El clima en {lugar} es {datos['weather'][0]['description']}")
        print(f"La temperatura actual es {datos['main']['temp']} °C")
        print(f"La humedad relativa es {datos['main']['humidity']} %")
        print(f"La presión atmosférica es {datos['main']['pressure']} hPa")

        engine.say(f"El clima en {lugar} es {datos['weather'][0]['description']}")
        engine.say(f"La temperatura actual es {datos['main']['temp']} grados centígrados")
        engine.say(f"La humedad relativa es {datos['main']['humidity']} por ciento")
        engine.say(f"La presión atmosférica es {datos['main']['pressure']} hectopascales")
        engine.runAndWait()
      
      elif "abrir explorador" in texto.lower() or "abre explorador" in texto.lower():
        abrir_explorador()
        engine.say("Abriendo tu explorador de archivos")
        engine.runAndWait()
      
      elif "crear carpeta" in texto.lower() or "crea carpeta" in texto.lower():
        nombre = texto.split()[-1]
        crear_carpeta(nombre)
        engine.say(f"Creando la carpeta {nombre} en el escritorio")
        engine.runAndWait()
        
      elif "Abrir carpeta" in texto.lower():
        nombre = texto.split()[-1]
        abrir_carpeta(nombre)
        engine.say(f"Abriendo la carpeta {nombre} en el explorador de archivos")
        engine.runAndWait()
        
      else:
          engine.say("no he entendido lo que has dicho ")
          engine.runAndWait()
          print("No he entendido lo que has dicho")
 
    except sr.UnknownValueError:
   
      print("No se ha podido entender el audio")
      engine.say("No se ha podido entender el audio")
      engine.runAndWait()


ventana = tk.Tk()
ventana.title("BotAssist")
ventana.geometry("300x250")

imagen = Image.open("paisaje.jpg")
foto = ImageTk.PhotoImage(imagen)

marco = tk.Frame(ventana)
marco.pack()

etiqueta = tk.Label(marco, image=foto)
etiqueta.pack()

tamaño_icono = (45, 45) 
icono_hablar_img = Image.open("Mic.png")
icono_hablar_img = icono_hablar_img.resize(tamaño_icono)
icono_hablar = ImageTk.PhotoImage(icono_hablar_img)

icono_ayuda_img = Image.open("interrogacion.jpg")
icono_ayuda_img = icono_ayuda_img.resize(tamaño_icono)
icono_ayuda = ImageTk.PhotoImage(icono_ayuda_img)

boton = tk.Button(marco, image=icono_hablar, command=escuchar_y_procesar)
boton_1 = tk.Button(marco, image=icono_ayuda, command=abrir_segunda)

boton.place(x=10, y=60)
boton_1.place(x=230, y=180)

etiqueta_hablar = tk.Label(marco, text="Presiona para hablar")
#etiqueta_ayuda = tk.Label(marco, text="Ayuda")

etiqueta_hablar.place(x=10, y=37)
#etiqueta_ayuda.place(x=200, y=157)

ventana.mainloop()

saludar_usuario()
