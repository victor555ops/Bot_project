import speech_recognition as sr
import webbrowser
import tkinter as tk
import pyttsx3
import datetime 
import requests 
import json 
import sys 
import subprocess
import pywhatkit

edad_bot = 1
nombre_bot = 2
respuesta = 3


r = sr.Recognizer()

engine = pyttsx3.init()

engine.setProperty("language", "es-ES")
engine.setProperty("rate", 150)

def obtener_hora():
  
  hora = datetime.datetime.now().time()
 
  hora_str = hora.strftime("%H:%M")
  
  return hora_str



def abrir_segunda():
    
  ventana.withdraw()
  
  ventana_segunda = tk.Toplevel()
  ventana_segunda.title("Segunda ventana")
  ventana.configure(bg="")
  ventana_segunda.geometry("550x300")
   
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
    audio = r.listen(source)
 
    try:
    
      texto = r.recognize_google(audio, language="es-ES")

      print("Has dicho: " + texto)
     
      if "youtube" in texto.lower() or "reproduce" in texto.lower():
      
        musica = texto.lower().split("youtube")[1].strip()
        musica = texto.lower().split("reproduce")[1].strip()
        
        engine.say("Reproduciendo" + musica)
        engine.runAndWait()
      
        print("Reproduciendo" + musica)
        pywhatkit.playonyt(musica)
   
      elif "abrir word" in texto.lower() or "Cerrar word" in texto.lower():
      
        webbrowser.open("C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE")

        engine.say("Abriendo Word")
        engine.runAndWait()
        
        engine.say("Cerrando Word")
        engine.runAndWait()
      
        print("Abriendo Word")
        print("Cerrando Word")
      
      
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
      
      elif "nombre" in texto.lower() or "Cual es tu nombre" in texto.lower():
        
        engine.say("No tengo un nombre definido, pero usted puede decirme como guste o ponerme un nombre, sera un placer ayudarlo.")
        engine.runAndWait()
        
        print(f"No tengo un nombre definido, pero usted puede decirme como guste o ponerme un nombre, sera un placer ayudarlo.")
        
      elif "gracias" in texto.lower():
        
        engine.say("Claro, es un placer atenderte")
        engine.runAndWait()
      
      elif "clima" in texto.lower() or "tiempo" in texto.lower():

      
        api_key = "ab613795bd25da5f26bb66bfcc72d475"

        lugar = "Veracruz"

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
      
      else:
        
        engine.say("no he entendido lo que has dicho ")
        engine.runAndWait()
        print("No he entendido lo que has dicho")
 
    except sr.UnknownValueError:
   
      print("No se ha podido entender el audio")



ventana = tk.Tk()

ventana.title("Bot Project")
ventana.configure(bg="peach puff")
ventana.geometry("235x150")

etiqueta = tk.Label(ventana, text="Bienvenido")

etiqueta.pack()

boton = tk.Button(ventana, text="Presiona para hablar", command=escuchar_y_procesar)

boton.pack()

boton_1 = tk.Button(ventana, text="Ayuda", command=abrir_segunda)

boton_1.pack()


ventana.mainloop()

saludar_usuario()
