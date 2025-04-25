import re
import time
import speech_recognition as sr
import pyttsx3
import pyautogui
from number_parser import parser

# Inicialización
engine = pyttsx3.init()
r = sr.Recognizer()

def speak(text):
    engine.say(text)
    engine.runAndWait()

def abrir_word_y_nuevo_doc():
    print("🧠 Ejecutando operación especial: abrir Word y nuevo documento...")
    speak("Abriendo Word para hacer una operación matemática.")
    time.sleep(1)
    pyautogui.press('win')
    time.sleep(0.5)
    pyautogui.write('word', interval=0.05)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey('enter')
    time.sleep(0.5)
    pyautogui.hotkey('alt')
    pyautogui.write('b2')
    pyautogui.write('y')
    print("📝 Word debería estar listo.")

def convertir_numeros_es(texto):
    # “5 coma 5” o “5 con 5” → “5,5”
    texto = re.sub(r'(\d+)\s*(?:coma|con)\s*(\d+)', r'\1,\2', texto, flags=re.IGNORECASE)

    # Separar tokens
    tokens = re.split(r'(\s+)', texto)
    resultado = []

    for tok in tokens:
        if tok.isspace():
            resultado.append(tok)
            continue
        if re.fullmatch(r'[\d\.,]+', tok):
            resultado.append(tok.replace('.', ''))
        elif re.search(r'[\d\.,]', tok):
            resultado.append(tok)
        else:
            resultado.append(parser.parse(tok))

    texto2 = ''.join(resultado)

    # Eliminar comas de miles, pero dejar las decimales
    texto2 = re.sub(r'(?<=\d),(?=\d{3}\b)', '', texto2)
    return texto2

def escuchar_y_teclear():
    with sr.Microphone() as source:
        print("🎙️ Habla ahora...")
        audio = r.listen(source)

    try:
        texto = r.recognize_google(audio, language="es-ES")
        print(f"🔊 Has dicho: {texto}")

        texto_lower = texto.lower().strip()

        # Acción especial: abrir Word
        if texto_lower == "operación matemáticas" or texto_lower == "operación matemática":
            abrir_word_y_nuevo_doc()
            return

        # Acción especial: pulsar Enter
        if texto_lower == "aceptar":
            print("✅ Acción: presionando Enter")
            speak("Aceptado")
            pyautogui.press('enter')
            return

        # Procesamiento normal
        texto_convertido = convertir_numeros_es(texto)
        print(f"📝 Convertido a: {texto_convertido}")
        speak(texto_convertido)

        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        pyautogui.write(texto_convertido, interval=0)

        print("✅ Texto tecleado.")
    except sr.UnknownValueError:
        print("😕 No entendí lo que dijiste.")
    except sr.RequestError as e:
        print(f"🚫 Error en el servicio de reconocimiento: {e}")

if __name__ == "__main__":
    print("▶️ Arranca y pon el foco en el campo de destino.")
    while True:
        escuchar_y_teclear()
