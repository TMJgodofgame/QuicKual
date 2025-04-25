import re
import time
import speech_recognition as sr
import pyautogui
from number_parser import parser

# Inicialización
r = sr.Recognizer()

def abrir_word_y_nuevo_doc():
    print("🧠 Ejecutando operación especial: abrir Word y nuevo documento...")
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
    texto = re.sub(r'(\d+)\s*(?:coma|con)\s*(\d+)', r'\1,\2', texto, flags=re.IGNORECASE)
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

        if texto_lower == "operación matemáticas" or texto_lower == "operación matemática":
            abrir_word_y_nuevo_doc()
            return

        if texto_lower == "aceptar":
            print("✅ Acción: presionando Enter")
            pyautogui.press('enter')
            return

        texto_convertido = convertir_numeros_es(texto)
        print(f"📝 Convertido a: {texto_convertido}")

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
