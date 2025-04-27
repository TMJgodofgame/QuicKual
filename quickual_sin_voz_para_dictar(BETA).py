import re
import time
import speech_recognition as sr
import pyautogui
import pyperclip
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
    time.sleep(2)
    pyautogui.hotkey('enter')
    time.sleep(1)
    print("📝 Word debería estar listo.")

def quickual():
    pyautogui.hotkey('down', 'down', 'down', 'down', 'down')
    time.sleep(1)
    pyautogui.hotkey('alt')
    pyautogui.write('b2')
    pyautogui.write('y')
    print("Quickual debería estar listo.")

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

def ejecutar_movimiento_direccion(texto):
    texto = parser.parse(texto)  # Convertir palabras a números

    patrones = {
        'derecha': 'right',
        'izquierda': 'left',
        'arriba': 'up',
        'abajo': 'down',
        'borrar':'backspace',
        'suprimir': 'delete'
       }
    for palabra, tecla in patrones.items():
        match = re.search(r'(\d+)\s+' + palabra, texto)
        if match:
            cantidad = int(match.group(1))
            print(f"➡️ Moviendo {palabra} {cantidad} veces...")
            for _ in range(cantidad):
                pyautogui.press(tecla)
            return True
    return False

def escribir_texto(texto):
    pyperclip.copy(texto)  # Copia el texto
    pyautogui.hotkey('ctrl', 'v')  # Pega el texto

def escuchar_y_teclear():
    with sr.Microphone() as source:
        print("🎙️ Habla ahora...")
        audio = r.listen(source)

    try:
        texto = r.recognize_google(audio, language="es-ES")
        print(f"🔊 Has dicho: {texto}")

        texto_lower = texto.lower().strip()

        if ejecutar_movimiento_direccion(texto_lower):
            return

        if texto_lower in ["operación matemáticas", "operación matemática", "operación mate"]:
            abrir_word_y_nuevo_doc()
            return

        if texto_lower == "aceptar":
            print("✅ Acción: presionando Enter")
            pyautogui.press('enter')
            return

        if texto_lower in ["otra operación", "otra operación matemática", "otra vez"]:
            quickual()
            return

        # Convertir números en el texto
        texto_convertido = convertir_numeros_es(texto)
        print(f"📝 Convertido a: {texto_convertido}")

        # Borrar lo que había y escribir el nuevo texto
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        escribir_texto(texto_convertido)

        print("✅ Texto escrito correctamente.")
    except sr.UnknownValueError:
        print("😕 No entendí lo que dijiste.")
    except sr.RequestError as e:
        print(f"🚫 Error en el servicio de reconocimiento: {e}")

if __name__ == "__main__":
    print("▶️ Arranca y pon el foco en el campo de destino.")
    while True:
        escuchar_y_teclear()
