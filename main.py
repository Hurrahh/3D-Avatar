import speech_recognition as sr
import pyttsx3
import pyautogui
import os
import pyvda
from pyvda import VirtualDesktop, AppView
import win32gui
import time
import win32process
import win32con
import psutil
import subprocess
from langchain_ollama import OllamaLLM
from pixie import speech_to_text
from Huawei import predict

recognizer = sr.Recognizer()
engine = pyttsx3.init()

def llm(text):
    llm = OllamaLLM(model = 'llama3.2')
    response = llm.invoke(text)
    return response

apps = {
    "notepad": {
        "open": "notepad",
        "process": "notepad.exe"
    },
    "calculator": {
        "open": "calc",
        "process": "calculator.exe"
    },
    "vscode":{
        "open":"code",
        "process":"code.exe"
    }
}

def speak(text):
    engine.say(text)
    engine.runAndWait()

def listen_command(prompt="Listening..."):
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)
    try:
        command = recognizer.recognize_google(audio).lower()
        print(f"You said: {command}")
        return command
    except sr.UnknownValueError:
        speak("Sorry, I didn't get that.")
        return ""
    except sr.RequestError:
        speak("Network error.")
        return ""

def get_foreground_window():
    return win32gui.GetForegroundWindow()

def open_notepad():
    global notepad_pid
    os.system("start notepad")
    speak("Opening Notepad")
    time.sleep(1)
    hwnd = win32gui.GetForegroundWindow()
    _, notepad_pid = win32process.GetWindowThreadProcessId(hwnd)


def close_notepad():
    def enum_window_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "notepad" in title.lower():
                print(f"Closing window: {title}")
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                speak("Closed Notepad")

    win32gui.EnumWindows(enum_window_callback, None)


def open_application(app_name):
    if app_name in apps:
        command = apps[app_name]["open"]
        subprocess.Popen(command)
        speak(f"Opening {app_name}")
    else:
        speak(f"App {app_name} not found")

def close_application(app_name):
    if app_name in apps:
        proc_name = apps[app_name]["process"]
        found = False
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and proc_name.lower() in proc.info['name'].lower():
                proc.kill()
                found = True
        if found:
            speak(f"Closed {app_name}")
        else:
            speak(f"{app_name} was not running")
    else:
        speak(f"App {app_name} not found")

def gesture_control():
    speak("gesture control activated")
    response = predict.main()
    if response == "left":
        try:
            current = VirtualDesktop.current()
            VirtualDesktop(current - 1).go()
        except :
            print("No previous window")
    elif response == "right":
        try:
            current = VirtualDesktop.current()
            VirtualDesktop(current + 1).go()
        except:
            print("no forward window")

    elif response == "close":
        current_window = win32gui.GetForegroundWindow()
        win32gui.PostMessage(current_window, win32con.WM_CLOSE, 0, 0)
        speak("Closed current window")

    elif response == "":
        pass

def dictate_to_notepad():
    global typing_mode
    typing_mode = True
    speak("Dictation mode activated. Say 'end dictation' to stop.")

    punctuation_map = {
        "period": ".",
        "comma": ",",
        "question mark": "?",
        "exclamation mark": "!",
        "new line": "\n",
        "tab": "\t"
    }

    while typing_mode:
        command = listen_command()
        if not command:
            continue

        if "stop typing" in command:
            typing_mode = False
            speak("Dictation ended")
            break

        for word, symbol in punctuation_map.items():
            if word in command:
                command = command.replace(word, symbol)

        pyautogui.typewrite(command + " ")

def perform_action(command):
    # application control
    # if "open notepad" in command:
    #     open_notepad()
    # elif "close notepad" in command:
    #     close_notepad()
    if command.startswith("open "):
        app_name = command.replace("open ", "").strip()
        open_application(app_name)

    elif command.startswith("close "):
        app_name = command.replace("close ", "").strip()
        close_application(app_name)

    elif "activate gesture control" in command:
        gesture_control()

    elif 'start typing' in command:
        dictate_to_notepad()

    elif "take screenshot" in command:
        pyautogui.screenshot("screenshot.png")
        speak("Screenshot taken")

    elif "scroll down" in command:
        pyautogui.scroll(-500)

    # Virtual Desktop
    elif "switch window" in command:
        speak("Switching window")
        pyautogui.hotkey('alt', 'tab')

    elif "switch tab" in command:
        speak("Switching tab")
        pyautogui.hotkey('ctrl', 'tab')

    elif "previous tab" in command:
        speak("Going to previous tab")
        pyautogui.hotkey('ctrl', 'shift', 'tab')

    elif "next desktop" in command:
        current = VirtualDesktop.current()
        VirtualDesktop(current + 1).go()
        speak("Switched to next desktop")

    elif "previous desktop" in command:
        current = VirtualDesktop.current()
        VirtualDesktop(current - 1).go()
        speak("Switched to previous desktop")

    elif "move window to next desktop" in command:
        current = VirtualDesktop.current()
        current.move(VirtualDesktop(current + 1))
        speak("Moved window to next desktop")

    elif "new desktop" in command:
        pyvda.CreateDesktop()
        speak("New desktop created")

    elif "remove desktop" in command:
        current = pyvda.GetCurrentDesktopNumber()
        pyvda.RemoveDesktop(current)
        speak("Removed current desktop")

    elif "exit" in command:
        speak("Goodbye!")
        exit()

    else:
        speak("Command not recognized.")

while True:
    predict.main()
    cmd = listen_command()
    if cmd:
        perform_action(cmd)
