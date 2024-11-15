import speech_recognition as sr
import os
import subprocess
import pywhatkit as kit
import pyautogui
import webbrowser
import keyboard
import streamlit as st
import google.generativeai as genai
import pyttsx3
import base64
import threading

recognizer = sr.Recognizer()
engine = pyttsx3.init()
lock = threading.Lock()

GOOGLE_API_KEY = "AIzaSyDZ6yDuQgQWxzc5Qq24Dpf_BkvcOjx_SP8"
genai.configure(api_key=GOOGLE_API_KEY)

geminiModel = genai.GenerativeModel("gemini-pro")
chat = geminiModel.start_chat(history=[])

def get_gemini_response(query):
    try:
        instantResponse = chat.send_message(query, stream=True)
        response_text = ' '.join([outputChunk.text for outputChunk in instantResponse if hasattr(outputChunk, 'text')])
        return response_text
    except AttributeError as e:
        st.error(f"Error in Gemini API response: {str(e)}")
    except Exception as e:
        st.error(f"Unexpected error: {str(e)}")

st.sidebar.title('Help Menu')
st.sidebar.write('This is JARVIS, your personal assistant. Here are some ways to use it:')
st.sidebar.markdown("""
- **Open Application**: Type an app name to open it.
- **Open Website**: Enter a URL to open a website.
- **Chat**: Ask JARVIS anything!
- **Tell Joke**: Hear a joke to lighten the mood.
- **Text-to-Speech**: Convert your text to speech output.
""")

st.sidebar.title('Features Menu')
st.sidebar.write('JARVIS offers the following features:')
st.sidebar.markdown("""
- Application Launcher
- Website Launcher
- Joke Teller
- Text-to-Speech Converter
""")

background_image_path = os.path.join(os.path.dirname(__file__), "jarvis.png")
logo_image_path = os.path.join(os.path.dirname(__file__), "jarvis1.jpg")

def image_to_base64(image_path):
    try:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    except FileNotFoundError:
        st.error(f"Image file '{image_path}' not found.")
        return None

background_image_base64 = image_to_base64(background_image_path)
logo_image_base64 = image_to_base64(logo_image_path)

if background_image_base64:
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url("data:image/png;base64,{background_image_base64}");
            background-size: cover;
            background-position: center;
            color: #FFFFFF;
            font-family: 'Arial', sans-serif;
        }}
        .stButton > button {{
            background-color: #ff6f91;
            color: white;
            font-size: 16px;
            border-radius: 8px;
            transition: all 0.2s ease-in-out;
        }}
        h1, h2, h3, h4, h5, h6, p {{
            color: white;
        }}
        .stButton > button:hover {{
            background-color: #ff4d7e;
        }}
        .stSidebar {{
            background: #020024;
            color: white;
        }}
        .chatbox {{
            border: 2px solid #ff6f91;
            border-radius: 10px;
            padding: 10px;
            margin: 10px 0;
            background-color: #f8f9fa;
            max-height: 300px;
            overflow-y: scroll;
        }}
        .user-msg {{
            text-align: right;
            color: #007bff;
            margin-bottom: 5px;
        }}
        .jarvis-msg {{
            text-align: left;
            color: #28a745;
            margin-bottom: 5px;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

if logo_image_base64:
    st.markdown(
        f"""
        <div style="text-align: center;">
            <img src="data:image/jpg;base64,{logo_image_base64}" alt="JARVIS Logo" style="width: 200px;">
        </div>
        """,
        unsafe_allow_html=True
    )

st.title('JARVIS - Your Personal Assistant ')

if 'chat_history' not in st.session_state:
    st.session_state['chat_history'] = []

universal_app_names = {
    'word': 'winword.exe',
    'excel': 'excel.exe',
    'powerpoint': 'powerpnt.exe',
    'chrome': 'chrome.exe',
    'notepad': 'notepad.exe',
    'vscode': 'code.exe',
    'paint': 'mspaint.exe',
    'outlook': 'outlook.exe',
    'firefox': 'firefox.exe',
    'edge': 'msedge.exe',
    'calculator': 'calc.exe',
    'command prompt': 'cmd.exe',
    'control panel': 'control.exe',
    'file explorer': 'explorer.exe',
    'onenote': 'onenote.exe',
    'teams': 'teams.exe',
    'skype': 'skype.exe',
    'spotify': 'spotify.exe',
    'notepad++': 'notepad++.exe',
    'sublime text': 'sublime_text.exe',
    'pycharm': 'pycharm64.exe',
    'android studio': 'studio64.exe',
    'eclipse': 'eclipse.exe',
    'putty': 'putty.exe',
    'mysql workbench': 'mysqlworkbench.exe',
    'xampp': 'xampp-control.exe',
    'task manager': 'taskmgr.exe',
    'windows media player': 'wmplayer.exe',
    'microsoft store': 'WinStore.App.exe',
    'visual studio': 'devenv.exe',
    'netbeans': 'netbeans.exe',
    'figma': 'figma.exe',
    'android emulator': 'emulator.exe',
    'xbox game bar': 'gamebar.exe',
    'teamviewer': 'TeamViewer.exe',
    'anydesk': 'AnyDesk.exe',
    'settings':'SystemSettings.exe',
    'camera': 'camera.exe',
    'alarm': 'alarmclock.exe',
    'calendar': 'wlcalendar.exe',
    'microsoft edge canary': 'msedge_canary.exe',
    'paint 3d': 'paint3d.exe',
    'wordpad': 'wordpad.exe',
    'snipping tool': 'SnippingTool.exe',
    'microsoft whiteboard': 'Whiteboard.exe',
    'google drive': 'googledrivesync.exe',
    'dropbox': 'Dropbox.exe',
    'onedrive': 'OneDrive.exe',
    'powershell': 'powershell.exe',
    'power automate': 'PowerAutomate.exe',
    'onenote 2016': 'ONENOTE.exe',
    'visual studio code insiders': 'code-insiders.exe',
    'microsoft access': 'msaccess.exe',
    'microsoft publisher': 'MSPUB.exe',
    'adobe acrobat': 'Acrobat.exe',
    'chatgpt': 'ChatGPT.exe',
}

def open_application(app_name):
    app_path = universal_app_names.get(app_name.lower())
    if app_path:
        try:
            os.startfile(app_path)
        except Exception as e:
            st.error(f"Error opening {app_name}: {str(e)}")
    else:
        st.error(f"Application '{app_name}' not found.")

def open_website(url):
    try:
        webbrowser.open(f"{url}.com")
        st.write(f"Opening website: {url}")  # Simulate opening a website
    except Exception as e:
        st.error(f"Error opening website {url}: {str(e)}")

option = st.selectbox('Select an Option:', [
    'Open Application',
    'Open Website',
    'Chat',
])

if option == 'Open Application':
    app_name = st.text_input('Enter Application Name:')
    if st.button('Open Application'):
        open_application(app_name)

elif option == 'Open Website':
    url = st.text_input('Enter Website URL:')
    if st.button('Open Website'):
        open_website(url)

elif option == 'Chat':
    user_message = st.text_input('Enter your message:')
    if st.button('Send'):
        response = get_gemini_response(user_message)
        if response:
            st.write("**JARVIS:** " + response)
            st.session_state.chat_history.append(('User', user_message))
            st.session_state.chat_history.append(('JARVIS', response))

def speak(text: str) -> None:
    def run_speech():
        engine.say(text)
        engine.runAndWait()

    # Run the speech synthesis in a separate thread
    speech_thread = threading.Thread(target=run_speech)
    speech_thread.start()

def take_voice_command() -> str:
    try:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source)
            print("Listening...")
            audio = recognizer.listen(source)
            try:
                print("Recognizing...")
                command = recognizer.recognize_google(audio)
                print(f"User  said: {command}")
            except sr.UnknownValueError:
                speak("Sorry, I didn't get that. Please repeat.")
                return None
            except sr.RequestError:
                speak("Sorry, I am having trouble connecting to the service.")
                return None
            return command.lower()
    except Exception as e:
        speak(f"An error occurred: {str(e)}")
        return None

def open_app(app_name: str) -> None:
    app_paths = {
        'word': 'C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE',
        'excel': 'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE',
        'powerpoint': 'C:/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE',
        'chrome': 'C:/Program Files/Google/Chrome/Application/chrome.exe',
        'notepad': 'C:/Windows/system32/notepad.exe',
        'vscode': 'C:/Users/YourUsername/AppData/Local/Programs/Microsoft VS Code/Code.exe',
        'paint': 'C:/Windows/system32/mspaint.exe',
        'outlook': 'C:/Program Files/Microsoft Office/root/Office16/OUTLOOK.EXE',
        'firefox': 'C:/Program Files/Mozilla Firefox/firefox.exe',
        'edge': 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
        'calculator': 'C:/Windows/system32/calc.exe',
        'cmd': 'C:/Windows/system32/cmd.exe',
        'control_panel': 'C:/Windows/system32/control.exe',
        'explorer': 'C:/Windows/explorer.exe',
        'onenote': 'C:/Program Files/Microsoft Office/root/Office16/ONENOTE.EXE',
        'teams': 'C:/Program Files/Microsoft Teams/current/Teams.exe',
        'skype': 'C:/Program Files (x86)/Microsoft/Skype for Desktop/Skype.exe',
        'spotify': 'C:/Users/YourUsername/AppData/Roaming /Spotify/Spotify.exe',
        'zoom': 'C:/Users/YourUsername/AppData/Roaming/Zoom/bin/Zoom.exe',
        'vlc': 'C:/Program Files/VideoLAN/VLC/vlc.exe',
        'notepad++': 'C:/Program Files/Notepad++/notepad++.exe',
        'sublime_text': 'C:/Program Files/Sublime Text 3/sublime_text.exe',
        'pycharm': 'C:/Program Files/JetBrains/PyCharm 2023.1.2/bin/pycharm64.exe',
        'android studio': 'C:/Program Files/Android/Android Studio/bin/studio64.exe',
        'eclipse': 'C:/Program Files/eclipse/eclipse.exe',
        'mysql_workbench': 'C:/Program Files/MySQL/MySQL Workbench 8.0/mysqlworkbench.exe',
        'task manager': 'C:/Windows/system32/taskmgr.exe',
        'windows media player': 'C:/Program Files/Windows Media Player/wmplayer.exe',
        'microsoft_store': 'C:/Windows/SystemApps/Microsoft.WindowsStore/WinStore.App.exe',
        'visual_studio': 'C:/Program Files (x86)/Microsoft Visual Studio/2019/Community/Common7 /IDE/devenv.exe',
        'photos': 'C:/Program Files/WindowsApps/Microsoft.Windows.Photos.exe',
        'snipping_tool': 'C:/Windows/system32/SnippingTool.exe',
        'teamviewer': 'C:/Program Files (x86)/TeamViewer/TeamViewer.exe',
        'anydesk': 'C:/Program Files/AnyDesk/AnyDesk.exe',
        'powershell': 'C:/Windows/system32/WindowsPowerShell/v1.0/powershell.exe',
        'settings': 'C:/Windows/ImmersiveControlPanel/SystemSettings.exe',
        'sticky_notes': 'C:/Program Files/WindowsApps/Microsoft.MicrosoftStickyNotes/StickyNotes.exe',
        'bing search': 'C:/Program Files/WindowsApps/Microsoft.BingNews/Video.UI.exe',
        'voice recorder': 'C:/Program Files/WindowsApps/Microsoft.WindowsSoundRecorder/SoundRecorder.exe',
        'clock': 'C:/Program Files/WindowsApps/Microsoft.WindowsAlarms/AlarmsClock.exe',
        'camera': 'C:/Program Files/WindowsApps/Microsoft.WindowsCamera/WindowsCamera.exe',
        'mail': 'C:/Program Files/WindowsApps/microsoft.windowscommunicationsapps/WindowsMail.exe'
    }

    if app_name in app_paths:
        speak(f"Opening {app_name}")
        os.startfile(app_paths[app_name])
    else:
        speak(f"Sorry, I don't have the path for {app_name}")

def close_app(app_name: str) -> None:
    processes = {
        'word': 'WINWORD.EXE',
        'excel': 'EXCEL.EXE',
        'settings':'SystemSettings.exe',
        'powerpoint': 'POWERPNT.EXE',
        'chrome': 'chrome.exe',
        'notepad': 'notepad.exe',
        'vscode': 'Code.exe',
        'paint': 'mspaint.exe',
        'outlook': 'OUTLOOK.EXE',
        'firefox': 'firefox.exe',
        'edge': 'msedge.exe',
        'calculator': 'calc.exe',
        'command prompt': 'cmd.exe',
        'control panel': 'control.exe',
        'file explorer': 'explorer.exe',
        'onenote': 'ONENOTE.EXE',
        'teams': 'Teams.exe',
        'skype': 'Skype.exe',
        'spotify': 'Spotify.exe',
        'vlc': 'vlc.exe',
        'notepad++': 'notepad++.exe',
        'sublime text': 'sublime_text.exe',
        'pycharm': 'pycharm64.exe',
        'android studio': 'studio64.exe',
        'mysql workbench': 'mysqlworkbench.exe',
        'virtualbox': 'VirtualBox.exe',
        'opera': 'launcher.exe',
        'task manager': 'taskmgr.exe',
        'windows media player': 'wmplayer.exe',
        'microsoft store': 'WinStore.App.exe',
        'visual studio': 'devenv.exe',
        'photos': 'Microsoft.Photos.exe',
        'snipping tool': 'SnippingTool.exe',
        'dropbox': 'Dropbox.exe',
        'onedrive': 'OneDrive.exe',
        'teamviewer': 'TeamViewer.exe',
        'anydesk': 'AnyDesk.exe',
        'powershell': 'powershell.exe',
    }

    if app_name in processes:
        speak(f"Closing {app_name}")
        os.system(f"taskkill /f /im {processes[app_name]}")
    else:
        speak(f"Sorry, I don't know how to close {app_name}")

def open_website(site_name: str) -> None:
    try:
        if "search" in site_name:
            search_query = site_name.replace("search", "").strip()
            kit.search(search_query)
            speak(f"Searching {search_query}.")
        else:
            webbrowser.open(f"https://{site_name}.com")
            speak(f"{site_name} opened.")
    except Exception as e:
        speak(f"An error occurred while opening the website: {str(e)}")

def play_music(song: str) -> None:
    try:
        kit.playonyt(song)
        speak(f"Playing {song} on YouTube.")
    except Exception as e:
        speak(f"An error occurred while playing {song}: {str(e)}")

watch_later = []
def save_to_watch_later(song: str) -> None:
    try:
        watch_later.append(song)
        speak(f"Saved {song} to watch later.")
    except Exception as e:
        speak(f"An error occurred while saving {song} to watch later: {str(e)}")

def play_from_watch_later() -> None:
    try:
        if watch_later:
            play_music(watch_later[0])
        else:
            speak("Your watch later list is empty.")
    except Exception as e:
        speak(f"An error occurred while playing from watch later: {str(e)}")

def control_volume(action: str) -> None:
    try:
        if action == "up":
            pyautogui.press("volumeup")
        elif action == "down":
            pyautogui.press("volumedown")
            speak(f"Volume {action}.")
    except Exception as e:
        speak(f"An error occurred while controlling volume: {str(e)}")

def close_tabs(action: str) -> None:
    try:
        if action == "all":
            keyboard.press_and_release('ctrl+shift+w')
        elif action == "single":
            keyboard.press_and_release('ctrl+w')
        speak(f"Closed {action} tab(s).")
    except Exception as e:
        speak(f"An error occurred while closing {action} tab(s): {str(e)}")

def control_system(action: str) -> None:
    try:
        if action == "shutdown":
            os.system("shutdown /s /t 1")
        elif action == "restart":
            os.system("shutdown /r /t 1")
        elif action == "lock":
            keyboard.press_and_release("win+l")
        speak(f"System will {action} now.")
    except Exception as e:
        speak(f"An error occurred while trying to {action} the system: {str(e)}")

def make_note() -> None:
    try:
        note_text = take_voice_command()
        if note_text:
            with open("note.txt", "w") as file:
                file.write(note_text)
            subprocess.Popen(["notepad.exe", "note.txt"])
            speak("Note saved in Notepad.")
    except Exception as e:
        speak(f"An error occurred while making the note: {str(e)}")

def main() -> None:
    speak("Jarvis at your service. What can I do for you?")
    while True:
        command = take_voice_command()
        if command is None:
            continue
        if "open" in command:
            if "website" in command:
                speak("opening app")
                threading.Thread(target=open_website, args=(command.replace("open website", "").strip(),)).start()
            else:
                threading.Thread(target=open_app, args=(command.replace("open", "").strip(),)).start()
        elif "close" in command:
            threading.Thread(target=close_app, args=(command.replace("close", "").strip(),)).start()
        elif "play music" in command:
            song = command.replace("play music", "").strip()
            threading.Thread(target=play_music, args=(song,)).start()
        elif "volume" in command:
            if "up" in command:
                threading.Thread(target=control_volume, args=("up",)).start()
            elif "down" in command:
                threading.Thread(target=control_volume, args=("down",)).start()
        elif "close tab" in command:
            if "all" in command:
                threading.Thread(target=close_tabs, args=("all",)).start()
            elif "single" in command:
                threading.Thread(target=close_tabs, args=("single",)).start()
        elif "shutdown" in command or "restart" in command or "lock" in command:
            threading.Thread(target=control_system, args=(command,)).start()
        elif "make note" in command:
            threading.Thread(target=make_note).start()
        elif "exit" in command or "quit" in command:
            speak("Goodbye!")
            break

if __name__ == "__main__":
    main()