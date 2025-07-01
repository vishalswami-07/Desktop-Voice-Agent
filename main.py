# Voice Assistant App using Streamlit + LangChain + Custom Tools + Voice + Memory (with edge-tts)

import streamlit as st
import speech_recognition as sr
import os
import pyautogui
import psutil
import json
from datetime import datetime
from langchain.agents import initialize_agent, Tool
from langchain_openai import AzureChatOpenAI
from langchain.agents.agent_types import AgentType
from pathlib import Path
from dotenv import load_dotenv
import threading
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import webbrowser
import subprocess
import shutil
import winreg
import asyncio
import edge_tts
import tempfile
import asyncio
from datetime import datetime
import platform

os.environ['CURL_CA_BUNDLE'] = 'huggingface.co.crt'

# ------------------- ENV -------------------
load_dotenv()
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2024-08-01-preview")

# ------------------- TTS using edge-tts -------------------
async def speak_async(text):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmpfile:
            output_path = tmpfile.name

        communicate = edge_tts.Communicate(text, voice="en-IN-NeerjaNeural")
        await communicate.save(output_path)

        # Play the file cross-platform
        if platform.system() == "Windows":
            os.startfile(output_path)
        elif platform.system() == "Darwin":
            subprocess.call(["afplay", output_path])
        else:
            subprocess.call(["mpg123", output_path])
    except Exception as e:
        print(f"Error in TTS: {e}")

def speak(text):
    try:
        asyncio.run(speak_async(text))
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(speak_async(text))

# ------------------- Microphone -------------------
def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        st.info("🎙️ Listening...")
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio)
            st.success(f"You said: {command}")
            return command
        except sr.UnknownValueError:
            st.error("Sorry, I didn't understand that.")
            return ""

# ------------------- Custom Tools -------------------
def open_application(app_name):
    app_name = app_name.strip().lower()
    common_apps = {
        "chrome": "chrome.exe", "notepad": "notepad.exe", "calculator": "calc.exe",
        "paint": "mspaint.exe", "vs code": "code.exe", "visual studio code": "code.exe",
        "word": "winword.exe", "excel": "excel.exe", "powerpoint": "powerpnt.exe", "cmd": "cmd.exe"
    }
    exe = common_apps.get(app_name, app_name + ".exe")
    try:
        path = shutil.which(exe)
        if path:
            os.system("start " + exe) if exe.lower() == "cmd.exe" else subprocess.Popen(path)
            return f"I've opened {app_name}."
        # fallback via registry
        reg_paths = [r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths",
                     r"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths"]
        for base in reg_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base) as base_key:
                    for i in range(winreg.QueryInfoKey(base_key)[0]):
                        subkey_name = winreg.EnumKey(base_key, i)
                        if app_name in subkey_name.lower():
                            with winreg.OpenKey(base_key, subkey_name) as subkey:
                                target_path, _ = winreg.QueryValueEx(subkey, None)
                                os.system("start cmd") if exe.lower() == "cmd.exe" else subprocess.Popen(target_path)
                                return f"I've opened {app_name}."
            except: continue
        return f"Sorry, I couldn't find an app named '{app_name}'."
    except Exception as e:
        return f"Failed to open {app_name}. Error: {str(e)}"

def get_current_time(_): return datetime.now().strftime("%A, %d %B %Y %I:%M %p")

def open_notepad(_): os.system("notepad"); return "Opened Notepad."

def take_screenshot(_):
    path = os.path.join(os.getcwd(), "screenshot.png")
    pyautogui.screenshot(path)
    return f"Screenshot saved to {path}"

def set_volume_level(level_str):
    try:
        level = max(0, min(100, int(''.join(filter(str.isdigit, level_str)))))
        dev = AudioUtilities.GetSpeakers()
        interface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(level / 100.0, None)
        return f"Volume set to {level}%"
    except Exception as e:
        return f"Failed to set volume: {e}"

def get_cpu_usage(_): return f"Current CPU usage is {psutil.cpu_percent(interval=1)}%"

def open_vscode(_): os.system("code"); return "Opened Visual Studio Code."

def mute_volume(_): os.system("nircmd.exe mutesysvolume 1"); return "System volume muted."
def unmute_volume(_): os.system("nircmd.exe mutesysvolume 0"); return "System volume unmuted."

def search_file(query):
    root = Path.home()
    matches = list(root.rglob(f"*{query.strip()}*"))
    return f"Found: {matches[0]}" if matches else "No file found."

def web_search(query):
    url = f"https://www.google.com/search?q={query.strip().replace(' ', '+')}"
    webbrowser.open(url)
    return f"Performed a web search for: {query}"

MEMORY_PATH = "memory.json"
def save_reminder(text):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    memory = json.load(open(MEMORY_PATH)) if os.path.exists(MEMORY_PATH) else {}
    memory[now] = text
    json.dump(memory, open(MEMORY_PATH, "w"), indent=2)
    return f"Reminder saved: {text}"

def list_reminders(_):
    if not os.path.exists(MEMORY_PATH): return "No reminders yet."
    memory = json.load(open(MEMORY_PATH))
    return "\n".join([f"{k}: {v}" for k, v in memory.items()])

# ------------------- LangChain Agent -------------------
def get_agent():
    llm = AzureChatOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_version=AZURE_OPENAI_API_VERSION,
        temperature=0,
    )
    tools = [
        Tool(name="Open Notepad", func=open_notepad, description="Open Notepad."),
        Tool(name="Take Screenshot", func=take_screenshot, description="Take a screenshot."),
        Tool(name="Check CPU Usage", func=get_cpu_usage, return_direct=True, description="Get CPU usage."),
        Tool(name="Open VS Code", func=open_vscode, description="Open Visual Studio Code."),
        Tool(name="Mute Volume", func=mute_volume, description="Mute the system."),
        Tool(name="Unmute Volume", func=unmute_volume, description="Unmute the system."),
        Tool(name="Search File", func=lambda q: search_file(q), description="Find a file."),
        Tool(name="Save Reminder", func=lambda q: save_reminder(q), description="Save a reminder."),
        Tool(name="List Reminders", func=list_reminders, description="List reminders."),
        Tool(name="Open Application", func=open_application, description="Open an installed app."),
        Tool(name="Get Current Time", func=get_current_time, description="Get system time."),
        Tool(name="Set Volume Level", func=set_volume_level, description="Set system volume."),
        Tool(name="Web Search", func=web_search, description="Search the web.")
    ]
    return initialize_agent(tools, llm, agent=AgentType.ZERO_SHOT_REACT_DESCRIPTION, handle_parsing_errors=True, verbose=True)

# ------------------- Streamlit UI -------------------
st.set_page_config(page_title="🧠 Voice Assistant AI", layout="centered")
st.title("🧠Desktop Voice Assistant")

if "agent" not in st.session_state:
    st.session_state.agent = get_agent()

if st.button("🎙️ Start Listening"):
    query = listen()
    if query:
        with st.spinner("Thinking..."):
            response = st.session_state.agent.run(query)
            st.session_state.response = response

# Show and speak response
if "response" in st.session_state:
    st.success(st.session_state.response)
    speak(st.session_state.response)
    del st.session_state.response

st.markdown("---")
st.markdown("Built with **LangChain**, **Azure OpenAI**, **Streamlit**, and **edge-tts** 🎧")

# Sidebar with system info
with st.sidebar:
    st.header("🖥️ System Status")
        
    try:
            # CPU and Memory
        cpu_percent = psutil.cpu_percent(interval=0.1)
        memory = psutil.virtual_memory()
            
        st.metric("CPU Usage", f"{cpu_percent}%")
        st.metric("Memory Usage", f"{memory.percent}%")
            
            # Battery (if available)
        battery = psutil.sensors_battery()
        if battery:
            st.metric("Battery", f"{battery.percent}%")
            
            # Current time
        st.metric("Current Time", datetime.now().strftime('%I:%M %p'))
            
    except Exception as e:
        st.error(f"Could not load system info: {e}")
        
    st.header("🔧 Settings")
    st.info("Voice agent is ready!")
        
    if st.button("Test TTS"):
        st.session_state.agent.speak("Text to speech is working correctly!")
