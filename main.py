import os
import sys
import time
import json
import subprocess
import shutil
import webbrowser
import datetime
import threading
from pathlib import Path

import streamlit as st
import speech_recognition as sr
import pyttsx3
import psutil
import pyautogui
import keyboard

# Audio control
try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    AUDIO_AVAILABLE = True
except ImportError:
    AUDIO_AVAILABLE = False
    st.warning("Audio control not available. Install pycaw for volume control.")


class WindowsVoiceAgent:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.tts_engine = pyttsx3.init()
        self.setup_tts()
        self.command_history = []
        
    def setup_tts(self):
        """Configure text-to-speech settings"""
        voices = self.tts_engine.getProperty('voices')
        if voices:
            # Try to use a female voice if available
            for voice in voices:
                if 'female' in voice.name.lower() or 'zira' in voice.name.lower():
                    self.tts_engine.setProperty('voice', voice.id)
                    break
        
        self.tts_engine.setProperty('rate', 180)  # Speed of speech
        self.tts_engine.setProperty('volume', 0.8)  # Volume level
    
    def listen_for_command(self, timeout=5):
        """Listen for voice command with timeout"""
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=10)
            
            command = self.recognizer.recognize_google(audio).lower()
            self.command_history.append(command)
            return command
            
        except sr.WaitTimeoutError:
            return "timeout"
        except sr.UnknownValueError:
            return "unclear"
        except sr.RequestError as e:
            return f"error: {str(e)}"
    
    def speak(self, text):
        """Convert text to speech"""
        self.tts_engine.say(text)
        self.tts_engine.runAndWait()
    
    # === SYSTEM CONTROL FUNCTIONS ===
    
    def control_volume(self, action, amount=10):
        """Control system volume"""
        if not AUDIO_AVAILABLE:
            return "Audio control not available"
        
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            
            current = volume.GetMasterVolumeLevelScalar()
            
            if action == "set":
                # Set to specific percentage
                new_volume = max(0.0, min(1.0, amount/100))
                volume.SetMasterVolumeLevelScalar(new_volume, None)
                return f"Volume set to {int(new_volume*100)}%"
            elif action == "up":
                new_volume = min(1.0, current + amount/100)
                volume.SetMasterVolumeLevelScalar(new_volume, None)
                return f"Volume increased to {int(new_volume*100)}%"
            elif action == "down":
                new_volume = max(0.0, current - amount/100)
                volume.SetMasterVolumeLevelScalar(new_volume, None)
                return f"Volume decreased to {int(new_volume*100)}%"
            elif action == "mute":
                volume.SetMute(1, None)
                return "Volume muted"
            elif action == "unmute":
                volume.SetMute(0, None)
                return "Volume unmuted"
            else:
                return f"Current volume: {int(current*100)}%"
                
        except Exception as e:
            return f"Volume control error: {str(e)}"
    
    def manage_applications(self, action, app_name=None):
        """Manage running applications"""
        try:
            if action == "list":
                processes = []
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        if proc.info['name'] and not proc.info['name'].startswith('System'):
                            processes.append(proc.info['name'])
                    except:
                        continue
                unique_processes = list(set(processes))[:10]  # Top 10
                return f"Running applications: {', '.join(unique_processes)}"
            
            elif action == "close" and app_name:
                killed = False
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        if app_name.lower() in proc.info['name'].lower():
                            proc.terminate()
                            killed = True
                    except:
                        continue
                return f"Closed {app_name}" if killed else f"Could not find {app_name}"
            
            elif action == "open" and app_name:
                common_apps = {
                    'notepad': 'notepad.exe',
                    'calculator': 'calc.exe',
                    'paint': 'mspaint.exe',
                    'cmd': 'cmd.exe',
                    'powershell': 'powershell.exe',
                    'browser': 'chrome.exe',
                    'chrome': 'chrome.exe',
                    'firefox': 'firefox.exe',
                    'edge': 'msedge.exe'
                }
                
                exe_name = common_apps.get(app_name.lower(), f"{app_name}.exe")
                try:
                    subprocess.Popen(exe_name)
                    return f"Opened {app_name}"
                except:
                    return f"Could not open {app_name}"
                    
        except Exception as e:
            return f"Application management error: {str(e)}"
    
    def file_operations(self, action, source=None, destination=None, file_type=None):
        """Handle file operations"""
        try:
            if action == "list":
                desktop = Path.home() / "Desktop"
                files = list(desktop.glob("*"))[:10]
                file_names = [f.name for f in files if f.is_file()]
                return f"Desktop files: {', '.join(file_names)}" if file_names else "No files on desktop"
            
            elif action == "create" and source:
                file_path = Path.home() / "Desktop" / source
                if not source.endswith('.txt'):
                    source += '.txt'
                    file_path = Path.home() / "Desktop" / source
                
                file_path.touch()
                return f"Created file: {source}"
            
            elif action == "delete" and source:
                desktop = Path.home() / "Desktop"
                file_path = desktop / source
                if file_path.exists():
                    file_path.unlink()
                    return f"Deleted file: {source}"
                else:
                    return f"File not found: {source}"
            
            elif action == "move" and source and destination:
                src_path = Path(source)
                dst_path = Path(destination)
                if src_path.exists():
                    shutil.move(str(src_path), str(dst_path))
                    return f"Moved {source} to {destination}"
                else:
                    return f"Source file not found: {source}"
                    
        except Exception as e:
            return f"File operation error: {str(e)}"
    
    def system_info(self, info_type):
        """Get system information"""
        try:
            if info_type == "time":
                return f"Current time: {datetime.datetime.now().strftime('%I:%M %p')}"
            
            elif info_type == "date":
                return f"Today's date: {datetime.datetime.now().strftime('%B %d, %Y')}"
            
            elif info_type == "battery":
                battery = psutil.sensors_battery()
                if battery:
                    percent = battery.percent
                    plugged = "plugged in" if battery.power_plugged else "on battery"
                    return f"Battery: {percent}% ({plugged})"
                else:
                    return "Battery information not available"
            
            elif info_type == "cpu":
                cpu_percent = psutil.cpu_percent(interval=1)
                return f"CPU usage: {cpu_percent}%"
            
            elif info_type == "memory":
                memory = psutil.virtual_memory()
                return f"Memory usage: {memory.percent}% ({memory.used // (1024**3)}GB used)"
                
        except Exception as e:
            return f"System info error: {str(e)}"
    
    def web_operations(self, action, query=None):
        """Handle web-related operations"""
        try:
            if action == "search" and query:
                search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
                webbrowser.open(search_url)
                return f"Searching for: {query}"
            
            elif action == "youtube" and query:
                youtube_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
                webbrowser.open(youtube_url)
                return f"Searching YouTube for: {query}"
            
            elif action == "open":
                common_sites = {
                    'google': 'https://www.google.com',
                    'youtube': 'https://www.youtube.com',
                    'gmail': 'https://www.gmail.com',
                    'github': 'https://www.github.com',
                    'stackoverflow': 'https://stackoverflow.com'
                }
                
                site = query.lower() if query else 'google'
                url = common_sites.get(site, f"https://www.{site}.com")
                webbrowser.open(url)
                return f"Opened {site}"
                
        except Exception as e:
            return f"Web operation error: {str(e)}"
    
    def keyboard_shortcuts(self, shortcut):
        """Execute keyboard shortcuts"""
        try:
            shortcuts = {
                'copy': 'ctrl+c',
                'paste': 'ctrl+v',
                'cut': 'ctrl+x',
                'undo': 'ctrl+z',
                'save': 'ctrl+s',
                'select all': 'ctrl+a',
                'alt tab': 'alt+tab',
                'task manager': 'ctrl+shift+esc',
                'run dialog': 'win+r',
                'minimize all': 'win+m',
                'desktop': 'win+d'
            }
            
            if shortcut in shortcuts:
                keyboard.send(shortcuts[shortcut])
                return f"Executed: {shortcut}"
            else:
                return f"Unknown shortcut: {shortcut}"
                
        except Exception as e:
            return f"Keyboard shortcut error: {str(e)}"
    
    def process_command(self, command):
        """Process voice command and execute appropriate action"""
        command = command.lower().strip()
        
        # Volume control
        if any(word in command for word in ['volume', 'sound']):
            # Extract percentage if specified
            import re
            percentage_match = re.search(r'(\d+)%?', command)
            target_percentage = int(percentage_match.group(1)) if percentage_match else None
            
            if ('to' in command or 'set' in command) and target_percentage:
                # Set to specific percentage: "volume to 80%" or "set volume to 80%"
                return self.control_volume('set', target_percentage)
            elif 'up' in command or 'increase' in command:
                if target_percentage:
                    # "volume up to 80%" - set to that percentage
                    return self.control_volume('set', target_percentage)
                else:
                    # "volume up" - increase by 10%
                    return self.control_volume('up')
            elif 'down' in command or 'decrease' in command or 'lower' in command:
                if target_percentage:
                    # "volume down to 30%" - set to that percentage
                    return self.control_volume('set', target_percentage)
                else:
                    # "volume down" - decrease by 10%
                    return self.control_volume('down')
            elif 'mute' in command:
                return self.control_volume('mute')
            elif 'unmute' in command:
                return self.control_volume('unmute')
            else:
                return self.control_volume('get')
        
        # Application management
        elif any(word in command for word in ['open', 'start', 'launch']):
            if 'notepad' in command:
                return self.manage_applications('open', 'notepad')
            elif 'calculator' in command:
                return self.manage_applications('open', 'calculator')
            elif 'browser' in command or 'chrome' in command:
                return self.manage_applications('open', 'chrome')
            elif 'paint' in command:
                return self.manage_applications('open', 'paint')
            else:
                # Extract app name
                words = command.split()
                if 'open' in words:
                    idx = words.index('open')
                    if idx + 1 < len(words):
                        app_name = words[idx + 1]
                        return self.manage_applications('open', app_name)
        
        elif 'close' in command and any(word in command for word in ['app', 'application', 'program']):
            words = command.split()
            if 'close' in words:
                idx = words.index('close')
                if idx + 1 < len(words):
                    app_name = words[idx + 1]
                    return self.manage_applications('close', app_name)
        
        elif 'list' in command and 'app' in command:
            return self.manage_applications('list')
        
        # System information
        elif 'time' in command:
            return self.system_info('time')
        elif 'date' in command:
            return self.system_info('date')
        elif 'battery' in command:
            return self.system_info('battery')
        elif 'cpu' in command:
            return self.system_info('cpu')
        elif 'memory' in command or 'ram' in command:
            return self.system_info('memory')
        
        # Web operations
        elif 'search' in command:
            query = command.replace('search', '').replace('for', '').strip()
            if 'youtube' in command:
                return self.web_operations('youtube', query.replace('youtube', '').strip())
            else:
                return self.web_operations('search', query)
        
        elif 'youtube' in command:
            query = command.replace('youtube', '').strip()
            return self.web_operations('youtube', query)
        
        elif 'website' in command or 'site' in command:
            site = command.replace('website', '').replace('site', '').replace('open', '').strip()
            return self.web_operations('open', site)
        
        # File operations
        elif 'file' in command:
            if 'list' in command:
                return self.file_operations('list')
            elif 'create' in command:
                filename = command.replace('create', '').replace('file', '').strip()
                return self.file_operations('create', filename)
            elif 'delete' in command:
                filename = command.replace('delete', '').replace('file', '').strip()
                return self.file_operations('delete', filename)
        
        # Keyboard shortcuts
        elif any(shortcut in command for shortcut in ['copy', 'paste', 'cut', 'save', 'undo']):
            for shortcut in ['copy', 'paste', 'cut', 'save', 'undo', 'select all']:
                if shortcut in command:
                    return self.keyboard_shortcuts(shortcut)
        
        elif 'alt tab' in command:
            return self.keyboard_shortcuts('alt tab')
        elif 'task manager' in command:
            return self.keyboard_shortcuts('task manager')
        elif 'minimize' in command:
            return self.keyboard_shortcuts('minimize all')
        elif 'desktop' in command and 'show' in command:
            return self.keyboard_shortcuts('desktop')
        
        # Default response
        else:
            return f"I heard: '{command}', but I'm not sure how to help with that. Try commands like 'volume up', 'open notepad', 'what time is it', or 'search for cats'."

# Streamlit Interface
def main():
    st.set_page_config(page_title="Windows Voice Agent", page_icon="🗣️", layout="wide")
    
    st.title("🗣️ Windows Voice Agent")
    st.markdown("**Your personal voice-controlled Windows assistant**")
    
    # Initialize session state
    if 'agent' not in st.session_state:
        st.session_state.agent = WindowsVoiceAgent()
    if 'listening' not in st.session_state:
        st.session_state.listening = False
    if 'conversation' not in st.session_state:
        st.session_state.conversation = []
    
    # Control buttons
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("🎤 Listen", type="primary"):
            st.session_state.listening = True
    
    with col2:
        if st.button("🔄 Clear History"):
            st.session_state.conversation = []
            st.session_state.agent.command_history = []
    
    with col3:
        if st.button("ℹ️ Help"):
            st.info("""
            **Available Commands:**
            - Volume: "volume up/down/mute"
            - Apps: "open notepad/calculator/browser"
            - System: "what time is it", "battery status"
            - Web: "search for cats", "open youtube"
            - Files: "list files", "create file test"
            - Shortcuts: "copy", "paste", "alt tab"
            """)
    
    with col4:
        voice_enabled = st.checkbox("🔊 Voice Response", value=True)
    
    # Listening indicator
    if st.session_state.listening:
        with st.spinner("🎧 Listening for your command..."):
            command = st.session_state.agent.listen_for_command()
            
            if command == "timeout":
                response = "I didn't hear anything. Please try again."
            elif command == "unclear":
                response = "I couldn't understand that. Please speak clearly."
            elif command.startswith("error"):
                response = f"Speech recognition error: {command}"
            else:
                response = st.session_state.agent.process_command(command)
                
                # Add to conversation
                st.session_state.conversation.append({
                    'time': datetime.datetime.now().strftime('%H:%M:%S'),
                    'command': command,
                    'response': response
                })
                
                # Speak response if enabled
                if voice_enabled and not command.startswith("error"):
                    threading.Thread(target=st.session_state.agent.speak, args=(response,)).start()
            
            st.session_state.listening = False
            st.rerun()
    
    # Display conversation history
    if st.session_state.conversation:
        st.subheader("📝 Conversation History")
        
        for entry in reversed(st.session_state.conversation[-10:]):  # Show last 10
            with st.container():
                st.markdown(f"**[{entry['time']}] You said:** {entry['command']}")
                st.markdown(f"**Assistant:** {entry['response']}")
                st.divider()
    
    # Manual command input
    st.subheader("⌨️ Manual Command Input")
    manual_command = st.text_input("Type a command:", placeholder="e.g., 'volume up' or 'open notepad'")
    
    if st.button("Execute Command") and manual_command:
        response = st.session_state.agent.process_command(manual_command)
        st.success(response)
        
        # Add to conversation
        st.session_state.conversation.append({
            'time': datetime.datetime.now().strftime('%H:%M:%S'),
            'command': manual_command,
            'response': response
        })
        
        if voice_enabled:
            threading.Thread(target=st.session_state.agent.speak, args=(response,)).start()
    
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
            st.metric("Current Time", datetime.datetime.now().strftime('%I:%M %p'))
            
        except Exception as e:
            st.error(f"Could not load system info: {e}")
        
        st.header("🔧 Settings")
        st.info("Voice agent is ready!")
        
        if st.button("Test TTS"):
            st.session_state.agent.speak("Text to speech is working correctly!")

if __name__ == "__main__":
    main()