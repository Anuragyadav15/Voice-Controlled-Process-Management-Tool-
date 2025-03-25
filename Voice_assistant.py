import os
import subprocess
import speech_recognition as sr
import pyttsx3
import psutil
import keyboard
import time
import webbrowser
import screen_brightness_control as sbc
import platform
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Initialize text-to-speech engine
engine = pyttsx3.init()

# Initialize volume control (Windows only)
if platform.system() == "Windows":
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume_control = cast(interface, POINTER(IAudioEndpointVolume))
    except:
        volume_control = None
        print("Volume control initialization failed - some features may not work")
else:
    volume_control = None

# Dictionary mapping app names to their launch commands
app_commands = {
    "chrome": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "notepad": "notepad.exe",
    "spotify": "spotify",
    "calculator": "calc.exe",
    "terminal": "gnome-terminal",
    "code": "code",
}

def speak(text):
    """Function to speak text"""
    try:
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"Error in speech synthesis: {e}")

def close_application(target):
    """Function to close an application"""
    target = target.lower()
    closed = False
    message = f"{target} not found"
    
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if target in proc.info['name'].lower():
                proc.kill()
                closed = True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    
    if closed:
        message = f"Closed {target}"
    return message

def set_volume(level):
    """Set volume to specific level (0-100)"""
    try:
        if platform.system() == "Windows" and volume_control:
            # Convert 0-100 to 0.0-1.0 range
            volume_level = min(max(level, 0), 100)
            volume_control.SetMasterVolumeLevelScalar(volume_level/100, None)
            return f"Volume set to {volume_level}%"
        elif platform.system() == "Linux":
            level = min(max(level, 0), 100)
            subprocess.run(["amixer", "-D", "pulse", "sset", "Master", f"{level}%"])
            return f"Volume set to {level}%"
        else:
            return "Volume control not supported on this system"
    except Exception as e:
        return f"Failed to set volume: {str(e)}"

def adjust_volume(change):
    """Adjust volume up or down by percentage"""
    try:
        if platform.system() == "Windows" and volume_control:
            current = volume_control.GetMasterVolumeLevelScalar() * 100
            new_level = min(max(current + change, 0), 100)
            volume_control.SetMasterVolumeLevelScalar(new_level/100, None)
            return f"Volume {'increased' if change > 0 else 'decreased'} to {int(new_level)}%"
        elif platform.system() == "Linux":
            subprocess.run(["pactl", "set-sink-volume", "@DEFAULT_SINK@", f"{'+' if change > 0 else '-'}{abs(change)}%"])
            return f"Volume {'increased' if change > 0 else 'decreased'} by {abs(change)}%"
        else:
            return "Volume adjustment not supported on this system"
    except Exception as e:
        return f"Failed to adjust volume: {str(e)}"

def toggle_mute():
    """Toggle mute state"""
    try:
        if platform.system() == "Windows" and volume_control:
            is_muted = volume_control.GetMute()
            volume_control.SetMute(not is_muted, None)
            return "Volume muted" if not is_muted else "Volume unmuted"
        elif platform.system() == "Linux":
            subprocess.run(["amixer", "-D", "pulse", "sset", "Master", "toggle"])
            return "Toggled mute"
        else:
            return "Mute control not supported on this system"
    except Exception as e:
        return f"Failed to toggle mute: {str(e)}"

def set_brightness(level):
    """Set brightness to specific level (0-100)"""
    try:
        level = min(max(level, 0), 100)
        sbc.set_brightness(level)
        return f"Brightness set to {level}%"
    except Exception as e:
        return f"Failed to set brightness: {str(e)}"

def adjust_brightness(change):
    """Adjust brightness up or down by percentage"""
    try:
        current = sbc.get_brightness()[0]
        new_level = min(max(current + change, 0), 100)
        sbc.set_brightness(new_level)
        return f"Brightness {'increased' if change > 0 else 'decreased'} to {new_level}%"
    except Exception as e:
        return f"Failed to adjust brightness: {str(e)}"

def control_bluetooth(state):
    """Function to control Bluetooth"""
    if platform.system() == "Windows":
        try:
            subprocess.run(["powershell", "-command", f"Start-Process -FilePath 'ms-settings:bluetooth' -Verb runAs"])
            return f"Bluetooth turned {state}"
        except Exception as e:
            return f"Failed to control Bluetooth: {str(e)}"
    else:
        return "Bluetooth control is currently only supported on Windows"

def control_hotspot(state):
    """Function to control hotspot"""
    if platform.system() == "Windows":
        try:
            if state == "on":
                subprocess.run(["netsh", "wlan", "set", "hostednetwork", "mode=allow", "ssid=Hotspot", "key=password123"])
                subprocess.run(["netsh", "wlan", "start", "hostednetwork"])
            else:
                subprocess.run(["netsh", "wlan", "stop", "hostednetwork"])
            return f"Hotspot turned {state}"
        except Exception as e:
            return f"Failed to control hotspot: {str(e)}"
    else:
        return "Hotspot control is currently only supported on Windows"

def listen_to_command():
    """Function to listen to voice commands"""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        try:
            audio = recognizer.listen(source, timeout=5)
            command = recognizer.recognize_google(audio)
            print(f"You said: {command}")
            return command.lower()
        except sr.UnknownValueError:
            print("Sorry, I didn't understand that.")
            return None
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            return None
        except sr.WaitTimeoutError:
            print("Listening timed out.")
            return None

def interpret_command(command):
    """Function to interpret voice commands"""
    if not command:
        return "unknown", None
        
    command = command.lower()
    
    # Website commands
    if "open" in command and "website" in command:
        website = command.replace("open", "").replace("website", "").strip()
        return "open_website", website
    
    # App commands
    elif "open" in command:
        return "open", command.replace("open", "").strip()
    elif "close" in command:
        return "close", command.replace("close", "").strip()
    
    # System commands
    elif "list processes" in command:
        return "list", None
    elif "open photo" in command:
        return "open_photo", None
    elif "open settings" in command:
        return "open_settings", None
    
    # Network commands
    elif "turn on wifi" in command:
        return "wifi_on", None
    elif "turn off wifi" in command:
        return "wifi_off", None
    elif "turn on hotspot" in command:
        return "hotspot_on", None
    elif "turn off hotspot" in command:
        return "hotspot_off", None
    
    # Bluetooth commands
    elif "turn on bluetooth" in command:
        return "bluetooth_on", None
    elif "turn off bluetooth" in command:
        return "bluetooth_off", None
    
    # Brightness commands
    elif "set brightness" in command or "brightness" in command:
        try:
            level = int(''.join(filter(str.isdigit, command)))
            return "set_brightness", level
        except:
            return "brightness_error", None
    elif "increase brightness" in command:
        return "increase_brightness", None
    elif "decrease brightness" in command:
        return "decrease_brightness", None
    
    # Volume commands
    elif "set volume" in command or "volume" in command:
        try:
            level = int(''.join(filter(str.isdigit, command)))
            return "set_volume", level
        except:
            return "volume_error", None
    elif "increase volume" in command:
        return "increase_volume", None
    elif "decrease volume" in command:
        return "decrease_volume", None
    elif "mute" in command:
        return "mute_volume", None
    elif "unmute" in command:
        return "unmute_volume", None
    
    # Termination command
    elif "terminate" in command:
        return "terminate", None
    else:
        return "unknown", None

def execute_command(action, target):
    """Function to execute commands based on interpreted action"""
    if action == "open":
        if target in app_commands:
            try:
                subprocess.Popen(app_commands[target], shell=True)
                return f"Opened {target}"
            except Exception as e:
                return f"Failed to open {target}: {str(e)}"
        else:
            return f"Application '{target}' not found in the command list."
    
    elif action == "open_website":
        try:
            if not target.startswith(('http://', 'https://')):
                target = f"http://{target}.com"
            if "chrome" in app_commands:
                webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(app_commands["chrome"]))
                webbrowser.get('chrome').open_new(target)
            else:
                webbrowser.open_new(target)
            return f"Opening {target} in your browser"
        except Exception as e:
            return f"Failed to open website: {str(e)}"
    
    elif action == "close":
        return close_application(target)
    
    elif action == "list":
        processes = [proc.info['name'] for proc in psutil.process_iter(['pid', 'name'])]
        return "Running processes: " + ", ".join(processes[:10]) + "..."
    
    elif action == "open_photo":
        photos_dir = os.path.expanduser("~/Pictures")
        if os.path.exists(photos_dir):
            os.startfile(photos_dir)
            return "Opened photos directory."
        return "Photos directory not found."
    
    elif action == "open_settings":
        if platform.system() == "Windows":
            subprocess.Popen("start ms-settings:", shell=True)
        else:
            subprocess.Popen("gnome-control-center", shell=True)
        return "Opened system settings."
    
    elif action == "wifi_on":
        if platform.system() == "Windows":
            subprocess.run(["netsh", "interface", "set", "interface", "Wi-Fi", "enable"], shell=True)
        else:
            subprocess.run(["nmcli", "radio", "wifi", "on"])
        return "Wi-Fi turned on."
    
    elif action == "wifi_off":
        if platform.system() == "Windows":
            subprocess.run(["netsh", "interface", "set", "interface", "Wi-Fi", "disable"], shell=True)
        else:
            subprocess.run(["nmcli", "radio", "wifi", "off"])
        return "Wi-Fi turned off."
    
    elif action == "hotspot_on":
        return control_hotspot("on")
    
    elif action == "hotspot_off":
        return control_hotspot("off")
    
    elif action == "bluetooth_on":
        return control_bluetooth("on")
    
    elif action == "bluetooth_off":
        return control_bluetooth("off")
    
    elif action == "set_brightness":
        return set_brightness(target)
    
    elif action == "increase_brightness":
        return adjust_brightness(10)
    
    elif action == "decrease_brightness":
        return adjust_brightness(-10)
    
    elif action == "brightness_error":
        return "Please specify a brightness level between 0 and 100"
    
    elif action == "set_volume":
        return set_volume(target)
    
    elif action == "increase_volume":
        return adjust_volume(10)
    
    elif action == "decrease_volume":
        return adjust_volume(-10)
    
    elif action == "mute_volume":
        return toggle_mute()
    
    elif action == "unmute_volume":
        return toggle_mute()
    
    elif action == "volume_error":
        return "Please specify a volume level between 0 and 100"
    
    elif action == "terminate":
        return "terminate"
    
    else:
        return "I didn't understand that command. Please try again."

def main():
    """Main function to run the voice assistant"""
    active = True
    print("voice command in now activated")
    speak("voice command in now activated")

    while True:
        if active:
            command = listen_to_command()
            if command:
                action, target = interpret_command(command)
                response = execute_command(action, target)
                print(response)
                speak(response)

                if action == "terminate":
                    print("Terminating the program...")
                    speak("Goodbye!")
                    break
        else:
            time.sleep(0.1)

        if keyboard.is_pressed("ctrl+alt"):
            active = not active
            status = "resumed" if active else "paused"
            print(f"Voice command {status}.")
            speak(f"Voice command {status}")

if __name__ == "__main__":
    # Install required packages if missing
    try:
        import pycaw
    except ImportError:
        print("Installing required packages...")
        subprocess.run(["pip", "install", "pycaw", "comtypes", "screen-brightness-control"])
    
    main()
