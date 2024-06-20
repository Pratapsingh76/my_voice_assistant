import speech_recognition as sr
from datetime import datetime as dt
import pyttsx3
import json
from googlesearch import search
from bs4 import BeautifulSoup
import requests
import webbrowser
import subprocess
import os
import pyautogui
import datetime
import webbrowser
from pathlib import Path
import time
import ctypes
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume




# Function to recognize speech using Google Web Speech API
def recognize_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    try:
        print("Recognizing...")
        text = recognizer.recognize_google(audio)
        print("You said:", text)
        return text.lower()
    except sr.UnknownValueError:
        print("Sorry, I couldn't understand what you said.")
        return None
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
        return None

# Function to greet the user based on the time of the day
def wish_me():
    hour = dt.now().hour
    if 0 <= hour < 12:
        speak("Good Morning Sir!")
    elif 12 <= hour < 18:
        speak("Good Afternoon Sir!")
    else:
        speak("Good Evening Sir!")
    speak("I Am Your Voice Assistant Program. Please Tell Me How Can I Help You?...")

# Function to speak text
def speak(audio):
    engine = pyttsx3.init("sapi5")
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)
    rate = engine.getProperty('rate')
    engine.setProperty('rate', 160)
    print(f"You :{audio}.")
    engine.say(audio)
    engine.runAndWait()

# Function to load question and answer pairs from JSON file
def load_question_answers():
    with open("question_answers.json", "r") as file:
        question_answers = json.load(file)
    return question_answers

# Function to decrease the volume
# Function to prompt user and increase the volume by the specified amount
def increase_volume():
    try:
        increase_amount = int(input("By how much do you want to increase the volume? "))
        if increase_amount < 0:
            print("Please enter a positive number.")
            return
    except ValueError:
        print("Please enter a valid number.")
        return
    
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    
    # Get the current volume range
    volume_range = volume.GetVolumeRange()
    
    # Calculate the current volume level
    current_volume = volume.GetMasterVolumeLevel()
    
    # Calculate the target volume level
    target_volume = min(current_volume + increase_amount, volume_range[1])
    
    # Set the new volume level
    volume.SetMasterVolumeLevel(target_volume, None)
    
    print(f"Volume increased by {increase_amount}.")
# Function to prompt user and decrease the volume by the specified amount
def decrease_volume():
    try:
        decrease_amount = int(input("By how much do you want to decrease the volume? "))
        if decrease_amount < 0:
            print("Please enter a positive number.")
            return
    except ValueError:
        print("Please enter a valid number.")
        return
    
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    
    # Get the current volume range
    volume_range = volume.GetVolumeRange()
    
    # Calculate the current volume level
    current_volume = volume.GetMasterVolumeLevel()
    
    # Calculate the target volume level
    target_volume = max(current_volume - decrease_amount, volume_range[0])
    
    # Set the new volume level
    volume.SetMasterVolumeLevel(target_volume, None)
    
    print(f"Volume decreased by {decrease_amount}.")


 # open email id
def open_email():
    try:
        # URL for opening the default email client
        email_url = "https://mail.google.com/mail/u/0/#inbox"
        
        # Open the default email client
        webbrowser.open(email_url)
        print("Opening default email client...")
    except Exception as e:
        print(f"An error occurred: {e}")


# close browser tasks
def close_browser_tab():
    try:
        # Simulate the keyboard shortcut to close the active tab
        pyautogui.hotkey('ctrl', 'w')
        print("Closing the active browser tab...")
    except Exception as e:
        print(f"An error occurred: {e}")

# open my computer
def open_program(program_path):
    try:
        # Check if the specified program path exists
        if os.path.exists(program_path):
            # Open the program using its default application
            os.startfile(program_path)
            print(f"Opening {program_path}...")
        else:
            print(f"Error: {program_path} does not exist.")
    except Exception as e:
        print(f"An error occurred: {e}")

#   open webwhatsapp
def open_whatsapp_web():
    # URL for WhatsApp Web
    whatsapp_web_url = "https://web.whatsapp.com/"
    
    # Open WhatsApp Web in the default web browser
    webbrowser.open(whatsapp_web_url)

# Function to perform Google search and return the first result URL
def search_google(query):
    try:
        # Perform Google search
        search_results = search(query, stop=1, pause=2)
        # Get the first search result
        first_result = next(search_results)
        return first_result
    except Exception as e:
        print("Error occurred during Google search:", e)
        return None

# Function to open the control panel
def open_control_panel():
    subprocess.Popen(["control"])

#close cotrol panel
def close_control_panel():
    try:
        # Simulate the Alt + F4 keyboard shortcut to close the Control Panel window
        pyautogui.hotkey('alt', 'f4')
        print("Closing Control Panel...")
    except Exception as e:
        print(f"An error occurred: {e}")


# Function to restart the computer
def restart_computer():
    subprocess.Popen(["shutdown", "/r"])

# Function to shut down the computer
def shut_down_computer():
    subprocess.Popen(["shutdown", "/s"])
    

# Function to create a new folder
def create_folder_on_desktop(folder_name):
    try:
        # Get the path to the desktop folder
        desktop_path = Path.home() / "Desktop"
        
        # Create the new folder
        new_folder_path = desktop_path / folder_name
        os.makedirs(new_folder_path)
        
        print(f"Folder '{folder_name}' created successfully on the desktop.")
    except Exception as e:
        print(f"An error occurred: {e}")

# take screenshot
# Function to capture a screenshot
def take_screenshot():
    try:
        # Capture the screenshot
        screenshot = pyautogui.screenshot()
        
        # Define the file name with timestamp
        file_name = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".png"
        
        # Save the screenshot to the current directory
        screenshot.save(file_name)
        
        print("Screenshot saved as", file_name)
    except Exception as e:
        print("Error occurred while taking screenshot:", e)
# Function to open Microsoft Excel
def open_excel():
    try:
        subprocess.Popen(["start", "excel"])
        print("Microsoft Excel opened successfully.")
    except Exception as e:
        print("Error occurred while opening Microsoft Excel:", e)

# Function to open Microsoft Word
def open_word():
    try:
        subprocess.Popen(["start", "winword"])
        print("Microsoft Word opened successfully.")
    except Exception as e:
        print("Error occurred while opening Microsoft Word:", e)
# Function to main functionality
def main():
    speak("Welcome to Voice Assistant!")
    wish_me()
    question_answers = load_question_answers()

    while True:
        command = recognize_speech()
        if command:
            if command in question_answers:
                speak(question_answers[command])
            elif "exit" in command:
                speak("Goodbye Sir!")
                break
            elif "today date" in command:
                today_date = dt.now().strftime("%B %d, %Y")
                speak(f"Today's date is {today_date}.")
            elif "current time" in command:
                current_time = dt.now().strftime("%I:%M %p")
                speak(f"The current time is {current_time}.")
            elif "open control panel" in command:
                open_control_panel()
            elif "restart computer" in command:
                restart_computer()
            elif "shutdown computer" in command:
                shut_down_computer()
            elif "create new folder" in command:
                speak("Please tell me the name of the folder.")
                folder_name = recognize_speech()
                if folder_name:
                    folder_name = "New folder"
                    create_folder_on_desktop(folder_name)
            elif "capture screenshot" in command:
                take_screenshot()
            elif "open word" in command:
                open_word()
            elif "open excel" in command:
                open_excel()
            elif "open email" in command:
                open_email()
            elif "open whatsapp" in command:
                open_whatsapp_web()
            elif "close browser task" in command:
                close_browser_tab()
            elif "close control panel" in command:
                close_control_panel()
            elif "increase volume" in command:
                increase_volume()
            elif "decrease volume" in command:
                decrease_volume()
            elif "open my computer" in command:
                program_path = "D:/"  # Example path to Microsoft Word
                open_program(program_path)
            else:
                # Search on Google if the command is not recognized
                speak("Let me search that for you.")
                search_result = search_google(command)
                if search_result:
                    # Open the first search result link in the default web browser
                    webbrowser.open(search_result)
                    speak("Here is what I found on the web:")
                    speak("Please wait while I fetch the content.")
                else:
                    speak("Sorry, I couldn't find any relevant information.")

        else:
            pass
        list = []

if __name__ == "__main__":
    main()
