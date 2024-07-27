import tkinter as tk
import speech_recognition as sr # type: ignore
import webbrowser as web
import pyttsx3 # type: ignore 
import subprocess
from threading import Thread
import sys
import os
import pyautogui # type: ignore
import datetime
import pyowm # type: ignore
import geopy # type: ignore
from geopy.geocoders import Nominatim # type: ignore
import requests, json # type: ignore



class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.insert(tk.END, message)

    def flush(self):
        pass

def main():
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)

    r = sr.Recognizer()

    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)

        print("What do you want to use : Web browser or the desktop ?\n")
        engine.say(" What do you want to use : Web browser or the desktop ?")
        engine.runAndWait()

        audio = r.listen(source)

        try:
            dest = r.recognize_google(audio)
            if 'web' in dest.lower():
                print("How can i help you ?")
                engine.say('How can I help you ?')
                engine.runAndWait()

                audio = r.listen(source)

                dest = r.recognize_google(audio)
                engine.say('opening' + dest)
                engine.runAndWait()
                web.open(dest)
            
                if 'thank' in dest.lower():
                    print('I hope you liked my assistance. Have a nice day\n')
                    engine.say(" I hope you liked my assistance. Have a nice day")
                    engine.runAndWait()
                    exit()

                elif 'screenshot' in dest.lower():
                    print("Give a name to your screenshot")
                    engine.say(" Opening Give a name to your screenshot")
                    engine.runAndWait()
                    audio = r.listen(source)
                    dest = r.recognize_google(audio) 
                    pyautogui.screenshot(dest + ".png")  

            elif 'desktop' in dest.lower():
                print("How can i help you ?")
                engine.say('How can I help you ?')
                engine.runAndWait()

                audio = r.listen(source)

                dest = r.recognize_google(audio)

                if 'calculator' in dest.lower():
                    print("Opening Calculator")
                    engine.say(" Opening calculator")
                    engine.runAndWait()
                    subprocess.Popen(['calc.exe'])

                elif 'word' in dest.lower():
                    print("Opening Microsoft word")
                    engine.say(" Opening Microsoft word")
                    engine.runAndWait()

                    doc_path = r'C:\\path\\to\\your\\document.docx'
                    command = ['start', 'WINWORD.EXE', doc_path]
                    subprocess.Popen(command, shell=True)

                elif 'presentation' in dest.lower():
                    print("Opening Power Point Presentation")
                    engine.say(" Opening Power Point Presentation")
                    engine.runAndWait()
                    ppt_path = r'C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE'
                    subprocess.Popen([ppt_path])

                elif 'spot' in dest.lower():
                    print("Opening Spotify")
                    engine.say(" Opening spotify")
                    engine.runAndWait()
                    ppt_path = r'C:\\Users\\oishanu\AppData\\Roaming\\Spotify\\Spotify.exe'
                    subprocess.Popen([ppt_path])

                elif 'discord' in dest.lower():
                    print("Opening Discord")
                    engine.say(" Opening Discord")
                    engine.runAndWait()
                    path = r'C:\\Users\\oishanu\\OneDrive\\Desktop\\Discord.lnk'
                    subprocess.Popen([path])

                elif 'excel' in dest.lower():
                    print("Opening Microsoft Excel")
                    engine.say(" Opening Microsoft Excel")
                    engine.runAndWait()
                    path = r'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE'
                    subprocess.Popen([path])

                elif 'zoom' in dest.lower():
                    print("Opening Zoom")
                    engine.say(" Opening Zoom")
                    engine.runAndWait()
                    path = r'C:\\Users\\oishanu\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe'
                    subprocess.Popen([path])
                
                elif 'team' in dest.lower():
                    print("Opening Microsoft Teams")
                    engine.say(" Opening Microsoft Teams")
                    engine.runAndWait()
                    path = r'C:\\ProgramData\\oishanu\\Microsoft\\Teams\\current\\Teams.exe'
                    subprocess.Popen([path])

                elif 'camera' in dest.lower():
                    print("Opening Camera")
                    engine.say(" Opening Camera")
                    engine.runAndWait()
                    os.startfile("microsoft.windows.camera:") 

                elif 'notepad' in dest.lower():
                    print("Opening Notepad")
                    engine.say(" Opening Notepad")
                    engine.runAndWait()
                    subprocess.Popen(['notepad.exe'])

                elif 'screenshot' in dest.lower():
                    print("Give a name to your screenshot")
                    engine.say(" Opening Give a name to your screenshot")
                    engine.runAndWait()
                    audio = r.listen(source)
                    dest = r.recognize_google(audio) 
                    pyautogui.screenshot(dest + ".png") 

                elif 'chrome' in dest.lower():
                    print("Opening Google Chrome")
                    engine.say(" Opening Google Chrome")
                    engine.runAndWait()                   
                    subprocess.Popen(["C:\Program Files\Google\Chrome\Application\chrome.exe"])
                    
                elif 'whatsapp' in dest.lower():
                    print("Opening What's App")
                    engine.say(" Opening What's App")
                    engine.runAndWait()                   
                    subprocess.Popen(["C:\\Users\\oishanu\AppData\\Roaming\\WhatsApp\\WhatsApp.exe"])

                elif 'time' in dest.lower():
                    current_datetime = datetime.datetime.now()
                    current_time = current_datetime.strftime("%I:%M %p")
                    current_date = current_datetime.strftime("%B %d, %Y")
                    spoken_time = current_time.replace(":", " ").replace("AM", "a.m.").replace("PM", "p.m.")
                    print("The time right now is",current_time)
                    print("Today's date is",current_date)
                    engine.say("The current time is " + spoken_time)
                    engine.say("Today's date is " + current_date)
                    engine.runAndWait()
                    
                elif 'weather' in dest.lower():
                    weather()               

            elif 'thank' in dest.lower():
                print('I hope you liked my assistance. Have a nice day\n')
                engine.say(" I hope you liked my assistance. Have a nice day")
                engine.runAndWait()
                exit()

            elif 'time' in dest.lower():
                current_datetime = datetime.datetime.now()
                current_time = current_datetime.strftime("%I:%M %p")
                current_date = current_datetime.strftime("%B %d, %Y")
                spoken_time = current_time.replace(":", " ").replace("AM", "a.m.").replace("PM", "p.m.")
                print("The time right now is",current_time)
                print("Today's date is",current_date,'\n')
                engine.say("The current time is " + spoken_time)
                engine.say("Today's date is " + current_date)
                engine.runAndWait() 
                
            elif 'weather' in dest.lower():
                weather()

            elif 'screenshot' in dest.lower():
                    print("Give a name to your screenshot")
                    engine.say(" Give a name to your screenshot")
                    engine.runAndWait()
                    audio = r.listen(source)
                    dest = r.recognize_google(audio) 
                    pyautogui.screenshot(dest + ".png")
                    print("Screenshot named",dest,"has been saved to your desktop\n")
                    engine.say(" Screenshot named" + dest + "has been saved to your desktop")
                    engine.runAndWait()  
                
        except sr.UnknownValueError:
            print("Sorry, I could not understand what you said.\n")
            engine.say(' Sorry, I could not understand what you said.')
            engine.runAndWait()
        except sr.RequestError as e:
            print("Could not request results from Google Speech Recognition service; {0}".format(e))
            engine.say(' Could not request results from Google Speech Recognition service')
            engine.runAndWait()
        except web.Error as e:
            print("Error opening the web browser:", e)
            engine.say(' Error opening the web browser')
            engine.runAndWait()


def voice_recognition():
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice',voices[0].id)
    print("Mr.Sharky is ready to assist you")
    print("\nPlease do not close the window before 5 seconds\n")
    engine.say(' Mr.Sharky is ready to assist you. Please do not close the window before 5 seconds')
    engine.runAndWait()
    r = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            r.adjust_for_ambient_noise(source)
            audio = r.listen(source, timeout=5)
            dest = r.recognize_google(audio)
            if 'hello' in dest.lower():
                while True:
                    if __name__ == "__main__":
                        main()
        except sr.WaitTimeoutError:
            print("Sorry, you could not respond in time. Mr. Sharky will see you later. Sayonara\n")
            engine.say(' Sorry you could not respond in time, Mr. Sharky will see you later. Sayonara')
            engine.runAndWait()
            exit()

        except sr.UnknownValueError:
            print("Sorry, I could not understand what you said.")
            engine.say(' Sorry, I could not understand what you said.')
            engine.runAndWait()

window = tk.Tk()

def start_voice_recognition():
    voice_thread = Thread(target = voice_recognition)
    voice_thread.start()

def stop_voice_recognition(): 
    window.destroy()
    window.quit()

def start_gui():
    window.title("Voice Assistant")

    output_text = tk.Text(window)
    output_text.pack(fill=tk.BOTH, expand=True)

    sys.stdout = StdoutRedirector(output_text)

    start_button = tk.Button(window, text="Start Voice Assistant", command = start_voice_recognition)
    start_button.pack()

    stop_button = tk.Button(window, text="Stop Voice Assistant", command = stop_voice_recognition)
    stop_button.pack()

    window.mainloop()

def weather():
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)

    r = sr.Recognizer()

    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        print("Tell me the name of the city")
        engine.say(' Tell me the name of the city')
        engine.runAndWait()
        audio1 = r.listen(source,timeout=5)
        CITY = r.recognize_google(audio1)

    BASE_URL = "https://api.openweathermap.org/data/2.5/weather?"
    API_KEY = "c983f5e2ff8596a344dbee0b9ea7f6b2"

    URL = BASE_URL + "q=" + CITY + "&appid=" + API_KEY

    response = requests.get(URL)

    if response.status_code == 200:
        data = response.json()

        main = data['main']

        temperature_kelvin = main['temp']
        temperature_celsius = temperature_kelvin - 273.15
        humidity = main['humidity']
        report = data['weather']

        print('\n',f"{CITY:-^30}")
        engine.say('Giving you the weather report of' + CITY)
        engine.runAndWait()
        
        a = f"Temperature: {temperature_celsius:.2f} Celcius"
        b = f"Humidity: {humidity}%"
        c = f"Weather Report: {report[0]['description']}"
        
        print(a)
        print(b)
        print(c,'\n')
        engine.say( a)
        engine.say( b)
        engine.say( c)
        engine.runAndWait()

    else:
        # showing the error message
        print("Error in the HTTP request")

if __name__ == "__main__":
    start_gui()