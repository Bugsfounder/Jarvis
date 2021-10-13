import pyttsx3
import datetime
from requests.models import codes
from wikipedia.wikipedia import random
from win32com.client import Dispatch
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random
import datetime
import smtplib
import re
from googlesearch import googlesearch

# GETTING AVAILABLE VOICES FROM SYSTEMS FOR JARVIS
engine = pyttsx3.init('sapi5')
voices = engine.getProperty("voices")  # GET VOICE LIST FROM SYSTEM
engine.setProperty('voice', voices[1].id)  # SET A VOICE FOR JARVIS


# CREATING A SPEAK FUNCTION THAT TAKE AUDIO STRINT AND SPEAK IT
def speak(audio):
    """
    This Function is take string and say the string
    params: audio
    """
    engine.say(audio)  # SAY THE AUDIO
    engine.runAndWait()  # WAIT THE PROGRAM WHILE SAYING

    # speak = Dispatch('SAPI.SpVoice')
    # speak.Speak(audio)


# CREATING A WISH ME FUNCTION THAT WISH USER WHILE START THE FUNCTION
def wishMe():
    """
    This function is just with the user
    """
    hour = int(datetime.datetime.now().hour)  # GETTING HOUR AS INTEGER

    # IF HOUR IS GREATER TAN 0 OR LESS THAN 12 SAY GOOD MORNING TO USER
    if hour >= 0 and hour < 12:
        speak("Good Morning!")

    # IF HOUR IS GREATER TAN 12 OR LESS THAN 18 SAY GOOD AFTERNOON TO USER
    elif hour >= 12 and hour <= 18:
        speak("Good Afternoon!")
    # OTHERWISE SAY GOOD EVENING
    else:
        speak("Good Evening!")

    speak("I am Maahi. Please tell me How may i help you")


# CREATING A TAKECOMMAND FUNCTION TO TAKE AUDIO FROM MICROPHONE AND RETURN STRING OUTPUT
def takeCommand():
    """
    This function is just take arguments from microphone as user voice and return a string output
    """
    r = sr.Recognizer()
    with sr.Microphone(device_index=1) as source:
        # print("Listenning....")
        # speak("Listenning....")
        r.pause_threshold = 1
        # r.energy_threshold = 700
        audio = r.listen(source)
    try:
        # print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User Said: {query}")

    except Exception as e:
        # print("Please, Say That Again...")
        # speak("Please, Say That Again Clearly...")
        return "None"

    return query


def sendEmail(to, content):
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.ehlo()
    server.starttls()

    # SAVE YOUR EMAIL PASSWORD IN A FILE AND READ FROM THERE FOR SECURITY
    server.login("bugsfounder2021@gmail.com", "FOUNDER2021")
    server.sendmail("bugsfounder2021@gmail.com", to, content)
    server.close()


if __name__ == '__main__':
    wishMe()
    while True:
        query = takeCommand().lower()

        # LOGIC FOR EXECUTING TASKS BASED ON QUERY
        if 'wikipedia' in query:
            speak("Taking Action on: "+query)
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("Acording to Wikipedia")
            # print(results)
            speak(results)
        elif 'open youtube' in query:
            speak("Taking Action on: "+query)
            speak("Openning Youtube")
            webbrowser.open("youtube.com")

        elif 'open stackoverflow' in query or 'open stack overflow' in query:
            speak("Taking Action on: "+query)
            speak("Openning stackoverflow")
            webbrowser.open("stackoverflow.com")

        elif 'open google' in query:
            speak("Taking Action on: "+query)
            speak("Openning Google")
            # webbrowser.open("google.com")
            webbrowser.open(
                "https://www.google.com/search?q=pygame+documentation&rlz=1C1RXQR_enIN973IN973&oq=pygame+&aqs=chrome.1.69i57j69i59j0i433i512l2j0i512l6.2639j0j7&sourceid=chrome&ie=UTF-8")

        elif 'play music' in query:
            speak("Taking Action on: "+query)
            speak("Searching for the musing for Play")
            music_dir = 'F:\\manisha folder\\Moods\songs\\favourite☺☺☻☻☻'
            songs = os.listdir(music_dir)
            usingIndex = random.randint(0, len(songs)-1)
            os.startfile(os.path.join(music_dir, songs[usingIndex]))

        elif "the time" in query:
            speak("Taking Action on: "+query)
            date = datetime.datetime.now().strftime("%H:%M:%S")
            speak("The Time is: "+str(date))
        elif 'open code' in query:
            speak("Taking Action on: "+query)
            open("Visual Studio Code")
            codePath = "C:\\Users\\127.0.0.1\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
        elif 'send email' in query:
            try:
                speak("Taking Action on: "+query)
                speak("What Should i say!")
                content = takeCommand()
                to = "manishakumari200307@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent")

            except Exception as e:
                speak("Sorry manisha i am not able to send email")
        elif "exit" in query:
            break
        elif "who are you" in query:
            speak("My name is maahi and i am a desktop assistant of manisha")
        elif "how are you" in query:
            speak("I am fine, how are you")
        elif "i am fine" in query:
            speak("Superb Whats you next plan")
        elif "search google for" in query:
            import pywhatkit as kt
            pattern = re.findall(r'search google for [\S+\s\S+]+', str(query))
            searchQ = [i for i in pattern]
            searchQuery = str(searchQ[0]).replace("search google for ", "")
            speak("Let's perform Google search!")
            target = searchQuery
            kt.search(target)
        elif "press" in query:
            import pyautogui
            key = query.replace("press ", "")
            pyautogui.press(key)
        elif "write" in query:
            import pyautogui
            key = query.replace("write ", " ")
            pyautogui.typewrite(key)
        else:
            print("say something")
