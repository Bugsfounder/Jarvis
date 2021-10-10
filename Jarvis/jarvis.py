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
        print("Listenning....")
        r.pause_threshold = 1
        r.energy_threshold = 700
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User Said: {query}")

    except Exception as e:
        print("Please, Say That Again...")
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
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("Acording to Wikipedia")
            print(results)
            speak(results)
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open stackoverflow' in query or 'open stack overflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'play music' in query:
            music_dir = 'F:\\manisha folder\\Moods\songs\\favourite☺☺☻☻☻'
            songs = os.listdir(music_dir)
            usingIndex = random.randint(0, len(songs)-1)
            os.startfile(os.path.join(music_dir, songs[usingIndex]))

        elif "the time" in query:
            date = datetime.datetime.now().strftime("%H:%M:%S")
            speak("The Time is: "+str(date))
        elif 'open code' in query:
            codePath = "C:\\Users\\127.0.0.1\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
        elif 'send email' in query:
            try:
                speak("What Should i say!")
                content = takeCommand()
                to = "manishakumari200307@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent")

            except Exception as e:
                print(e)
                speak("Sorry manisha i am not able to send email")
        elif "exit" in query:
            break
