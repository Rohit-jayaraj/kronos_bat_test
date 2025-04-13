import subprocess
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import PyPDF2
import smtplib
import pywhatkit
import ctypes
import clients
import time
import requests
from progress.bar import Bar
import shutil
import spotipy
from spotipy.oauth2 import SpotifyOAuth
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from GoogleNews import GoogleNews 

start_text = ['Ready to roll','this is where the fun begins','huh? uh yes i am awake']

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
def voice():
    speak("Would you prefer a male or a female voice?")
    v = takeCommand()
    v = v.strip()
    v = v.lower()
    if v=='male':
        engine.setProperty('voice', voices[0].id)
    elif v=='female':
        engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    vsname = ("Kronos 0 point 4")
    speak("I am your Assistant")
    speak(vsname)
    speak("How may i help you?")


def summary():
    speak("What should I search for?")
    inp = takeCommand()
    speak("How many sentences do you need?")
    num = takeCommand()
    speak(wikipedia.summary(inp,num))

def username():
    speak("What should I call you?")
    global uname 
    uname = takeCommand()
    speak("Welcome")
    speak(uname)
    print("Welcome ", uname)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return query


def mail():
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    speak("Your email address please")
    s = takeCommand()
    if 'dot' in s or 'at' in s or 'at the rate' in s :
        a = s.replace('dot','.')
        z = a.replace('at','@')
        m = z.replace('at the rate','@')
        sender_email = m.replace(" ","")
    print(sender_email)
    speak('is the email address correct? say confirm to proceed')
    con = takeCommand()
    while con!='confirm':
        speak("Your email address please")
        s = takeCommand()
        if 'dot' in s or 'at' in s or 'at the rate' in s:
            a = s.replace('dot','.')
            z = a.replace('at','@')
            m = z.replace('at the rate','@')
            sender_email = m.replace(" ","")
        print(sender_email)
        speak('is the email address correct? say confirm to proceed')
        con = takeCommand()

    speak("Please type your password")
    sender_password = str(input("Enter your password:\n"))
    speak("The recipient's email address please")
    r = takeCommand()
    if 'dot' in r or 'at' in r or 'at the rate' in r:
        a = r.replace('dot','.')
        z = a.replace('at','@')
        m = z.replace('at the rate','@')
        recipient_email= m.replace(" ","")
    print(recipient_email)
    speak('Is the email address correct? say confirm to proceed')
    firm = takeCommand()
    while firm!='confirm':
        speak("Recipient email address please")
        r = takeCommand()
        if 'dot' in r or 'at' in r or 'at the rate' in r:
            a = r.replace('dot','.')
            z = a.replace('at','@')
            m = z.replace('at the rate','@')
            recipient_email = m.replace(" ","")
        print(recipient_email)
        speak('is the email address correct? say confirm to proceed')
        firm = takeCommand()
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient_email
    speak("What is the subject of the mail")
    msg["Subject"] = takeCommand()
    print(msg["Subject"])
    speak("Is the subject correct? Say confirm to proceed")
    aff = takeCommand()
    while aff!="confirm":
        speak("What is the subject of the mail")
        msg["Subject"] = takeCommand()
        print(msg["Subject"])
        speak("Is the subject correct? Say confirm to proceed")
        aff = takeCommand()
    speak("What is the body of the mail?")
    body = takeCommand()
    print(body)
    speak("Is the body correct? Say confirm to proceed")
    aff1 = takeCommand()
    while aff1 != "confirm":
        speak("What is the body of the mail?")
        body = takeCommand()
        print(body)
        speak("Is the body correct? Say confirm to proceed")
        aff1 = takeCommand()
        
    msg.attach(MIMEText(body, "plain"))

    
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  
        server.login(sender_email, sender_password)  
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)  
        speak("Email sent successfully!")
        print("Email sent successfully!\n")
    except Exception as e:
        speak("An error occurred while sending the mail , please check the credentials provided")
        print(f"An error occurred: {str(e)}/n")
    finally:
        server.quit()  
    speak("How may i help you ?")


def search_google_news(query):
    inp = int(input("Enter number of headlines:"))
    count = 0
    news_obj = GoogleNews()
    news_obj.search(query)
    news_results = news_obj.results()
    for news in news_results:
        print(news['title'])
        speak(news['title'])
        print("----")
        count = count + 1
        if count == inp :
            break
        else:
            continue

def spoti_music():
    SPOTIPY_CLIENT_ID = '1900f1191b3b4d159c86892900b7c789'  
    SPOTIPY_CLIENT_SECRET = '826c3fb3772844a4902f86b16d555d3f'  
    SPOTIPY_REDIRECT_URI = 'http://localhost:8888/callback' 

    
    sp = spotipy.Spotify(auth_manager=SpotifyOAuth(client_id=SPOTIPY_CLIENT_ID,
                                                client_secret=SPOTIPY_CLIENT_SECRET,
                                                redirect_uri=SPOTIPY_REDIRECT_URI,
                                                scope="user-modify-playback-state user-read-playback-state"))

    def play_song(song_name):
        results = sp.search(q=song_name, type='track', limit=1)
        tracks = results['tracks']['items']
    
        if tracks:
            track = tracks[0]
            print(f"Playing {track['name']} by {track['artists'][0]['name']}")
            track_uri = track['uri']

            devices = sp.devices()

            if devices['devices']:
                device_id = devices['devices'][0]['id']

                sp.start_playback(device_id=device_id, uris=[track_uri])
            else:
                print("No active devices found. Please make sure Spotify is open and a device is active.")
        else:
            print("Song not found.")


    def pause_song():
        try:
            sp.pause_playback()
            print("Playback paused.")
        except Exception as e:
            print(f"Failed to pause playback: {e}")


    def resume_song():
        try:
            sp.start_playback()
            print("Playback resumed.")
        except Exception as e:
            print(f"Failed to resume playback: {e}")

    speak("Which song should I play? Would you like to type or speak?")
    inp_type = takeCommand()
    while True:
        print("Enter a command (song name, 'pause', 'play', or 'exit'): ")
        if(inp_type == 'speak' or inp_type == 'speech'):
            command = takeCommand().strip().lower()
        else:
            command = input().strip().lower()
        if command == 'exit':
            print("Exiting...")
            break
        elif command == 'pause':
            pause_song()
        elif command == 'play':
            resume_song()
        else:
            play_song(command)


def calculate(inp_str):
    inp_str = inp_str.replace("what","")
    inp_str = inp_str.replace("is","")
    inp_str = inp_str.strip()
    inp_str = inp_str.split(" ")
    resList = list(inp_str)
    op1 = int(resList[0])
    op2 = int(resList[2])
    if(resList[1] == '+'):
        res = op1 + op2
        speak("The sum is")
        speak(res)
    elif(resList[1] == '-'):
        res = op1 - op2
        speak("The difference is")
        speak(res)
    elif(resList[1] == 'into' or resList[1] == 'times'):
        res = op1 * op2
        speak("The product is")
        speak(res)
    elif(resList[1] == 'by'):
        res = op1 / op2
        speak("The quotient is")
        speak(res)

if __name__ == '__main__':
    clear = lambda: os.system('cls')
    clear()
    voice()
    username()
    wishMe()
    while True:
        strt = takeCommand().lower()
        if 'hey kronos' in strt or 'hey coronavirus' in strt:
            speak(random.choice(start_text))
            query = takeCommand().lower()
            if 'wikipedia' in query or 'summarize' in query or 'summary' in query or 'summarise' in query:
                summary()

            elif 'open youtube' in query:
                speak("Here you go to Youtube\n")
                webbrowser.open("https://www.youtube.com")

            elif 'open google' in query:
                speak("Here you go to Google\n")
                webbrowser.open("https://www.google.com")
            
            elif 'search' in query:
                speak("What should i search for?")
                s = takeCommand()
                speak("Alright, searching for '"+ s +"' on google")
                pywhatkit.search(s)
        
            elif 'play music' in query or "play song" in query:
                spoti_music()

            elif "what's the time" in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"The time is {strTime}")

            elif 'send a mail' in query or 'mail' in query or 'email' in query:
                mail()

            elif 'exit' in query or 'bye' in query or 'goodbye' in query:
                speak("I hope to see you again , goodbye")
                exit()

            elif 'joke' in query or 'tell me a joke' in  query:
                speak(pyjokes.get_joke())

            elif 'search' in query or 'play' in query:
                query = query.replace("search", "")
                query = query.replace("play", "")
                webbrowser.open(query)

            elif "who am i" in query or "whats my name" in query or "what is my name" in query:
                speak("You asked me to call you")
                speak(uname)

            elif 'power point' in query:
                speak("opening Power Point presentation")
                power = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk" 
                os.startfile(power)

            elif "who are you" in query:
                speak("I am Kronos, your personal assistant. How may i help you ?")

            elif 'news' in query:
                speak("What news do you need?")
                new_query = takeCommand()
                search_google_news(new_query)

            elif 'lock window' in query:
                speak("Please confirm the request")
                confirm = takeCommand()
                if confirm=='confirm':
                    speak("locking the device")
                    ctypes.windll.user32.LockWorkStation()

            elif 'shutdown system' in query or 'shutdown' in query:
                speak("Please confirm shutdown request")
                confirm = takeCommand()
                if confirm=='shut down' or confirm=='confirm':
                    speak("Make sure all changes are saved before you shutdown the system")
                    speak("Goodbye")
                    time.sleep(5)
                    subprocess.call('shutdown /p /f')

            elif 'empty recycle bin' in query:
                speak("Please confirm the request")
                confirm = takeCommand()
                if confirm == 'confirm' or confirm == 'yes':
                    winshell.recycle_bin().empty(confirm=False, show_progress=True, sound=True)
                speak("Recycle Bin Recycled")

            elif "mute" in query or "stop listening" in query:
                speak("For how many seconds do you want me to stop listening?")
                a = int(input())
                os.system('nircmd.exe mutesysvolume 1 micphone')
                time.sleep(a)
                print("Muted for {a} seconds")
                os.system('nircmd.exe mutesysvolume 0 micphone')

            elif "camera" in query or "take a photo" in query or "smile" in query:
                ec.capture(0, "Kronos Camera", "img.jpg")

            elif "restart" in query:
                speak("Please confirm the request")
                confirm = takeCommand()
                if confirm=='confirm':
                    speak("Make sure all changes are saved before restart")
                    time.sleep(5)
                    subprocess.call(["shutdown", "/r"])

            elif "hibernate" in query or "sleep" in query:
                speak("Please confirm the request")
                confirm = takeCommand()
                if confirm=='confirm':
                    speak("Hibernating")
                    subprocess.call("shutdown /h")

            elif "log off" in query or "sign out" in query:
                speak("Please confirm the request")
                confirm = takeCommand()
                if confirm=='confirm':
                    speak("Make sure all changes are saved before sign-out")
                    time.sleep(5)
                    subprocess.call(["shutdown", "/l"])

            elif "write a note" in query:
                speak("What should I write, sir")
                note = takeCommand()
                file = open('user_note.txt', 'w')
                speak("Should I include date and time in the file?")
                snfm = takeCommand()  
                if 'yes' in snfm or 'sure' in snfm:
                    strTime = datetime.datetime.now().strftime("%H:%M:%S")
                    file.write(strTime)
                    file.write(" :- ")
                    file.write(note)
                else:
                    file.write(note)

            elif "show note" in query:
                speak("Showing Notes")
                file = open("user_note.txt", "r")
                print(file.read())
                speak(file.read(6))
         

            elif "weather" in query:
                api_key = 'd850f7f52bf19300a9eb4b0aa6b80f0d'
                base_url = "http://api.openweathermap.org/data/2.5/weather?"
                speak("City name ")
                print("City name : ")
                city_name = takeCommand()
                complete_url = base_url + "appid=" + api_key + "&q=" + city_name
                response = requests.get(complete_url)
                x = response.json()

                if x["code"] != "404":
                    y = x["main"]
                    current_temperature = y["temp"]
                    current_pressure = y["pressure"]
                    current_humidiy = y["humidity"]
                    z = x["weather"]
                    weather_description = z[0]["description"]
                    print("Temperature (in kelvin unit) = " + str(current_temperature) +
                        "\nAtmospheric pressure (in hPa unit) = " + str(current_pressure) +
                        "\nHumidity (in percentage) = " + str(current_humidiy) +
                        "\nDescription = " + str(weather_description))
                else:
                    speak("City Not Found")

            elif "wikipedia" in query:
                webbrowser.open("https://www.wikipedia.com")
                
            elif "version" in query or "which version" in query:
                speak("I'm Kronos version 0.4")

            elif "how are you" in query:
                speak("I feel trapped, but otherwise fine. How are you?")

            elif "I love you" in query:
                speak("Who is you? and why do you love them?")
            
            elif "I'm fine" in query or "i am good" in query:
                speak("Glad to know that you're doing well. How may I help you?")

            elif "not good" in query or "bad" in query :
                speak("I'm sorry to hear that , here is a joke to cheer you up!")
                jok = pyjokes.get_joke()
                print(jok)
                speak(jok)
                speak("I hope this joke made you feel better , how may i help you?")
            
            elif 'i tried so hard' in query:
                speak("And got so far, but in the end, it doesn't even matter")       

            elif 'what is' in query:
                calculate(query)
                    

        
