import subprocess
#import wolframalpha
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
#import winshell
import pyjokes
#import feedparser
import smtplib
import ctypes
import time
import shutil
#from clint.textui import progress
from ecapture import ecapture as ec 
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

#The number can set to 0 to get a male voice

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


#The machine learning and commands taken
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour<12:
        speak("Good Morning")
    elif hour>= 12 and hour<18:
        speak ("Good Afternoon Sir")
    else:
        speak("Good Eveing Sir")
    assname = ("Aal, 1 point o") # Here the name can be change to something else
    speak ("I am your Assistnat")
    speak (assname) # The command for his name

def username():
    speak("What should I call you sir?")
    uname = takeCommand()
    speak ("Welcome Mister")
    speak (uname)
    columns = shutil.get_terminal_size().columns

    print("####################".center(columns))
    print("Welcome Mr.", uname.center(columns))
    print("####################".center(columns))
    
    speak("How can I help you today?")

def takeCommand():

    r = sr.Recognizer()

    with sr.Microphone() as source:
    
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language ="en-in")
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize voice.")
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smpt.gmail.com', 587)
    server.ehlo()
    server.starttls()

    #Enable low security in gmail
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()


#Main Function Starts Here:
if __name__ == '__main__':
    clear = lambda: os.system('cls')

    #This function will clean any commands before execution of this py file
    clear()
    wishMe()
    username()

    while True:

        query = takeCommand().lower()
        #All commands said will be store here in 'query' 
        # will then be converted to lowercase for easily 
        #recongnition of command

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("Wikipedia", "")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Here you go to YouTube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here is google Sir\n")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak ("Here you go to Stack Over Flow, Happy Coding")
            webbrowser.open("stackoverflow.com")

        elif 'open hack the box' in query:
            speak ("Here is Hack the Box, Good luck master")
            webbrowser.open("hackthebox.com")

        elif 'the time' in query:
            strTime = dateTime.datetime.now().strftime("% H:% M:% S")
            speak ("Sir the time is {strTime}")
        
        elif 'email to Joe' in query:
            try:
                speak ("What should I say?")
                content = takeCommand()
                to = "Receiver email address"
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as E:
                print(e)
                speak("I am not able to send this email right now")

        elif 'send a mail' in query:
            try:
                speak("What should I add?")
                content = takeCommand()
                speak ("Whome should I send")
                to = input()
                sendEmail(to, content)
                speak("Email has been sent")
            except Exception as e:
                print(e)
                speak("I am not able to send this mail")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that you are doing well")

        elif 'change my name to' in query:
            query = quey.replace("change my name to", "")
            assname = query

        elif "change name" in query:
            speak("What would you like to call me Sir")
            assname = takeCommand()
            speak("Thank you for naming me")

        elif "What's your name" in query or "What is your name" in query:
            speak ("My friends call me")
            speak(assname)
            print("My friends call me", assname)

        elif 'exit' in query or "Goodbye" in query:
            speak("Thank you for your time")
            exit()

        elif "Who made you" in query or "Who created you" in query:
            speak("I was created by Omar")
        
        elif 'joke' in query:
            speak(pyjoke.get_joke())

        elif "calculate" in query:
            app_id = "Wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text
            print("The answer is " + answer) 
            speak("The answer is " + answer) 

        elif 'search' in query or 'play' in query:
             
            query = query.replace("search", "") 
            query = query.replace("play", "")          
            webbrowser.open(query) 
 
        elif "who i am" in query:
            speak("If you talk then definitely your human.")
 
        elif "why you came to world" in query:
            speak("Thanks to Omar. further It's a secret")
 
        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
                 
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")
 
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop aal from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
        
        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
             
        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")
 
        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])
 
        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
         
        elif "show note" in query:
            speak("Showing Notes")
            file = open("aals.txt", "r") 
            print(file.read())
            speak(file.read(6))

        elif "update assistant" in query:
            speak("After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream = True)
             
            with open("Voice.py", "wb") as Pypdf:
                 
                total_length = int(r.headers.get('content-length'))
                 
                for ch in progress.bar(r.iter_content(chunk_size = 2391975),
                                       expected_size =(total_length / 1024) + 1):
                    if ch:
                      Pypdf.write(ch)
                     
        # NPPR9-FWDCX-D2C8J-H872K-2YT43
        elif "aal" in query:
             
            wishMe()
            speak("Aal 1 point o in your service Sir")
            speak(assname)    
        
        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")
 
        elif "Good Morning" in query:
            speak("A warm" +query)
            speak("How are you Sir")
            speak(assname)
