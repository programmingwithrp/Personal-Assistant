# What you can do with it
# open Google, instagram , facebook , youtube ,  visual studio code , Pycharm , Notepad, open folder in your pc with given path , any topic wikipedia
# you can shutdown
# if it will not recognize any query then it will give you google link and also it will Automaticlly open that link..


# APP NAME: Rishi
#
# APPID: WYY24K-VA3VVGAG88
#
# USAGE TYPE: Personal/Non-commercial Only



# for input voice
import speech_recognition as sr

# import pyaudio
import pyttsx3

# for opening webbrower any link or your pc path location..
import webbrowser

# direct fatch data from google
from googlesearch import search

# for os operation..
import os

# for searching wikipedia...
import wikipedia

# for taking exact time..
import time
from datetime import datetime
import datetime

import wolframalpha

import cv2

import pywhatkit as kit

import sys
# All Functions




engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)



# we can Use below three function to speak ....

# def SpeakText(command):
#     # Initialize the engine
#     engine = pyttsx3.init()
#     engine.say(command)
#     engine.runAndWait()


# def speak(str):
#     from win32com.client import Dispatch
#     speak = Dispatch("SAPI.SpVoice")
#     speak.Speak(str)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wish():
    # nowtime = datetime.datetime.today()
    # nowtime = datetime.now().time()
    nowtime = datetime.datetime.now()
    if nowtime.hour >= 0 and nowtime.hour <= 11:
        print("Good Morning Rishi Sir...")
        speak("Good Morning Rishi Sir...")
        print("How may I help You..")
        speak("How may I help You..")
    elif nowtime.hour == 12 and nowtime.minute == 0:
        print("Good Noon Rishi sir...")
        speak("Good Noon Rishi sir...")
        print("how may i help you")
        speak("how may i help you")

    elif nowtime.hour >= 12 and nowtime.hour <= 15:
        print("Good Afternoon Rishi Sir...")
        speak("Good Afternoon Rishi Sir...")
        print("How may I help You..")
        speak("How may I help You..")
    elif nowtime.hour >= 16 and nowtime.hour <= 19:
        print("Good Evening Rishi Sir...")
        speak("Good Evening Rishi Sir...")
        print("How may I help You..")
        speak("How may I help You..")
    else:
        speak("Good Night Rishi Sir...")
        speak("Sleep well sir..")


def takecommand():
    # Initialize the recognizer
    r = sr.Recognizer()

    with sr.Microphone() as source2:
        # wait for a second to let the recognizer
        # adjust the energy threshold based on
        # the surrounding noise level
        r.adjust_for_ambient_noise(source2, duration=0.2)

        # listens for the user's input
        print("listening....")
        r.pause_threshold = 1
        our_audio = r.listen(source2)

        try:
            print("recognizing...")
            # Using google to recognize audio
            YourText = r.recognize_google(our_audio, language='en-in')
            print(f"did you say this {YourText}")
        except Exception as e:
            print(e)
            print("Say that again please")
            return None
        return YourText


# register chrome brower woth their part
webbrowser.register('chrome',
                    None,
                    webbrowser.BackgroundBrowser(
                        "C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))

if __name__ == '__main__':
    wish()

    while True:

        AudioToText = takecommand().lower()

        try:

            if "google" in AudioToText:
                webbrowser.get('chrome').open("https://www.google.com")

            elif "close chrome" in AudioToText:
                os.system('taskkill /f /im chrome.exe')

            elif "time" in AudioToText:
                currentime = time.asctime( time.localtime(time.time()) )
                print(currentime)
                speak(f"Date and Time : {currentime}")

            elif "send message in whatsapp" in AudioToText:

                kit.sendwhatmsg("+919100090525", "hi i am rishi from python program... I am not your Rishi", 12, 5)

            elif "facebook" in AudioToText:
                # speak("you are speaking facebook right..")
                speak("Opening Facebook")
                webbrowser.get('chrome').open('https://www.facebook.com/')
            elif "instagram" in AudioToText:
                # speak("you are speaking instagram right..")
                speak("Opening instagram")
                webbrowser.get('chrome').open('https://www.instagram.com/')

            elif "youtube" in AudioToText or "video" in AudioToText:
                speak("Opening youtube")
                webbrowser.get('chrome').open('https://www.youtube.com/')

            elif "visual studio code" in AudioToText or "code" in AudioToText:
                webbrowser.open(
                    "C://Users//patel//AppData//Local//Programs//Microsoft VS Code//Code.exe")

            elif "pycharm" in AudioToText:
                webbrowser.open(
                    "C://Users//patel//AppData//Local//Programs//PyCharm Community Edition 2020.1.1//bin//pycharm64.exe")
            elif "notepad" in AudioToText:
                webbrowser.open(
                    "C://Users//patel//AppData//Roaming//Microsoft//Windows//Start Menu//Programs//Accessories//Notepad.lnk")
            elif "marvel cinema" in AudioToText:
                webbrowser.open("E://marvel sequense")


            elif "command" in AudioToText:
                os.system('start cmd')

            elif "close command" in AudioToText:
                os.system('KILL')
            # elif "close" in AudioToText:
            # try:
            #     os.system('TASKKILL /F /IM C://Users/patel//PycharmProjects//pythonProject//project//assistant.py')
            # except Exception as e:
            #     print(e)
            # os.stop()

            elif "wikipedia" in AudioToText:

                print(wikipedia.summary(AudioToText, sentences=2))



            elif "take a photo" in AudioToText:

                cap = cv2.VideoCapture(0)
                while True:
                    ret,img = cap.read()
                    cv2.imshow('webcam',img)
                    k = cv2.waitKey(3)
                    if k == 27:
                        break
                cap.release()
                cv2.destroyAllWindows()

            elif "tell me" in AudioToText:
                speak(
                    "i can answer to computational and geographical question do you want to ask now..")
                question = takecommand().lower()
                print(f"Your question : {question}?")
                print("Wait a moment sir..")
                speak("wait a moment sir..")
                try:
                    app_id = 'WYY24K-VA3VVGAG88'
                    client = wolframalpha.Client(app_id)
                    res = client.query(question)
                    answer = next(res.results).text
                    print(f"Your Answer : {answer}.")
                    speak(answer)
                except:
                    speak("Sorry Sir , I can not find answer from my data..")

            elif ("shut down" and "down") in AudioToText:
                # listens for the user's input
                print("Do you want to shutdown.....")
                speak("Do you want to shutdown sir.....")

                try:
                    lastans = takecommand().lower()

                except:
                    lastans = "no"

                if "yes" in lastans:
                    print("shutting down your computer..")
                    speak("shutting down your computer")
                    os.system("shutdown /s /t 30")

                elif "no" in lastans:
                    print("Thank you sir")
                    speak("Keep it up sir")




            # SpeakText(AudioToText)
            # speak(AudioToText)

            else:
                print("I have some information link about your query..")
                speak("I have some information link about your query..")
                print("May be it will help you.")
                speak("May be it will help you.")
                speak("please wait a moment sir..")
                for j in search(AudioToText, tld="co.in", num=1, stop=1, pause=2):
                    print(j)
                    speak("opening for you...")
                    webbrowser.get('chrome').open(j)

        except Exception as e:
            print(e)

        finally:

            speak("thank you for using me..")
            print("Do you want to continue.....")
            speak("Do you want to continue sir.....")

            try:
                loopans = takecommand().lower()
            except Exception as e:
                loopans = "no"
                print("exeption")

            if "continue" in loopans:

                print("Ok sir \n")
                speak("ok Sir")


            else:

                print("Thank you sir")
                speak("Thank you sir")
                print("Keep it up sir")
                speak("Keep it up sir")
                break
