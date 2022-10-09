import pyttsx3
import speech_recognition as sr
import webbrowser
import datetime
import wikipedia
import pyjokes
import ctypes
import os
import subprocess
import wolframalpha
from tkinter import *


class VirtualAssistant:
    def __init__(self, root):
        # root = Tk()
        self.root = root
        self.root.title("project")
        #self.root.geometry("380x420")
        self.root.minsize(380, 420)
        self.root.maxsize(380, 420)

        filname = PhotoImage(file="assit.png")
        Label(self.root, image=filname).pack(pady=20)
        btn = Button(self.root, text="Start", font=("times new roman", 15, "bold"), fg="black", bg="white", bd=3,
                     relief=RIDGE, command=self.Take_query)
        btn.pack(pady=20)

        self.root.mainloop()

    def takeCommand(self):
        r = sr.Recognizer()

        with sr.Microphone() as source:
            print('Listening')

            r.pause_threshold = 0.7
            audio = r.listen(source)

            try:
                print("Recognizing")

                Query = r.recognize_google(audio, language='en-in')
                print("command is =", Query)

            except Exception as e:
                print(e)
                print("Say that again sir")
                return "None"

            return Query

    def speak(self, audio):
        engine = pyttsx3.init()

        voices = engine.getProperty('voices')
        engine.setProperty('voice', voices[1].id)
        engine.say(audio)
        engine.runAndWait()

    def tellDay(self):
        day = datetime.datetime.today().weekday() + 1

        Day_dict = {1: 'Monday', 2: 'Tuesday',
                    3: 'Wednesday', 4: 'Thursday',
                    5: 'Friday', 6: 'Saturday',
                    7: 'Sunday'}

        if day in Day_dict.keys():
            day_of_the_week = Day_dict[day]
            print(day_of_the_week)
            self.speak("The day is " + day_of_the_week)

    def tellTime(self):
        time = str(datetime.datetime.now())

        print(time)
        hour = time[11:13]
        min = time[14:16]
        self.speak("The time is sir" + hour + "Hours and" + min + "Minutes")

    def Hello(self):
        # self.wishMe()
        hour = int(datetime.datetime.now().hour)
        if hour >= 0 and hour < 12:
            self.speak("Good Morning Abhishek !")

        elif hour >= 12 and hour < 18:
            self.speak("Good Afternoon Abhishek !")

        else:
            self.speak("Good Evening Abhishek !")
        self.speak("I am your assistant. Tell me how may I help you")

    def Take_query(self):

        self.Hello()
        while (True):

            query = self.takeCommand().lower()
            if "open youtube" in query:
                self.speak("Opening youtube ")
                webbrowser.open("www.youtube.com")

            elif "open google" in query:
                self.speak("Opening Google ")
                webbrowser.open("www.google.com")

            elif "open facebook" in query:
                self.speak("Opening facebook")
                webbrowser.open("www.facebook.com")

            elif "which day it is" in query:
                self.tellDay()

            elif "what is time" in query:
                self.tellTime()

            elif "how are you" in query:
                self.speak("i am fine.")

            elif "bye" in query or "exit" in query:
                self.speak("Thank you for visiting. Bye")
                exit()

            elif "from wikipedia" in query:

                self.speak("Checking the wikipedia ")
                query = query.replace("wikipedia", "")

                result = wikipedia.summary(query, sentences=4)
                self.speak("According to wikipedia")
                self.speak(result)

            elif "who are you" in query:
                self.speak("I am Your Assistant")

            elif "joke" in query:
                self.speak(pyjokes.get_joke())

            elif "play music" in query or "play song" in query:
                self.speak("Here you go with music")
                # music_dir = "G:\\Song
                music_dir = "C:\\Users\\ABHISEK\\Music"
                songs = os.listdir(music_dir)
                print(songs)
                random = os.startfile(os.path.join(music_dir, songs[1]))

            elif "who made you" in query or "who created you" in query:
                self.speak("I have been created by Abhishek Kumar Singh.")

            elif "lock window" in query:
                self.speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

            elif "shut down my laptop" in query:
                self.speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call("shutdown /s /t 1")

            elif "restart my laptop" in query:
                self.speak("Hold On a Sec ! Your system is on its way to restart")
                subprocess.call(["shutdown", "/r"])

            elif 'search' in query:
                self.speak("Wait ! Searching")
                query = query.replace("search", "")
                webbrowser.open(query)

            elif 'open powerpoint presentation' in query:
                self.speak("opening Power Point presentation")
                power = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
                os.startfile(power)

            elif 'open microsoft word' in query:
                self.speak("opening microsoft word")
                power = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
                os.startfile(power)

            elif 'open excel' in query:
                self.speak("opening excel")
                power = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
                os.startfile(power)

            elif "where is" in query:
                query = query.replace("where is", "")
                location = query
                self.speak("User asked to Locate")
                self.speak(location)
                webbrowser.open("https://www.google.com/maps/place/" + location + "")

            elif "what is" in query or "who is" in query:
                client = wolframalpha.Client("266UX7-TEL9VQUJA4")
                res = client.query(query)

                try:
                    print(next(res.results).text)
                    self.speak(next(res.results).text)
                except StopIteration:
                    print("No results")

            elif "open email" in query:
                self.speak("Opening email")
                webbrowser.open("https://mail.google.com/mail/u/0/#inbox")

            elif "write a note" in query:
                self.speak("What should i write, sir")
                note = self.takeCommand()
                file = open('jarvis.txt', 'w')

                file.write(note)
                self.speak("note saved")

            elif "show note" in query:
                self.speak("Showing Notes")
                file = open("jarvis.txt", "r")
                print(file.read())

            elif "calculate" in query:

                try:
                    app_id = "266UX7-TEL9VQUJA4"
                    client = wolframalpha.Client(app_id)
                    indx = query.lower().split().index('calculate')
                    query = query.split()[indx + 1:]
                    res = client.query(' '.join(query))
                    answer = next(res.results).text
                    print("The answer is " + answer)
                    self.speak("The answer is " + answer)
                except:
                    print("something wrong")

root = Tk()
obj = VirtualAssistant(root)
root.mainloop()
