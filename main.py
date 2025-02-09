import win32com.client
import os
import speech_recognition as sr
import webbrowser
import datetime

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    speaker.Speak(f"{text}")

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recogizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"

if __name__ == "__main__":
    say("Jarvis AI")
    while True:
        print("Listening...")
        query = takecommand()
        # todo: Add more sites
        sites = [["Youtube", "https://www.youtube.com"], ["wikipedia" ,"https://www.wikipedia.com"],["Google" ,"https://www.Google.com"],["Typing test","https://www.edclub.com/sportal/"],["Github","https://github.com/dashboard"],["Chat gpt","https://chatgpt.com/"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir....")
                webbrowser.open(site[1])
        # todo: Add music
        if "play the music" in query.lower():
            musicpath = "C:\\Users\\admin\\Music\\goku_black_theme.mp3"
            say(f"Playing the music sir....")
            os.startfile(musicpath)
        # todo: timing feature
        if "the time" in query.lower():
            strfTime = datetime.datetime.now().strftime("%I:%M %p")
            say(f"Sir the time is  {strfTime}")
        # todo: add app
        if "open Chrome".lower() in query.lower():
            os.startfile("\\Users\Public\Desktop\Google Chrome.lnk ")
            say(f"Opening the Chrome sir....")
        if "open vs code".lower() in query.lower():
            os.startfile("C:\\Users\\admin\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code.lnk")
            say(f"Opening the vscode sir....")
        if "Shutdown".lower() in query.lower():
            os.startfile("\\Users\\admin\Desktop\SlideToShutDown.lnk")
            say(f"Shut downing the system sir....")
        if "open Instagram".lower() in query.lower():
            os.startfile("C:\\Users\\admin\Desktop\Instagram.lnk")
            say(f"Opening the Instagram sir....")
        # say(query)
