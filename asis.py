from gtts import gTTS
import speech_recognition as sr
import os
import re
import webbrowser
import requests
#from weather import Weather
import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")
#import pyttsx3
#listener = sr.Recognizer()
#engine = pyttsx3.init()
#voices = engine.getProperty('voices')
#engine.setProperty('voice', voices[1].id)


def talkToMe(audio):
    #"speaks audio passed as argument"
    #print(audio)
    #engine.say(audio)
    #engine.runAndWait()
    #"speaks audio passed as argument"
    print(audio)
    speak.Speak(audio)

def myCommand():
    #"listens for commands"

    r = sr.Recognizer()

    with sr.Microphone() as source:
        speak_command = 'Please give some command to me...'
        talkToMe(speak_command)
        
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)

    try:
        command = r.recognize_google(audio).lower()
        talkToMe('You said: ' + command + '\n')

    #loop back to continue to listen for commands if unrecognizable speech is received
    except sr.UnknownValueError:
        talkToMe('Your last command couldn\'t be heard ! I can understand commands like send email, open gmail, open website xyz.com and tell me a joke')
        #speak.Speak('Your last command couldn\'t be heard')
        command = myCommand()

    return command


def assistant(command):
    #"if statements for executing commands"
    message = 'Ask me to do something, I am not here for chitchat ! I can understand commands like send email, open gmail, open website xyz.com and tell me a joke'
    yourself = 'Give me your money'
    if 'hello' in command:
        talkToMe(message)

    elif 'hi' in command:
        talkToMe(message)

    elif 'hey' in command:
        talkToMe(yourself)
    
    elif 'plagiarism' in command:
        talkToMe('Plagiarism is the representation of another authors language, thoughts, ideas, or expressions as one is own original work. In educational contexts, there are differing definitions of plagiarism depending on the institution. Plagiarism is considered a violation of academic integrity and a breach of journalistic ethics. It is subject to sanctions such as penalties, suspension, expulsion from school or work, substantial fines and even incarceration. Recently, cases of extreme plagiarism have been identified in academia')

    elif 'open plagiarism' in command:
        #reg_ex = re.search('open gmail (.*)', command)
        url = 'https://www.grammarly.com/'
        webbrowser.open(url)
        talkToMe('Done!')

    elif 'open website' in command:
        reg_ex = re.search('open website (.+)', command)
        if reg_ex:
            domain = reg_ex.group(1)
            url = 'https://www.' + domain +".com"
            webbrowser.open(url)
            talkToMe('Done!')
        else:
            pass

    elif 'open notepad' in command:
        os.system('notepad')
        talkToMe('Done for you!')

    elif 'whats up' in command:
        talkToMe('Just doing my thing')

    elif 'tell me love' in command:
        talkToMe('I Love You So Much and Nothing is not for you !')

    elif 'tell me a joke' in command:
        res = requests.get(
                'https://icanhazdadjoke.com/',
                headers={"Accept":"application/json"}
                )
        if res.status_code == requests.codes.ok:
            talkToMe('Here is an awesome joke for you- ')
            talkToMe(str(res.json()['joke']))
        else:
            talkToMe('oops!I ran out of jokes')

    #this option is not funtioning as proxy need to set to bypass HMCL IT settings
    elif 'weather forecast in' in command:
        reg_ex = re.search('weather forecast in (.+)', command)
        if reg_ex:
            city = reg_ex.group(1)
            weather = Weather()
            location = weather.lookup_by_location(city)
            forecasts = location.forecast()
            for i in range(0,3):
                talkToMe('On %s will it %s. The maximum temperture will be %.1f degree.'
                         'The lowest temperature will be %.1f degrees.' % (forecasts[i].date(), forecasts[i].text(), (int(forecasts[i].high())-32)/1.8, (int(forecasts[i].low())-32)/1.8))


    elif 'bye' in command:
        talkToMe('See you later Take care my master!')
        exit()

    else:
            talkToMe('commands yang dapat dipahami, open gmail, open website xyz.com and tell me a joke')

talkToMe('Hey, I am alvian junior for your laptop, tell me your commands !')

#loop to continue executing multiple commands
while True:
    assistant(myCommand())