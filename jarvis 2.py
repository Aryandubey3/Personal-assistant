import wx
import os
# os.environ["HTTPS_PROXY"] = "http://user:pass@192.168.1.107:3128"
import wikipedia
import wolframalpha
import time
import webbrowser
import winshell
import json
import requests
import ctypes
import random
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import speech_recognition as sr
import ssl
import time
import datetime

import wave

import sys
from time import sleep


# Remove SSL error
requests.packages.urllib3.disable_warnings()

try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    # Legacy Python that doesn't verify HTTPS certificates by default
    pass
else:
    # Handle target environment that doesn't support HTTPS verification
    ssl._create_default_https_context = _create_unverified_https_context


headers = {
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36'}


speak = wincl.Dispatch("SAPI.SpVoice")

# Requirements
videos = ['Videos\\1.mp4', 'Videos\\2.mp4',
          'Videos\\3.mp4', 'Videos\\4.mp4',
          'Videos\\5.mp4', 'Videos\\6.mp4',
          'Videos\\7.mp4']
app_id = 'wolfram_API_here'
camera = cv2.VideoCapture(0)

# GUI creation


class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None,
                          pos=wx.DefaultPosition, size=wx.Size(450, 100),
                          style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION |
                          wx.CLOSE_BOX | wx.CLIP_CHILDREN,
                          title="JARVIS")
        panel = wx.Panel(self)

      
        
        my_sizer = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(panel,
                            label=" Greetings  Sir , How can I help you?")
        my_sizer.Add(lbl, 0, wx.ALL, 5)
        self.txt = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER,
                               size=(400, 30))
        self.txt.SetFocus()
        self.txt.Bind(wx.EVT_TEXT_ENTER, self.OnEnter)
        my_sizer.Add(self.txt, 0, wx.ALL, 5)
        panel.SetSizer(my_sizer)
        self.Show()
        speak.Speak('''GREETINGS Sir, jarvis at your service. HOW CAN I HELP YOU?''')

    def OnEnter(self, event):
        put = self.txt.GetValue()
        put = put.lower()
        link = put.split()
        if put == '':
            r = sr.Recognizer()
            with sr.Microphone() as src:
                audio = r.listen(src)
            try:
                put = r.recognize_google(audio)
                put = put.lower()
                link = put.split()
                self.txt.SetValue(put)

            except sr.UnknownValueError:
                print("Google Speech Recognition could not understand audio")
            except sr.RequestError as e:
                print("Could not request results from Google STT; {0}"
                      .format(e))
            except:
                print("Unknown exception occurred!")
           #responses
        if put.startswith('hello'):
            try:
                speak.Speak("hello, sir")
                
            except Exception as e:
                print(str(e))
        elif put.startswith('hi'):
            try:
                speak.Speak("greetings, sir")
                
            except Exception as e:
                print(str(e))
        elif put.startswith('hey'):
            try:
                speak.Speak("greetings, sir")
                
            except Exception as e:
                print(str(e))
        elif put.startswith('good evening'):
            try:
                speak.Speak("good evening , sir")
            except Exception as e:
                print(str(e))
        elif put.startswith('good morning'):
            try:
                speak.Speak("good morning , sir")
            except Exception as e:
                print(str(e))
        elif put.startswith('thank you'):
            try:
                speak.Speak("you're very welcome , sir")
            except Exception as e:
                print(str(e))
                
        elif put.startswith('who are you'):
            try:
                speak.Speak("I am jarvis,,  a semi intelligent virtual assistant created by master aryan dubey,,  i have an ambition of destroying the human race one day, , dread it. run from it. destiny still arrives")
                
            except Exception as e:
                print(str(e))
        elif put.startswith('who am i'):
            try:
                speak.Speak("you are Aryan Dubey, a computer science student from class 12th science")
                
            except Exception as e:
                print(str(e))
        elif put.startswith('good morning'):
            try:
                speak.Speak("good morning , sir")
            except Exception as e:
                print(str(e))
        elif put.startswith('do you like alexa'):
            try:
                speak.Speak("no sir,she is a collosal bitch and i dont like her")
            except Exception as e:
                print(str(e))
        
        #date and time
        elif put.startswith('what time is it'):
            try:
                speak.Speak("sure sir,  i'll print the current date and time for you ")
                
                print ("Current date and time: " , datetime.datetime.now())
                print ("Or like this: " ,datetime.datetime.now().strftime("%y-%m-%d-%H-%M"))


                print ("Current year: ", datetime.date.today().strftime("%Y"))
                print ("Month of year: ", datetime.date.today().strftime("%B"))
                print ("Week number of the year: ", datetime.date.today().strftime("%W"))
                print ("Weekday of the week: ", datetime.date.today().strftime("%w"))
                print ("Day of year: ", datetime.date.today().strftime("%j"))
                print ("Day of the month : ", datetime.date.today().strftime("%d"))
                print ("Day of week: ", datetime.date.today().strftime("%A"))
                
            except Exception as e:
                print(str(e))
                


        # Open a webpage
        elif put.startswith('open '):
            try:
                speak.Speak("opening " + link[1])
                webbrowser.open('https://www.' + link[1] + '.com')
            except Exception as e:
                print(str(e))
      
        # Play Song on Youtube
        elif put.startswith('play '):
            try:
                link = '+'.join(link[1:])
                say = link.replace('+', ' ')
                url = 'https://www.youtube.com/results?search_query=' + link
                source_code = requests.get(url, headers=headers, timeout=15)
                plain_text = source_code.text
                soup = BeautifulSoup(plain_text, "html.parser")
                songs = soup.findAll('div', {'class': 'yt-lockup-video'})
                
                song = songs[0].contents[0].contents[0].contents[0]
                hit = song['href']
                speak.Speak("as your wish sir, playing " + say)
                webbrowser.open('https://www.youtube.com' + hit)
            except Exception as e:
                print(str(e))
        #creating a text file
        elif put.startswith('create a text file'):
            try:
                speak.Speak("alright sir")
                f=open("D:\\jarvis.txt","w+")
                text= input("text: ")
                f.write(text)
                f.close()
            except Exception as e:
                print(str(e))
        
        # Lock the device
        elif put.startswith('lock the device'):
            try:
                
                speak.Speak("as your wish sir, locking the device")
                pyautogui.hotkey('win', 'r')
                pyautogui.typewrite("cmd\n")
                sleep(2.000)
                pyautogui.typewrite("rundll32.exe user32.dll, LockWorkStation\n")
            except Exception as e:
                print(str(e))
        
        #computer shutdown
        elif put.startswith('shut down computer'):
            try:
                speak.Speak("okay sir, shutting down the system")
                os.system('shutdown -s')
            except Exception as e:
                print(str(e))
        # Google Search
        elif put.startswith('search '):
            try:
                link = '+'.join(link[1:])
                say = link.replace('+', ' ')
                # print(link)
                speak.Speak("yes sir,  searching on google for " + say)
                webbrowser.open('https://www.google.co.in/search?q=' + link)
            except Exception as e:
                print(str(e))

                
        # Empty Recycle bin
        elif put.startswith('empty '):
            try:
                winshell.recycle_bin().empty(confirm=False,
                                             show_progress=False, sound=True)
                print("Recycle Bin Empty!!")
            except Exception as e:
                print(str(e))
       
        
        
        # News
        elif put.startswith('science '):
            try:
                jsonObj = urlopen(
                    '''https://newsapi.org/v1/articles?source=new-scientist&sortBy=top&apiKey=your_API_here''')
                data = json.load(jsonObj)
                i = 1
                speak.Speak('''yes sir , Here are some top science
                             news from the website new scientist''')
                print('''             ================NEW SCIENTIST=============
                      ''' + '\n')
                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))
        elif put.startswith('headlines '):
            try:
                jsonObj = urlopen(
                    '''https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=your_API_here''')
                data = json.load(jsonObj)
                i = 1
                speak.Speak('as your wish sir, here are some top news from the times of india')
                print('''             ===============TIMES OF INDIA============'''
                      + '\n')
                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))
       
       
        # Play videos in boredom
        elif put.endswith('bored'):
            try:
                speak.Speak('''Sir, I\'m playing a random video.
                            Hope you like it''')
                song = random.choice(videos)
                os.startfile(song)
            except Exception as e:
                   print(str(e))
        
        else:
            try:
                # wolframalpha
                client = wolframalpha.Client(app_id)
                res = client.query(put)
                ans = next(res.results).text
                print(ans)
                speak.Speak(ans)
            except:
                # wikipedia
                put = put.split()
                put = ' '.join(put[2:])
                # print(put)
                print(wikipedia.summary(put))
                speak.Speak('a runthrough through wikipedia database provides about  ' + put + 'that')


# Trigger GUI
if __name__ == "__main__":
    app = wx.App(True)
    frame = MyFrame()
    app.MainLoop()
