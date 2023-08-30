# SATYAM SINGH
# Please ensure the installation of all the provided modules by executing the following command:"pip install module_name"

import speech_recognition as sr
import pyttsx3
import datetime
import os
import webbrowser
import time
import yagmail
import psutil
from fpdf import FPDF
import wikipedia
import datetime
import pyautogui
from mss import mss
from speech_recognition import AudioFile
import nltk
from nltk.sentiment import SentimentIntensityAnalyzer
import cv2
import qrcode
import pywhatkit
import sqlite3

# Please make sure to update the email address before sending the email. Additionally, set your email account password in an environmental variable named "EMAIL_PASS." If you've chosen a different variable name, kindly adjust the password retrieval line accordingly.
sender='satyamsingh20942@gmail.com'
password= os.environ.get('emaipass')

engine= pyttsx3.init('sapi5')
# You can change voice speed of jarvis by changing the numerical value of "new_voicerate".
new_voicerate= 125
engine.setProperty('rate',new_voicerate)
voices= engine.getProperty('voices')
# Here now you can change the voices in place of voices[i] you can input i according to the voices that you have saved in registry.
# Attached are the voice files labeled as 'voice 1', 'voice 2', 'voice 3', 'voice 4', and 'voice 5'.
engine.setProperty('voice',voices[5].id)

def speak(audio):
    '''
    This function helps the program to communicate with the user using voice 
    '''
    engine.say(audio)
    engine.runAndWait()

def greet(): 
    '''
    This Function is used to greet the user at first when used 
    '''
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<= 12:
        speak("good morning sir!!")
    elif hour>=12 and hour<= 16:
        speak("good afternoon sir!")
    elif hour>=16 and hour<= 18:
        speak("good evening sir!!")
    elif hour>=19 and hour<=24:
        speak("good night sir!!")
        
def commands():
    '''
    The command function will help Jarvis to listen the user and then decode with the help of the google recognizer.
    '''
    reco= sr.Recognizer()
    # Adjust the 'device_index' parameter to align with the designated microphone number being utilized for input. 
    # For assistance in determining the microphone number, kindly consult the code snippet provided within the file titled 'Locating the Device Microphone ID'.
     
    with sr.Microphone(device_index=2) as source:
        print("listening...")
        reco.pause_threshold=1
        reco.energy_threshold=300
        audio= reco.listen(source)     
    try:
        print("recognising...")
        query= reco.recognize_google(audio, language='en-in')

    except Exception as e:
        speak("say that again please ")
        # commands()
        return "None"
    return query  
        
def quit():
    '''
    This function facilitates program termination.
    '''
    speak('quiting sir! ')
    exit()

# SHUTDOWN THE PC
def shutdowninst():
    '''
    This function will initiate an immediate shutdown of your computer without additional prompt or delay.
    '''
    os.system('shutdown /s /t 0')
def shutdown():
    '''
    This function will initiate a 30-second timer before the shutdown process begins. This interval allows for the appropriate management of unsaved work and any necessary preparations.
    '''
    os.system('shutdown /s /t 30')
    
# RESTART THE PC
def restartinst():
    '''
    This function will initiate an immediate restart of your computer without additional prompt or delay.
    '''
    os.system('shutdown /r /t 0')  
def restart():
    '''
    This function will initiate a 30-second timer before the restart process begins. This interval allows for the appropriate management of unsaved work and any necessary preparations.
    '''
    os.system('shutdown /r /t 10')
    
def app(appname):
    '''
    This function facilitates the initiation of the application. However, prior to utilizing this function, please adjust the file location accordingly to ensure its seamless functionality. You have the flexibility to include additional apps based on your requirements.
    '''
    if 'telegram' in appname:
        os.startfile('"C:\\Users\\satya\\AppData\\Roaming\\Telegram Desktop\\Telegram.exe"')
    if 'visual' in appname:
        os.startfile('"C:\\Users\\satya\\AppData\Local\\Programs\\Microsoft VS Code\\Code.exe"')
    if 'chrome' in appname:
        os.startfile('"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"')
    if 'microsoftedge' in appname:
        os.startfile('"C:\\Program Files (x86)\\Microsoft\\Edge\Application\\msedge.exe"')
    if 'word' in appname:
        os.startfile('"C:\\Program Files\\Microsoft Office\\root\Office16\\WINWORD.EXE"')
    if 'excel' in appname:
        os.startfile('"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"')
    
def website(web):
    speak("Write domain name")
    domain= input("Kindly furnish the precise domain extension of the website (e.g., com, in, etc.) for our records. Please omit the period before the domain extension.:- ")
    try:
        webbrowser.open_new_tab(f'www.{web}.{domain}')
    except Exception as e:
        speak("sorry didn't found the website.")            
                  
def about():
    '''
    This function facilitates the retrieval of device specifications, encompassing essential metrics like "CPU Count","CPU Frequency","CPU Stats","CPU Percentage","Battery Health","Battery Percentage","Total Utilized Storage (in GB)", and "Available Free Space on your device".    
    '''
    battery= psutil.sensors_battery()
    windows= psutil.WINDOWS
    Mac=psutil.MACOS
    linux= psutil.LINUX
    C_cunt=psutil.cpu_count()
    C_feq= psutil.cpu_freq()
    C_stat= psutil.cpu_stats()
    C_per=psutil.cpu_percent()
    battery_per= battery.percent
    battery_heal= battery.health
    disk_part= psutil.disk_partitions()
    path=psutil.disk_usage('/')
    used= path.used/(1024**3)
    free= path.free/(1024**3)
    full= float(used)+ float(free)
    if windows== True:
        speak('You are working on Windows')
    elif Mac== True:
        speak('You are working on Mac')
    elif linux== True:
        speak('You are working on Linux')
    speak(f'Sure sir, the cpu count is {psutil.cpu_count},the frequency of the cpu is {psutil.cpu_freq},the statistics of the cpu is {psutil.cpu_stats}')
    speak(f'cpu percentage is {psutil.cpu_percent},battery percantage is {battery.percent},battery health is {battery.health}')  
    speak(f'The total used space on the device is {used}Gb, and the total free space available is {free}Gb, out of total {full} Gb')  
    speak('Here is the stats printed  ')
    print(f'CPU Count= {C_cunt} \n CPU Frequency= {C_feq} \n CPU Stats= {C_stat} \n CPU percentage= {C_per} \n Battery percentage= {battery_per} \n battery health= {battery_heal} \n Disk partitions={disk_part} \n The used space of the disk is {used} in GB \n  The free space of the disk is {free} in GB')   
    
def pdfgen():
    '''
    This function is used for generating the pdf file.
    '''
    speak('would you want portrait(P) or the  landscape(L)')
    orientati= commands()
    if 'portrait' or 'P' or 'p' in orientati:
        oriet='P'
    elif 'landscape' or 'L' or 'l' in orientati:
        oriet='L'
    pdf= FPDF(orientation=oriet,unit='pt', format='A4')
    pdf.add_page()
   
    pdf.set_font(family='courier',style='B',size=36)
    speak('heading of the pdf')
    heading=input('Heading:- ')
    description_wiki=wikipedia.summary(heading,sentences=15)
    
    pdf.cell(w=0,h=50,txt=heading,align='C',ln=1)
    pdf.set_font(family='courier',style='B',size=18)
    pdf.cell(w=0,h=20,txt='description',ln=1)
    pdf.set_font(family='Times',size=14)
    pdf.multi_cell(w=0,h=15,txt=description_wiki)
    pdf.add_page()
    speak('would you like to add an image if yes press y.')
    time.sleep(0.5)
    speak('and add this to the folder location and name as serial numerical number with jpg extension')
    addimgornot= input('y/n ')
    if 'y' or 'Y' in addimgornot:
        speak('how many images you want to add in your pdf.')
        time.sleep(0.5)
        speak(' do not exceed it to 7')
    
        numberofimages=int(input())
        for i in range (numberofimages):
            pdf.image(f'{i}.jpg',x=240 ,w=80,h=80)

    pdf.output('output.pdf')
            
# ------------------------------------For music by spotify---------------------------------------
def playmusic(song_name):
    '''
    This function will help in playing the song on spotify.
    '''
    i=0
    os.system("spotify")
    time.sleep(5)
    pyautogui.hotkey('ctrl','l')
    time.sleep(5)
    pyautogui.write(song_name)
    time.sleep(5)
    pyautogui.press('Enter')
    time.sleep(1)
    pyautogui.press('Tab')
    for i in range(2):
        time.sleep(1)
        pyautogui.press('Enter')
def artist(artist_name):
    '''
    This function serves the purpose of playing music solely based on the artist's name. However, it is designed to function seamlessly exclusively with playlists curated by Spotify or those originating from verified artists. Kindly note that during its current phase of development, optimal functionality is not assured for playlists sourced from other origins.
    '''
    os.system('spotify')
    time.sleep(5)
    pyautogui.hotkey('ctrl','l')
    time.sleep(2)
    pyautogui.write(artist_name)
    pyautogui.press('Enter')
    time.sleep(2)
    pyautogui.press('Tab')
    time.sleep(2)
    pyautogui.press('Enter')
    time.sleep(1)
    pyautogui.press('Enter')
# -------------------------coming soon for artist------------------------------------------
def spotify_playlist():
    '''
    This function is intended for the playback of playlists that are exclusively crafted within the Spotify platform.
    '''
    speak('make sure that the playlist is made by spotify')
    playlist=input("enter the name of the playlist")
    os.system('spotify')
    time.sleep(1)
    pyautogui.hotkey('ctrl','l')
    time.sleep(1)
    pyautogui.write(playlist)
    time.sleep(1)
    pyautogui.press('Enter')
    time.sleep(0.7)
    pyautogui.press('Tab')
    time.sleep(0.7)
    pyautogui.press('Enter')
    time.sleep(0.7)
    j=2
    for i in range(j):
        pyautogui.press('Tab')
    pyautogui.press('Enter')
    time.sleep(0.7)
    
def take_screenshot(N,T):
    '''
    This function facilitates the capture of a specified number of screenshots at defined intervals.
    '''
    for i in range(N):
        time.sleep(T)
        with mss() as screen:
            screen.shot(output=f'screenshot{i}.jpg')

def contentusingattachment(list,content):
    '''
    This function proves instrumental in the context of incorporating attachments within email correspondence.
    '''
    i=int(input("number of attachment- "))
    speak("enter the file name with extension")
    for k in range(i):
        file= input(f"file {k} - ")
        list.append(file)
    list.append(content)
    return list
def playsongyoutube():
    '''
    This function enables the playback of a song on YouTube in fullscreen mode.
    '''
    speak("which song you would like to  listen")
    songname=input("Enter the name of the song:- ")
    # i=0
    # webbrowser.open("www.youtube.com")
    # time.sleep(4)
    # pyautogui.press('/')
    # time.sleep(2)
    # pyautogui.write(f"play {songname} song")
    # time.sleep(1)
    # pyautogui.press("Enter")
    # for i in range(9):
    #     pyautogui.press("Tab")
    #     time.sleep(0.7)
    # pyautogui.press("Enter")
    # time.sleep(0.5)
    pywhatkit.playonyt(songname)
    time.sleep(3)
    fullscreen()
    
def fullscreen():
    pyautogui.press('f')
def webcameraqrcodedetector():
    '''
    This function acquires video input from a webcam, performs QR code detection, and subsequently launches the URLs encoded within the QR codes in a web browser.
    '''
    video = cv2.VideoCapture(0)
    success,frame= video.read()
    detect = cv2.QRCodeDetector()
    # 萨蒂亚姆辛格
    while True:
        success, frame = video.read()
        if not success:
            break
        url, coords, pixel = detect.detectAndDecode(frame)
        if url:
            webbrowser.open(url)
            break
        
        cv2.imshow('frame', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    video.release()
    cv2.destroyAllWindows()          

def qrcodedetector(image):
    '''
    This function necessitates an image file containing the QR code, which, when provided, initiates the automatic launch of the associated website in a web browser. The corresponding URL will be displayed within the terminal to facilitate further proceedings.
    '''
    image= cv2.imread(image)
    detector=cv2.QRCodeDetector(image)
    url= detector.detectAndDecode(detector)[0]
    print(f'This is the url of the qrcode {url}')
    speak("would you want to open the url in browser ")
    openbrowser= commands().lower()
    if 'yes' in openbrowser:
        webbrowser.open(url)

def wikiandfile(topic):
    '''
    This function facilitates the extraction of information from Wikipedia, offering the flexibility to convert the acquired information into various formats such as text or PDF files, based on your preferences and requirements.
    '''
    time.sleep(0.4)
    speak('do you want to create a pdf or a txt file')
    command= commands().lower()
    if 'text' in command:
        with open(f'{topic}.txt','w',unicode='utf-8') as f:
            filecontent= wikipedia.page(topic).content
            f.write(filecontent)
            f.close()
    else:
        print(wikipedia.page(topic).content)
        
def sqltocsv(file):
    webbrowser.open('https://inloop.github.io/sqlite-viewer/')
    
    
if __name__ == '__main__':
    greet()
    print("today",datetime.datetime.today())
    while True:
        query= commands().lower()
        print(query)
        if 'quit' in query:
            quit()
        elif 'are you there' in query:
            speak("always there for you sir")
            
        # elif 'hlo' or 'hii' in query:
        #     speak("hlo Sir, what i may do for you ")
            
        elif "how are you" in query:
            battery=psutil.sensors_battery()
            percentage= battery.percent
            if percentage>=40 :
                speak("I am fine, what about you sir!!")
            else :
                speak("fine, but the battery is running low. but with you till last sir!!!")
                
        elif 'open' in query :
            ''' openig a website or a application from your laptop '''
            speak("would you want to open a website orr an app sir")
            comm= commands().lower()
            if 'app' in comm:
                speak("Sure sir, which app you wants to open")
                appname= commands().lower()
                app(appname)
            elif 'website' in comm:
                speak('which website you wants to open sir, speak only the website name')
                web= commands().lower()
                website(web)
                
        elif 'wait a minute' in query :
            '''giving some time between the next command'''
            speak('sure sir')
            time.sleep(40)
            # brea= commands().lower()
            # if 'are you there jarvis' in brea:
            #     speak('always there for you sir')
        

        elif 'about' in query:
            about()
            
        elif 'shutdown' in query:
            speak('you want instant shutdown')
            instr= commands().lower()
            if 'yes' in instr:
                shutdowninst()
            else:
                speak("Your system is going to be shutdown in 10 seconds,Make sure you save all your unsaved works")
                shutdown()
                
        elif 'restart' in query :
            speak('you want instant restart')
            instr= commands().lower()
            if 'yes' in instr:
                restartinst()
            else:
                speak("Your system is going to be restart in 10 seconds,Make sure you save all your unsaved work")
                restart()
        
        elif 'send' and ('mail' or 'email') in query :
            '''Used input function here in oredr to not having a single mistake'''
            speak('sure sir to whome you wants to send mail')
            speak('enter the email id of the recipient- ')
            reciever_mail= input("Enter the mail id of the recipient:- ")
            speak('recievers mail id is successfully taken')
            time.sleep(2)
            speak('enter the subject of the mail- ')
            subject= input("Subject:-")
            speak('subject taken successfully')
            time.sleep(2)
            speak("enter the content of the mail sir")
            content= (input('Enter the content of the mail- '))
            speak("would you want to attch some file in mail,if yes then type ")
            list=[]
            attachment=input("want to attach some files if yes type Y/y")
            if attachment=='y' or attachment=='Y':
                contentusingattachment(list,content)
            elif attachment=='n' or attachment=='N':
                list.append(content)
            yag= yagmail.SMTP(user='satyamsingh20942@gmail.com' , password= password)
            yag.send(to= reciever_mail, subject=subject, contents=list)
            speak('Email sent successfully!')
        elif 'generate' and 'pdf' in query:
            pdfgen()
        elif 'play'in query and (('song' in query) or ('music' in query)) :
            speak("sure sir")
            time.sleep(0.5)
            speak("Wold you like to listen the song on youtube or not")
            ifyes= commands().lower()
            if 'yes' in ifyes or 'youtube' in ifyes:
                playsongyoutube()
            else:
                speak("want to listen by artist or any playlist by spotify or a particular song")
                print("want to listen by artist or any playlist by spotify or a particular song")
                command= commands().lower()
                if 'playlist' in command:
                   
                    spotify_playlist()
                elif 'song' in command:
                    song_name=input("Enter the name of the song - ")
                    playmusic(song_name)
        
        elif 'take' and 'screenshot' in query:
            speak("sure sir!, How many screen shots you need")
            numberofss=int(input("Enter the number of screenshots you need:- "))
            time_of_sleep= int(input("time interval you need between your screenshot :- "))
            take_screenshot(numberofss,time_of_sleep)
            
        elif 'mood' in query:
            speak("would you like to import a file or want to speak")
            findingmode= commands().lower()
            recognizer= sr.Recognizer()
            if 'file' in findingmode:
                speak("Enter the name of the file")
                file=input()
                with AudioFile(file) as audio_file:
                    audio= recognizer.record(audio_file)
                text= recognizer.recognize_google(audio)
                nltk.download('vader_lexicon')
                analyzer= SentimentIntensityAnalyzer(text)
                scores= analyzer.polarity_scores()
                if (scores['neg']>scores['pos']):
                    speak("There is negetavity in the speech")
                elif (scores['neg']< scores['pos']):
                    speak("There is positivity in the speech")
                else:
                    speak("the mood of the person is normal")
            else:
                speak("speak your thoughts on the microphone")
                speech= commands().lower()
                nltk.download('vader_lexicon')
                analyzer= SentimentIntensityAnalyzer(speech)
                scores=analyzer.polarity_scores()
                if (scores['neg']>scores['pos']):
                    speak("There is negetavity in the speech")
                elif (scores['neg']< scores['pos']):
                    speak("There is positivity in the speech")
                else:
                    speak("the mood of the person is normal")
                    
        elif ('qrcode' and 'scan' in query) or ('scan' in query):
            speak("import an image ")
            time.sleep(1)
            speak('Are you interested in providing input via a camera or not?')
            camornot=input("Enter 'Y'or 'y' for giving the input through the camera")
            if camornot=='y' or camornot=='Y':
                webcameraqrcodedetector()
            else:
                speak("enter  the image name which you want to scan  ")
                iamge= input('image file name')
                qrcodedetector()

        elif 'wiki' in query or 'wikipedia' in query:
            speak("Sure sir what is the topic ")
            topic= input("Enter the topic of the wiki:- ")
            wikiandfile(topic)
            
        elif 'qrcode' in query and 'create' in query:
            speak('lets create a qrcode')
            time.sleep(0.2)
            speak('enter the url of which you want to create a qrcode')
            url=input("Enter the url for creating the qrcode :- ")
            image= qrcode.make(url)
            image.save('qr.jpg')
        # elif 'sql' in query:
        #     speak('sure! ready to retrive the data from sql file')
        #     time.sleep(1)
        #     speak("Enter the name of the file sir!")
        #     filename= input("Enter the name of the file:- ")
        #     speak("Please confirm if you wish to convert to CSV, or create an Excel file.")
        #     wish= input("for CSV file type 'CSV' for excel type 'EXCEL'.")
        #     wish_lower=wish.lower()
        #     if wish == 'csv':
        #         sqltocsv(filename)
        
            
        else :
            continue
        
        
            
            
            