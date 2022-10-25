import os
import tkinter

import pyttsx3
import speech_recognition as sr
import wolframalpha
from tkinter import *
import tkinter as tk
from time import sleep
r = Tk()

root = Tk()
root.title("Desktop Assistant")
root.geometry("450x500")
root.maxsize()
canvas1 = tk.Canvas(root, width=500, height=500)
canvas1.pack()
label1 = tk.Label(root, text=' Desktop Assistant ', font=("Helvetica", 25), fg="orange")
canvas1.create_window(225, 50, window=label1)
import logging    # first of all import the module

logging.basicConfig(filename='../Backend/std.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
logging.warning('This message will get logged on to a file')
logger=logging.getLogger()
logger.setLevel(logging.DEBUG)
def help_toplevel():
    top2 = Toplevel()
    top2.title("Toplevel2")
    top2.geometry("800x430")

def main_voice():
    assistant_speaks("What can i do for you?")
    text = get_audio()
    text.lower()
    if "exit" in str(text) or "bye" in str(text) or "sleep" in str(text):
        assistant_speaks("Ok bye " + '.')
    if text == "note down":
       process_text(text)
    if "open Google" in str(text) or "open" in str(text) or "open Word" in str(text) or "open exel" in str(text) or "open power point" in str(text)  :
        process_text(text)
    if text == "can you answer a question":
        process_text(text)
    if text == "save clipboard to notepad":
        process_text(text)
    if text == "shutdown":
        process_text(text)
    if text == "thank you":
        assistant_speaks("Your welcome")
def save():
    global gv_language
    global gv_phrase_time
    language = (var2.get())
    phrase_time = (var1.get())
    print(var1.get())
    print(var2.get())
    gv_phrase_time = phrase_time
    gv_language = language
#    logger.info("The phrase time got set into"+gv_phrase_time)
#    logger.info("The language got set into" + gv_language)

def settings_toplevel():
    global var1
    global var2

    var1 = tk.StringVar()
    var2 = tk.StringVar()
    top1 = Toplevel(root)
    top1.title("Toplevel1")
    top1.geometry("500x450")
    settings_lable = tk.Label(top1, text='Settings', font=("Helvetica", 25), fg="orange").place(x=205, y=50)
    phrase_time_lable = tk.Label(top1, text=' Phrase Time: ').place(x=100, y=140)
    phrase_textbox = tk.Entry(top1, textvariable=var1).place(x=270, y=140)  # create 3nd entry box
    language_textbox = tk.Entry(top1, textvariable=var2).place(x=270, y=180)
    Language_lable = tk.Label(top1, text=' Language: ').place(x=100, y=180)
    settings_buttton_exit = tk.Button(top1, text='EXIT', command=top1.destroy, bg='orange').place(x=235, y=225)
    save_button = tk.Button(top1, text="SAVE", fg="orange", command=save).place(x=235, y=250)
    point1 = tk.Label(top1, text='For English en-IN').place(x=100, y=245)
    point1 = tk.Label(top1, text='For Hindi hi-IN').place(x=100, y=275)
    point1 = tk.Label(top1, text= 'For Telugu te-IN').place(x=100, y=300)
    point1 = tk.Label(top1, text= 'For Tamil ta-IN'  ).place(x=100, y=325)
    # settings_lable.pack()
    # phrase_time_lable.pack()
    # phrase_textbox.pack()
    # Language_lable.pack()
    # settings_buttton_exit.pack()
    # language_textbox.pack()
    # save_button.pack()
    top1.mainloop()
    gv_phrase_time = int(var1.get())
    gv_language = str(var2.get())



num = 1
def shutdown():
    logger.info("Device was kept into sleep")
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
def assistant_speaks(output):
    print("PerSon : ", output)
    SpeakText(output)
def open_chrome():
    os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"')
    assistant_speaks("Chrome Opened")
    logger.info("Chrome Opened using button")

def open_word():
    os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk"')
    assistant_speaks("Word Opened")
    logger.info("Word Opened using button")
def open_excel():
    os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")
    assistant_speaks("Execel Opened")
    logger.info("Excel Opened using button")
def open_ppt():
    os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk")
    assistant_speaks("Power Point Opened")
    logger.info("Power Point Opened using button")
def question():
    assistant_speaks('Yes, What is the question')
    question = get_audio()
    if question == "what is the time":
        app_id = "8R34XL-W64U38LAW4"
        client = wolframalpha.Client(app_id)
        res = client.query(question)
        answer = next(res.results).text
        assistant_speaks(answer)
        logger.info("Time asked Output:"+answer)

    if question == "what is the weather":
        app_id = "8R34XL-W64U38LAW4"
        client = wolframalpha.Client(app_id)
        res = client.query(question)
        answer = next(res.results).text
        assistant_speaks(answer + "In your area")
        logger.info("Weather asked Output:"+answer)
    else:
        app_id = "8R34XL-W64U38LAW4"
        client = wolframalpha.Client(app_id)
        res = client.query(question)
        answer = next(res.results).text
        assistant_speaks(answer)
        logger.info("The question is:"+question)
        logger.info("Question Asked Output:"+ answer)
def notepad():
    speak = ('Notepad function enabled. Phrase time for 2 minuits')
    logger.info(speak)
    assistant_speaks(speak)
    print(speak)
    assistant_speaks("File name")
    file_name = get_audio()
    logger.info("file name as:"+file_name)
    file = open(file_name, 'w')
    assistant_speaks("File name saved")
    assistant_speaks("Text recording")
    text = get_audio_for_notepad()
    file.write(text)
    file.close()
    assistant_speaks("Text saved")
    logger.info("Text saved into the file:"+text)
def clipboard():
    while not r.selection_get(selection="CLIPBOARD"):
        sleep(0.1)
    result = r.selection_get(selection="CLIPBOARD")
#    r.clipboard_clear()
#    r.option_clear()
    #r.destroy()

    assistant_speaks("The file is save in a notepad with the file name 'Clipboard'")
    file = open("../Backend/Clipboard", 'w')
    text = result
    file.write(result)
    file.close()
    assistant_speaks("Text saved")
def read_clipboard():
    while not r.selection_get(selection="CLIPBOARD"):
        sleep(0.1)
    result = r.selection_get(selection="CLIPBOARD")
    assistant_speaks(result)

def process_text(input):
    try:
        if "open Google" in input:
            os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"')
            assistant_speaks("Chrome Opened")
            logger.info("Chrome opened using voice command")
        if "open Word" in input:
            os.startfile('"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk"')
            assistant_speaks("Word Opened")
            logger.info("Word opened using voice command")
        if "open Excel" in input:
            os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")
            assistant_speaks("Execel Opened")
            logger.info("Excel opened using voice command")
        if "open PowerPoint" in input:
            os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk")
            assistant_speaks("Power Point Opened")
            logger.info("Power Point opened using voice command")

        if "who made you" in input or "by whom are you designed " in input:
            speak = "I have been designed by B.Siddharth."
            logger.info("question asked who made you or by whom are you designed")
            #assistant_speaks(speak)
            print(speak)
            return

        if input == "can you answer a question":
            assistant_speaks('Yes, What is the question')
            question = get_audio()
            if question == "what is the time":
                app_id = "8R34XL-W64U38LAW4"
                client = wolframalpha.Client(app_id)
                res = client.query(question)
                answer = next(res.results).text
                assistant_speaks(answer)
                logger.info("Time asked Output:"+answer)


            if question == "what is the weather":
                app_id = "8R34XL-W64U38LAW4"
                client = wolframalpha.Client(app_id)
                res = client.query(question)
                answer = next(res.results).text
                assistant_speaks(answer + "In your area")
                logger.info("Weather asked Output:" + answer)
            else:
                app_id = "8R34XL-W64U38LAW4"
                client = wolframalpha.Client(app_id)
                res = client.query(question)
                answer = next(res.results).text
                assistant_speaks(answer)
                logger.info("The question is:" + question)
                logger.info("Question Asked Output:" + answer)
        if input == "note down":
            speak = ('Notepad function enabled. Phrase time for 2 minuits')
            logger.info(speak)
            assistant_speaks(speak)
            print(speak)
            assistant_speaks("File name")
            file_name = get_audio()
            logger.info("file name as:" + file_name)
            file = open(file_name, 'w')
            assistant_speaks("File name saved")
            assistant_speaks("Text recording")
            text = get_audio_for_notepad()
            file.write(text)
            file.close()
            assistant_speaks("Text saved")
            logger.info("Text saved into the file:" + text)
        if input == "shutdown":
            logger.info("Device was kept into sleep")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        if input == "save clipboard to notepad":
            clipboard()
    except:
        print('except')
        logger.warning("none of the requests matched")


def SpeakText(command):
    # Initialize the engine
    engine = pyttsx3.init()
    engine.say(command)
    engine.runAndWait()

def get_audio():
    rObject = sr.Recognizer()
    audio = ''
    while(1):
        with sr.Microphone() as source2:
            audio = rObject.listen(source2, phrase_time_limit=gv_phrase_time)
            try:
                text = rObject.recognize_google(audio, language=gv_language)
                print("You : ", text)
                return text
            except:
                assistant_speaks("Could not understand your audio, Please try again !")
                logger.warning("Could not understand your audio")
                return 0
def get_audio_for_notepad():
    rObject = sr.Recognizer()
    audio = ''
    while(1):
        with sr.Microphone() as source2:
            audio = rObject.listen(source2, phrase_time_limit=120)
            try:
                text = rObject.recognize_google(audio, language='en-US')
                print("You : ", text)
                return text
            except:
                assistant_speaks("Could not understand your audio, Please try again !")
                return 0
NOTEPAD = tk.Button(root,command = notepad,text = "Notepad", bg = "orange")
canvas1.create_window(100, 100, window=NOTEPAD)  # button to call the 'values' command above
HELP = tk.Button(root, text='HELP', command=help_toplevel, bg='orange')  # button to call the 'values' command above
canvas1.create_window(225, 100, window=HELP)
chrome = tk.Button(root,text = "Open Chrome",command=open_chrome, bg='orange')
canvas1.create_window(300, 100, window=chrome)
word = tk.Button(root,text = "Open Word",command=open_word, bg='orange')
canvas1.create_window(100, 200, window=word)
excel = tk.Button(root,text = "Open Excel",command=open_excel, bg='orange')
canvas1.create_window(225, 200, window=excel)
question_123 = tk.Button(root,text = "Anewer a Question",command=question, bg='orange')
canvas1.create_window(350, 200, window=question_123)
SETTINGS = tk.Button(root, text='SETTINGS',command = settings_toplevel,bg='orange')  # button to call the 'values' command above
canvas1.create_window(100, 300,window=SETTINGS)
voice_button = tk.Button(root, text='🎙️Voice' ,command =main_voice ,bg='orange')
canvas1.create_window(225, 300,window=voice_button)
shutdown = tk.Button(root, text='Shutdown',command = shutdown,bg='orange')  # button to call the 'values' command above
canvas1.create_window(350, 300,window=shutdown)
clip = tk.Button(root, text='Save clipboard to notepad',command = clipboard,bg='orange')  # button to call the 'values' command above
canvas1.create_window(100, 400,window=clip)
clip = tk.Button(root, text='Read clipboard',command = read_clipboard,bg='orange')  # button to call the 'values' command above
canvas1.create_window(255, 400,window=clip)
root.mainloop()











