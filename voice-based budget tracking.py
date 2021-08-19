import pyttsx3
import datetime
import speech_recognition as sr
import smtplib
import webbrowser as wb
import os
import pyautogui
import psutil
engine = pyttsx3.init()
import xlwt
from xlwt import Workbook

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0, 0, 'Date')
sheet1.write(0, 1, 'Time')
sheet1.write(0, 2, 'Money Added')
sheet1.write(0, 3, 'Money Spent')
sheet1.write(0, 4, 'Reason')
sheet1.write(0, 5, 'Total')
def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def takeCommand():
    r = sr.Recognizer()
    r.energy_threshold = 300
    with sr.Microphone() as source:
        print("Listening....")
        r.adjust_for_ambient_noise(source, duration=1)
        # audio=r.listen(source)
        audio = r.listen(source=source, timeout=7, phrase_time_limit=6)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language="en-in")
        print(query)
    except Exception as e:
        print(e)
        speak("Say that again please")
        return "None"
    return query


if __name__ == "__main__":
    total = 0
    net_spent=0
    i=1
    k=1
    while True:
        query = takeCommand().lower()
        res = [int(i) for i in query.split() if i.isdigit()]
        if res:
            if ("got" in query) and (str(res[0]).isdigit()):
                total += res[0]
                print("Money added: Rs "+str(res[0])+" total: Rs "+str(total))
                print(query)
                sheet1.write(k, 0,datetime.datetime.now().strftime("%x"))
                sheet1.write(k, 1,datetime.datetime.now().strftime("%X"))
                sheet1.write(k, 5, total)
                sheet1.write(k, 4, query)
                sheet1.write(k, 3, 0)
                sheet1.write(k, 2, res[0])
                i+=1
                k+=1
                res = []
                continue
                
            elif ("salary" in query) and (str(res[0]).isdigit()):
                total += res[0]
                print("Money added: Rs "+str(res[0])+" total: Rs "+str(total))
                print(query)
                sheet1.write(k, 0,datetime.datetime.now().strftime("%x"))
                sheet1.write(k, 1,datetime.datetime.now().strftime("%X"))
                sheet1.write(k, 5, total)
                sheet1.write(k, 4, query)
                sheet1.write(k, 3, 0)
                sheet1.write(k, 2, res[0])
                i+=1
                k+=1
                res = []    
                
            elif ("income" in query) and (str(res[0]).isdigit()):
                total += res[0]
                print("Money added: Rs "+str(res[0])+" total: Rs "+str(total))
                print(query)
                sheet1.write(k, 0,datetime.datetime.now().strftime("%x"))
                sheet1.write(k, 1,datetime.datetime.now().strftime("%X"))
                sheet1.write(k, 5, total)
                sheet1.write(k, 4, query)
                sheet1.write(k, 3, 0)
                sheet1.write(k, 2, res[0])
                i+=1
                k+=1
                res = [] 

            elif ("spen" in query) and (str(res[0]).isdigit()):
                total -= res[0]
                net_spent+=res[0]
                print("spent: Rs "+str(res[0])+" total: Rs "+str(total))
                print(query)
                sheet1.write(i, 0,datetime.datetime.now().strftime("%x"))
                sheet1.write(i, 1,datetime.datetime.now().strftime("%X"))
                sheet1.write(i, 5, total)
                sheet1.write(i, 4, query)
                sheet1.write(i, 3, res[0])
                sheet1.write(i, 2, 0)
                i+=1
                k+=1
                res = []
                
            elif ("expense" in query) and (str(res[0]).isdigit()):
                total -= res[0]
                net_spent+=res[0]
                print("spent: Rs "+str(res[0])+" total: Rs "+str(total))
                print(query)
                sheet1.write(i, 0,datetime.datetime.now().strftime("%x"))
                sheet1.write(i, 1,datetime.datetime.now().strftime("%X"))
                sheet1.write(i, 5, total)
                sheet1.write(i, 4, query)
                sheet1.write(i, 3, res[0])
                sheet1.write(i, 2, 0)
                i+=1
                k+=1
                res = [] 
                
            elif ("paid" in query) and (str(res[0]).isdigit()):
                total -= res[0]
                net_spent+=res[0]
                print("spent: Rs "+str(res[0])+" total: Rs "+str(total))
                print(query)
                sheet1.write(i, 0,datetime.datetime.now().strftime("%x"))
                sheet1.write(i, 1,datetime.datetime.now().strftime("%X"))
                sheet1.write(i, 5, total)
                sheet1.write(i, 4, query)
                sheet1.write(i, 3, res[0])
                sheet1.write(i, 2, 0)
                i+=1
                k+=1
                res = []    
                
            elif ("gave" in query) and (str(res[0]).isdigit()):
                total -= res[0]
                net_spent+=res[0]
                print("spent: Rs "+str(res[0])+" total: Rs "+str(total))
                print(query)
                sheet1.write(i, 0,datetime.datetime.now().strftime("%x"))
                sheet1.write(i, 1,datetime.datetime.now().strftime("%X"))
                sheet1.write(i, 5, total)
                sheet1.write(i, 4, query)
                sheet1.write(i, 3, res[0])
                sheet1.write(i, 2, 0)
                i+=1
                k+=1
                res = [] 

            else:
                pass

        elif "enough" in query:
            print("total amount: "+str(total)+" net_spent_amount: "+str(net_spent))
            wb.save('budget1.xls')
            quit()


