import cv2
import time
import face_recognition
import os
import pyttsx3
import pandas as pd
import datetime
from tkinter import *
from pathlib import Path
from PIL import ImageTk, Image
import random
import sys
#from datetime import datetime,timedelta

'''
                            Variables for the requirement check
                            ----------------------------------
'''
today = datetime.datetime.now().strftime("%d-%m")
month = datetime.datetime.now().strftime("%B")

pathofimages = "knownimgaes/"
pathofData = "Game/Data/"
pathofMonth = pathofData+str(month)
print(pathofMonth+"/"+today)

unknownimgpath = "unknownimgfolder/"
knownimgpaths = "knownimgaes/"

#import the cascade for face detection
face_cascade = cv2.CascadeClassifier('xmldata/haarcascade_frontalface_default.xml')
'''------------------------------------------------------------------------------------------------'''
def speak(word):
    engine = pyttsx3.init()
    engine.setProperty('rate',135)

    engine.say(word)
    engine.runAndWait()

def createformat():
    Student=[]
    for files in os.listdir(pathofimages):
        studentname = files.replace(".jpg","")
        Student.append(studentname)
        format = {"Name": Student,"Time":""}

    df = pd.DataFrame(format)
    df.transpose()
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(pathofMonth+"/"+today+'.xlsx')
    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

def checkcheckrequirements():
    if not os.path.exists(pathofMonth):
        os.makedirs(pathofMonth)
        speak("created month folder")
        print("created month folder")
    if not os.path.exists(pathofMonth+"/"+today):
        speak("created date file")
        print("created date file")
        f = open(pathofMonth+"/"+today+".xlsx",'w')
        f.close()
        speak("crated formats")
        createformat()
        speak("all requirements satisfied")
        print("created requirements")
    speak("All requirements satisfied.. Moving further")
def clearimgaefolder():
    for files in os.listdir(unknownimgpath):
        os.remove(unknownimgpath+files)
    speak("Clearing images")

def TakeSnapshotAndSave():
    speak("Taking snapshots please wait and stay stable")
    # access the webcam (every webcam has a number, the default is 0)
    cap = cv2.VideoCapture(0)

    num = 0 
    while num<1:
        time.sleep(2)
        speak("Capturing... look in camera please")
        time.sleep(1)
        # Capture frame-by-frame
        ret, frame = cap.read()
        # to detect faces in video
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, 1.3, 5)
        

        for (x,y,w,h) in faces:
            cv2.rectangle(frame,(x,y),(x+w,y+h),(255,0,0),2)
            roi_gray = gray[y:y+h, x:x]
            roi_color = frame[y:y+h, x:x+w]

        x = 0
        y = 20
        text_color = (0,255,0)

       
        num = num+1
        cv2.imwrite('unknownimgfolder/opencv'+str(num)+'.jpg',frame)
        image= cv2.imread('unknownimgfolder/opencv'+str(num)+'.jpg')
        cv2.imshow("Please press any key to continue",image)
        speak("press any key to continue once image is shown")
        cv2.waitKey(0) 

    # When everything done, release the capture
    
    cap.release()
    cv2.destroyAllWindows()
    

def compare():
    try:
        speak("comparing... please wait")
        time.sleep(1)
        for knownfile in os.listdir(knownimgpaths):
            for filename in os.listdir(unknownimgpath): 
                known_image = face_recognition.load_image_file(knownimgpaths+knownfile)
                known_encoding = face_recognition.face_encodings(known_image)[0]
                unknown_image = face_recognition.load_image_file(unknownimgpath+filename)
                unknown_encoding = face_recognition.face_encodings(unknown_image)[0]
                results = face_recognition.compare_faces([known_encoding], unknown_encoding,tolerance= 0.5)
                if results[0] == False:
                    continue
                if results[0] == True:
                    return knownfile
    except IndexError:
        return None

def checkname(name):
    try:
        today = datetime.datetime.now()
        month = today.strftime("%B")
        ndate = today.strftime("%d-%m")
        path = Path("Game/Data/"+str(month)+"/"+str(ndate)+".xlsx")
        print(path)
        if os.path.exists(path):
            df = pd.read_excel(path, index_col=0, engine='openpyxl')
            #print(df)
            Time = df.loc[name][0]
            print(Time)
            if str(Time) == "nan":
                return True
            else:
                return False
    except KeyError:
        return False
def markattendence(name):
    today = datetime.datetime.now().strftime("%d-%m")
    print("Marking your Time ")
    speak("Marking your Time ")
    df = pd.read_excel(pathofMonth+"/"+str(today)+".xlsx",engine='openpyxl')
    currenttimeonly = datetime.datetime.now().strftime("%H:%M:%S")
    yearNow = datetime.datetime.now().strftime("%Y")
    writeInd = []
    for index, item in df.iterrows():
        # print(index, item['Birthday'])
        bday = item['Name']
        if name.lower() == bday.lower() :
            writeInd.append(index)
    for i in writeInd:
        df.loc[i, 'Time'] = str(currenttimeonly)

    df.to_excel(pathofMonth+"/"+str(today)+".xlsx", index=False,engine='openpyxl')  
def facerecognitionsystem():
    global fbox
    clearimgaefolder()
    TakeSnapshotAndSave()
    final = compare()
    if final != None:
        Student = final.replace(".jpg","")
        with open("Game/Student.txt","w")as f:
            f.write(Student)
            f.close()
        print(Student)
        print("Identified you... You are "+ Student)
        speak("Identified you.. You are "+ Student)
        print(checkname(Student))
        if checkname(Student) == False:
            speak("Sorry Your Quota is Completed. You Can't play game today. Comeback tommorrows")
            sys.exit()
            
        else:
            markattendence(Student)
            speak("Succesfully marked your time")
            fbox.insert(END,"Your Name is: "+Student,"Marked your Time: True")
            speak("Starting Game")
            os.startfile("Game\Snake.py")
            sys.exit()
            
    else:
        print("You are not identified")
        speak("You are not Identified are you sure your image is in Database folder")
        speak("Please try again")
        fbox.insert(END,"Your Name is: Not known","Marked your Time: False","please try again later")
        # speak("next student please")
    
    
    
if __name__ == "__main__":
    checkcheckrequirements()
    root = Tk()
    root.title("Game Limitation Program")
    root.config(bg="Yellow")
    root.geometry("655x800")
    frame = Frame(root, width=100, height=150)
    frame.pack(side="bottom")
    frame.place(anchor='center', relx=0.5, rely=0.5)
    image = Image.open("sources/logo2.jpg")
    resize_image = image.resize((200, 200))
    img = ImageTk.PhotoImage(resize_image)
    label = Label(frame, image = img)
    label.pack(side="bottom")
    Heading0 = Label(root,bg="blue",text="JNV KATNI",font="Comicsans 24 bold",fg="Red")
    Heading0.pack(fill="x",padx=2,pady=5)
    Heading1 = Label(root,bg="black",text="Addiction free Game",font="Comicsans 18 bold",fg="Red")
    Heading1.pack(fill="x",padx=2,pady=5)
    Heading2 = Label(root,bg="blue",text="To Start Please click Submit Button ",font="Comicsans 10 bold",fg="Red")
    Heading2.pack(fill="x",padx=2,pady=5)
    

    f1=Frame(root,bg="yellow")
    f1.pack(padx=4)
    Button(f1,bg="red",text="Submit",font="Algerian 10 bold",command=facerecognitionsystem).grid(row=0,column=2,padx=10,pady=10)
    Button(f1,bg="red",text="Show info",font="Algerian 10 bold",command=exit).grid(row=0,column=1,padx=10,pady=10)
    Label(root,bg="yellow",fg="red",text="Your Response",font="Algerian 20 bold").pack()
    final = Frame(root,bg="yellow",borderwidth=5,relief=SUNKEN)
    fbox = Listbox(final,height=3,width=30)
    fbox.pack()
    final.pack(side=TOP,padx=15,pady=15)
    ending = Frame(root,bg="green",borderwidth=5,relief=SUNKEN)
    Label(ending,text="Created By ~ $!ddH@nt Sharma",font="Comicsans 10 bold").pack()
    ending.pack(side=BOTTOM,fill= X)
    root.mainloop()
