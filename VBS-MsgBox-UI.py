#Import the required Libraries
from cgitb import text
from distutils.cmd import Command
from msilib.schema import RadioButton
from operator import index
import string
from tkinter import *
from tkinter import ttk
import os
from tkinter import font

#apastrophe :)
Apostrophe = '"'

#Die verschiedenen aussuch möglichkeiten (Page1)
variations = ["None", "Critical", "Question", "Exclamation", "Information", "HelpButton"]


#Fenster definieren
win= Tk()
win.title(" VBS-MsgBox-UI ")
win.iconbitmap("")

#Grösse des fensters
win.geometry("500x600")

#notebook definiren (ka was das ist lol xD)
my_notebook = ttk.Notebook(win)
my_notebook.pack(pady=15)

#verschiedene tabs
frame1 = Frame(my_notebook)
frame1.pack()
frame2 = Frame(my_notebook)
frame2.pack()
frame3 = Frame(my_notebook)
frame3.pack()

#hinzufügen der tabs
my_notebook.add(frame1, text="Msg Box")
my_notebook.add(frame2, text="Text to speech")
my_notebook.add(frame3, text="Taskkill")

#-----------------------------------------------------------------


#füllvorgang
def FillFile():
   #alle variabeln müssen aufgerufen werden (page 1)
   global entry1, entry2, entry3, var1

   #die variabeln müssen jezt exrahiert werden (page 1)
   string = entry1.get()
   string2 = entry2.get()
   string3 = entry3.get()
   stringy = var1.get()

   #die verschiedenen möglichkeiten die man aussuchen kann (page 1)
   if (stringy == 0):
        vas = "vbOKOnly"
   elif (stringy == 1):
        vas = "vbCritical"
   elif (stringy == 2):
        vas = "vbQuestion"
   elif (stringy == 3):
        vas = "vbExclamation"
   elif (stringy == 4):
        vas = "vbInformation"
   elif (stringy == 5):
        vas = "vbMsgBoxHelpButton"


   #erstellen der datei (page 1)
   FileVBS = open(string + ".vbs", "w+")
   FileVBS.write("MsgBox " + Apostrophe + string3 + '"' + ', ' + vas + ", " + Apostrophe + string2 + Apostrophe)
   FileVBS.close()


   #um den titel zur filename zu verändern (page 1)
   label.configure(text="Filename= " + string)

   #um die erstellte datei zu starten (page 1)
   os.startfile(string + ".vbs")

#funktion zum löschen der erstellten datei (page 1)
def DeleteFile():
    global entry1
    string = entry1.get()
    os.remove(string + ".vbs")
    label.configure(text="VBS MsgBox UI")

#Der titel (page 1)
label=Label(frame1, text="VBS MsgBox UI", font=("Impact 30 bold"))
label.pack()

#----- Lable und Entry sektion (page 1)
entryLabel1= Label(frame1, text="File Name")
entryLabel1.pack()
entry1= Entry(frame1, width= 40)
entry1.pack()

entryLabel2= Label(frame1, text="Title")
entryLabel2.pack()
entry2= Entry(frame1, width= 40)
entry2.pack()

entryLabel3= Label(frame1, text="Text")
entryLabel3.pack()
entry3= Entry(frame1, width= 40,)
entry3.pack()

CheckBoxLabel = Label(frame1, text="Msg Box Type")
CheckBoxLabel.pack()
#----- Lable und Entry sektion fertig (page 1)

#Die variable welche speichert wass man unten ausgewählt hat (page 1)
var1 = IntVar()

#Die runden Buttons zum auswählen (page 1)
for index in range(len(variations)):
    checkbox1 = Radiobutton(frame1, text=variations[index], variable=var1, value=index) 
    checkbox1.pack()


#Die Zwei knöpfe welche unten sthehen (page 1)
ttk.Button(frame1, text= "Okay", width= 20, command= FillFile).pack(pady=20)
ttk.Button(frame1, text= "Delete File", width= 20, command= DeleteFile).pack(pady=20)


#-----------------------------------------------------------------

def FillFile2():
     #alle variabeln müssen aufgerufen werden (page 2)
   global entryP1, entryP2, entryP3, entryP4, Spe, tone

   #die variabeln müssen jezt exrahiert werden (page 2)
   stringy = entryP1.get()
   stringy2 = entryP2.get()
   stringy3 = entryP3.get()
   stringy4 = entryP4.get()

   Spe = Tk.DubleVar()
   tone = Tk.DubleVar()



   print(Spe)
   print(tone)

     #erstellen der datei (page 2)
   FileVBS2 = open(stringy + ".vbs", "w+")
   FileVBS2.write('''
Set Stimme = CreateObject("SAPI.spVoice")
Set Stimme.Voice = Stimme.GetVoices.Item(0)
''' + "Stimme.Rate = " + stringy3 + '''
''' + "Stimme.Volume = " + stringy4 + '''
''' + "Stimme.Speak " '"' + stringy2 + '"')
   FileVBS2.close()
   os.startfile(stringy + ".vbs")

def DeleteFile2():
    global entryP1
    stringy = entryP1.get()
    os.remove(stringy + ".vbs")

#Der titel (page 2)
label=Label(frame2, text="VBS Speech to Text UI", font=("Impact 30 bold"))
label.pack()

#----- Lable und Entry sektion (page 2)
entryLabelP1= Label(frame2, text="File Name")
entryLabelP1.pack()
entryP1= Entry(frame2, width= 40)
entryP1.pack()

entryLabelP2= Label(frame2, text="Output")
entryLabelP2.pack()
entryP2= Entry(frame2, width= 40)
entryP2.pack()


entryLabelP3= Label(frame2, text="Speed (-10 to 10)")
entryLabelP3.pack()
entryP3= Entry(frame2, width= 40)
entryP3.insert(END, "1")
entryP3.pack()

entryLabelP4= Label(frame2, text="Volume (0 to 100)")
entryLabelP4.pack()
entryP4= Entry(frame2, width= 40)
entryP4.insert(END, "100")
entryP4.pack()

#----- Lable und Entry sektion (page 2)

#Die Zwei knöpfe welche unten sthehen (page 2)
ttk.Button(frame2, text= "Okay", width= 20, command= FillFile2).pack(pady=20)
ttk.Button(frame2, text= "Delete File", width= 20, command= DeleteFile2).pack(pady=20)


#-----------------------------------------------------------------


def FillFile3():
     

     TaskkillFileName = entryA1.get()
     TaskkillApplication = entryA2.get()

     FileVBS2 = open(TaskkillFileName + ".bat", "w+")
     FileVBS2.write("TASKKILL /F /IM " + TaskkillApplication)
     FileVBS2.close()
     os.startfile(TaskkillFileName + ".bat")

def DeleteFile3():
     global entryA1
     stringy23 = entryA1.get()
     os.remove(stringy23 + ".bat")


     




#Der titel (page 3)
label = Label(frame3, text="Taskkill", font=("Impact 30 bold"))
label.pack()

entryLabelA1= Label(frame3, text="File Name")
entryLabelA1.pack()
entryA1= Entry(frame3, width= 40)
entryA1.pack()

entryLabelA2= Label(frame3, text="Application (please add file extension at the end)")
entryLabelA2.pack()
entryA2= Entry(frame3, width= 40)
entryA2.pack()

ttk.Button(frame3, text= "Okay", width= 20, command= FillFile3).pack(pady=20)
ttk.Button(frame3, text= "Delete File", width= 20, command= DeleteFile3).pack(pady=20)


#Meine Credits :)
credits = Label(win, text="Made by PianoNic")
credits.pack()


win.mainloop()
