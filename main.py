from tkinter import *
from tkinter import ttk
import os
import subprocess

# Constants
VARIATIONS = ["None", "Critical", "Question", "Exclamation", "Information", "HelpButton"]

# Create the main window
win = Tk()
win.title("VBS Toolbox Suite")
win.geometry("500x600")

# Create a notebook for tabbed interface
notebook = ttk.Notebook(win)
notebook.pack(pady=15)

# Create frames for each tab
frame1 = Frame(notebook)
frame2 = Frame(notebook)
frame3 = Frame(notebook)

# Add frames to notebook
notebook.add(frame1, text="Msg Box")
notebook.add(frame2, text="Text to speech")
notebook.add(frame3, text="Taskkill")

# Frames for buttons in each tab
frame1_buttons = Frame(frame1)
frame1_buttons.pack(side=BOTTOM, pady=10)

frame2_buttons = Frame(frame2)
frame2_buttons.pack(side=BOTTOM, pady=10)

frame3_buttons = Frame(frame3)
frame3_buttons.pack(side=BOTTOM, pady=10)

# Functions for Msg Box tab
def fill_file():
    filename = entry_filename.get()
    title = entry_title.get()
    text = entry_text.get()
    box_type = var_box_type.get()

    box_types = {
        0: "vbOKOnly",
        1: "vbCritical",
        2: "vbQuestion",
        3: "vbExclamation",
        4: "vbInformation",
        5: "vbMsgBoxHelpButton"
    }

    vbs_content = f'MsgBox "{text}", {box_types[box_type]}, "{title}"'
    
    with open(f"{filename}.vbs", "w+") as file:
        file.write(vbs_content)
    
    show_buttons(filename, frame1_buttons)

def delete_file():
    filename = entry_filename.get()
    os.remove(f"{filename}.vbs")
    hide_buttons(frame1_buttons)

def open_explorer_msgbox():
    filename = entry_filename.get()
    subprocess.run(['explorer', f"/select,{os.path.abspath(f'{filename}.vbs')}"])

def execute_msgbox():
    filename = entry_filename.get()
    subprocess.run(['cmd', '/c', f'start {filename}.vbs'])

def show_buttons(filename, button_frame):
    ttk.Button(button_frame, text="Execute", width=20, command=execute_msgbox).pack(side=LEFT, padx=5)
    ttk.Button(button_frame, text="Open Explorer", width=20, command=open_explorer_msgbox).pack(side=LEFT, padx=5)
    ttk.Button(button_frame, text="Delete File", width=20, command=delete_file).pack(side=LEFT, padx=5)

def hide_buttons(button_frame):
    for widget in button_frame.winfo_children():
        widget.pack_forget()

# Widgets setup for Msg Box tab
label_filename = Label(frame1, text="VBS MsgBox", font=("Impact", 30, "bold"))
label_filename.pack()

entry_filename = Entry(frame1, width=40)
entry_title = Entry(frame1, width=40)
entry_text = Entry(frame1, width=40)
entry_labels = ["File Name", "Title", "Text"]

for label_text, entry in zip(entry_labels, [entry_filename, entry_title, entry_text]):
    Label(frame1, text=label_text).pack()
    entry.pack()

Label(frame1, text="Msg Box Type").pack()

var_box_type = IntVar()
for index, variation in enumerate(VARIATIONS):
    Radiobutton(frame1, text=variation, variable=var_box_type, value=index).pack()

ttk.Button(frame1, text="Generate", width=20, command=fill_file).pack(pady=10)

# Functions for Text to Speech tab
def fill_file_tts():
    filename = entry_tts_filename.get()
    output_text = entry_output.get()
    speed = entry_speed.get()
    volume = entry_volume.get()

    vbs_content = f"""
Set Stimme = CreateObject("SAPI.spVoice")
Set Stimme.Voice = Stimme.GetVoices.Item(0)
Stimme.Rate = {speed}
Stimme.Volume = {volume}
Stimme.Speak "{output_text}"
"""
    with open(f"{filename}.vbs", "w+") as file:
        file.write(vbs_content)
    
    show_buttons_tts(filename, frame2_buttons)

def delete_file_tts():
    filename = entry_tts_filename.get()
    os.remove(f"{filename}.vbs")
    hide_buttons_tts(frame2_buttons)

def open_explorer_tts():
    filename = entry_tts_filename.get()
    subprocess.run(['explorer', f"/select,{os.path.abspath(f'{filename}.vbs')}"])

def execute_tts():
    filename = entry_tts_filename.get()
    subprocess.run(['cmd', '/c', f'start {filename}.vbs'])

def show_buttons_tts(filename, button_frame):
    ttk.Button(button_frame, text="Execute", width=20, command=execute_tts).pack(side=LEFT, padx=5)
    ttk.Button(button_frame, text="Open Explorer", width=20, command=open_explorer_tts).pack(side=LEFT, padx=5)
    ttk.Button(button_frame, text="Delete File", width=20, command=delete_file_tts).pack(side=LEFT, padx=5)

def hide_buttons_tts(button_frame):
    for widget in button_frame.winfo_children():
        widget.pack_forget()

# Widgets setup for Text to Speech tab
label_tts = Label(frame2, text="VBS Speech to Text", font=("Impact", 30, "bold"))
label_tts.pack()

entry_tts_filename = Entry(frame2, width=40)
entry_output = Entry(frame2, width=40)
entry_speed = Entry(frame2, width=40)
entry_volume = Entry(frame2, width=40)

entry_speed.insert(END, "1")
entry_volume.insert(END, "100")

tts_labels = ["File Name", "Output", "Speed (-10 to 10)", "Volume (0 to 100)"]

for label_text, entry in zip(tts_labels, [entry_tts_filename, entry_output, entry_speed, entry_volume]):
    Label(frame2, text=label_text).pack()
    entry.pack()

ttk.Button(frame2, text="Generate", width=20, command=fill_file_tts).pack(pady=10)

# Functions for Taskkill tab
def fill_file_taskkill():
    filename = entry_taskkill_filename.get()
    application = entry_application.get()

    with open(f"{filename}.bat", "w+") as file:
        file.write(f"TASKKILL /F /IM {application}")
    
    show_buttons_taskkill(filename, frame3_buttons)

def delete_file_taskkill():
    filename = entry_taskkill_filename.get()
    os.remove(f"{filename}.bat")
    hide_buttons_taskkill(frame3_buttons)

def open_explorer_taskkill():
    filename = entry_taskkill_filename.get()
    subprocess.run(['explorer', f"/select,{os.path.abspath(f'{filename}.bat')}"])

def execute_taskkill():
    filename = entry_taskkill_filename.get()
    subprocess.run(['cmd', '/c', f'start {filename}.bat'])

def show_buttons_taskkill(filename, button_frame):
    ttk.Button(button_frame, text="Execute", width=20, command=execute_taskkill).pack(side=LEFT, padx=5)
    ttk.Button(button_frame, text="Open Explorer", width=20, command=open_explorer_taskkill).pack(side=LEFT, padx=5)
    ttk.Button(button_frame, text="Delete File", width=20, command=delete_file_taskkill).pack(side=LEFT, padx=5)

def hide_buttons_taskkill(button_frame):
    for widget in button_frame.winfo_children():
        widget.pack_forget()

# Widgets setup for Taskkill tab
label_taskkill = Label(frame3, text="Taskkill", font=("Impact", 30, "bold"))
label_taskkill.pack()

entry_taskkill_filename = Entry(frame3, width=40)
entry_application = Entry(frame3, width=40)

taskkill_labels = ["File Name", "Application (please add file extension at the end)"]

for label_text, entry in zip(taskkill_labels, [entry_taskkill_filename, entry_application]):
    Label(frame3, text=label_text).pack()
    entry.pack()

ttk.Button(frame3, text="Generate", width=20, command=fill_file_taskkill).pack(pady=10)

# Credits
credits = Label(win, text="Made by PianoNic")
credits.pack()

# Run the main loop
win.mainloop()