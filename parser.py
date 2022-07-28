# Intel Confidential

# Ronny Z. Valtonen

#####################################
# BASIC PREP                        #
# pip install pandas (1.4.2)        #
# python 2.7.18                     #
# Linux: sudo pacman -S tk          #
# MacOS: brew insatll python-tk     #
#####################################

# This program is for use to automatically parse XML PCMark10 Scores for easy reading.
import pandas as pd
import numpy as np

# Program
import sys
import os
import subprocess
import xml.etree.cElementTree as et

# UI
from tkinter import filedialog
from turtle import clear
import pandas
from tkinter import *

def main(): 
    # Declare a window
    window = Tk()

    # Set the window title name
    window.title("DTT Parser")
    # Set a width and height
    window.configure(width = 500, height = 300)

    # Set a window colour
    window.configure(bg = 'gray18')

    # Move the window into the center of the screen
    winWidth = window.winfo_reqwidth()
    winHeight = window.winfo_reqheight()
    posRight = int(window.winfo_screenwidth() / 2 - winWidth / 2)
    posDown = int(window.winfo_screenheight() / 2 - winHeight / 2)
    window.geometry("+{}+{}".format(posRight, posDown))

    # Ask user for the csv file
    window.filename = filedialog.askopenfilename(initialdir= "/", title = "Select Participant Log File", filetypes = (("XML Files", "*.xml"),("All Files","*.*")))
    print(window.filename)

    canvas= Canvas(window, width= 1000, height= 750, bg="White")

    tree = et.parse(window.filename)

    root = tree.getroot()

    # Empty arrays for entering data.
    PC10Score = []
    Essentials = []
    Productivity = []
    DigContentCreation = []
    AppStartup = []
    VideoConfrence = []
    WebBrowsing = []
    Spreadsheet = []
    Writing = []
    PhotoEditing = []
    RenderVisual = []
    VideoEditing = []

    print("Beginning parse")

    ### For loops to grab the root name of each score catagory ##
    for score in root.iter('PCMark10Score'):
        print(score.text)
        canvas.create_text(150, 50, text="PC10Score:  " + score.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        PC10Score.append(score.text)

    for ess in root.iter('EssentialsScore'):
        print(ess.text)
        canvas.create_text(150, 80, text="Essentials:  " + ess.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        Essentials.append(ess.text)

    for product in root.iter('ProductivityScore'):
        print(product.text)
        canvas.create_text(150, 110, text="Productivity:  " + product.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        Productivity.append(product.text)

    for dig in root.iter('DigitalContentCreationScore'):
        print(dig.text)
        canvas.create_text(150, 130, text="Digital Content Creation:  " + dig.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        DigContentCreation.append(dig.text)

    for app in root.iter('AppStartupScore'):
        print(app.text)
        canvas.create_text(150, 150, text="App Startup:  " + app.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        AppStartup.append(app.text)

    for video in root.iter('VideoConferencingScore'):
        print(video.text)
        canvas.create_text(150, 170, text="Video Confrencing:  " + video.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        VideoConfrence.append(video.text)

    for web in root.iter('WebBrowsingScore'):
        print(web.text)
        canvas.create_text(150, 190, text="Web Browsing:  " + web.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        WebBrowsing.append(web.text)

    for spread in root.iter('SpreadsheetsScore'):
        print(spread.text)
        canvas.create_text(150, 210, text="Spreadsheets:  " + spread.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        Spreadsheet.append(spread.text)

    for write in root.iter('WritingScore'):
        print(write.text)
        canvas.create_text(150, 230, text="Writing:  " + write.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        Writing.append(write.text)

    for photo in root.iter('PhotoEditingScore'):
        print(photo.text)
        canvas.create_text(150, 250, text="Photo Editing:  " + photo.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        PhotoEditing.append(photo.text)

    for render in root.iter('RenderingAndVisualizationScore'):
        print(render.text)
        canvas.create_text(150, 270, text="Rendering and Visualization:  " + render.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        RenderVisual.append(render.text)

    for videoedit in root.iter('VideoEditingScore'):
        print(videoedit.text)
        canvas.create_text(150, 290, text="Video Editing:  " + videoedit.text, fill="black", font=('Helvetica 15 bold'))
        canvas.pack()
        VideoEditing.append(videoedit.text)



    window.mainloop()
    print("Parse complete.")

    # Restart the program when user exists so you can pick a new file.
    subprocess.call([sys.executable, os.path.realpath(__file__)] + sys.argv[1:])

if __name__ == "__main__":
    main()


