# Intel Confidential

# Ronny Z. Valtonen

#####################################
# BASIC PREP                        #
# pip install openpyxl (3.0.10)     #
# pip install pandas (1.4.2)        #
# python 2.7.18                     #
# Linux: sudo pacman -S tk          #
# MacOS: brew install python-tk     #
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

def PC_mark10(file):
            window = Tk()

            # Set the window title name
            window.title("DTT Parser")
            # Set a width and height
            window.configure(width = 500, height = 300)

            # Set a window colour
            window.configure(bg = 'gray18')

            canvas = Canvas(window, width= 1000, height= 750, bg="White")
            # Set the tree to parse the file selected by the user
            tree = et.parse(file)
    
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

            all_scores = []
    
            print("Beginning parse")
    
            #btn = Button(win, text = 'Next Benchmark', bd = '5', command = canvas.update_idletasks)
            #btn.pack(side = 'bottom')

            # Open the next file info, if no more, end the program.
            close_benchmark = Button(window, text = "Next Benchmark", command = window.quit)
            close_benchmark.pack(pady=20)
            
            # End the program entirely
            exit_program = Button(window, text = "Exit application", command = window.destroy)
            exit_program.pack(pady=20)

            ### For loops to grab the root name of each score catagory ##
            for score in root.iter('PCMark10Score'):
                print(score.text)
                canvas.create_text(150, 50, text="PC10Score:  " + score.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                PC10Score.append(score.text)
                all_scores.append(score.text)
    
            for ess in root.iter('EssentialsScore'):
                print(ess.text)
                canvas.create_text(150, 80, text="Essentials:  " + ess.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                Essentials.append(ess.text)
                all_scores.append(ess.text)
    
            for product in root.iter('ProductivityScore'):
                print(product.text)
                canvas.create_text(150, 110, text="Productivity:  " + product.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                Productivity.append(product.text)
                all_scores.append(product.text)
    
            for dig in root.iter('DigitalContentCreationScore'):
                print(dig.text)
                canvas.create_text(150, 130, text="Digital Content Creation:  " + dig.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                DigContentCreation.append(dig.text)
                all_scores.append(dig.text)
    
            for app in root.iter('AppStartupScore'):
                print(app.text)
                canvas.create_text(150, 150, text="App Startup:  " + app.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                AppStartup.append(app.text)
                all_scores.append(app.text)
    
            for video in root.iter('VideoConferencingScore'):
                print(video.text)
                canvas.create_text(150, 170, text="Video Confrencing:  " + video.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                VideoConfrence.append(video.text)
                all_scores.append(video.text)
    
            for web in root.iter('WebBrowsingScore'):
                print(web.text)
                canvas.create_text(150, 190, text="Web Browsing:  " + web.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                WebBrowsing.append(web.text)
                all_scores.append(web.text)
    
            for spread in root.iter('SpreadsheetsScore'):
                print(spread.text)
                canvas.create_text(150, 210, text="Spreadsheets:  " + spread.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                Spreadsheet.append(spread.text)
                all_scores.append(spread.text)
    
            for write in root.iter('WritingScore'):
                print(write.text)
                canvas.create_text(150, 230, text="Writing:  " + write.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                Writing.append(write.text)
                all_scores.append(write.text)
    
            for photo in root.iter('PhotoEditingScore'):
                print(photo.text)
                canvas.create_text(150, 250, text="Photo Editing:  " + photo.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                PhotoEditing.append(photo.text)
                all_scores.append(photo.text)
    
            for render in root.iter('RenderingAndVisualizationScore'):
                print(render.text)
                canvas.create_text(150, 270, text="Rendering and Visualization:  " + render.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                RenderVisual.append(render.text)
                all_scores.append(render.text)
    
            for videoedit in root.iter('VideoEditingScore'):
                print(videoedit.text)
                canvas.create_text(150, 290, text="Video Editing:  " + videoedit.text, fill="black", font=('Helvetica 15 bold'))
                canvas.pack()
                VideoEditing.append(videoedit.text)
                all_scores.append(videoedit.text)
            
            print("Parse complete.")

            col1 = "PC Mark 10"
            col2 = "Scores"

            name_list = ["Overall Score", "Essentials", "Productivity", "Digital Content Creation", 
            "App Startup", "Video Conferencing", "Web Browsing", "Spreadsheets", "Writing", "Photo Editing",
            "Render and Visualization", "Video Editing"]

            data = pd.DataFrame({col2:all_scores})

            # use variables
            data.to_excel('pcmark10_data.xlsx', sheet_name='pcmark 10 scores', index=False)

def crossmark(file):
            window = Tk()

            # Set the window title name
            window.title("DTT Parser")
            # Set a width and height
            window.configure(width = 500, height = 300)

            # Set a window colour
            window.configure(bg = 'gray18')

            canvas = Canvas(window, width= 1000, height= 750, bg="White")

            close_benchmark = Button(window, text = "Next Benchmark", command = window.quit)
            close_benchmark.pack(pady=20)

            exit_program = Button(window, text = "Exit application", command = window.destroy)
            exit_program.pack(pady=20)

            list = []

            col = 'Score'

            with open(file, 'r', errors = 'replace') as f:
                content = f.readlines()
    
            print("Parsing Crossmark file...")
            canvas.create_text(150, 290, text="Overall Score:  " + content[17], fill="black", font=('Helvetica 15 bold'))
            print("Overall score: " + content[17])
            list.append(content[17])
            canvas.pack()
    
            canvas.create_text(150, 310, text="Productivity:  " + content[19], fill="black", font=('Helvetica 15 bold'))
            print("Productivity: " + content[19])
            list.append(content[19])
            canvas.pack()
    
            canvas.create_text(150, 330, text="Creativity:  " + content[21], fill="black", font=('Helvetica 15 bold'))
            print("Creativity: " + content[21])
            list.append(content[21])
            canvas.pack()
    
            canvas.create_text(150, 350, text="Responsiveness:  " + content[23], fill="black", font=('Helvetica 15 bold'))
            print("Responsiveness: " + content[23])
            list.append(content[23])
            canvas.pack()

            score_name = ['Overall Score', 'Productivity', 'Creativity', 'Responsiveness']
            col1 = 'Crossmark'
            data = pd.DataFrame({col1:score_name, col:list})
            data.to_excel('cross_mark data.xlsx', sheet_name='data', index=False)

            print("Parse complete.")


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
    # winWidth = window.winfo_reqwidth()
    # winHeight = window.winfo_reqheight()
    # posRight = int(window.winfo_screenwidth() / 2 - winWidth / 2)
    # posDown = int(window.winfo_screenheight() / 2 - winHeight / 2)
    # window.geometry("+{}+{}".format(posRight, posDown))

    # Ask user for the csv file
    window.filename = filedialog.askopenfilenames(initialdir= "/", title = "Select Participant Log File")

    #check if excel sheet is there, if so, clean, if not make it

    
    # Automatically detect which benchmark was selected.
    file: str
    for file in window.filename:
        # If it's a PCMark10 benchmark, call the proper function
        if "PCMark10" in file:
            PC_mark10(file)
    

        # If it's a Crossmark benchmark, call the proper function.
        if "Crossmark" in file:
            crossmark(file)
    
            # Restart the program when user exists so you can pick a new file.
            # subprocess.call([sys.executable, os.path.realpath(__file__)] + sys.argv[1:])
    
        else:
            print("Not a valid file type.")
            
        window.mainloop()

if __name__ == "__main__":
    main()
