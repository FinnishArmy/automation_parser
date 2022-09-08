# Intel Confidential

# Ronny Z. Valtonen

#####################################
# BASIC PREP                        #
# pip install openpyxl (3.0.10)     #
# pip install xlsxwriter (3.0.3)    #
# pip install pandas (1.4.2)        #
# python 3.9.12                     #
# Linux: sudo pacman -S tk          #
# MacOS: brew install python-tk     #
#####################################

# This program is for use to automatically parse XML PCMark10 Scores for easy reading.
import pandas as pd
import numpy as np

# Program
import sys
import os
import json
import subprocess
import xml.etree.cElementTree as et
from os.path import exists
import xlsxwriter
import csv

# UI
from tkinter import filedialog
import pandas
from tkinter import *


# PC Mark 10 Benchmark
def PC_mark10(file, sheet):
            window = Tk()

            # Set the window title name
            window.title("PCMark10")
            # Set a width and height
            window.configure(width = 200, height = 200)

            # Set a window colour
            window.configure(bg = 'gray18')

            canvas = Canvas(window, width= 500, height= 250, bg="White")
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
            close_benchmark = Button(window, text = "Next Benchmark", command = window.quit).pack(pady=5)
            
            # End the program entirely
            exit_program = Button(window, text = "Exit application", command = window.destroy).pack(pady=5)

            ### For loops to grab the root name of each score catagory ##
            for score in root.iter('PCMark10Score'):
                print(score.text)
                canvas.create_text(10, 10, text="PC10Score:  " + score.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                PC10Score.append(score.text)
                all_scores.append(score.text)
    
            for ess in root.iter('EssentialsScore'):
                print(ess.text)
                canvas.create_text(10, 30, text="Essentials:  " + ess.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                Essentials.append(ess.text)
                all_scores.append(ess.text)
    
            for product in root.iter('ProductivityScore'):
                print(product.text)
                canvas.create_text(10, 50, text="Productivity:  " + product.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                Productivity.append(product.text)
                all_scores.append(product.text)
    
            for dig in root.iter('DigitalContentCreationScore'):
                print(dig.text)
                canvas.create_text(10, 70, text="Digital Content Creation:  " + dig.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                DigContentCreation.append(dig.text)
                all_scores.append(dig.text)
    
            for app in root.iter('AppStartupScore'):
                print(app.text)
                canvas.create_text(10, 90, text="App Startup:  " + app.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                AppStartup.append(app.text)
                all_scores.append(app.text)
    
            for video in root.iter('VideoConferencingScore'):
                print(video.text)
                canvas.create_text(10, 110, text="Video Confrencing:  " + video.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                VideoConfrence.append(video.text)
                all_scores.append(video.text)
    
            for web in root.iter('WebBrowsingScore'):
                print(web.text)
                canvas.create_text(10, 130, text="Web Browsing:  " + web.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                WebBrowsing.append(web.text)
                all_scores.append(web.text)
    
            for spread in root.iter('SpreadsheetsScore'):
                print(spread.text)
                canvas.create_text(10, 150, text="Spreadsheets:  " + spread.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                Spreadsheet.append(spread.text)
                all_scores.append(spread.text)
    
            for write in root.iter('WritingScore'):
                print(write.text)
                canvas.create_text(10, 170, text="Writing:  " + write.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                Writing.append(write.text)
                all_scores.append(write.text)
    
            for photo in root.iter('PhotoEditingScore'):
                print(photo.text)
                canvas.create_text(10, 190, text="Photo Editing:  " + photo.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                PhotoEditing.append(photo.text)
                all_scores.append(photo.text)
    
            for render in root.iter('RenderingAndVisualizationScore'):
                print(render.text)
                canvas.create_text(10, 210, text="Rendering and Visualization:  " + render.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                RenderVisual.append(render.text)
                all_scores.append(render.text)
    
            for videoedit in root.iter('VideoEditingScore'):
                print(videoedit.text)
                canvas.create_text(10, 230, text="Video Editing:  " + videoedit.text, fill="black", font=('Helvetica 15 bold'), anchor='w')
                canvas.pack()
                VideoEditing.append(videoedit.text)
                all_scores.append(videoedit.text)
            
            print("Parse complete.")

            col1 = "PC Mark 10"
            col2 = "Scores"

            name_list = ["Overall Score", "Essentials", "Productivity", "Digital Content Creation", 
            "App Startup", "Video Conferencing", "Web Browsing", "Spreadsheets", "Writing", "Photo Editing",
            "Render and Visualization", "Video Editing"]


            # Add the data to a new worksheet inside the workbook created in the main function.
            my_worksheet = sheet.add_worksheet()
            my_worksheet.set_column('A:A', 20)
            my_worksheet.write('A1', 'Overall Score')
            my_worksheet.write('B1', all_scores[0])
            my_worksheet.write('A2', 'Essentials')
            my_worksheet.write('B2', all_scores[1])
            my_worksheet.write('A3', 'Productivity')
            my_worksheet.write('B3', all_scores[2])
            my_worksheet.write('A4', 'Digital Content Creation')
            my_worksheet.write('B4', all_scores[3])
            my_worksheet.write('A5', 'App Startup')
            my_worksheet.write('B5', all_scores[4])
            my_worksheet.write('A6', 'Video Conferencing')
            my_worksheet.write('B6', all_scores[5])
            my_worksheet.write('A7', 'Web Browsing')
            my_worksheet.write('B7', all_scores[6])
            my_worksheet.write('A8', 'Spreadsheets')
            my_worksheet.write('B8', all_scores[7])
            my_worksheet.write('A9', 'Writing')
            my_worksheet.write('B9', all_scores[8])
            my_worksheet.write('A10', 'Photo Editing')
            my_worksheet.write('B10', all_scores[9])
            my_worksheet.write('A11', 'Render and Visual')
            my_worksheet.write('B11', all_scores[10])
            my_worksheet.write('A12', 'Video Editing')
            my_worksheet.write('B12', all_scores[11])

            #mcp_power(sheet)



# Crossmark Benchmark
def crossmark(file, sheet):
            window = Tk()

            # Set the window title name
            window.title("Crossmark")
            # Set a width and height
            window.configure(width = 500, height = 300)

            # Set a window colour
            window.configure(bg = 'gray18')

            canvas = Canvas(window, width= 500, height= 250, bg="White")

            close_benchmark = Button(window, text = "Next Benchmark", command = window.quit).pack(pady=5)

            exit_program = Button(window, text = "Exit application", command = window.destroy).pack(pady=5)

            list = []

            col = 'Score'

            with open(file, 'r', errors = 'replace') as f:
                content = f.readlines()
    
            print("Parsing Crossmark file...")
            canvas.create_text(10, 30, text="Overall Score:  " + content[17], fill="black", font=('Helvetica 15 bold'), anchor='w')
            print("Overall score: " + content[17])
            list.append(content[17])
    
            canvas.create_text(10, 50, text="Productivity:  " + content[19], fill="black", font=('Helvetica 15 bold'), anchor='w')
            print("Productivity: " + content[19])
            list.append(content[19])
    
            canvas.create_text(10, 70, text="Creativity:  " + content[21], fill="black", font=('Helvetica 15 bold'), anchor='w')
            print("Creativity: " + content[21])
            list.append(content[21])
    
            canvas.create_text(10, 90, text="Responsiveness:  " + content[23], fill="black", font=('Helvetica 15 bold'), anchor='w')
            print("Responsiveness: " + content[23])
            list.append(content[23])
            canvas.pack()

            score_name = ['Overall Score', 'Productivity', 'Creativity', 'Responsiveness']
            col1 = 'Crossmark'

            my_worksheet = sheet.add_worksheet()
            my_worksheet.set_column('A:A', 20)
            my_worksheet.write('A1', 'Overall Score')
            my_worksheet.write('A2', 'Productivity')
            my_worksheet.write('A3', 'Creativity')
            my_worksheet.write('A4', 'Responsiveness')
            my_worksheet.write('B1', list[0])
            my_worksheet.write('B2', list[1])
            my_worksheet.write('B3', list[2])
            my_worksheet.write('B4', list[3])

            print("Parse complete.")


# Get the average power.
def mcp_power(file):
            window = Tk()

            # Set the window title name
            window.title("Power")
            # Set a width and height
            window.configure(width = 350, height = 350)

            # Set a window colour
            window.configure(bg = 'gray18')

            canvas = Canvas(window, width= 250, height= 250, bg="White")

            close_benchmark = Button(window, text = "Next Benchmark", command = window.quit).pack(pady=5)

            exit_program = Button(window, text = "Exit application", command = window.destroy).pack(pady=5)

            power = []

            # Convert the file to a string, split the csv data at the commas
            # convert the list into a string and grab line 332
            # remove the brackets, commans and print out the value
            file: str
            with open(file) as fd:
                for line in fd.readlines():
                    power.append(line.split(','))

            mcp = []
            mcp.append(power[-2])
            
            s = ''.join(str(x) for x in mcp)

            canvas.create_text(120, 50, text="MCP Power AVG:  " + s.strip('[]').strip("'").split(',')[3].strip(), fill="black", font=('Helvetica 15 bold'))
            canvas.pack()

            print(s.strip('[]').strip("'").split(',')[3].strip())

def cinebench_multicore(file, sheet):
    print("Cinebench MultiThread Detected")
    window = Tk()

    # Set the window title name
    window.title("Cinebench")
     # Set a width and height
    window.configure(width = 350, height = 350)

    # Set a window colour
    window.configure(bg = 'gray18')

    canvas = Canvas(window, width= 250, height= 250, bg="White")

    close_benchmark = Button(window, text = "Next Benchmark", command = window.quit).pack(pady=5)

    exit_program = Button(window, text = "Exit application", command = window.destroy).pack(pady=5)

    score = []

    # Convert the file to a string, split the csv data at the commas
    # convert the list into a string and grab line 332
    # remove the brackets, commans and print out the value
    file: str
    with open(file) as fd:
        for line in fd.readlines():
            score.append(line.split(','))

    result = (str(score[-2]))
    final_string = result[5:9]
    print(final_string)

    canvas.create_text(90, 50, text="Multi: " + final_string, fill="black", font=('Helvetica 15 bold'), anchor='w')
    canvas.pack()






def cinebench_singlecore(file, sheet):
    print("Cinebench SingleThread Detected")
    window = Tk()

    # Set the window title name
    window.title("Cinebench")
     # Set a width and height
    window.configure(width = 350, height = 350)

    # Set a window colour
    window.configure(bg = 'gray18')

    canvas = Canvas(window, width= 250, height= 250, bg="White")

    close_benchmark = Button(window, text = "Next Benchmark", command = window.quit).pack(pady=5)

    exit_program = Button(window, text = "Exit application", command = window.destroy).pack(pady=5)

    score = []

    # Convert the file to a string, split the csv data at the commas
    # convert the list into a string and grab line 332
    # remove the brackets, commans and print out the value
    file: str
    with open(file) as fd:
        for line in fd.readlines():
            score.append(line.split(','))

    result = (str(score[-2]))
    final_string = result[5:9]
    print(final_string)

    canvas.create_text(90, 50, text="Single: " + final_string, fill="black", font=('Helvetica 15 bold'), anchor='w')
    canvas.pack()

def touch_xprt(file, sheet):
    window = Tk()

    # Set the window title name
    window.title("Touch Xprt 2016")
    # Set a width and height
    window.configure(width = 200, height = 200)

    # Set a window colour
    window.configure(bg = 'gray18')

    canvas = Canvas(window, width= 500, height= 500, bg="White")
    # Set the tree to parse the file selected by the user
    tree = et.parse(file)

    # Buttons
    close_benchmark = Button(window, text = "Next Benchmark", command = window.quit).pack(pady=5)
    exit_program = Button(window, text = "Exit application", command = window.destroy).pack(pady=5)

    root = tree.getroot()
    value_height = 70

    for next_neighbor in root.iter('OverAllScore'):
        workload = next_neighbor.attrib
        main_score = list(workload.values())[0]
        canvas.create_text(90, value_height, text = "Overall Score:  " + main_score, fill="black", font=('Helvetica 15 bold'), anchor='w')
        canvas.pack()

        print(main_score)

    # Loop through the XML file's elements
    for neighbor in root.iter('WorkLoad'):
        value_height = value_height + 20
        workloads = neighbor.attrib

        # Convert the dict to a list and grab the values of each element
        workload_to_list = list(workloads.values())[0]
        workload_values = list(workloads.values())[1]
        print(workload_to_list)
        print(workload_values)

        canvas.create_text(90, value_height, text = workload_to_list + ":  " + workload_values, fill="black", font=('Helvetica 15 bold'), anchor='w')
        canvas.pack()



def geekbench(file, sheet):
    window = Tk()

    # Set the window title name
    window.title("Geekbench")
    # Set a width and height
    window.configure(width = 200, height = 200)

    # Set a window colour
    window.configure(bg = 'gray18')

    canvas = Canvas(window, width= 500, height= 250, bg="White")
    print("TEST")
    f = open(file)

    data = json.load(f)

    multi_score = data["multicore_score"]
    single_score = data["score"]
    print(multi_score)
    print(single_score)

    close_benchmark = Button(window, text = "Next Benchmark", command = window.quit).pack(pady=5)

    exit_program = Button(window, text = "Exit application", command = window.destroy).pack(pady=5)

    canvas.create_text(90, 50, text="Multicore Score: " + str(multi_score), fill="black", font=('Helvetica 15 bold'), anchor='w')
    canvas.create_text(90, 80, text="Singlecore Score: " + str(single_score), fill="black", font=('Helvetica 15 bold'), anchor='w')
    canvas.pack()


# Automatically detect which benchmark was selected.
def pick_file(window, workbook):

    file: str
    for file in window.filename:
        # If it's a PCMark10 benchmark, call the proper function
        if "PCMark10_result" in file: #Instead check the type
            PC_mark10(file, workbook)
    
        # If it's a Crossmark benchmark, call the proper function.
        if "Crossmark.txt" in file:
            crossmark(file, workbook)

        # If it's a power data file, call the proper function.
        if "summary" in file:
            mcp_power(file)

        # If it's a cinebench multithreaded txt, call the proper function.
        if "MultiThreaded" in file:
            cinebench_multicore(file, workbook)

        # If it's a cinebench singlethreaded txt, call the proper function.
        if "SingleThreaded" in file:
            cinebench_singlecore(file, workbook)
        
        if "TouchXPRT" in file:
            touch_xprt(file, workbook)

        if "Geekbench" in file:
            geekbench(file, workbook)
        

        window.mainloop()



# Driver
def main():
    # Declare a window
    window = Tk()

    # Set the window title name
    window.title("DTT Parser")
    # Set a width and height
    window.configure(width = 500, height = 300)

    # Set a window colour
    window.configure(bg = 'gray18')

    exit_program = Button(window, text = "Exit application", command = window.destroy)
    exit_program.pack(pady=20)

    # Move the window into the center of the screen
    winWidth = window.winfo_reqwidth()
    winHeight = window.winfo_reqheight()
    posRight = int(window.winfo_screenwidth() / 2 - winWidth / 2)
    posDown = int(window.winfo_screenheight() / 2 - winHeight / 2)
    window.geometry("+{}+{}".format(posRight, posDown))

    # Ask user for the csv file
    window.filename = filedialog.askopenfilenames(initialdir= "/", title = "Select Participant Log File")

    #check if excel sheet is there, if so, clean, if not make it
    file_exists = os.path.exists('benchmarks.xlsx')

    if file_exists != True:
        workbook = xlsxwriter.Workbook('benchmarks.xlsx')

    else:
        print("Excel sheet already exists! Deleting previous sheet")
        os.remove('benchmarks.xlsx')
        workbook = xlsxwriter.Workbook('benchmarks.xlsx')


    # Automatically detect which benchmark was selected.
    pick_file(window, workbook)

    workbook.close()

if __name__ == "__main__":
    main()