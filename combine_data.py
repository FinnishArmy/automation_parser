# Intel Confidential

# Ronny Z. Valtonen

from openpyxl import *
import openpyxl
from openpyxl.styles import *
from openpyxl.cell import *
from openpyxl.utils import *
from openpyxl.worksheet.dimensions import *
from os.path import exists

#####################################
# BASIC PREP                        #
# pip install openpyxl (3.0.10)     #
# pip install pillow     (3.0.3)    #
# python 3.9.12                     #
# Linux: sudo pacman -S tk          #
# MacOS: brew install python-tk     #
#####################################

def template(cd_data):
    # Setup basic layout
    cd_data['A1'] = "[Insert DUT]" 
    cd_data['A2'] = "Power Slider Mode"
    cd_data['A3'] = "EPP (for reference)"
    cd_data['A4'] = "[Insert DUT Info]"
    cd_data['A5'] = "Run"
    cd_data['A6'] = "PCMark10 Score"
    cd_data['A7'] = "Essentials"
    cd_data['A8'] = "Productivity"
    cd_data['A9'] = "Dig Content Creation"
    cd_data['A10'] = "App Startup"
    cd_data['A11'] = "Video Confrencing"
    cd_data['A12'] = "Web Browsing"
    cd_data['A13'] = "Spreadsheet"
    cd_data['A14'] = "Writing"
    cd_data['A15'] = "Photo Editing"
    cd_data['A16'] = "Render and Visual"
    cd_data['A17'] = "Video Editing"
    cd_data['A18'] = "DAQ MCP Power"
    cd_data['A19'] = "Perf/Watt"
    cd_data['C3'] = "[Insert DC/AC Mode]"
    cd_data['C3'].alignment = Alignment(horizontal='center')
    cd_data['B5'] = "R1"
    cd_data['C5'] = "R2"
    cd_data['D5'] = "R3"
    cd_data.column_dimensions['A'].width = 17

    int = 0
    for cell in cd_data['A']:
        if cell.value != None:
            int = int+1

            if (6<=int<=9):
                cd_data['A' + str(int)].fill = PatternFill('solid', start_color="00FF9900")

            if (10<=int<=17):
                cd_data['A' + str(int)].fill = PatternFill('solid', start_color="00FFCC99")
            
            if (18<=int<=19):
                cd_data['A' + str(int)].fill = PatternFill('solid', start_color="00C0C0C0")

def open_excel():
    # Open the benchmakrs excel sheet, but make sure it exists.
    path = 'benchmarks.xlsx'

    file_exists = exists(path)

    if file_exists == True:
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        cell_obj = sheet_obj.cell(row=2, column=1)

        # Crossmark
        if cell_obj.value == 'Productivity':
            print("Crossmark data detected, compiling information")

        # PCMark10
        if cell_obj.value == 'Essentials':
            print("PCMark10 data detected, compiling information")
            second_col = sheet_obj['B']

            # Loop through the column and get the values.
            for x in range(len(second_col)):
                pc_Data = second_col[x].value
                print(second_col[x].value)



        else:
            print("No data exists in the 'benchmarks' excel file.")
            

    else:
        print("Benchmark data does not exist.")


def main():
    cd = Workbook()

    # Make the workbook active
    cd_data = cd.active

    # Make template
    template(cd_data)

    open_excel()

    # Save the workbook
    cd.save("combined_Data.xlsx")




if __name__ == "__main__":
    main()