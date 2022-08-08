# Intel Confidential

# Ronny Z. Valtonen

from cgi import test
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

def open_excel():
    pass


def template(cd_data):
    # Setup the headers in a dict
    headers_a = (
        ('[Insert DUT]', None), ('Power Slider Mode', None, '[Insert AC/DC]'), ('EPP (for reference)', None), 
        ('[Insert DUT Info]', None), ('Run', 'R1', 'R2', 'R3'), ('PCMark10 Score', None), ('Essentials', None), 
        ('Productivity', None), ('Dig Content Creation', None), ('App Startup', None), ('Video Confrencing', None),
        ('Web Browsing', None), ('Spreadsheet', None), ('Writing', None), ('Photo Editing', None), 
        ('Render and Visual', None), ('Video Editing', None), ('DAQ MCP Power', None), ('Perf/Watt', None), 
        ('Crossmark Score', None), ('Productivity', None), ('Creativity', None), ('Responsiveness', None), 
        ('DAQ MCP Power', None))

    # Go through each dict entry and add it to the worksheet
    for i in headers_a:
        print(i)
        cd_data.append(i)

    # Fix up alignment
    cd_data['C4'].alignment = Alignment(horizontal='center')
    cd_data.column_dimensions['A'].width = 17

    int = 0

    # Set the colours and make it magical
    for cell in cd_data['A']:
        if cell.value != None:
            int = int+1

            if (6<=int<=9):
                cd_data['A' + str(int)].fill = PatternFill('solid', start_color="00FF9900")

            if (20<=int<=20):
                cd_data['A' + str(int)].fill = PatternFill('solid', start_color="00FF9900")

            if (21<=int<=23):
                cd_data['A' + str(int)].fill = PatternFill('solid', start_color="00FFCC99")

            if (10<=int<=17):
                cd_data['A' + str(int)].fill = PatternFill('solid', start_color="00FFCC99")
            
            if (18<=int<=19):
                cd_data['A' + str(int)].fill = PatternFill('solid', start_color="00C0C0C0")

        # Open the benchmakrs excel sheet, but make sure it exists.
    path = 'benchmarks.xlsx'

    file_exists = exists(path)

    if file_exists == True:
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        cell_obj = sheet_obj.cell(row=2, column=1)

        second_col = sheet_obj['B']

        # Crossmark
        if cell_obj.value == 'Productivity':
            print("Crossmark data detected, compiling information")

            cross_Data = []
            # Loop through the column and get the values.
            for x in range(len(second_col)):
                cross_Data.append(second_col[x].value)
                print(second_col[x].value)

            # Move to the next column

            check_data_b = cd_data['B']
            check_data_c = cd_data['C']

            if check_data_b[21].value != None:
                print("Run1 already is filled")
                if check_data_c[21].value != None:
                    print("Run2 already is filled")

                if check_data_c[21].value == None:
                    print("Population Run2")
                    cd_data['C20'] = cross_Data[0]
                    cd_data['C21'] = cross_Data[1]
                    cd_data['C22'] = cross_Data[2]
                    cd_data['C23'] = cross_Data[3]

            # Write to this column
            if check_data_b[21].value == None:
                print("Populating Run1")
                cd_data['B20'] = cross_Data[0]
                cd_data['B21'] = cross_Data[1]
                cd_data['B22'] = cross_Data[2]
                cd_data['B23'] = cross_Data[3]
                
                
                
                


            # Paste info into next free column in cd_data
            # .append to not overwrite info




        # PCMark10
        if cell_obj.value == 'Essentials':
            print("PCMark10 data detected, compiling information")

            # Loop through the column and get the values.
            for x in range(len(second_col)):
                pc_Data = second_col[x].value
                print(second_col[x].value)
            

    else:
        print("Benchmark data does not exist.")
    


def main():
    cd = Workbook()

    # Make the workbook active
    cd_data = cd.active

    # Make template
    template(cd_data)

    # Save the workbook
    cd.save("combined_Data.xlsx")




if __name__ == "__main__":
    main()