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


def template(cd_data):

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
                # Grab the row column value
                this = second_col[x].value.strip()

                 # Add to our tuple with a None value afterwards to help later with appending
                cross_Data += (this,) 
                
                # Remove the inital None temp value each iteration
                #result = [i[1:] for i in cross_Data]



            # Move to the next column

            check_data_b = cd_data['B']
            check_data_c = cd_data['C']

            index = 0

            # Check if Column for Run1 is clear for crossmark
            if check_data_b[20].value != None:
                print("Run1 already is filled")

                # If Run 2 is full, write to column of Run 3
                if check_data_c[20].value != None:
                    for r3 in cross_Data:
                        if index <= 3:
                            index += 1
                            cd_data.append({'D': cross_Data[index-1]})
                cd_data.move_range("B25:B28", rows=-5, cols=0)

                # Column for Run 2 writing
                if check_data_c[20].value == None:
                    print("Populating Run2")
                    for r2 in cross_Data:
                        if index <= 3:
                            index += 1
                            print(index)
                            cd_data.append({'C': cross_Data[index-1]})
                cd_data.move_range("B25:B28", rows=-5, cols=1)


            # Column for Run 1 writing
            if check_data_b[20].value == None:
                print("Populating Run1")
                for p in cross_Data:
                    if index <= 3:
                        index += 1
                        #print(cross_Data[index-1])
                        print(index)
                        cd_data.append({'C': cross_Data[index-1]})
            cd_data.move_range("B25:B28", rows= -5, cols=0)



        # PCMark10
        if cell_obj.value == 'Essentials':
            print("PCMark10 data detected, compiling information")

            # Loop through the column and get the values.
            for x in range(len(second_col)):
                pc_Data = second_col[x].value
                #print(second_col[x].value)
            

    else:
        print("Benchmark data does not exist.")
    


def main():

    file_exists = exists("combined_Data.xlsx")

    if file_exists == True:
        cd = load_workbook(filename= "combined_Data.xlsx")
        cd_data = cd.active

    if file_exists == False:
        cd = Workbook()
        cd_data = cd.active
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



    # Make template
    template(cd_data)

    # Save the workbook
    cd.save("combined_Data.xlsx")




if __name__ == "__main__":
    main()