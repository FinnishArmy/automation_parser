# Intel Confidential

# Ronny Z. Valtonen

from openpyxl import *
from openpyxl.styles import *
from openpyxl.cell import *
from openpyxl.utils import *
from openpyxl.worksheet.dimensions import *

#####################################
# BASIC PREP                        #
# pip install openpyxl (3.0.10)     #
# pip install pillow     (3.0.3)    #
# python 3.9.12                     #
# Linux: sudo pacman -S tk          #
# MacOS: brew install python-tk     #
#####################################

def main():
    cd = Workbook()

    # Make the workbook active
    cd_data = cd.active

    # Setup basic layout
    cd_data['A1'] = "[Insert DUT]"
    cd_data['A6'].fill = PatternFill('solid', start_color="00FF9900")
    cd_data['A7'].fill = PatternFill('solid', start_color="00FF9900")
    cd_data['A8'].fill = PatternFill('solid', start_color="00FF9900")
    cd_data['A9'].fill = PatternFill('solid', start_color="00FF9900")
    cd_data['A2'] = "Power Slider Mode"
    cd_data['A3'] = "EPP (for reference)"
    cd_data['A4'] = "[Insert DUT Info]"
    cd_data['A5'] = "Run"
    cd_data['A6'] = "PCMark10 Score"
    cd_data['A7'] = "Essentials"
    cd_data['A8'] = "Productivity"
    cd_data['A9'] = "Dig Content Creation"
    cd_data['A10'] = "App Startup"
    cd_data['A10'].fill = PatternFill('solid', start_color="00FFCC99")
    cd_data['A11'].fill = PatternFill('solid', start_color="00FFCC99")
    cd_data['A12'].fill = PatternFill('solid', start_color="00FFCC99")
    cd_data['A13'].fill = PatternFill('solid', start_color="00FFCC99")
    cd_data['A14'].fill = PatternFill('solid', start_color="00FFCC99")
    cd_data['A15'].fill = PatternFill('solid', start_color="00FFCC99")
    cd_data['A16'].fill = PatternFill('solid', start_color="00FFCC99")
    cd_data['A17'].fill = PatternFill('solid', start_color="00FFCC99")
    cd_data['A11'] = "Video Confrencing"
    cd_data['A12'] = "Web Browsing"
    cd_data['A13'] = "Spreadsheet"
    cd_data['A14'] = "Writing"
    cd_data['A15'] = "Photo Editing"
    cd_data['A16'] = "Render and Visual"
    cd_data['A17'] = "Video Editing"
    cd_data['A18'] = "DAQ MCP Power"
    cd_data['A19'] = "Perf/Watt"
    cd_data['A18'].fill = PatternFill('solid', start_color="00C0C0C0")
    cd_data['A19'].fill = PatternFill('solid', start_color="00C0C0C0")

    cd_data['B5'] = "R1"
    cd_data['C5'] = "R2"
    cd_data['D5'] = "R3"
    cd_data.column_dimensions['A'].width = 17



    # Save the workbook
    cd.save("combined_Data.xlsx")




if __name__ == "__main__":
    main()