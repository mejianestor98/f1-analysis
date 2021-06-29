import PyPDF2
import xlsxwriter
import tkinter as tk
from tkinter import filedialog


#===============================FILE SELECTION============================

dialog = tk.Tk()

file_path = filedialog.askopenfilename()

path_parts = file_path.split('/')

file_name = path_parts[len(path_parts) - 1].split('.')[0]

dialog.destroy()

#=========================================================================

#==============================HELPER FUNCTIONS===========================

def is_lap_time(text):

    try:
        result = text.split(':')[1].split(':')
        return True

    except:
        return False

def text_to_lap_time(laptime):



    parts = laptime.split(':')

    if len(parts) == 2:

        minutes = int(parts[0]) * 60
        parts2 = parts[1].split('.')
        seconds = int(parts2[0])
        miliseconds = float(f'0.{parts2[1]}')

        result = minutes + seconds + miliseconds

    elif len(parts) == 3:

        result = laptime


    return result

#=========================================================================

#===============================FILE READING==========================

# creating an object 
file = open(file_path, 'rb')

# creating a pdf reader object
fileReader = PyPDF2.PdfFileReader(file)

# print the number of pages in pdf file
pages = fileReader.numPages

pagesData = []

for i in range(pages):
    pageObj = fileReader.getPage(i)
    pagesData.append(pageObj.extractText().split('\n'))

#====================================================================

#==============================DATA CLEANING=========================


x = 0
driver_control = 0

driverData = []

#FP and Q = 2, R = 1
offset = 2

for page in pagesData : 

    for j in range(len(page)):
        if (page[j].isdigit()) and (not is_lap_time(page[j + offset])) and (len(page[j + offset].split(' ')) == 2) and (not "Formula" in page[j + offset]) and (not page[j + offset] == ' '):
            print(f"found driver {page[j]} - {page[j+offset]}! driver count: {x}")
            driverData.append([])
            driver_control = x
            driverData[driver_control].append(page[j])
            x += 1

        else :
            driverData[driver_control].append(page[j])

danny = driverData[1]

drivers = []
driver_times = []

for driver in driverData:

    drivers.append(f"{driver[0]} - {driver[offset].split(' ')[1][0:3]}")

    first_vector = []
    second_vector = []

    #True to first, False to Second
    write_to = False

    for y in range(len(driver)) :

        try:
                
            if driver[y] == "TIME":

                if int(driver[y + 1]) > 1:
                    write_to = False
                else :
                    write_to = True

            if is_lap_time(driver[y]):

                if write_to :
                    first_vector.append(text_to_lap_time(driver[y]))
                else :
                    second_vector.append(text_to_lap_time(driver[y]))
        except:
            print(f"Error with driver: {driver}")
        
    driver_times.append(first_vector + second_vector)

#=====================================================================

#==============================EXCEL EXPORT===========================

workbook = xlsxwriter.Workbook(f'{file_name}.xlsx')

worksheet = workbook.add_worksheet("LapTimes")
worksheet.freeze_panes(2, 1)

worksheet.write(2, 0, "TOD")

for i in range(len(drivers)) :

    worksheet.write(1, i + 1, drivers[i])

for i in range(len(driver_times)):

    for j in range(len(driver_times[i])):

        worksheet.write(j + 2, i + 1, driver_times[i][j])

workbook.close()