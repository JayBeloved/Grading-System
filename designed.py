"""
This is a Grading Python written in Python.
It grades students in mass from an excel file.

TO DO :
- Build GUI for application using tkinter


"""
############################################################################################################################
# Import Required Libraries and Module
import pandas as pds
import numpy as np
import os

############################################################################################################################
from tkinter import *  # import neccessary default modules from tkinter
from tkinter import ttk
from tkinter import messagebox # message box to display the answers when the buttons are clicked
from turtle import color
############################################################################################################################

# Create a tkinter window and Set the width and height of the window
win = Tk()
win.geometry("1100x800")

win.title("GRADING SYSTEM BY GST 202 PROGRAMMING CLASS 2021/2022")

############################################################################################################################

#Define Grade Function
def grade():
    # Get path inputted by the user
    filepath = path_entry.get()


    # Get Excel File
    if filepath:
        my_data = pds.read_excel(f'{filepath}')
        # Create a new Column in the excel file called Grade
        my_data2 = my_data
        my_data2['Grade'] = np.nan

        # Check the User score column and group them into seperate list
        a_list = my_data2['Score'] > 69
        b_list = (my_data2['Score'] > 59) & (my_data2['Score'] < 70)
        c_list = (my_data2['Score'] > 49) & (my_data2['Score'] < 60)
        d_list = (my_data2['Score'] > 44) & (my_data2['Score'] < 50)
        e_list = (my_data2['Score'] > 39) & (my_data2['Score'] < 45)
        f_list = my_data2['Score'] < 40

        # Update The Table with the Grade of each students
        my_data2.loc[a_list, 'Grade'] = "A"
        my_data2.loc[b_list, 'Grade'] = "B"
        my_data2.loc[c_list, 'Grade'] = "C"
        my_data2.loc[d_list, 'Grade'] = "D"
        my_data2.loc[e_list, 'Grade'] = "E"
        my_data2.loc[f_list, 'Grade'] = "F"

        my_data.update(my_data2)

        grade_success = Label(win,
                              text="Scores Graded Successfully, Click the Export Button to Continue",
                              background='Green',
                              foreground='White',
                              font=('Georgia 11'))
        grade_success_canvas = canvas_1.create_window(350, 500, anchor="sw", window=grade_success)

        export_button = Button(
            win, text="Export and Open Graded File",
            width=50, background='#6e7c7c', foreground='Black',
            command=export_file)
        export_button_canvas = canvas_1.create_window(400, 550, anchor="sw", window=export_button)

############################################################################################################################


def export_file():
    # Get path inputted by the user
    filepath = path_entry.get()

    # Get Excel File
    if filepath:
        my_data = pds.read_excel(f'{filepath}')
        # Create a new Column in the excel file called Grade
        my_data2 = my_data
        my_data2['Grade'] = np.nan

        # Check the User score column and group them into seperate list
        a_list = my_data2['Score'] > 69
        b_list = (my_data2['Score'] > 59) & (my_data2['Score'] < 70)
        c_list = (my_data2['Score'] > 49) & (my_data2['Score'] < 60)
        d_list = (my_data2['Score'] > 44) & (my_data2['Score'] < 50)
        e_list = (my_data2['Score'] > 39) & (my_data2['Score'] < 45)
        f_list = my_data2['Score'] < 40

        # Update The Table with the Grade of each students
        my_data2.loc[a_list, 'Grade'] = "A"
        my_data2.loc[b_list, 'Grade'] = "B"
        my_data2.loc[c_list, 'Grade'] = "C"
        my_data2.loc[d_list, 'Grade'] = "D"
        my_data2.loc[e_list, 'Grade'] = "E"
        my_data2.loc[f_list, 'Grade'] = "F"

        my_data.update(my_data2)
        # Export to a new excel file
        new_file_name = "graded_" + filepath.split("\\")[-1]
        my_data.to_excel(f"./Exports/{new_file_name}", index=False)

        # End Process
        export_success = Label(win,
                               text="File Exported Successfully and saved in the Exports Folder",
                               background='Green',
                               foreground='White',
                               font=('Georgia 11'))
        export_success_canvas = canvas_1.create_window(350, 600, anchor="sw", window=export_success)

        # Open the Exported File
        os.startfile(f".\Exports\{new_file_name}")

        # Add Refresh Button
        refresh_button = Button(
            win, text="Refresh",
            width=50, background='#6e7c7c', foreground='Black',
            command=export_file)
        refresh_button_canvas = canvas_1.create_window(
            400, 650, anchor="sw", window=refresh_button)

###########################################################################################################
# Create Canvas for Application Background

# Add Background Image
bg_image = PhotoImage(file="bg.png")

canvas_1 = Canvas(win, width=1100, height=800)
canvas_1.pack(fill= "both", expand= True)

canvas_1.create_image(0, 0, image= bg_image, anchor="nw")

############################################################################################################################
# Add Introduction
canvas_1.create_text(
    550, 30, text="THIS IS A GRADING PROGRAM. IT SAVES YOU TIME IN GRADING YOUR STUDENTS SCORES",
    font=('Georgia 15'),
    fill="Black" 
    )

# Create Steps Labels
step1 = Label(
    win,
    text="\n STEP 1 \n\
    Store your students details and score in an excel file \n\
(make sure the header \
of the score/mark column is 'Score' ) \n",
    font=('Georgia 10'),
    background='#6e7c7c',
    foreground="White",
    width=45,
)

step2 = Label(
    win,
    text="\n STEP 2 \n\
Save the file, copy the path to the file e.g \n\
'C:Desktop\Grading System\Imports\gst_202.xlsx') \n\
The path should end with the name of \n\
the excel file and the file extension xls or xlsx\n",
    font=('Georgia 10'),
    background='#6e7c7c',
    foreground="White",
    width=40,
)

step3 = Label(
    win,
    text="\n STEP 3 \n\
    Paste the copied path in to the text box labeled \n\
    'File Path' and click the import button\n",
    font=('Georgia 10'),
    background='#6e7c7c',
    foreground="White",
    width=40,
)

step4 = Label(
    win,
    text="\n STEP 4 \n\
    If file input is successful\n\
    Grade button becomes visible.",
    font=('Georgia 10'),
    background='#6e7c7c',
    foreground="White",
    width=40
)

step5 = Label(
    win,
    text="\n STEP 5 \n\
    Click the Grade Scores button\n \
    The score in the file will.\n\
    be automatically graded.",
    font=('Georgia 10'),
    background='#6e7c7c',
    foreground="White",
    width=40
)

step6 = Label(
    win,
    text="\n STEP 6 \n\
    Click the Export and Open button\n \
    To export the result to a new excel file\n\
    and open in Excel once it is done.",
    font=('Georgia 10'),
    background='#6e7c7c',
    foreground="White",
    width=40
)

# Display Instrctions in the Canvas
step1_canvas = canvas_1.create_window(20, 100, anchor="w", window=step1)
step2_canvas = canvas_1.create_window(400, 115, anchor="w", window=step2)
step3_canvas = canvas_1.create_window(750, 100, anchor="w", window=step3)
step4_canvas = canvas_1.create_window(20, 230, anchor="w", window=step4)
step5_canvas = canvas_1.create_window(400, 230, anchor="w", window=step5)
step6_canvas = canvas_1.create_window(750, 230, anchor="w", window=step6)
############################################################################################################################
# Add separator into canvas
canvas_1.create_line(0, 300, 1100, 300, fill="White", width=15)

# Add Input Field and Label into canvas for path entry
path_label = Label(win, text="Paste File Path Here : ",
                   width=20, foreground='Blue', font=('Georgia 12'))
path_label_canvas = canvas_1.create_window(20, 350, anchor="w", window=path_label)

path_entry = Entry(win, width=72, font=('Georgia 12'))
path_entry_canvas = canvas_1.create_window(940,350, anchor="e", window=path_entry)
############################################################################################################################
# Define Import Function
def import_file():
    # Get path inputted by the user
    filepath = path_entry.get()

    # Get Excel File
    if filepath:
        my_data = pds.read_excel(f'{filepath}')
        import_success = Label(win,
                                text="File Imported Successfully, Click the Grade Button to Continue",
                                background='Green',
                                foreground='White',
                                font=('Georgia 11'))
        import_success_canvas = canvas_1.create_window(350,410, anchor="sw", window=import_success)

        grade_button = Button(
            win, text="Grade Scores",
            width=50, background='#6e7c7c', foreground='Black',
            command=grade)
        grade_button_canvas = canvas_1.create_window(400, 450, anchor="sw", window=grade_button)
        
    else:
        messagebox.showerror("Empty Path", "The File Path is empty or file does not exist.\n\
    Kindly paste the correct path to the excel file.")

############################################################################################################################
# Add Import button to Canvas
import_button = Button(
    win, text="Import Excel File",
    width=15, background='#6e7c7c', foreground='Black',
    command=import_file)

import_button_canvas = canvas_1.create_window(1060, 350, anchor="e", window=import_button)
############################################################################################################################


############################################################################################################################


############################################################################################################################

# Call the GUI of the application
win.mainloop()





