"""
This is a Grading Python written in Python.
It grades students in mass from an excel file.

TO DO :
- Build GUI for application using tkinter


"""
############################################################################################################################
# Import Required Libraries and Module
from doctest import master
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

        grade_success_frame = Frame(win)
        grade_success = Label(grade_success_frame,
                            text="Scores Graded Successfully, Click the Export Button to Continue",
                            background='Green',
                            foreground='White',
                            font=('Georgia 11'))
        grade_success.pack(fill="both", expand=True)
        grade_success_frame.grid(row=9, column=0, columnspan=5, sticky="ew")

        export_button = Button(
            win, text="Export and Open Graded File",
            width=50, background='Black', foreground='White',
            command=export_file)
        export_button.grid(row=10, column=3, padx=5, pady=5)

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
        export_success_frame = Frame(win)
        export_success = Label(export_success_frame,
                                text="File Exported Successfully and saved in the Exports Folder",
                                background='Green',
                                foreground='White',
                                font=('Georgia 11'))
        export_success.pack(fill="both", expand=True)
        export_success_frame.grid(
            row=11, column=0, columnspan=5, sticky="ew")

        # Open the Exported File
        os.startfile(f"./Exports/{new_file_name}")

###########################################################################################################

# Create Labels for introduction and instructions
placeholder2 = Label(win, text="             ") # Invisible label to occupy space to ensure introduction label is aligned.

introduction = Label(
    win,
    text="THIS IS A GRADING PROGRAM. IT SAVES YOU TIME IN GRADING YOUR STUDENTS SCORES",
    background='Grey',
    font=('Georgia 15')
    
)

step1 = Label(
    win,
    text="\n STEP 1 \n\
    Store your students details and score in an excel file \n\
(make sure the header \
of the score/mark column is 'Score' ) \n",
    font=('Georgia 10'),
    background='#28614a'
)

step2 = Label(
    win,
    text="\n STEP 2 \n\
Save the file, copy the path to the file e.g \n\
'C:Desktop\Grading System\Imports\gst_202.xlsx') \n\
The path should end with the name of \n\
the excel file and the file extension xls or xlsx\n",
    font=('Georgia 10'),
    borderwidth=1
)

step3 = Label(
    win,
    text="\n STEP 3 \n\
    Paste the copied path in to the text box labeled \n\
    'File Path' and click the import button\n",
    font=('Georgia 10'),
    borderwidth=1
)

step4 = Label(
    win,
    text="\n STEP 4 \n\
    If file input is successful\n\
    Grade button becomes visible.",
    font=('Georgia 10'),
    borderwidth=1
)

step5 = Label(
    win,
    text="\n STEP 5 \n\
    Click the Grade Scores button\n \
    The score in the file will.\n\
    be automatically graded.",
    font=('Georgia 10'),
    borderwidth=1
)

step6 = Label(
    win,
    text="\n STEP 6 \n\
    Click the Export and Open button\n \
    To export the result to a new excel file\n\
    and open in Excel once it is done.",
    font=('Georgia 10'),
    borderwidth=1
)
############################################################################################################################
# Display the Labels on the application

placeholder2.grid(row=0,column=0,columnspan=2)
introduction.grid(row=0, column=2, columnspan=3, pady=10)
step1.grid(row=1, column=2)
step2.grid(row=1, column=3)
step3.grid(row=1, column=4)
step4.grid(row=2, column=2)
step5.grid(row=2, column=3)
step6.grid(row=2, column=4)

############################################################################################################################

# Create a Separator
status_frame = Frame(win)
status = Label(status_frame, text="  ", background='Grey')
status.pack(fill="both", expand=True)
status_frame.grid(row=4, column=0, columnspan=5, sticky="ew")




# Create Input Box to paste file path
path_label = Label(win, text="Enter File Path Below :\n\n\n", width= 20, foreground='Blue', font=('Georgia 12'))
path_label.grid(row=5, column=2, sticky=N, columnspan=3, padx=5, pady=5)

path_entry = Entry(win, width=150)
path_entry.grid(row=5, column=2, sticky=W, columnspan=3, padx=5, pady=5)



# Define the Import function for the import button
def import_file():

    # Get path inputted by the user
    filepath = path_entry.get()

    # Get Excel File
    if filepath:
        my_data = pds.read_excel(f'{filepath}')
        import_success_frame = Frame(win)
        import_success = Label(import_success_frame,
                            text="File Imported Successfully, Click the Grade Button to Continue",
                            background='Green',
                            foreground='White',
                            font=('Georgia 11'))
        import_success.pack(fill="both", expand=True)
        import_success_frame.grid(row=7, column=0, columnspan=5, sticky="ew")

        grade_button = Button(
            win, text="Grade Scores",
            width=30, background='Black', foreground='White', 
            command=grade)
        grade_button.grid(row=8, column=3, padx=5, pady=5)
    else:
        messagebox.showerror("Empty Path", "The File Path is empty or file does not exist.\n\
Kindly paste the correct path to the excel file.")



# Create and add the import button to the GUI
import_button = Button(
    win, text="Import Excel File",
    width=30, background='Black', foreground='White', 
    command=import_file)

import_button.grid(row=6, column=3, padx=5, pady=5)


# Call the GUI of the application
win.mainloop()


