"""
This is a Grading Python written in Python.
It grades students in mass from an excel file.

"""

# Import Required Libraries and Module
import pandas as pds
import numpy as np
################################


# Define Grading Function
def grade():
    # Get Excel File
    file_path = input("Kindly paste the absolute path to the excel file: \n")

    if file_path:
        my_data = pds.read_excel(f'{file_path}')

        print("\n File imported successfully \n")

        check = input("Continue with grading? (Y for Yes | N for No): ")

        # Check user response
        if check == "Y" or check == "y":
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

            # Grading Completed. Prompt User to view file or close program
            rsp = input("Grading Completed, you can check your excel file or type V to view result now: ")
            if rsp == "V" or rsp == "v":
                print(my_data)
                print('\n')

                # Prompt User to export data to new excel file

                exp = input("Do you want to export the result into a new excel file? (Y for Yes | N for No): ")
                if exp == "Y" or exp == "y":
                    new_file_name = "graded_" + file_path.split("\\")[-1]
                    my_data.to_excel(f"./Exports/{new_file_name}", index=False)

                print("\n Thanks for using the program \n See You soon !!!!")
        else:
            print(
                "Alright, You can come back and use the program anytime. I'll be waiting. \n")


# Program Introduction
run_app = 1

while run_app == 1:
    print("This is a grading program. Saves you time in grading your students score \n\
    HOW TO USE \n \
    1 - Store your students details and score in an excel file (make sure the header of the score/mark column is 'Score' ) \n \
    2 - Save the file, copy the path to the file e.g 'C:Desktop\Grading System\Imports\gst_202.xlsx' \n \
        The path should end with the name of the excel file and the file extension xls or xlsx \n \
    3 - Run the grading program, respond with Y when prompted. \n \
    4 - Paste the copied file path from step 2 and click enter to import your excel file \n \
    5 - Type 'Y' to continue with the grading. Once grading is done type 'V' to view results.\n \
        TYPE X to Exit Program \n ")

    # Prompt User to run grading 
    start = input(
        "Would you like to Grade your students scores? (Y for Yes | N for No):  ")
    if start == "Y" or start == "y":
        grade()
    else:
        end = input("Do you want to exit the program? (Type X to exit): ")
        #######
        if end == "X" or end == "x":
            run_app = 0
            print("Bye for now \n")
            exit()



