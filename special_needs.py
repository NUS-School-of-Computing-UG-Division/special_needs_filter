# Importing necessary pacakages for the program
import pandas as pd
import os

# Asks the user to input the directory and the filename
# Returns pandas dataframe
def getFile(): 
    input_loc = input("Enter the directory path of input file : ")
    input_filename = input("Enter the file name: ")
    input_file_add = os.path.join(input_loc, input_filename)
    file = pd.read_excel(input_file_add,header=1)
    return file


try:
    # Trying to get the file or print error.
    file = getFile()

    # Asking the user for output directory and filename
    output_directory = input("Enter the directory path to save the output file: ")
    output_name = input("Enter the output file name which will be followed by course name: ")
    if (not os.path.exists(output_directory)):
        raise Exception("Sorry directory does not exist")

    # Filtering the file which have School of Computing and Multi Disciplinary Programme
    new_file = file.query("`COURSE HOST FACULTY` in ('School of Computing', 'Multi Disciplinary Programme')")
    # Dropping the unnecessary columns
    unncessary_col = ["Unnamed: 0","ACAD PLAN DESCR", "EXAM DATE","EXAM DAY","START TIME", "END TIME", "ORIGINAL DURATION","TOTAL EXTRA TIME GIVEN /hr","TOTAL BREAK TIME GIVEN /hr", "Actual Extra Time","Actual Break Time","REVISED DURATION","Throughout Candidature (Y/N)","Start Date","End Date","Degree Level","Primary Component","Non-Graded Components"]
    new_file = new_file.drop(unncessary_col,axis = 1,errors= "ignore")
    
    # Droppping the courses which are higher than 5000
    new_file['Last_4_Digits'] = new_file['COURSE CODE'].str.extract(r'(\d+)')[0].astype(int)

    # Filter rows where 'Last_4_Digits' are greater than 5000
    new_file = new_file[new_file['Last_4_Digits'] < 5000]

    # Take input from the user for course codes to exclude
    codes_to_exclude = input("Enter course codes to exclude (comma-separated) : ").split(',')
    codes_to_exclude = [code.strip() for code in codes_to_exclude]
    new_file = new_file[~new_file["COURSE CODE"].isin(codes_to_exclude)]

    for course in new_file["COURSE CODE"].unique():    
        course_file = new_file[new_file["COURSE CODE"] == course]
        output_filename = f" {output_name} {course}.xlsx" 
        output_file_location = os.path.join(output_directory, output_filename);
        course_file.to_excel(output_file_location,engine = 'openpyxl')  
except Exception as X:
    print(X)


