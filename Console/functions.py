import xlrd
import pandas as pd
import numpy as np
import os
import os.path
import msvcrt
from colorama import init, Fore, Back, Style

FILEDIFFNAME = "Excel_diff"

# Allows color
init(convert=True)


def clear(): return os.system('cls')


def intro():
    print(f"{Style.RESET_ALL}{Fore.GREEN}Excel Data Analyzer Tool{Fore.RESET}")
    print(
        f"Produced for {Fore.LIGHTBLUE_EX}Change {Fore.RED}Healthcare{Fore.RESET}")
    print(f"Made by {Fore.YELLOW}Chaitanya Yedumbaka{Fore.RESET}")
    print("\n")


def ask():
    file1path = input("Please Provide the Path to Your Excel File #1:")

    # Checks Whether the File #1 Path is an actual File Path
    while os.path.isfile(file1path) == False:
        print(f"\n{Fore.RED}INVALID FILE PATH")
        file1path = input(
            f"Please Provide the CORRECT Path to Your Excel File #1:")
        print(Fore.RESET, end='')

    file2path = input("\nPlease Provide the Path to Your Excel File #2:")

    # Checks Whether the File #2 Path is an actual File Path
    while os.path.isfile(file2path) == False:
        print(f"\n{Fore.RED}INVALID FILE PATH")
        file2path = input(
            "Please Provide the CORRECT Path to Your Excel File #2:")
        print(Fore.RESET, end='')

    # Makes sure the inputed File Path hasn't already been used
    if file2path == file1path:
        print(f"\n{Fore.RED}PATH ALREADY USED")
        file2path = input(
            "Please Provide a Path That Hasn't Already Been Used:")
        print(Fore.RESET, end='')

    outputpath = input(
        "\nPlease Provide a Path to Where You Want To Save the Output File:")

    # Checks Whether the Output File Path Location is an actual File Directory
    while os.path.isdir(outputpath) == False:
        print(f"\n{Fore.RED}INVALID FILE DIRECTORY")
        outputpath = input(
            "Please Provide an ACTUAL Path to Where You Want To Save the Output File:")
        print(Fore.RESET, end='')

    print("\n[Analyzing]...")

    # Pandas reads the excel file through its path
    xls1 = pd.ExcelFile(file1path)
    xls2 = pd.ExcelFile(file2path)

    # Variables (s = Sheets #) (i = File #)
    s = 0
    i = 0

    # If the file name already exists, it will make a new one instead of overwritting the file
    while os.path.isfile(f"{outputpath}{FILEDIFFNAME}({i}).xlsx") == True:
        i += 1

    # A for loop that loops through all the sheets inside the excel files and compares them.
    # Then gives an output and saves it
    for sh1, sh2 in zip(xls1.sheet_names, xls2.sheet_names):
        s += 1
        # Converts ExcelFile data into excel dataframe
        data1 = xls1.parse(sh1)
        # Turns the excel file data into a panda dateframe
        df1 = pd.DataFrame(data1)

        # Converts ExcelFile data into excel dataframe
        data2 = xls2.parse(sh2)
        # Turns the excel file data into a panda dateframe
        df2 = pd.DataFrame(data2)

        # Replaces NaN with Empty Space
        df1 = df1.fillna('')
        df2 = df2.fillna('')

        # Makes the two dataframes match
        df1.equals(df2)

        # Compares the two dataframes
        comparsion_values = df1.values == df2.values

        rows, cols = np.where(comparsion_values == False)

        for item in zip(rows, cols):
            df1.iloc[item[0], item[1]] = '{} --> {}'.format(
                df1.iloc[item[0], item[1]], df2.iloc[item[0], item[1]])

        print(
            f"\n{Fore.YELLOW}Excel File #1: {sh1} compared with Excel File #2: {sh2}{Fore.RESET}")
        print(f"\n{df1}")

        # If it is the first sheet, it will create the excel file and write into it
        if (s == 1):
            with pd.ExcelWriter(f"{outputpath}{FILEDIFFNAME}({i}).xlsx", mode='w') as writer:
                df1.to_excel(writer, sheet_name=f'Sheet{s}',
                             index=False, header=True)

        # If it is the second or later sheet, it will create a new sheet and write into it
        if (s >= 2):
            with pd.ExcelWriter(f"{outputpath}{FILEDIFFNAME}({i}).xlsx", mode='a') as writer:
                df1.to_excel(writer, sheet_name=f'Sheet{s}',
                             index=False, header=True)

    # Asks if the user wants to open the output file or not
    viewAns = input(
        "\nThis preview has been saved in an excel file, would you like to view it? (y/n):")

    # Removes the cap sensitivity
    viewAns.casefold()

    while (viewAns != "y" and viewAns != "yes" and viewAns != "n" and viewAns != "no") == True:
        print(f"{Fore.RED}INVAILD INPUT")
        viewAns = input(
            'Please provide either "y" if you wish to view the output excel file or "n" if you wish to not view the output excel file:')

    # Removes color from text
    print(Fore.RESET, end='')

    # Removes the cap sensitivity
    viewAns.casefold()

    # Opens the output file
    if (viewAns == "y" or viewAns == "yes"):
        os.system(f"start EXCEL.EXE {outputpath}{FILEDIFFNAME}({i}).xlsx")
