from tkinter import *
from tkinter import font, filedialog, messagebox as mb
# from tkinter.ttk import Button, Label, Entry, Style, Notebook, Frame
from tkinter.ttk import *
from functools import partial
import xlrd
import pandas as pd
import numpy as np
import os
import os.path
import msvcrt
from colorama import init, Fore, Back
from PIL import ImageTk, Image

# Allows color
# init(convert=True)

window_height = 400
window_width = 576

FILEDIFFNAME = "Excel_diff"

fontSizeSmall = 7
fontSizeRegular = 10
fontSizeLarge = 12

defaultFont = None


def openFileDialog(ver):
    filename = filedialog.askopenfilename(
        initialdir="/", title="Select Excel File", filetypes=(("Excel Documents", "*.xlsx"), ("All Files", "*.*")))

    if (ver == 1):
        file1Entry.delete(0, END)
        file1Entry.insert(0, filename)
    elif (ver == 2):
        file2Entry.delete(0, END)
        file2Entry.insert(0, filename)


def openDirDialog():
    dirlocation = filedialog.askdirectory(
        initialdir="/", title="Select Output File Directory")

    outputpathEntry.delete(0, END)
    outputpathEntry.insert(0, f"{dirlocation}/")


def openOutputFile():
    if os.path.isdir(outputpathEntry.get()) == False:
        mb.showerror("Cannot Open File",
                     "Cannot retrieve the file because it hasn't been created yet")
        return

    outputpath = outputpathEntry.get()
    os.system(f"start EXCEL.EXE {outputpath}{FILEDIFFNAME}({i}).xlsx")


def compareCommand():
    if os.path.isfile(file1Entry.get()) == False or os.path.isfile(file2Entry.get()) == False or os.path.isdir(outputpathEntry.get()) == False:
        mb.showerror("Missing/Incorrect Inputs",
                     "Missing/Incorrect File 1, File 2, or Output Paths")
        return

    file1path = file1Entry.get()
    file2path = file2Entry.get()
    outputpath = outputpathEntry.get()
    nb.select(outputFrame)

    outputEntry.delete('1.0', END)

    # Pandas reads the excel file through its path
    xls1 = pd.ExcelFile(file1path)
    xls2 = pd.ExcelFile(file2path)

    # Variables (s = Sheets #) (i = File #)
    s = 0
    global i
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

        outputEntry.insert(
            END, f"Excel File #1: {sh1} compared with Excel File #2: {sh2}")
        outputEntry.insert(END, f"\n\n{df1}\n\n")

        # print(
        #     f"\n{Fore.YELLOW}Excel File #1: {sh1} compared with Excel File #2: {sh2}{Fore.RESET}")
        # print(f"\n{df1}")

        # If it is the first sheet, it will create the excel file and write into it
        if (s == 1):
            with pd.ExcelWriter(f"{outputpath}{FILEDIFFNAME}({i}).xlsx", mode='w') as writer:
                df1.to_excel(
                    writer, sheet_name=f"Sheet{s}", index=False, header=True)

        # If it is the second or later sheet, it will create a new sheet and write into it
        if (s >= 2):
            with pd.ExcelWriter(f"{outputpath}{FILEDIFFNAME}({i}).xlsx", mode='a') as writer:
                df1.to_excel(writer, sheet_name=f'Sheet{s}',
                             index=False, header=True)


def cmdCompareCommand():
    if mb.askquestion("Open Console Interface Tool?",
                      "Do you wish to open the Console Interface Version of Excel Data Analyzer?") == "yes":
        root.destroy()
        os.system("python ./Console/main.py")


root = Tk()
root.title("Excel Data Analyzer Tool")
root.resizable(width=0, height=0)
root.iconbitmap("./Images/icon.ico")

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x_coordinate = int(screen_width / 2) - (window_width / 2)
y_coordinate = int(screen_height / 2) - (window_height / 2)

root.geometry("%dx%d+%d+%d" %
              (window_width, window_height, x_coordinate, y_coordinate))

fontSmall = font.Font(family=defaultFont, size=fontSizeSmall)
fontRegular = font.Font(family=defaultFont, size=fontSizeRegular)
fontLarge = font.Font(size=fontSizeLarge)

buttonStyle = Style()
buttonStyle.configure("W.TButton", font=fontRegular)

nb = Notebook(root)
nb.grid(row=0, column=0)

# Compare Frame (Holds the Compare Window Content)
compareFrame = Frame(root)
compareFrame.grid(padx=5, pady=5)
nb.add(compareFrame, text="Compare")

file1Label = Label(compareFrame, text="File 1:",
                   font=(defaultFont, fontSizeRegular))
file1Label.grid(row=0, column=0)

file1Entry = Entry(compareFrame, width=50, font=(defaultFont, fontSizeRegular))
file1Entry.grid(row=0, column=1)
file1Entry.insert(0, "<Select File>")

openFile1Button = Button(compareFrame, text="Browse",
                         style="W.TButton", command=partial(openFileDialog, 1))
openFile1Button.grid(row=0, column=2, padx=5, pady=5)


compareButton = Button(compareFrame, text="Compare",
                       style="W.TButton", command=compareCommand)
compareButton.grid(row=1, column=1, padx=10, pady=10)


file2Label = Label(compareFrame, text="File 2:",
                   font=(defaultFont, fontSizeRegular))
file2Label.grid(row=2, column=0)

file2Entry = Entry(compareFrame, width=50, font=(defaultFont, fontSizeRegular))
file2Entry.grid(row=2, column=1)
file2Entry.insert(0, "<Select File>")

openFile2Button = Button(compareFrame, text="Browse",
                         style="W.TButton", command=partial(openFileDialog, 2))
openFile2Button.grid(row=2, column=2, padx=5, pady=5)


outputpathLabel = Label(compareFrame, text="Output:",
                        font=(defaultFont, fontSizeRegular))
outputpathLabel.grid(row=3, column=0, padx=5, pady=5)

outputpathEntry = Entry(compareFrame, width=50,
                        font=(defaultFont, fontSizeRegular))
outputpathEntry.grid(row=3, column=1, padx=5, pady=5)
outputpathEntry.insert(0, "<Select File Directory>")

outputpathEntryButton = Button(compareFrame, text="Browse",
                               style="W.TButton", command=openDirDialog)
outputpathEntryButton.grid(row=3, column=2, padx=5, pady=5)

cmdCompareButton = Button(
    compareFrame, text="Console Interface?", style="W.TButton", command=cmdCompareCommand)
cmdCompareButton.place(relx=0.01, rely=0.985, anchor=SW)

companyLabel = Label(root, text="Made For Change Healthcare Use")
companyLabel.place(relx=1, rely=0, anchor=NE)

nameLabel = Label(compareFrame, text="Made By Chaitanya Yedumbaka")
nameLabel.place(relx=1, rely=1, anchor=SE)


# Output Frame (Holds the Output Window Content)
outputFrame = Frame(root)
outputFrame.grid(padx=5, pady=5)
nb.add(outputFrame, text="Output")

outputLabel = Label(outputFrame, text="Output",
                    font=(defaultFont, fontSizeRegular))
outputLabel.grid(row=0, column=0, padx=5, pady=5)

outputEntry = Text(outputFrame, width=80, height=17,
                   font=(defaultFont, fontSizeRegular))
outputEntry.grid(row=1, column=0, padx=5, pady=5)

outputOpenLabel = Label(
    outputFrame, text="This preview has been saved in an excel file, if you wish to view it press Open.", font=(defaultFont, fontSizeRegular))
outputOpenLabel.grid(row=2, column=0, padx=5, pady=5)

outputOpenButton = Button(outputFrame, text="Open",
                          style="W.TButton", command=openOutputFile)
outputOpenButton.grid(row=3, column=0, padx=5, pady=(0, 5))

nameLabel = Label(outputFrame, text="Made By Chaitanya Yedumbaka")
nameLabel.place(relx=1, rely=1, anchor=SE)

# How To Use Frame (Contains the How To Use Window Content)
howtoFrame = Frame(root)
howtoFrame.grid(padx=5, pady=5)
nb.add(howtoFrame, text="How To Use")

howtoUseLabel = Label(howtoFrame, text="How To Use Excel Data Analyzer Tool", font=(
    defaultFont, fontSizeLarge, "underline"))
howtoUseLabel.grid(row=0, column=0, padx=5, pady=(10, 5))

tabRulesLabel = Label(
    howtoFrame, font=(defaultFont, fontSizeRegular), text="1. In order to use this tool, please check you are in the Compare Page by looking at the top left of the window.", wraplength=565)
tabRulesLabel.grid(row=1, column=0, padx=5, pady=(5, 0))

# tabImgLoad = ImageTk.PhotoImage(Image.open("./Images/TabImage.jpg"))
# tabImg = Label(howtoFrame, image=tabImgLoad)
# tabImg.grid(row=2, column=0, padx=5)

file1RulesLabel = Label(
    howtoFrame, font=(defaultFont, fontSizeRegular), text="2. Once you are in the compare page, please specify File 1's Location Path by either pressing Browse or manually typing it in.", wraplength=565)
file1RulesLabel.grid(row=2, column=0, padx=5, pady=(5, 0))

# file1ImgLoad = ImageTk.PhotoImage(Image.open("./Images/File1Image.jpg"))
# file1Img = Label(howtoFrame, image=file1ImgLoad)
# file1Img.grid(row=4, column=0)

file2RulesLabel = Label(
    howtoFrame, font=(defaultFont, fontSizeRegular), text="3. After specifying File 1's Path, please specifiy File 2's Location Path by eithering pressing Browse or manually typing it in.", wraplength=565)
file2RulesLabel.grid(row=3, column=0, padx=5, pady=(5, 0))

# file2ImgLoad = ImageTk.PhotoImage(Image.open("./Images/File2Image.jpg"))
# file2Img = Label(howtoFrame, image=file2ImgLoad)
# file2Img.grid(row=6, column=0)

outputPathRulesLabel = Label(
    howtoFrame, font=(defaultFont, fontSizeRegular), text="4. After specifying both Files' Paths, please input the Output Location Path which where the output file will be saved at by eithering pressing Browse or manually typing it in. Please make sure it is a directory/folder and not a file.", wraplength=565)
outputPathRulesLabel.grid(row=4, column=0, padx=5, pady=(5, 0))

# outputPathImgLoad = ImageTk.PhotoImage(Image.open("./Images/OutputImage.jpg"))
# outputPathImg = Label(howtoFrame, image=outputPathImgLoad)
# outputPathImg.grid(row=8, column=0)

compareButtonRulesLabel = Label(howtoFrame, font=(defaultFont, fontSizeRegular),
                                text="5. After specifying all fields, press Compare in order to start the comparing process", wraplength=565)
compareButtonRulesLabel.grid(row=5, column=0, padx=5, pady=(5, 0))

outputTabRulesLabel = Label(howtoFrame, font=(defaultFont, fontSizeRegular),
                            text="6. Once the comparsion is completed, you will be brought to the Output Page. In the output page, you can see a preview of the comparsions.", wraplength=565)
outputTabRulesLabel.grid(row=6, column=0, padx=5, pady=(5, 0))

openOutputRulesLabel = Label(howtoFrame, font=(defaultFont, fontSizeRegular),
                             text="7. If you wish to see the actual output file, press Open at the bottom of the page and the file will open.", wraplength=565)
openOutputRulesLabel.grid(row=7, column=0, padx=5, pady=(5, 0))

restartRulesLabel = Label(howtoFrame, font=(defaultFont, fontSizeRegular),
                          text="8. Once you are done, if you wish to restart the tool, just go back to the Compare Page and redo the steps.", wraplength=565)
restartRulesLabel.grid(row=8, column=0, padx=5, pady=(5, 0))

creditsRulesLabel = Label(howtoFrame, font=(defaultFont, fontSizeRegular),
                          text="Thank You for using the tool.", wraplength=565)
creditsRulesLabel.place(relx=0.5, rely=1, anchor=S)

root.mainloop()
