
import glob
import pandas as pd
from tkinter import Frame, Label, StringVar, Entry, Button, filedialog
from tkinter import RAISED, LEFT, X
import tkinter.messagebox as Msb


class Run():
    def __init__(self, frame):
        self.parent = frame
        self.stopValue = False
        self.Run()
    def Run(self):
        self.parent.title("Excel Tool")

        self.parent.resizable(0, 0)

        self.frame1 = Frame(self.parent, relief=RAISED, borderwidth=1)
        self.frame1.pack(fill=X, padx=0, pady=0)

        self.inputFolderLabel = Label(
            self.frame1, text="Input Folder:", anchor='w', width=10, fg="Green")
        self.inputFolderLabel.pack(side=LEFT, padx=5, pady=0)

        self.folderPath = StringVar(self.parent, value='')
        self.inputFolderPath= Entry(self.frame1, width=30,
                               textvariable=self.folderPath)
        self.inputFolderPath.pack(side=LEFT, padx=5, pady=5)

        browseButton = Button(self.frame1, text="Browse", width= 10, command=self.BrowseButton)
        browseButton.pack(side=LEFT, padx=5, pady=5)

        self.frame3 = Frame(self.parent, relief=RAISED, borderwidth=1)
        self.frame3.pack(fill=X, padx=0, pady=0)

        self.outputFolderLabel = Label(
            self.frame3, text="Out Folder:", anchor='w', width=10, fg="Green")
        self.outputFolderLabel.pack(side=LEFT, padx=5, pady=0)

        self.outFolderPath = StringVar(self.parent, value='')
        self.inputOutFolderPath= Entry(self.frame3, width=30,
                               textvariable=self.outFolderPath)
        self.inputOutFolderPath.pack(side=LEFT, padx=5, pady=5)

        browseOutButton = Button(self.frame3, text="Browse", width= 10, command=self.BrowseOutButton)
        browseOutButton.pack(side=LEFT, padx=5, pady=5)

        self.frame2 = Frame(self.parent, relief=RAISED, borderwidth=1)
        self.frame2.pack(fill=X, padx=0, pady=0)

        self.keywordLabel = Label(
            self.frame2, text="Keyword:", anchor='w', width= 10, fg="Green")
        self.keywordLabel.pack(side=LEFT, padx=5, pady=0)

        self.keyword = StringVar(self.parent, value='Database')
        self.inputKeyword= Entry(self.frame2, width=30,
                               textvariable=self.keyword)
        self.inputKeyword.pack(side=LEFT, padx=5, pady=5)

        browseButton = Button(self.frame2, text="Start", width= 10, command=self.Combine)
        browseButton.pack(side=LEFT, padx=5, pady=5)

    def BrowseButton(self):
        # Allow user to select a directory and store it in global var
        # called folder_path
        filename = filedialog.askdirectory()
        self.folderPath.set(filename)
        print(filename)
    
    def BrowseOutButton(self):
        # Allow user to select a directory and store it in global var
        # called folder_path
        filename = filedialog.askdirectory()
        self.outFolderPath.set(filename)
        print(filename)
    
    def RedFile(self):
        keyword = self.keyword.get()
        localFile = self.folderPath.get().replace("/","\\")
        file_list = glob.glob(localFile + "**/**/**/*"+keyword+".xlsx")
        return file_list

    def Combine(self):
        output = self.outFolderPath.get()
        files = self.RedFile()
        dfCombined = pd.read_excel(files[0])
        files = files[1:]
        for file in files:
            dfCombined = dfCombined.merge(pd.read_excel(file), how="right")
        print(dfCombined.shape)
        dfCombined.to_excel(output + '/data.xlsx', index=False)
        mess = "Completed merging %d files with '%s' in the name" %(len(files)+1, self.keyword.get())
        Msb.showinfo(title='Notification', message=mess)
