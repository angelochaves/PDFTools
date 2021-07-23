import os
import PyPDF2
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import font
from tkinter import messagebox
from tkinter.font import BOLD, Font

def searchFile():
    global fileName
    fileName = filedialog.askopenfile(initialdir = "/",title = "Select file",filetypes = (("PDF files","*.pdf"),("all files","*.*")))
    if fileName == None:
        return
    ttk.Label(tab1, text=str(fileName.name), justify=CENTER, font=('Helvetica', 10, BOLD), wraplength=300).grid(column = 0, row = 2, padx = 0, pady = 5, columnspan=3)
    ttk.Button(tab1, text="Extract", command=extract).grid(column = 1, row = 6, padx = 30, pady = 10)


def showTargetPage():
    if pageOptions.get() == 999:
        numPages.set('')
        numPages.set('########')
    if pageOptions.get() < 999:
        numPages.set('')


def extract():
    pdf = open(fileName.name, 'rb')
    reader = PyPDF2.PdfFileReader(pdf)
    pages = reader.getNumPages()

    outFile = open("output.txt", "a")


    if pageOptions.get() == 1 and int(numPages.get()) <= (pages - 1) and numPages.get().isdigit():
        readPage = reader.getPage(int(numPages.get()))
        outFile.write(readPage.extractText())
        pdf.close()
        outFile.close()
        return

    elif pageOptions.get() == 50 and int(numPages.get()) <= (pages - 1) and numPages.get().isdigit():
        pages = int(numPages.get()) + 1

    elif pageOptions.get() != 999:        
        tk.messagebox.showinfo("Something's wrong...", "Please inform a valid page")
        outFile.close()
        os.remove("output.txt")
        return

    for page in range(0, pages):
        readPage = reader.getPage(page)
        outFile.write(readPage.extractText())
        outFile.write("\n\n")
    pdf.close()
    outFile.close()
    

def addFile():
    addedFile = filedialog.askopenfile(initialdir = "/", title = "Select file", filetypes = (("PDF files","*.pdf"), ("all files","*.*")))
    if addedFile == None:
        return
    listFiles.insert(END, str(addedFile.name))


def delFile():
    curSel = listFiles.curselection()
    if curSel == ():
        messagebox.showinfo("Something's wrong...", "Please select a file to delete")
        return
    listFiles.delete(curSel[0])


def upFile():
    curSel = listFiles.curselection()
    if curSel == ():
        messagebox.showinfo("Something's wrong...", "Please select a file to go up 1 position")
        return
    if curSel[0] > 0:
        listFiles.insert(curSel[0]-1, listFiles.get(curSel[0]))
        listFiles.delete(curSel[0]+1)
        listFiles.selection_set(curSel[0]-1)

    
def downFile():
    curSel = listFiles.curselection()
    if curSel == ():
        messagebox.showinfo("Something's wrong...", "Please select a file to go down 1 position")
        return
    if curSel[0] < (listFiles.size() - 1):
        listFiles.insert(curSel[0]+2, listFiles.get(curSel[0]))
        listFiles.delete(curSel[0])
        listFiles.selection_set(curSel[0]+1)


def appender():
    appender = PyPDF2.PdfFileMerger(strict=True)
    if listFiles.size() <= 0:
        messagebox.showinfo("Something's wrong...", "Please add files to merge/append")
        return
    for i in range(0, listFiles.size()):
        doc = PyPDF2.PdfFileReader(open(listFiles.get(i), 'rb'))
        appender.append(doc)
    appender.write('output.pdf')
    messagebox.showinfo("Sucess!", "Files merged sucessfully!")
    for i in range(listFiles.size()-1, -1, -1):
        listFiles.delete(i)


def composeAdd():
    global fileToCompose
    fileToCompose = filedialog.askopenfile(initialdir = "/",title = "Select file",filetypes = (("PDF files","*.pdf"),("all files","*.*")))
    if fileToCompose == None:
        return
    ############# ESTA SOBREPONDO O TEXTO QUANDO CARREGA OUTRO ARQUIVO NA SEQUENCIA
    ttk.Label(tab3, text=str(fileToCompose.name), justify=CENTER, font=('Helvetica', 10, BOLD), wraplength=300).grid(column = 0, row = 1, padx = 10, pady = 5, columnspan=3)


def addPageRange():
    if (int(singlePage.get()) < 0) or (int(singlePage.get()) > (int(PyPDF2.PdfFileReader(open(fileToCompose.name, 'rb')).getNumPages()) - 1)):
        messagebox.showinfo("Something's wrong...", "Page(s) out of range")
        return
    listPages.insert(END, str(singlePage.get()))
    singlePageEntry.delete(0, END)


def addRange():
    
    if ((int(rangePageStart.get()) or int(rangePageEnd.get())) < 0) or ((int(rangePageStart.get()) or int(rangePageEnd.get())) > (int(PyPDF2.PdfFileReader(open(fileToCompose.name, 'rb')).getNumPages()) - 1)):
        messagebox.showinfo("Something's wrong...", "Page(s) out of range")
        return
    if int(rangePageStart.get()) < int(rangePageEnd.get()):
        for i in range(int(rangePageStart.get()), int(rangePageEnd.get()) + 1):
            listPages.insert(END, str(i))
    if int(rangePageStart.get()) > int(rangePageEnd.get()):
        for i in range(int(rangePageStart.get()), int(rangePageEnd.get()) - 1, -1):
            listPages.insert(END, str(i))
    if int(rangePageStart.get()) == int(rangePageEnd.get()):
        listPages.insert(END, str(rangePageStart.get()))
    rangePageStartEntry.delete(0, END)
    rangePageEndEntry.delete(0, END)


def deleteRange():
    curSelRange = listPages.curselection()
    if curSelRange == ():
        messagebox.showinfo("Something's wrong...", "Please select a page/range to delete")
        return
    listPages.delete(curSelRange[0])


def clearRange():
    for i in range(listPages.size()-1, -1, -1):
        listPages.delete(i)


def upRange():
    curSelRange = listPages.curselection()
    if curSelRange == ():
        messagebox.showinfo("Something's wrong...", "Please select a page/range to go up 1 position")
        return
    if curSelRange[0] > 0:
        listPages.insert(curSelRange[0]-1, listPages.get(curSelRange[0]))
        listPages.delete(curSelRange[0]+1)
        listPages.selection_set(curSelRange[0]-1)


def downRange():
    curSelRange = listPages.curselection()
    if curSelRange == ():
        messagebox.showinfo("Something's wrong...", "Please select a page/range to go down 1 position")
        return
    if curSelRange[0] < (listPages.size() - 1):
        listPages.insert(curSelRange[0]+2, listPages.get(curSelRange[0]))
        listPages.delete(curSelRange[0])
        listPages.selection_set(curSelRange[0]+1)


def sliceMerge():
    reader = PyPDF2.PdfFileReader(open(fileToCompose.name, 'rb'))
    writer = PyPDF2.PdfFileWriter()
    for i in range(0, listPages.size()):
        writer.addPage(reader.getPage(int(listPages.get(i))))
        print(listPages.get(i))
    writer.write(open('output.pdf', 'wb'))


root = tk.Tk()

root.title("PDF Tools")
tabControl = ttk.Notebook(root)

tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tab3 = ttk.Frame(tabControl)

tabControl.add(tab1, text='Extract Text')
tabControl.add(tab2, text='Append')
tabControl.add(tab3, text='Slice/Merge')

tabControl.pack(expand=1, fill="both")

# TAB 1 - EXTRACT
numPages = tk.StringVar(tab1)
pageOptions = tk.IntVar(tab1)

ttk.Label(tab1, text="Selected File", font=('Helvetica', 12, BOLD), width=25, anchor=CENTER).grid(column = 1, row = 1, padx = 0, pady = 5)
ttk.Radiobutton(tab1, text="All pages", variable=pageOptions, value=999, command=showTargetPage, width=15).grid(column = 0, row = 3, padx = 10, pady = 5)
ttk.Radiobutton(tab1, text="Up to page...", variable=pageOptions, value=50, command=showTargetPage, width=15).grid(column = 1, row = 3, padx = 10, pady = 5)
ttk.Radiobutton(tab1, text="Just page...", variable=pageOptions, value=1, command=showTargetPage, width=15).grid(column = 2, row = 3, padx = 10, pady = 5)
ttk.Label(tab1, text="Target Page", width=25, anchor=CENTER).grid(column = 0, row = 4, padx = 0, pady = 5, columnspan=3)
ttk.Entry(tab1, textvariable=numPages, width=10).grid(column = 0, row = 5, padx = 0, pady = 5, columnspan=3)
ttk.Button(tab1, text="Select File", command=searchFile, width=25).grid(column = 0, row = 0, padx = 0, pady = 10, columnspan=3)

# TAB 2 - MERGE
listFiles = tk.Listbox(tab2, width=60)
listFiles.pack()
#scroll = tk.Scrollbar(tab2)
#scrollbar.grid(column = 5, row = 0, padx = 0, pady = 2, rowspan=7, sticky="sn")
#listPages.config(yscrollcommand = scrollbar.set)
#scrollbar.config(command = listPages.yview)

addBtn = ttk.Button(tab2, text="Add File", command=addFile)
addBtn.pack(side=LEFT, padx=20, pady=10)
delBtn = ttk.Button(tab2, text="Delete File", command=delFile)
delBtn.pack(side=LEFT, padx=5, pady=10)
upBtn = ttk.Button(tab2, text="Go up", command=upFile)
upBtn.pack(side=LEFT, padx=5, pady=10)
downBtn = ttk.Button(tab2, text="Go down", command=downFile)
downBtn.pack(side=LEFT, padx=5, pady=10)
composeBtn = ttk.Button(tab2, text="Append", command=appender)
composeBtn.pack(side=LEFT, padx=5, pady=10)

# TAB 3 - SLICE
singlePage = tk.StringVar(tab3)
rangePageStart = tk.StringVar(tab3)
rangePageEnd = tk.StringVar(tab3)

ttk.Button(tab3, text="Select File", command=composeAdd, width=25).grid(column = 0, row = 0, padx = 10, pady = 5, columnspan=3)
listPages = tk.Listbox(tab3, width=10, height=12, justify=CENTER)
listPages.grid(column = 4, row = 0, padx = 10, pady = 2, rowspan=7)
scrollbar = tk.Scrollbar(tab3)
scrollbar.grid(column = 5, row = 0, padx = 0, pady = 2, rowspan=7, sticky="sn")
listPages.config(yscrollcommand = scrollbar.set)
scrollbar.config(command = listPages.yview)

ttk.Button(tab3, text="Delete", command=deleteRange, width=15).grid(column = 6, row = 0, padx = 10, pady = 3)
ttk.Button(tab3, text="Clear all", command=clearRange, width=15).grid(column = 6, row = 1, padx = 10, pady = 3)
ttk.Button(tab3, text="Go up", command=upRange, width=15).grid(column = 6, row = 2, padx = 10, pady = 3)
ttk.Button(tab3, text="Go down", command=downRange, width=15).grid(column = 6, row = 3, padx = 10, pady = 3)
ttk.Button(tab3, text="Slice/Merge", command=sliceMerge, width=15).grid(column = 6, row = 4, padx = 10, pady = 3)

ttk.Separator(tab3, orient='horizontal').grid(column = 0, row = 2, padx = 5, pady = 1, columnspan=3, sticky="ew")
ttk.Label(tab3, text="Page", width=10, anchor=CENTER).grid(column = 0, row = 3, padx = 5, pady = 1)
singlePageEntry = ttk.Entry(tab3, textvariable=singlePage, width=5)
singlePageEntry.grid(column = 1, row = 3, padx = 5, pady = 1)
ttk.Button(tab3, text="Add", command=addPageRange, width=10).grid(column = 2, row = 3, padx = 5, pady = 1)

ttk.Separator(tab3, orient='horizontal').grid(column = 0, row = 4, padx = 5, pady = 1, columnspan=3, sticky="ew")
ttk.Label(tab3, text="From page", width=15, anchor=CENTER).grid(column = 0, row = 5, padx = 5, pady = 1)
ttk.Label(tab3, text="To", width=10, anchor=CENTER).grid(column = 0, row = 6, padx = 5, pady = 1)
ttk.Label(tab3, text="", width=10, anchor=CENTER).grid(column = 0, row = 7, padx = 5, pady = 1)
rangePageStartEntry = ttk.Entry(tab3, textvariable=rangePageStart, width=5).grid(column = 1, row = 5, padx = 5, pady = 1)
rangePageEndEntry = ttk.Entry(tab3, textvariable=rangePageEnd, width=5).grid(column = 1, row = 6, padx = 5, pady = 1)
ttk.Button(tab3, text="Add", command=addRange, width=10).grid(column = 2, row = 5, padx = 5, pady = 1, rowspan=2)

root.mainloop()