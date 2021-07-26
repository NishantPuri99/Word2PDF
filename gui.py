from tkinter import *
from tkinter import filedialog
from Doc2Pdf import convert2pdf

root = Tk()
root.title('Convert2PDF | Convert your Docx files to PDFs')
root.minsize(690,100)
root.maxsize(820,250)
root.geometry("690x100")

inputLabelFieldHeader = Label(root,text="Use the button below to select your word document.")
inputLabelFieldHeader.grid(row=0,column=0,padx=10,pady=10)
inputFileNameString = ""

outputLabelFieldHeader = Label(root,text="Select Directory to Save the Converted PDF in.")
outputLabelFieldHeader.grid(row=0,column=2,padx=10,pady=10)
outputFileNameString = ""

def browseFilesForConversion():
    filePath = filedialog.askopenfilename()
    filePath = filePath.replace('/','\\')
    inputFileName = Label(root,text=f"Input File Path:\n{filePath}")
    inputFileName.grid(row=2,column=0,padx=10,pady=10)
    global inputFileNameString 
    inputFileNameString = filePath
    root.geometry("780x150")
    outputSelectButton['state'] = NORMAL

def converterPlaceHolder():
    callConverter()

inputSelectButton = Button(root,text="Select Word File",command=browseFilesForConversion)
inputSelectButton.grid(row=1,column=0,padx=10,pady=10)

convertButton = Button(root,text="Convert to PDF",command=converterPlaceHolder,state=DISABLED)
convertButton.grid(row=1,column=1,padx=10,pady=10)

def selectOutputPath():
    outputFilePath = filedialog.askdirectory()
    outputFilePath = outputFilePath.replace('/','\\')
    outputFileName = Label(root,text=f"Output File Path:\n{outputFilePath}")
    outputFileName.grid(row=2,column=2,padx=10,pady=10)
    global outputFileNameString 
    outputFileNameString = outputFilePath
    convertButton['state'] = NORMAL

outputSelectButton = Button(root,text="Select Output Directory",command=selectOutputPath,state=DISABLED)
outputSelectButton.grid(row=1,column=2,padx=10,pady=10)


def callConverter():
    print(inputFileNameString,outputFileNameString)
    convert2pdf(inputFileNameString,outputFileNameString)

root.mainloop()

# Check win32com
# Convert to exe
# push to github
