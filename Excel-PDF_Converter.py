import tkinter 
from tkinter import filedialog as fd
from tkinter import messagebox as mb
from win32com import client
import win32api


def callback():
	name = fd.askopenfilename()
	global filename
	filename=name.replace('/','\\')
	if filename:
		Window.geometry("500x150")
		tkinter.Label(Window, text='File has been selected,now create click on create PDF').pack(fill=tkinter.X,padx=5, pady=5)
		tkinter.Button(Window, text='Create PDF', command=createpdf).pack(fill=None,padx=5, pady=5)

def createpdf():
	app = client.DispatchEx("Excel.Application")
	app.Interactive = False
	app.Visible = False
	Workbook = app.Workbooks.Open(filename)
	output_file = r'C:\sample_output.pdf'
	try:
		Workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
	except Exception as e:
		print("Failed to convert in PDF format.Please confirm environment meets all the requirements  and try again")
		print(str(e))
	finally:
		mb.showinfo('Excel to PDF Converter', 'The file has been converted succesfully : "C:\sample_output.pdf"')
		Workbook.Close()
		app.quit()
		app.Exit()

Window = tkinter.Tk()
Window.geometry("500x90")
Window.title('Excel to PDF Converter')
errmsg = 'Error!'
tkinter.Label(Window, text="Click on Browse for selecting the Excel file").pack(fill=tkinter.X,padx=5, pady=5)
tkinter.Button(Window, text='Browse',
       command=callback).pack(fill=None,padx=5, pady=5)

Window.mainloop()
