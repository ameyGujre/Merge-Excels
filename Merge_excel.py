from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os
import pandas as pd
import time


root = Tk()
root.title('Xyclone_v1')

windowWidth = root.winfo_reqwidth()
windowHeight = root.winfo_reqheight()
positionRight = int(root.winfo_screenwidth()/2 - 400/2)
positionDown = int(root.winfo_screenheight()/2 - 600/2)

root.geometry("+{}+{}".format(positionRight, positionDown))


root.maxsize(600,350)

### GLobal Functions
all_xl_files_path = list()
template_cols = list()

##Browse function
def import_sheets():
    root.filename = filedialog.askopenfilenames(title="Select Files", filetypes=(("Excel file", "*.xlsx"),("Excel file", "*.xls")))
    all_xl_files_path.append(root.filename)
    label_text = str(len(root.filename))+" File(s) Selected"
    label2.config(text=label_text, fg="#3CB043")

def import_temp():
    root.filename = filedialog.askopenfilename(title='Select template excel', filetypes=(("Excel file", "*.xlsx"),("Excel file", "*.xls")))
    temp_df = pd.read_excel(root.filename)
    template_cols.append(temp_df.columns.values)
    label4.config(text=("Template file: "+ str(os.path.basename(root.filename))), font=("Calibri 10 italic"))

def consolidate_data():

    df_list = []


    for i in range(0,len(all_xl_files_path[0])):
        temp_df = pd.read_excel(all_xl_files_path[0][i])

        ##Selecting all the required columns
        try:
            temp_df = temp_df[template_cols[0]]
            df_list.append(temp_df)
        except KeyError as E:
            filename_temp = os.path.basename(all_xl_files_path[0][i])
            messagebox.showerror('Column name mismatched', f'Error: {filename_temp} has incorrect column headings')
            root.destroy()

    final_df = pd.concat(df_list)

    label_export.config(text=("Total records consolidated : "+str(final_df.shape[0])), fg='#3CB043')

    
    return final_df


def export():


    df_main = consolidate_data()

    export_filename = filedialog.asksaveasfilename(initialfile='Untitled.xlsx', defaultextension=".xlsx", filetypes=[("Excel file","*.xlsx")])
    df_main.to_excel(export_filename, index=False)
    
    
    time.sleep(2.5)
    root.destroy()

'''Frame 1 for importing the workbooks'''
##Creating frame1
frame1 = LabelFrame(root, text='Import Data', height=300, padx=10, pady=10)
frame1.grid(row=0,column=0, sticky='ew', padx=10, pady=10)
root.grid_columnconfigure(0,weight=1)

label1 = Label(frame1, text='Please browse all the workbook files      (*xls/*xlsx)', font=("Calibri"))
label1.grid(row=0,column=0, padx=2, pady=5)
##Label when file is uploaded
label2 = Label(frame1, text="0 File(s) selected", font=("Calibri", 10), fg='#FF0000')
label2.grid(row=1,column=0, padx=2, pady=1)

b = Button(frame1, text='Browse', pady=5, command=import_sheets, width=15)
b.grid(row=0, column=1, padx=70)



'''Frame 2 for importing the column template'''
frame2 = LabelFrame(root, text='Select Template Excel', height=300, padx=10, pady=10)
frame2.grid(row=1,column=0, sticky='ew', padx=10, pady=10)
root.grid_columnconfigure(0,weight=1)

label3 = Label(frame2, text='Select the template excel to take the column reference', font=("Calibri"))
label3.grid(row=0,column=0, padx=2, pady=5)


b2 = Button(frame2, text='Browse', pady=3, command=import_temp, width=15)
b2.grid(row=0, column=1, padx=50)

label4 = Label(frame2, text=" ", font=("Calibri", 10))
label4.grid(row=1,column=0, padx=2, pady=1)




'''Frame 3 for importing the column template'''
frame3 = LabelFrame(root, text='Export', height=300, padx=10, pady=10)
frame3.grid(row=2,column=0, sticky='ew', padx=10, pady=10)
root.grid_columnconfigure(0,weight=1)


label5 = Label(frame3, text='Export the consolidated file in the required location     ', font=("Calibri"))
label5.grid(row=0,column=0, padx=2, pady=5)

b3 = Button(frame3, text='Export', pady=3, command=export, width=15)
b3.grid(row=0, column=1, padx=50)

label_export = Label(frame3, text=" ", font=("Calibri", 10))
label_export.grid(row=1,column=0)

root.mainloop()