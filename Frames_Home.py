
from msedge.selenium_tools import Edge, EdgeOptions
from tkinter import messagebox
from tkinter import filedialog
from FramesEventCoding import cat_event_coding
from FramesEventCoding import pcs_event_coding
import os, getpass
import pandas
import tkinter as tk
import xlwings
#---------- XPath ------ one element ----------------

def open_edge():
    global driver
    edgedriver_path = os.path.join(os.path.abspath(os.getcwd()), "msedgedriver.exe")
    Edge_Options = EdgeOptions()
    Edge_Options.add_experimental_option("excludeSwitches", ["enable-automation"])
    Edge_Options.add_experimental_option('useAutomationExtension', False)
    Edge_Options.use_chromium = True
    Edge_Options.add_argument('no-sandbox')
    Edge_Options.add_argument("user-data-dir=C:/Users/" + getpass.getuser() + "/AppData/Local/Microsoft/Edge/User Data")
    try:
        driver = Edge(options=Edge_Options)
    except:
        messagebox.showinfo('Frames event coding tool',
                            'Please close Edge and then run the tool.')
    driver.set_page_load_timeout(60)
    print('Running Edgedriver....')
    driver.get('https://theframe.catlin.com/?windowId=1')
    driver.maximize_window()


def allocation_file():
    global filename
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select allocation File")

def run_cat():
    try:
        sht = xlwings.Book(filename).sheets['CAT']
    except:
        messagebox.showinfo('Frames event coding tool',
                            'Please select allocation file and then run the tool.')
        return
    cat_df = sht.range('A1').options(pandas.DataFrame,
                                 header=1,
                                 index=False,
                                 expand='table').value
    open_edge()
    try:
        cat_event_coding.cat_event_coding(driver=driver,cat_df=cat_df)
        messagebox.showinfo('Frames event coding tool',
                            'PCS Allocation Completed, Please check the CAT output file to see status.')
    except:
        messagebox.showinfo('Frames event coding tool',
                            'Frames is not responding. Please run the tool for remaining BPR')

def run_pcs():
    try:
        sht = xlwings.Book(filename).sheets['Property Claims']
    except:
        messagebox.showinfo('Frames event coding tool',
                            'Please select allocation file and then run the tool.')
        return
    pcs_df = sht.range('A1').options(pandas.DataFrame,
                                     header=1,
                                     index=False,
                                     expand='table').value
    open_edge()
    try:
        pcs_event_coding.pcs_event_coding(driver=driver,pcs_df=pcs_df)
        messagebox.showinfo('Frames event coding tool', 'PCS Allocation Completed, Please check the PCS output file to see status.')
    except:
        messagebox.showinfo('Frames event coding tool',
                            'Frames is not responding. Please run the tool for remaining BPR')

if __name__ == '__main__':
    window = tk.Tk()
    window.title('AXA XL')
    greeting = tk.Label(text="Frames event coding tool", font=('Arial',12), fg='yellow', bg='green', width=65)
    frm_open = tk.Frame(master=window, width=70)
    lbl_open = tk.Label(master=frm_open, text='Select allocation file',font=('Arial',10),
                        width=50, borderwidth=2, relief="ridge").grid(row=0, column=0, pady=5)
    btn_open = tk.Button(master=frm_open, text='Browse',font=('Arial',10), width=20, bg='lavender', command=allocation_file).grid(row=0,
                                                                                                column=1,pady=5)
    frm_extract = tk.Frame(master=window, width=70)
    lbl_cat = tk.Label(master=frm_extract, text='CAT event coding',font=('Arial',10),
                           width=50, borderwidth=2, relief="ridge").grid(row=1, column=0, pady=5)
    btn_cat = tk.Button(master=frm_extract,font=('Arial',10), text='Run', width=20, bg='misty rose', command=run_cat).grid(
        row=1, column=1, pady=5)
    lbl_pcs = tk.Label(master=frm_extract, text='PCS event coding',font=('Arial',10), width=50, borderwidth=2,
                          relief="ridge").grid(row=2, column=0, pady=5)
    btn_pcs = tk.Button(master=frm_extract, text='Run',font=('Arial',10), width=20, bg='floral white', command=run_pcs).grid(row=2,
                                                                                                column=1,pady=5)
    # lbl_status = tk.Label(text='status...', width=50)
    greeting.pack()
    frm_open.pack(fill=tk.BOTH)
    frm_extract.pack(fill=tk.BOTH)
    # lbl_status.pack(pady=5, side=tk.BOTTOM)
    window.mainloop()
