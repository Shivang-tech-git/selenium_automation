
from msedge.selenium_tools import Edge, EdgeOptions
from tkinter import messagebox
from tkinter import filedialog
# from FramesEventCoding import cat_event_coding
# from FramesEventCoding import pcs_event_coding
from FramesEventCoding import main_dataframe
from pathlib import Path
from datetime import datetime
import os, getpass
import pandas
import tkinter as tk
import xlwings
import traceback
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
    Edge_Options.add_argument("headless")
    Edge_Options.add_argument("disable-gpu")
    try:
        driver = Edge(options=Edge_Options)
    except:
        messagebox.showinfo('Frames event coding tool',
                            'Please close Edge and then run the tool.')
    driver.set_page_load_timeout(60)
    print('Running Edgedriver....')
    driver.get('https://theframe.catlin.com/?windowId=1')


def allocation_file():
    global filename
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select allocation File")

# def run_event(sht_name):
#     try:
#         # if output file checkbox is not selected
#         if output_selected.get() == 0:
#             # Check if allocation file exists
#             try:
#                 sht = xlwings.Book(filename).sheets[sht_name]
#             except:
#                 messagebox.showinfo('Frames event coding tool',
#                                 'Please select allocation file and then run the tool.')
#                 return
#             # work with allocation file
#             open_edge()
#             if sht_name == 'Property Claims':
#                 pcs_df = sht.range('A1').options(pandas.DataFrame, header=1, index=False, expand='table').value
#                 pcs_df['Status'] = ''
#                 pcs_df.to_csv('PCS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
#                 blank_pcs_df = pcs_df
#                 pcs_event_coding.pcs_event_coding(driver=driver, pcs_df=pcs_df, blank_pcs_df=blank_pcs_df)
#                 messagebox.showinfo('Frames event coding tool',
#                                     'PCS Allocation Completed, Please check the PCS output file to see status.')
#             elif sht_name == 'CAT':
#                 cat_df = sht.range('A1').options(pandas.DataFrame, header=1, index=False, expand='table').value
#                 cat_df['Status'] = ''
#                 cat_df.to_csv('CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')), index=False)
#                 blank_cat_df = cat_df
#                 cat_event_coding.cat_event_coding(driver=driver, cat_df=cat_df, blank_cat_df=blank_cat_df)
#                 messagebox.showinfo('Frames event coding tool',
#                                     'CAT Allocation Completed, Please check the CAT output file to see status.')
#
#         # if output file checkbox is selected
#         elif output_selected.get() == 1:
#             open_edge()
#             # work with output file
#             if sht_name == 'Property Claims':
#                 pcs_df = pandas.read_csv('PCS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')))
#                 blank_pcs_df = pcs_df.loc[pcs_df['Status'].isnull()]
#                 pcs_event_coding.pcs_event_coding(driver=driver, pcs_df=pcs_df, blank_pcs_df=blank_pcs_df)
#                 messagebox.showinfo('Frames event coding tool',
#                                     'PCS Allocation Completed, Please check the PCS output file to see status.')
#
#             elif sht_name == 'CAT':
#                 cat_df = pandas.read_csv('CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')))
#                 blank_cat_df = cat_df.loc[cat_df['Status'].isnull()]
#                 cat_event_coding.cat_event_coding(driver=driver, cat_df=cat_df, blank_cat_df=blank_cat_df)
#                 messagebox.showinfo('Frames event coding tool',
#                                     'CAT Allocation Completed, Please check the CAT output file to see status.')
#
#     except Exception as msg:
#         driver.quit()
#         messagebox.showerror('Frames event coding tool', traceback.format_exc())
#
# def run_refresh(event):
#     try:
#         if event == 'cat':
#             cat_df = pandas.read_csv('CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')))
#         elif event == 'pcs':
#             pcs_df = pandas.read_csv('PCS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')))
#     except:
#         messagebox.showerror('Frames event coding tool',
#                             'Error: Ouptut file does not exist.')
#         return
#     open_edge()
#     try:
#         if event == 'cat':
#             ue_cat_df = cat_df.loc[cat_df['Status'] == 'Undefined Error']
#             if ue_cat_df.shape[0] > 0:
#                 cat_event_coding.cat_event_refresh(driver=driver, cat_df=cat_df, ue_cat_df=ue_cat_df)
#                 messagebox.showinfo('Frames event coding tool',
#                                     'CAT Allocation Completed, Please check the CAT output file to see status.')
#             else:
#                 messagebox.showinfo('Frames event coding tool',
#                                     'No BPR found with undefined error status in CAT output file.')
#
#         elif event == 'pcs':
#             ue_pcs_df = pcs_df.loc[pcs_df['Status'] == 'Undefined Error']
#             if ue_pcs_df.shape[0] > 0:
#                 pcs_event_coding.pcs_event_refresh(driver=driver, pcs_df=pcs_df,ue_pcs_df=ue_pcs_df)
#                 messagebox.showinfo('Frames event coding tool',
#                                 'PCS Allocation Completed, Please check the PCS output file to see status.')
#             else:
#                 messagebox.showinfo('Frames event coding tool',
#                                     'No BPR found with undefined error status in PCS output file.')
#
#     except Exception as msg:
#         driver.quit()
#         messagebox.showerror('Frames event coding tool', traceback.format_exc())

def run_event(sht_name):
    if sht_name == 'Property Claims':
        selected_event = {'event':'PCS',
                          'output_filename':'PCS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                          'message':'PCS Allocation Completed, Please check the PCS output file to see status.'}
    elif sht_name == 'CAT':
        selected_event = {'event':'CAT',
                          'output_filename':'CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                          'message':'CAT Allocation Completed, Please check the CAT output file to see status.'}
    elif sht_name == 'Various':
        selected_event = {'event':'VARS',
                          'output_filename':'VARS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                          'message':'VARS Allocation Completed, Please check the VARS output file to see status.'}

    try:
        # if output file checkbox is not selected
        if output_selected.get() == 0:
            # Check if allocation file exists
            try:
                sht = xlwings.Book(filename).sheets[sht_name]
            except:
                messagebox.showinfo('Frames event coding tool',
                                'Please select allocation file and then run the tool.')
                return
            # work with allocation file
            open_edge()
            main_df = sht.range('A1').options(pandas.DataFrame, header=1, index=False, expand='table').value
            main_df['Status'] = ''
            main_df.to_csv(selected_event['output_filename'], index=False)
            blank_df = main_df
            main_dataframe.event_coding(selected_event['event'], driver=driver, main_df=main_df, blank_df=blank_df)
            messagebox.showinfo('Frames event coding tool',
                                selected_event['message'])

        # if output file checkbox is selected
        elif output_selected.get() == 1:
            open_edge()
            # work with output file
            main_df = pandas.read_csv(selected_event['output_filename'])
            blank_df = main_df.loc[main_df['Status'].isnull()]
            main_dataframe.event_coding(selected_event['event'], driver=driver, main_df=main_df, blank_df=blank_df)
            messagebox.showinfo('Frames event coding tool',
                                selected_event['message'])

    except Exception as msg:
        driver.quit()
        messagebox.showerror('Frames event coding tool', traceback.format_exc())

def run_refresh(event):
    if event == 'pcs':
        selected_event = {'event':'PCS',
                          'output_filename':'PCS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                          'message':'PCS Allocation Completed, Please check the PCS output file to see status.'}
    elif event == 'cat':
        selected_event = {'event':'CAT',
                          'output_filename':'CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                          'message':'CAT Allocation Completed, Please check the CAT output file to see status.'}
    elif event == 'vars':
        selected_event = {'event':'VARS',
                          'output_filename':'VARS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                          'message':'VARS Allocation Completed, Please check the VARS output file to see status.'}
    try:
        main_df = pandas.read_csv(selected_event['output_filename'])
    except:
        messagebox.showerror('Frames event coding tool',
                            'Error: Ouptut file does not exist.')
        return
    open_edge()
    try:
        ue_df = main_df.loc[main_df['Status'] == 'Undefined Error']
        if ue_df.shape[0] > 0:
            main_dataframe.event_refresh(event=selected_event['event'],driver=driver, main_df=main_df, ue_df=ue_df)
            messagebox.showinfo('Frames event coding tool',
                                selected_event['message'])
        else:
            messagebox.showinfo('Frames event coding tool',
                                'No BPR found with undefined error status in output file.')

    except Exception as msg:
        driver.quit()
        messagebox.showerror('Frames event coding tool', traceback.format_exc())

if __name__ == '__main__':
    window = tk.Tk()
    output_selected = tk.IntVar()
    window.title('AXA XL')
    # ------------------------------- Title ----------------------------
    greeting = tk.Label(text="Frames event coding tool", font=('Arial',12), fg='yellow', bg='green', width=60)
    frm_open = tk.Frame(master=window, width=60)
    # ------------------------------ Browse allocation file --------------------------
    lbl_open = tk.Label(master=frm_open, text='Select allocation file',font=('Arial',10),
                        width=30, borderwidth=2, relief="ridge").grid(row=0, column=0, pady=5)
    btn_open = tk.Button(master=frm_open, text='Browse',font=('Arial',10), width=15, bg='lavender',
                         command=allocation_file).grid(row=0,column=1,pady=5)
    use_output_file = tk.Checkbutton(master=frm_open, text='Use Output file',font=('Arial',10), width=15, bg='lavender',
                         variable=output_selected).grid(row=0,column=2,pady=5)
    # ------------------------------ CAT event code ----------------------------------
    frm_extract = tk.Frame(master=window, width=60)
    lbl_cat = tk.Label(master=frm_extract, text='CAT event coding',font=('Arial',10),
                           width=30, borderwidth=2, relief="ridge").grid(row=1, column=0, pady=5)
    btn_cat = tk.Button(master=frm_extract,font=('Arial',10), text='Run', width=15, bg='misty rose',
                        command=lambda : run_event('CAT')).grid(row=1, column=1, pady=5)
    btn_cat_ref = tk.Button(master=frm_extract, font=('Arial', 10), text='Refresh', width=15, bg='misty rose',
                        command=lambda : run_refresh('cat')).grid(row=1, column=2, pady=5)
    # ----------------------------- PCS event code -----------------------------------
    lbl_pcs = tk.Label(master=frm_extract, text='PCS event coding',font=('Arial',10), width=30, borderwidth=2,
                          relief="ridge").grid(row=2, column=0, pady=5)
    btn_pcs = tk.Button(master=frm_extract, text='Run',font=('Arial',10), width=15, bg='floral white',
                        command=lambda : run_event('Property Claims')).grid(row=2,column=1,pady=5)
    btn_pcs_refresh = tk.Button(master=frm_extract, text='Refresh', font=('Arial', 10), width=15, bg='floral white',
                        command=lambda : run_refresh('pcs')).grid(row=2,column=2, pady=5)
    # ----------------------------- vars event code -----------------------------------
    lbl_vars = tk.Label(master=frm_extract, text='Various event coding',font=('Arial',10), width=30, borderwidth=2,
                          relief="ridge").grid(row=3, column=0, pady=5)
    btn_vars = tk.Button(master=frm_extract, text='Run',font=('Arial',10), width=15, bg='floral white',
                        command=lambda : run_event('Various')).grid(row=3,column=1,pady=5)
    btn_vars_refresh = tk.Button(master=frm_extract, text='Refresh', font=('Arial', 10), width=15, bg='floral white',
                        command=lambda : run_refresh('vars')).grid(row=3,column=2, pady=5)
    # lbl_status = tk.Label(text='status...', width=50)
    greeting.pack()
    frm_open.pack(fill=tk.BOTH)
    frm_extract.pack(fill=tk.BOTH)
    # lbl_status.pack(pady=5, side=tk.BOTTOM)
    window.mainloop()
