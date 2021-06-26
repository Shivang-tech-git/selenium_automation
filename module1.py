# developed by scriptdeveloper with love for Geman brothers and sisters.
import pyodbc
import tkinter as tk
import sys, os
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
def open_chrome():
    global driver
    if getattr(sys, 'frozen', False): 
        chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
        driver = webdriver.Chrome(chromedriver_path)
    else:
        driver = webdriver.Chrome()
    driver.get("https://www.jameda.de")
    lbl_status.config(text="Fertig, bitte suchen Sie nach einem Arzt und fahren Sie mit Schritt 2 fort ....")
def extract_data():
    while True:
        try:
            next_page = driver.find_element_by_xpath('//div[@class = "sc-14z3gaz-1 dgmOEV"]')
            wait = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//div[@class = "sc-14z3gaz-1 dgmOEV"]')))
            if next_page.get_attribute('innerText') == "Mehr anzeigen":
                driver.execute_script("arguments[0].click();", driver.find_element_by_xpath('//div[@class = "sc-14z3gaz-1 dgmOEV"]/a'))
                continue
            if next_page.get_attribute('innerText') == "Nach oben":
                break
        except:
            break
    names_list = []
    address_list = []
    total_results = driver.find_elements_by_xpath('//ol[@id = "result-list"]/li')
    if total_results == 0:
        messagebox.showinfo(message="No results found.")
        return None
    for result in range(2,total_results.__len__() + 1):
        try:
            names_list.append(driver.find_element_by_xpath('//li[{0}]/descendant::h2[@class = "sc-1lwaqfx-2 gmiMet"][2]'.format(result)).get_attribute('innerText'))
        except:names_list.append('None')
        try:
            address_list.append(driver.find_element_by_xpath('//li[{0}]/descendant::div[@class = "sc-1fj7nwi-3 XCBsd"]'.format(result)).get_attribute('innerText'))
        except:address_list.append('None')
        db_path = os.path.abspath(os.getcwd()) + '\Jameda.accdb'
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';')
        cursor = conn.cursor()
    for item in range(0,names_list.__len__()):
         cursor.execute('''
                    INSERT INTO Table1 (Name, Address)
                    VALUES('{0}', '{1}')

                  '''.format(names_list[item].strip(), address_list[item].strip()))
         conn.commit()
    lbl_status.config(text="{0} Datensätze extrahiert, fertig, bitte suchen Sie nach Doktor und fahren Sie mit Schritt 2 fort .....".format(names_list.__len__()))
    messagebox.showinfo(message="{0} Anzahl der Datensätze.".format(names_list.__len__()))
def delete_data():
    db_path = os.path.abspath(os.getcwd()) + '\Jameda.accdb'
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';')
    cursor = conn.cursor()
    cursor.execute('''
                    DELETE FROM Table1
                   ''')
    conn.commit()
    messagebox.showinfo(message="Getan")
window = tk.Tk()
window.title('Jameda Search Application')
greeting = tk.Label(text="Jameda Search Application",fg='yellow',bg='green',width=70)
frm_open = tk.Frame(master=window,width=100)
lbl_open = tk.Label(master=frm_open,text='Schritt: 1 Öffnen Sie die Jemeda-Webseite und suchen Sie den Arzt',width=50, borderwidth=2, relief="ridge").grid(row=0, column=0, pady=5)
btn_open = tk.Button(master=frm_open,text='Run Schritt 1',width=20,bg='orange', command=open_chrome).grid(row=0, column=1, pady=5)
frm_extract = tk.Frame(master=window,width=100)
lbl_extract = tk.Label(master=frm_extract,text='Schritt: 2 Alle Adressen extrahieren und in MS Access einfügen',width=50, borderwidth=2, relief="ridge").grid(row=1, column=0, pady=5)
btn_extract = tk.Button(master=frm_extract,text='Run Schritt 2',width=20,bg='orange', command=extract_data).grid(row=1, column=1, pady=5)
lbl_delete = tk.Label(master=frm_extract,text='Daten vorher löschen.',width=50, borderwidth=2, relief="ridge").grid(row=2, column=0, pady=5)
btn_delete = tk.Button(master=frm_extract,text='löschen',width=20,bg='orange', command=delete_data).grid(row=2, column=1, pady=5)
lbl_status = tk.Label(text='Statusleiste...',width=50)
greeting.pack()
frm_open.pack(fill=tk.BOTH)
frm_extract.pack(fill=tk.BOTH)
lbl_status.pack(pady=5, side=tk.BOTTOM)
window.mainloop()

