import PyPDF2
import numpy
import pandas
import openpyxl
import re
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.numbers import FORMAT_NUMBER_00
import xml.etree.ElementTree as ET

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

import argparse

def toDouble(s, decimal_separator, thousands_separator):
    try:
        if thousands_separator != '':
            s = s.replace(thousands_separator, '')
        parts = s.split(decimal_separator)
        if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
            raise ValueError("The string does not represent a floating point number.")
        return float(s.replace(decimal_separator, '.'))
    except Exception as e:
        raise ValueError("The string does not represent a floating point number.") from e
    
def contains_substring(s, substrings):
    return any(substring in s for substring in substrings)
    
#bis hier fertig

class Buchung:
    def __init__(self, datum, text, wert):
        self.datum = datum
        self.text = text
        self.wert = wert
        self.lines = []
        
    def addLine(self, line):
        self.lines.append(line)
        
    def __str__(self):
        res = str(self.datum) + "     " + self.text + ": " + str(self.wert) + "\n"
        for line in self.lines:
            res += "    " + line + "\n" 
        return res

class Auszug:
    def __init__(self, pdffile):
        reader = PyPDF2.PdfReader(pdffile)
        self.lines = []
        for page in reader.pages:
            linesp = page.extract_text().split("\n") 
            for line in linesp:
                lines = line.strip()
                lines = lines[:-6]
                lines = lines.strip()
                self.lines.append(lines)
                
    def checkForBuchungBegin(self, line):
        # Check if the string starts with a day and month
        match = re.match(r'^\s*\d{1,2}\.\s*\d{1,2}', line)
        if match:
            # Extract the day and month
            day, month = map(int, match.group().split('.'))
            # Check if the day and month are valid
            if 1 <= day <= 31 and 1 <= month <= 12:
                # Return True and the date
                return True, datetime(datetime.now().year, month, day)
        # Return False and None if the string cannot be interpreted as a date
        return False, None
     
    def buildBuchungen(self, ignoreLinesMarker):
        self.buchungen = []
        for line in self.lines:
            p = self.checkForBuchungBegin(line)
            if p[0] == True:
                linesp = line.split(" ")
                lastpart = linesp[-1]
                lastchar = lastpart[-1]
                
                if lastchar == "-":
                    minus = True
                else:
                    minus = False
                    
                if minus == True:
                    lastpart = lastpart[:-1]
                    
                v = toDouble(lastpart, ",", ".")
                if(minus == True):
                    v = v * -1
                
                buchung = Buchung(p[1], line, v)
                self.buchungen.append(buchung)
               
            else:
                if len(self.buchungen) > 0:
                    if contains_substring(line, ignoreLinesMarker) == False:
                        self.buchungen[-1].addLine(line)
                        
    def combineBuchungen(self, config):
        # Initialize a dictionary to hold the Konto names and their Buchungen
        self.konto_buchungen = {konto: [] for konto in config.keys()}
        self.konto_buchungen["cannot assign"] = []
        self.konto_buchungen["cannot assign uniquely"] = []

        # Iterate over the Buchungen
        for buchung in self.buchungen:
            matches = []

            # Check if the text or any of the lines of the Buchung contain a marker
            for konto, markers in config.items():
                if any(marker in buchung.text for marker in markers) or any(marker in line for marker in markers for line in buchung.lines):
                    matches.append(konto)

            # Assign the Buchung to the corresponding Konto
            if len(matches) == 0:
                self. konto_buchungen["cannot assign"].append(buchung)        
            if len(matches) == 1:
                self.konto_buchungen[matches[0]].append(buchung)
            elif len(matches) > 1:
                self.konto_buchungen["cannot assign uniquely"].append(buchung)
                
def select_file(file_var):
    file_var.set(filedialog.askopenfilename())

    
def parseConfig(configFile):
    # Parse the XML file
    tree = ET.parse(configFile)
    root = tree.getroot()

    # Initialize an empty dictionary to hold the Konto names and their Markers
    konto_markers = {}

    # Check and iterate over the Konto elements in the KontoMarker element
    kontoMarkerElement = root.find('KontoMarker')
    if kontoMarkerElement is None:
        raise ValueError("The XML configuration does not contain a 'KontoMarker' tag.")

    for konto in kontoMarkerElement:
        # Get the name of the Konto
        name = konto.get('Name')

        # Get the Marker elements of the Konto
        markers = [marker.text for marker in konto.findall('Marker')]

        # Add the name and markers to the dictionary
        konto_markers[name] = markers
        
    ignoreMarker = []
    ignoreMarkerElement = root.find('IgnoreMarker')
    if ignoreMarkerElement is not None:
        for marker in ignoreMarkerElement:
            ignoreMarker.append(marker.text)

    return konto_markers, ignoreMarker
 
class MyApplication:
    def __init__(self, root):
        self.root = root
        self.setup_ui()
        self.auszug = None  # This will hold the auszug object
        
    def setup_ui(self):       
        root = tk.Tk()
        root.geometry('300x250')  # Set the size of the window

        kontoauszugFile = tk.StringVar()
        configFile = tk.StringVar()

        button1 = tk.Button(root, text='Select Kontoauszug', command=lambda: select_file(kontoauszugFile))
        button1.pack(padx=20, pady=10, ipadx=10, ipady=10)  # Add margins and padding

        button2 = tk.Button(root, text='Select Configuration', command=lambda: select_file(configFile))
        button2.pack(padx=20, pady=10, ipadx=10, ipady=10)  # Add margins and padding

        compute_button = tk.Button(root, text='Compute', command=lambda: self.compute(kontoauszugFile, configFile, error_label))
        compute_button.pack(padx=20, pady=10, ipadx=10, ipady=10)  # Add margins and padding

        error_label = tk.Label(root, text='', fg='red')
        error_label.pack(padx=20, pady=10)

        columns = ("Konto", "Summe der Buchungen")
        self.result_tree = ttk.Treeview(root, columns=columns, show="headings")
        for col in columns:
            self.result_tree.heading(col, text=col)
        self.result_tree.pack(pady=20, expand=True, fill="both")

        # For example, setup Treeview for displaying results
        columns = ("Konto", "Summe der Buchungen")
        self.result_tree = ttk.Treeview(self.root, columns=columns, show="headings")
        for col in columns:
            self.result_tree.heading(col, text=col)
        self.result_tree.pack(pady=20, expand=True, fill="both")
        self.result_tree.bind("<Double-1>", self.open_subtable)

    def compute(self, auszugFile, configFile, error_label):
        try:
            if not os.path.exists(auszugFile.get()):
                if auszugFile.get() == "":
                    raise ValueError("No Kontoauszug selected.")
            
                    raise ValueError("Kontoauszug " + auszugFile.get() + " does not exist.")
            if not os.path.exists(configFile.get()):
                if configFile.get() == "":
                    raise ValueError("No config file selected.")
                else:
                    raise ValueError("Config file " + configFile.get() + " does not exist.")
            
            kontomarkers, ignoremarkers = parseConfig(configFile.get())
               
            auszug = Auszug(auszugFile.get())
            buchungen = auszug.buildBuchungen(ignoremarkers)

            auszug.combineBuchungen(kontomarkers)
        
            self.update_result_tree(auszug.konto_buchungen)
                       
        except Exception as e:
            error_label.config(text=str(e))
            
    def update_result_tree(self, data):
        # Clear existing data
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        # Insert new data
        for konto, buchungen in data.items():
            summe = sum(buchung.wert for buchung in buchungen)
            self.result_tree.insert("", tk.END, values=(konto, summe))

    def open_subtable(self, event):
        selected_item = self.result_tree.selection()[0]
        konto = self.result_tree.item(selected_item)['values'][0]

        detail_window = tk.Toplevel(self.root)
        detail_window.title(f"Details for {konto}")

        detail_columns = ("Buchung Text", "Buchung Wert")
        detail_tree = ttk.Treeview(detail_window, columns=detail_columns, show="headings")
        for col in detail_columns:
            detail_tree.heading(col, text=col)
        detail_tree.pack(expand=True, fill="both")

        if self.auszug:
            buchungen = self.auszug.konto_buchungen[konto]
            for buchung in buchungen:
                detail_tree.insert("", tk.END, values=(buchung.text, buchung.wert))

if __name__ == "__main__":
    root = tk.Tk()
    app = MyApplication(root)
    root.mainloop()














# Function to update the Treeview with new data
'''
        
def open_subtable(event, auszug):
    selected_item = result_tree.selection()[0]  # Get selected item
    konto = result_tree.item(selected_item)['values'][0]  # Extract the Konto value

    # Create a new Toplevel window
    detail_window = tk.Toplevel(root)
    detail_window.title(f"Details for {konto}")

    # Setup Treeview in the new window
    detail_columns = ("Buchung Text", "Buchung Wert")
    detail_tree = ttk.Treeview(detail_window, columns=detail_columns, show="headings")
    for col in detail_columns:
        detail_tree.heading(col, text=col)
    detail_tree.pack(expand=True, fill="both")

    # Populate the Treeview with Buchung objects
    buchungen = auszug.konto_buchungen[konto]
    for buchung in buchungen:
        detail_tree.insert("", tk.END, values=(buchung.text, buchung.wert))''



#result_tree.bind("<Double-1>", lambda event, auszug=auszug: open_subtable(event, auszug))

root.mainloop()'''





    
