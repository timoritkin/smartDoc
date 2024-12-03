import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, font, ttk
from docxtpl import DocxTemplate
import openpyxl

# Change to the current script directory
os.chdir(sys.path[0])


def load_data(self):
    try:
        # Verify file path
        path = "patients data.xlsx"

        # Load workbook and active sheet
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active

        # Convert sheet values to a list
        list_values = list(sheet.values)

        # Check if the sheet is empty
        if not list_values:
            print("The Excel sheet is empty.")
            return

        # Normalize column headers by stripping spaces
        original_headers = list_values[0]
        normalized_headers = [header.strip() if header else "" for header in original_headers]

        # Print normalized column headers for debugging
        print("Normalized column names:")
        for header in normalized_headers:
            print(repr(header))

        # Clear existing Treeview contents
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        # Predefined columns in the desired order
        cols = ("קובץ WORD", "גיל", "שם פרטי", "שם משפחה", "תעודה מזהה")

        # Configure Treeview columns
        self.treeview['columns'] = cols
        for col in cols:
            self.treeview.heading(col, text=col, anchor="center")
            self.treeview.column(col, width=100, anchor="center")

        # Insert data rows, mapping to normalized columns
        for row in list_values[1:]:
            # Create a dictionary to map normalized data to predefined columns
            row_dict = dict(zip(normalized_headers, row))

            # Extract values in the desired order
            ordered_row = [
                row_dict.get("קובץ WORD", ""),
                row_dict.get("גיל", ""),
                row_dict.get("שם פרטי", ""),
                row_dict.get("שם משפחה", ""),
                row_dict.get("תעודה מזהה", "")
            ]

            # Insert the row with ordered values
            self.treeview.insert("", "end", values=ordered_row)

    except FileNotFoundError:
        print(f"Error: File '{path}' not found.")
    except PermissionError:
        print(f"Error: No permission to read '{path}'.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


def create_docx(f_name, l_name, id_num, age):
    # Load the template
    doc = DocxTemplate('template/Clalit mushlam template.docx')

    # Define the relative folder path (relative to the current script's location)
    folder_name = 'patients docx'  # Folder where you want to save the docx file

    # Get the current script's directory and join it with the relative folder path
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Get the directory of the script
    folder_path = os.path.join(script_dir, folder_name)  # Combine with the folder name

    # Ensure the folder exists, create it if it doesn't
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # Get the current date in the desired format (e.g., dd-mm-yyyy)
    current_date = datetime.now().strftime('%d-%m-%Y')  # Use hyphens instead of slashes

    # Prepare context for the document
    context = {'f_name': f_name, 'l_name': l_name, 'id': id_num, 'age': age}

    # Render the document with the provided data
    doc.render(context)

    # Save the document with a new name
    file_name = f'{f_name}_{l_name}_{id_num}_{current_date}_doc.docx'

    # Combine the folder path with the file name
    file_path = os.path.join(folder_path, file_name)
    doc.save(file_path)

    # Open the document automatically
    if sys.platform == "win32":  # For Windows
        os.startfile(file_path)
    elif sys.platform == "darwin":  # For macOS
        subprocess.run(["open", file_path])
    else:  # For Linux
        subprocess.run(["xdg-open", file_path])


class PatientForm:
    def __init__(self, root):
        self.treeview = None
        self.search_entry = None
        self.file_listbox = None
        self.search_label = None
        self.submit_button = None
        self.age_entry = None
        self.age_label = None
        self.id_label = None
        self.id_entry = None
        self.l_name_entry = None
        self.l_name_label = None
        self.f_name_entry = None
        self.f_name_label = None
        self.root = root
        self.root.title("SmartDoc")
        # Create Tab Control
        self.tab_control = ttk.Notebook(root)

        self.patient_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.patient_tab, text='פרטי מטופל')

        # Medical History Tab
        self.search_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.search_tab, text='חיפוש מטופל')

        # Pack the Tab Control
        self.tab_control.pack(expand=1, fill="both", padx=10, pady=5)
        # Patient Information Tab Contents
        self.create_patient_info_tab()
        # Medical History Tab Contents
        self.create_search_tab()

    def create_patient_info_tab(self):

        # padY_size=100
        padX_size = 20
        padX_age_size = 20
        mySticky = 'e'
        # Hebrew font configuration
        hebrew_font = font.Font(family="Arial", size=14)

        # Name Input with RTL support
        self.f_name_label = tk.Label(self.patient_tab, text="שם פרטי", font=hebrew_font, anchor='center')
        self.f_name_label.grid(row=0, column=1, padx=padX_size, pady=5, sticky='e')
        self.f_name_entry = tk.Entry(self.patient_tab, font=hebrew_font, width=30, justify='right')
        self.f_name_entry.grid(row=0, column=0, padx=padX_size, pady=5, sticky='w')

        self.l_name_label = tk.Label(self.patient_tab, text="שם משפחה", font=hebrew_font, anchor='center')
        self.l_name_label.grid(row=1, column=1, padx=padX_size, pady=5, sticky='e')
        self.l_name_entry = tk.Entry(self.patient_tab, font=hebrew_font, width=30, justify='right')
        self.l_name_entry.grid(row=1, column=0, padx=padX_size, pady=5, sticky='w')

        # id input
        self.id_label = tk.Label(self.patient_tab, text="תעודת זהות", font=hebrew_font, anchor='center')
        self.id_label.grid(row=2, column=1, padx=padX_size, pady=5, sticky='e')
        self.id_entry = tk.Entry(self.patient_tab, font=hebrew_font, width=30, justify='right')
        self.id_entry.grid(row=2, column=0, padx=padX_size, pady=5, sticky='w')

        # Age Input
        self.age_label = tk.Label(self.patient_tab, text="גיל", font=hebrew_font, anchor='center')
        self.age_label.grid(row=3, column=1, padx=padX_size, pady=5, sticky='e')
        self.age_entry = tk.Entry(self.patient_tab, font=hebrew_font, width=10, justify='right')
        self.age_entry.grid(row=3, column=0, padx=padX_age_size, pady=5, sticky='e')

        # Submit Button
        self.submit_button = tk.Button(self.patient_tab, text=" WORD צור קובץ ", font=hebrew_font,
                                       command=self.collect_data)
        self.submit_button.grid(row=4, column=0, columnspan=2, padx=padX_size, pady=10, sticky='we')

    def create_search_tab(self):
        hebrew_font = ("Arial", 14)

        # Configure column weights to make the layout responsive
        self.search_tab.columnconfigure(0, weight=1)  # Search button
        self.search_tab.columnconfigure(1, weight=3)  # Search entry
        self.search_tab.columnconfigure(2, weight=1)  # Label
        self.search_tab.rowconfigure(1, weight=3)  # Label

        self.search_label = tk.Label(self.search_tab, text="חיפוש מטופל", font=hebrew_font, anchor='center')
        self.search_label.grid(row=0, column=2, padx=10, pady=5, sticky='we')
        self.search_entry = tk.Entry(self.search_tab, font=hebrew_font, width=30, justify='right')
        self.search_entry.grid(row=0, column=1, padx=10, pady=5, sticky='we')

        # Submit Button
        self.submit_button = tk.Button(self.search_tab, text=" חיפוש", font=hebrew_font,
                                       command=self.collect_data)
        self.submit_button.grid(row=0, column=0, sticky='we', padx=10, pady=10)

        self.treeFrame = ttk.Frame(self.search_tab)
        self.treeFrame.grid(row=1, column=0, padx=10, pady=10, columnspan=3, sticky='nswe')

        self.treeScroll = ttk.Scrollbar(self.treeFrame)
        self.treeScroll.pack(side="right", fill="y")

        cols = ("קובץ WORD", "גיל", "שם פרטי", "שם משפחה", "תעודה מזהה")
        self.treeview = ttk.Treeview(self.treeFrame, show="headings",
                                     yscrollcommand=self.treeScroll.set, columns=cols, height=13)

        self.treeview.column("קובץ WORD", width=100)  # Corrected column name
        self.treeview.column("גיל", width=50)
        self.treeview.column("שם משפחה", width=100)
        self.treeview.column("שם פרטי", width=100)
        self.treeview.column("תעודה מזהה", width=100)

        self.treeview.pack()
        self.treeScroll.config(command=self.treeview.yview)
        load_data(self)

    def collect_data(self):
        first_name = self.f_name_entry.get()
        last_name = self.l_name_entry.get()
        ID = self.id_entry.get()
        age = self.age_entry.get()

        if not first_name or not last_name or not age or not ID:
            messagebox.showwarning("שגיאת קלט", " ! אנא מלא את כל השדות")
            return

        try:
            age = int(age)
        except ValueError:
            messagebox.showerror("שגיאת קלט", "!הגיל חייב להיות מספר")
            return

        create_docx(first_name, last_name, ID, age)
        # patient = Patient(first_name, last_name, ID, age)
