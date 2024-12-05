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


def open_word_document(event):
    # Get the selected item
    selected_item = event.widget.selection()
    if not selected_item:
        return

    # Get the Word file path from the first column
    word_file = event.widget.item(selected_item, 'values')[0]

    # Check if file exists before attempting to open
    if word_file and os.path.exists(word_file):
        try:
            # Use the default application to open the file
            if os.name == 'nt':  # Windows
                os.startfile(word_file)
            elif os.name == 'posix':  # macOS and Linux
                subprocess.run(['open', word_file], check=True)
            else:
                print("Unsupported operating system")
        except Exception as e:
            print(f"Error opening file: {e}")
    else:
        print("File not found")


def load_data(self):
    global path
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
        cols = ("קובץ", "תאריך ביקור", "גיל", "שם פרטי", "שם משפחה", "תעודה מזהה")

        # Configure Treeview columns
        self.treeview['columns'] = cols
        for col in cols:
            self.treeview.heading(col, text=col, anchor="center")
            self.treeview.column(col, width=100, anchor="center")
        self.treeview.bind('<Double-1>', open_word_document)

        # Insert data rows, mapping to normalized columns
        for row in list_values[1:]:
            # Create a dictionary to map normalized data to predefined columns
            row_dict = dict(zip(normalized_headers, row))

            # Extract values in the desired order
            ordered_row = [
                row_dict.get("קובץ", ""),
                row_dict.get("תאריך ביקור", ""),
                row_dict.get("גיל", ""),
                row_dict.get("שם פרטי", ""),
                row_dict.get("שם משפחה", ""),
                row_dict.get("תעודה מזהה", "")
            ]

            # Insert the row with ordered values
            self.treeview.insert("", "end", values=ordered_row)
            # After populating the treeview, store the original data
            for child in self.treeview.get_children():
                self.original_treeview_data.append(self.treeview.item(child)['values'])


    except FileNotFoundError:
        print(f"Error: File '{path}' not found.")
    except PermissionError:
        print(f"Error: No permission to read '{path}'.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


def create_docx(f_name, l_name, id_num, age, date):
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

    # Prepare context for the document
    context = {'f_name': f_name, 'l_name': l_name, 'id': id_num, 'age': age, 'date': date}

    # Render the document with the provided data
    doc.render(context)

    # Save the document with a new name
    file_name = f'{f_name}_{l_name}_{id_num}_{date}_doc.docx'

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

    return file_path


def insert_row(first_name, last_name, ID, age, time, docx):
    # Verify file path
    path = "patients data.xlsx"

    # Load workbook and active sheet
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    row_values = [ID, first_name, last_name, age, docx, time]
    sheet.append(row_values)
    workbook.save(path)


class PatientForm:

    def __init__(self, root):
        self.delete_button = None
        self.create_button = None
        self.treeview = None
        self.search_entry = None
        self.file_listbox = None
        self.search_label = None
        self.search_button = None
        self.age_entry = None
        self.age_label = None
        self.id_label = None
        self.id_entry = None
        self.l_name_entry = None
        self.l_name_label = None
        self.f_name_entry = None
        self.f_name_label = None
        self.original_treeview_data = []
        self.root = root
        self.style = ttk.Style(root)
        # self.root.call("source", "forest-light.tcl")
        # self.style.theme_use("forest-light")

        self.root.title("SmartDoc")
        # Set the background color of the root window
        self.root.configure(background="#6b92d1")  # Replace with your desired color
        # Create Tab Control
        self.tab_control = ttk.Notebook(root)

        self.patient_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.patient_tab, text='פרטי מטופל')

        # Medical History Tab
        self.search_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.search_tab, text='חיפוש מטופל')

        # Define a custom style with rounded corners for ttk.Entry
        self.style.configure("Rounded.TEntry",
                             relief="solid",  # Border style
                             borderwidth=2,  # Border width
                             background="#ffffff",  # Background color
                             foreground="#2A3335",  # Text color
                             padding=5)  # Padding inside the entry widget

        # You can optionally add a focus highlight color
        self.style.map("Rounded.TEntry",
                       foreground=[('focus', '#2A3335')],
                       background=[('focus', 'lightblue')])

        # Customize the tab appearance using ttk.Style
        style = ttk.Style()
        style.configure("TNotebook.Tab",
                        font=("Arial", 12, "bold"),  # Font style
                        padding=[10, 5],  # Tab padding
                        background="#355a96",  # Tab background color
                        foreground="#355a96")  # Tab text color

        style.map("TNotebook.Tab",
                  background=[('selected', 'red')],  # Change background when selected
                  foreground=[('selected', 'black')])  # Change text color when selected

        style.configure("Treeview", font=("Arial", 12))  # Change font and size
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"))  # Change header font and size
        style.configure('TNotebook', background='#6b92d1')

        style.configure('Custom.TButton',
                        background='#0A5EB0',
                        foreground='black',
                        font=('Arial', 12, 'bold'),
                        padding=10
                        )

        # Hover effect
        style.map('Custom.TButton',
                  background=[('active', '#0A5EB0')],
                  foreground=[('active', 'black')]
                  )
        style.configure('delete.TButton',
                        background='#FF2929',
                        foreground='black',
                        font=('Arial', 12, 'bold'),
                        padding=10
                        )

        # Hover effect
        style.map('Custom.TButton',
                  background=[('active', '#0A5EB0')],  # Darker green on hover
                  foreground=[('active', 'black')]
                  )


        # Pack the Tab Control
        self.tab_control.pack(expand=1, fill="both", padx=10, pady=5, )
        # Patient Information Tab Contents
        self.create_patient_info_tab()
        # Medical History Tab Contents
        self.create_search_tab()

    def search_data(self):
        search_term = self.search_entry.get().lower()

        # Clear existing items in treeview
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        # Reinsert items that match the search term
        seen_items = set()
        for child in self.original_treeview_data:
            # Convert all values to strings and check for search term
            if any(search_term in str(value).lower() for value in child):
                # Use tuple of values to check for duplicates
                item_tuple = tuple(child)
                if item_tuple not in seen_items:
                    self.treeview.insert('', 'end', values=child)
                    seen_items.add(item_tuple)

    def create_patient_info_tab(self):
        # Configure columns to allow proper space distribution
        self.patient_tab.grid_columnconfigure(0, weight=1, minsize=200)  # For the entry fields
        self.patient_tab.grid_columnconfigure(1, weight=1)  # For the labels, no expansion
        # padY_size=100
        padX_size = 10
        padX_age_size = 10
        # Hebrew font configuration
        hebrew_font = font.Font(family="Arial", size=14)
        self.logo_label = tk.Label(self.patient_tab, text="SmartDoc", font=hebrew_font, anchor='center')
        self.logo_label.grid(row=0, column=0,columnspan =2, padx=padX_size, pady=5, sticky='ew')  # align the label to the right

        self.f_name_label = tk.Label(self.patient_tab, text="שם פרטי", font=hebrew_font, anchor='center')
        self.f_name_label.grid(row=1, column=1, padx=padX_size, pady=5, sticky='ew')  # align the label to the right
        self.f_name_entry = ttk.Entry(self.patient_tab, font=hebrew_font, width=30, justify='right',
                                      style="Rounded.TEntry")
        self.f_name_entry.grid(row=1, column=0, padx=padX_size, pady=5, sticky='e')  # align the entry to the right

        self.l_name_label = tk.Label(self.patient_tab, text="שם משפחה", font=hebrew_font, anchor='center')
        self.l_name_label.grid(row=2, column=1, padx=padX_size, pady=5, sticky='ew')  # align the label to the right
        self.l_name_entry = ttk.Entry(self.patient_tab, font=hebrew_font, width=30, justify='right',
                                      style="Rounded.TEntry")
        self.l_name_entry.grid(row=2, column=0, padx=padX_size, pady=5, sticky='e')  # align the entry to the right

        # id input
        self.id_label = tk.Label(self.patient_tab, text="תעודת זהות", font=hebrew_font, anchor='center')
        self.id_label.grid(row=3, column=1, padx=padX_size, pady=5, sticky='ew')
        self.id_entry = ttk.Entry(self.patient_tab, font=hebrew_font, width=30, justify='right',
                                  style="Rounded.TEntry")
        self.id_entry.grid(row=3, column=0, padx=padX_size, pady=5, sticky='e')

        # Age Input
        self.age_label = tk.Label(self.patient_tab, text="גיל", font=hebrew_font, anchor='center')
        self.age_label.grid(row=4, column=1, padx=padX_size, pady=5, sticky='ew')
        self.age_entry = ttk.Entry(self.patient_tab, font=hebrew_font, width=10, justify='right',
                                   style="Rounded.TEntry")
        self.age_entry.grid(row=4, column=0, padx=padX_age_size, pady=5, sticky='e')

        # Submit Button
        self.create_button = ttk.Button(self.patient_tab, text=" WORD צור קובץ ", style='Custom.TButton',
                                        command=self.collect_data)
        self.create_button.grid(row=5, column=0, padx=padX_size, pady=5, sticky='e')

    def create_search_tab(self):
        hebrew_font = ("Arial", 14)

        # Configure column weights to make the layout responsive
        self.search_tab.columnconfigure(0, weight=1)  # Search button
        self.search_tab.columnconfigure(1, weight=3)  # Search entry
        self.search_tab.columnconfigure(2, weight=1)  # Label
        self.search_tab.rowconfigure(1, weight=1)  # Make treeFrame's row expandable

        self.search_label = tk.Label(self.search_tab, text="חיפוש מטופל", font=hebrew_font, anchor='center')
        self.search_label.grid(row=0, column=3, padx=10, pady=5, sticky='we')
        self.search_entry = ttk.Entry(self.search_tab, font=hebrew_font, width=30, justify='right',
                                      style="Rounded.TEntry")
        self.search_entry.grid(row=0, column=2, padx=10, pady=5, sticky='we')

        # Search Button
        self.search_button = ttk.Button(self.search_tab, text="חיפוש", style='Custom.TButton',
                                        command=self.search_data)
        self.search_button.grid(row=0, column=1, sticky='we', padx=10, pady=10)

        self.delete_button = ttk.Button(self.search_tab, text="איפוס", style='delete.TButton',
                                        command=self.delete_search_data)
        self.delete_button.grid(row=0, column=0, sticky='we', padx=10, pady=10)

        self.treeFrame = ttk.Frame(self.search_tab)
        self.treeFrame.grid(row=1, column=0, padx=10, pady=10, columnspan=4, sticky='nswe')

        self.treeScroll = ttk.Scrollbar(self.treeFrame)
        self.treeScroll.pack(side="right", fill="y")

        cols = ("קובץ", "תאריך ביקור", "גיל", "שם פרטי", "שם משפחה", "תעודה מזהה")
        self.treeview = ttk.Treeview(self.treeFrame, show="headings",
                                     yscrollcommand=self.treeScroll.set, columns=cols, height=13)
        self.treeview.column("קובץ", width=100)
        self.treeview.column("תאריך ביקור", width=100)
        self.treeview.column("גיל", width=50)
        self.treeview.column("שם משפחה", width=100)
        self.treeview.column("שם פרטי", width=100)
        self.treeview.column("תעודה מזהה", width=100)

        self.treeview.pack(fill="both", expand=True)
        self.treeScroll.config(command=self.treeview.yview)
        load_data(self)

    def delete_search_data(self):
        self.search_entry.delete(0, tk.END)
        self.search_data()

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

        # Get the current date in the desired format (e.g., dd-mm-yyyy)
        current_date = datetime.now().strftime('%d-%m-%Y')  # Use hyphens instead of slashes
        docx = create_docx(first_name, last_name, ID, age, current_date)
        insert_row(first_name, last_name, ID, age, docx, current_date)
        # Clear all entry widgets
        self.f_name_entry.delete(0, tk.END)
        self.l_name_entry.delete(0, tk.END)
        self.id_entry.delete(0, tk.END)
        self.age_entry.delete(0, tk.END)
        load_data(self)
        # patient = Patient(first_name, last_name, ID, age)


def main():
    root = tk.Tk()
    root.option_add('*Font', 'Arial 14')
    PatientForm(root)
    root.mainloop()


if __name__ == "__main__":
    main()
