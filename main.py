import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, font, ttk
from docxtpl import DocxTemplate
import db_setup as db
import customtkinter
from PIL import Image, ImageTk


# Change to the current script directory
os.chdir(sys.path[0])


def open_word_document(event):
    # Get the selected item
    selected_item = event.widget.selection()
    if not selected_item:
        return

    p_id = event.widget.item(selected_item, 'values')[4]
    path = db.get_docx_path(p_id)
    # Check if file exists before attempting to open
    if path and os.path.exists(path):
        try:
            # Use the default application to open the file
            if os.name == 'nt':  # Windows
                os.startfile(path)
            elif os.name == 'posix':  # macOS and Linux
                subprocess.run(['open', path], check=True)
            else:
                print("Unsupported operating system")
        except Exception as e:
            print(f"Error opening file: {e}")
    else:
        print("File not found")


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


def load_data(self):
    # Clear existing Treeview contents before inserting new data
    for item in self.treeview.get_children():
        self.treeview.delete(item)

    # Fetch data and populate the Treeview
    rows = db.fetch_data()

    # Keep track of inserted patient IDs to prevent duplicates
    inserted_patient_ids = set()

    for row in rows:
        # Assuming the patient_id is the last element in the row
        patient_id = row[-1]

        # Only insert if this patient_id hasn't been inserted before
        if patient_id not in inserted_patient_ids:
            self.treeview.insert("", tk.END, values=row)

            inserted_patient_ids.add(patient_id)

    print(f"Total rows inserted: {len(inserted_patient_ids)}")


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
        self.root.iconbitmap("logo/logo_icon.ico")  # Provide the path to your .ico file
        self.root.configure(bg="#E8ECD7")  # Use a color name or hex code


        # Tab Control Setup
        self.tab_control = customtkinter.CTkTabview(root, fg_color="#E8ECD7")

        # Add tabs with Hebrew names
        self.patient_tab = self.tab_control.add('המטופל פרטי')

        self.search_tab = self.tab_control.add('מטופל חיפוש')

        # Pack the Tab Control with proper expansion
        self.tab_control.pack(expand=True, fill="both", padx=10, pady=5)

        # Configure tabs
        self.create_patient_info_tab()
        self.create_search_tab()

    def search_data(self, event=None):
        search_term = self.search_entry.get()

        # Clear existing items in treeview
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        # Get search results from database
        results = db.search_patients(search_term)

        # Keep track of seen patient IDs to avoid duplicates
        seen_patient_ids = set()

        # Reinsert matching items
        for row in results:
            # If this patient hasn't been seen before, insert the row
            if row[-1] not in seen_patient_ids:
                self.treeview.insert('', 'end', values=row)
                seen_patient_ids.add(row[-1])

        # Optional: Show a message if no results found
        if len(seen_patient_ids) == 0:
            messagebox.showinfo("Search Results", "No matching records found.")


    def create_patient_info_tab(self):
        search_term = self.search_entry.get()

        # Clear existing items in treeview
        for item in self.patient_tree.get_children():
            self.patient_tree.delete(item)

        # Get search results from database
        results = db.search_patients(search_term)

        # Keep track of seen patient IDs to avoid duplicates
        seen_patient_ids = set()

        # Reinsert matching items
        for row in results:
            # If this patient hasn't been seen before, insert the row
            if row[-1] not in seen_patient_ids:
                self.treeview.insert('', 'end', values=row)
                seen_patient_ids.add(row[-1])

        # Optional: Show a message if no results found
        if len(seen_patient_ids) == 0:
            messagebox.showinfo("Search Results", "No matching records found.")

    def create_patient_info_tab(self):
        # Configure grid for proper layout
        self.patient_tab.grid_rowconfigure(0, weight=1)  # Allocate space for the label
        self.patient_tab.grid_columnconfigure(0, weight=1)  # Adjust columns
        self.patient_tab.grid_columnconfigure(1, weight=1)  # Adjust columns
        self.patient_tab.grid_rowconfigure(5, weight=1)  # Adjust columns

        # Set a Hebrew-friendly font
        hebrew_font = ("Arial", 14)
        padX_size = 10
        padX_age_size = 10
        # Load an image using Pillow
        image = Image.open("logo/SamartDoc.png")
        image = image.resize((250, 150))  # Resize the image if needed
        self.photo = ImageTk.PhotoImage(image)  # Keep a reference to the image
        # Logo Label

        self.logo_label = customtkinter.CTkLabel(
            self.patient_tab,
            image=self.photo,
            text=""  # Set text to an empty string to only show the image
        )

        self.logo_label.grid(row=0, column=0, columnspan=2)
        # Hebrew Labels and Entries with right alignment
        # First Name
        self.f_name_label = customtkinter.CTkLabel(
            self.patient_tab,
            text="שם פרטי",
            font=hebrew_font,
            anchor="e"  # Right align the text
        )
        self.f_name_label.grid(row=1, column=1, padx=10, pady=5, sticky='w')

        self.f_name_entry = customtkinter.CTkEntry(
            self.patient_tab,
            font=hebrew_font,
            width=250,
            justify='right'
        )
        self.f_name_entry.grid(row=1, column=0, padx=10, pady=5, sticky='e')

        # Last Name
        self.l_name_label = customtkinter.CTkLabel(
            self.patient_tab,
            text="שם משפחה",
            font=hebrew_font,
            anchor="e"
        )
        self.l_name_label.grid(row=2, column=1, padx=10, pady=5, sticky='w')
        self.l_name_entry = customtkinter.CTkEntry(
            self.patient_tab,
            font=hebrew_font,
            width=250,
            justify='right'

        )
        self.l_name_entry.grid(row=2, column=0, padx=10, pady=5, sticky='e')  # align the entry to the right

        # id input

        self.id_label = customtkinter.CTkLabel(
            self.patient_tab,
            text="תעודת זהות",
            font=hebrew_font,
            anchor="w"
        )
        self.id_label.grid(row=3, column=1, padx=padX_size, pady=5, sticky='w')
        self.id_entry = customtkinter.CTkEntry(
            self.patient_tab,
            font=hebrew_font,
            width=250,
            justify='right'
        )
        self.id_entry.grid(row=3, column=0, padx=10, pady=5, sticky='e')

        # Age Input
        self.age_label = customtkinter.CTkLabel(
            self.patient_tab,
            text="גיל",
            font=hebrew_font,
            anchor="e"
        )
        self.age_label.grid(row=4, column=1, padx=10, pady=5, sticky='w')
        self.age_entry = customtkinter.CTkEntry(
            self.patient_tab,
            font=hebrew_font,
            width=50,
            justify='right'
        )
        self.age_entry.grid(row=4, column=0, padx=padX_age_size, pady=5, sticky='e')

        # Submit Button
        self.create_button = customtkinter.CTkButton(self.patient_tab, text=" WORD צור קובץ ", width=250,
                                                     command=self.collect_data,

                                                     )
        self.create_button.grid(row=5, column=0, padx=padX_size, pady=5, sticky='e')


    def create_search_tab(self):
            hebrew_font = ("Arial", 14)

            # Configure column weights to make the layout responsive
            self.search_tab.columnconfigure(0, weight=1)  # Search button
            self.search_tab.columnconfigure(1, weight=1)  # Search entry
            self.search_tab.columnconfigure(2, weight=3)  # Label
            self.search_tab.rowconfigure(1, weight=1)  # Make treeFrame's row expandable

            self.search_label = customtkinter.CTkLabel(
                self.search_tab,
                text="חיפוש מטופל",
                font=hebrew_font,
                anchor="center"
            )
            self.search_label.grid(row=0, column=3, padx=10, pady=5, sticky='we')

            self.search_entry = customtkinter.CTkEntry(
                self.search_tab,
                font=hebrew_font,

                justify='right'
            )
            self.search_entry.grid(row=0, column=2, padx=10, pady=10, sticky='we')
            self.search_entry.bind("<Return>", self.search_data)

            # Search Button
            self.search_button = customtkinter.CTkButton(self.search_tab,
                                                         text="חיפוש",
                                                         width=100,
                                                         command=self.search_data)
            self.search_button.grid(row=0, column=1, sticky='we', padx=10, pady=10)

            self.delete_button = customtkinter.CTkButton(self.search_tab,
                                                         text="איפוס",
                                                         width=100,
                                                         fg_color="red",
                                                         hover_color="#AF1740",
                                                         command=self.delete_search_data)
            self.delete_button.grid(row=0, column=0, sticky='we', padx=10, pady=10)

            self.treeFrame = ttk.Frame(self.search_tab)
            self.treeFrame.grid(row=1, column=0, padx=10, pady=10, columnspan=4, sticky='nswe')

            self.treeScroll = ttk.Scrollbar(self.treeFrame)
            self.treeScroll.pack(side="right", fill="y")

            cols = ("תאריך ביקור", "גיל", "שם פרטי", "שם משפחה", "תעודה מזהה")
            self.treeview = ttk.Treeview(self.treeFrame, show="headings",
                                         yscrollcommand=self.treeScroll.set, columns=cols, height=13)
            # Configure each column
            for col in cols:
                # Set column heading with center alignment
                self.treeview.heading(col, text=col, anchor="center")

                # Set column width and data alignment
                self.treeview.column(col, width=100, anchor="center")
            # Bind the left-click event to the open_docx function
            self.treeview.bind("<Double-1>", open_word_document)
            # Bind the Enter key press event to the open_docx function
            self.treeview.bind("<Return>", open_word_document)
            self.treeScroll.config(command=self.treeview.yview)
            self.treeview.pack(fill="both", expand=True)

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
            age = str(age)
        except ValueError:
            messagebox.showerror("שגיאת קלט", "!הגיל חייב להיות מספר")
            return

        # Get the current date in the desired format (e.g., dd-mm-yyyy)
        current_date = datetime.now().strftime('%d-%m-%Y')  # Use hyphens instead of slashes
        docx = create_docx(first_name, last_name, ID, age, current_date)
        db.insert_patient_record(first_name, last_name, ID, age, docx, current_date)
        # Clear all entry widgets
        self.f_name_entry.delete(0, tk.END)
        self.l_name_entry.delete(0, tk.END)
        self.id_entry.delete(0, tk.END)
        self.age_entry.delete(0, tk.END)
        load_data(self)


def main():
    # Call the function to create the tables
    db.create_tables()
    root = tk.Tk()
    root.option_add('*Font', 'Arial 14')
    PatientForm(root)
    root.mainloop()


if __name__ == "__main__":
    main()
