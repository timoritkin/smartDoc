import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk
from docxtpl import DocxTemplate
import db_setup as db
import customtkinter as ctk
from PIL import Image
from tkcalendar import DateEntry
from customtkinter import CTkImage

hebrew_font = ("Arial", 16, "bold")
padX_size = 10
padY_size = (0, 20)
sticky_label = "w"
sticky_entry = "e"


def resource_path(relative_path):
    """Get the absolute path to a resource, compatible with PyInstaller."""
    try:
        # Use the temp folder path when running as a PyInstaller bundle
        base_path = sys._MEIPASS
    except AttributeError:
        # Use the current directory in normal execution
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def open_word_document(event):
    # Get the selected item
    selected_item = event.widget.selection()
    if not selected_item:
        return

    # Retrieve file information
    p_id = event.widget.item(selected_item, 'values')[5]
    visit_date = event.widget.item(selected_item, 'values')[0]

    # Get the document path from the database (adjust `db.get_docx_path` if necessary)
    path = db.get_docx_path(p_id, visit_date)

    # Resolve the full path for bundled environments
    if path:
        path = resource_path(path)

    # Check if the file exists before attempting to open it
    if path and os.path.exists(path):
        try:
            # Use the default application to open the file
            if os.name == 'nt':  # Windows
                os.startfile(path)
            elif os.name == 'posix':  # macOS and Linux
                subprocess.run(['open', path], check=True)
            else:
                messagebox.showerror("Error", "Unsupported operating system")
        except Exception as e:
            messagebox.showerror("Error", f"Error opening file: {e}")
    else:
        messagebox.showwarning("Warning", "הקובץ לא נמצא")


def create_docx(f_name, l_name, id_num, age, date, phone):
    # Load the template using resource_path
    template_path = resource_path('template/Clalit mushlam template.docx')
    doc = DocxTemplate(template_path)

    # Define the folder name for saving documents
    folder_name = 'patients docx'

    # Get the current script's directory (adjust for PyInstaller's bundle)
    script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
    folder_path = os.path.join(script_dir, folder_name)

    # Ensure the folder exists, create it if it doesn't
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # Prepare context for the document
    context = {'f_name': f_name, 'l_name': l_name, 'id': id_num, 'age': age, 'phone': phone}

    # Render the document with the provided data
    doc.render(context)

    # Save the document with a new name
    file_name = f'{f_name}_{l_name}_{id_num}_{date}_doc.docx'
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

    for row in rows:
        birthdate_str = row[2]  # Example: row[2] is the birthdate column in 'dd/mm/yyyy' format
        age = self.calculate_age(birthdate_str)
        row_with_replaced_age = list(row)  # Convert the tuple to a list
        row_with_replaced_age[2] = age  # Assuming the Age column is at index 2
        self.treeview.insert("", tk.END, values=row_with_replaced_age)


class PatientForm:

    def __init__(self, root):
        self.root = root
        self.root.title("SmartDoc")
        self.root.geometry("800x700")
        self.root.iconbitmap("logo/logo_icon.ico")  # Provide the path to your .ico file
        self.root.configure(bg="#E8ECD7")  # Use a color name or hex code

        # Configure the grid layout for the window
        self.root.grid_columnconfigure(0, weight=1)  # Main frame will expand
        self.root.grid_columnconfigure(1, weight=0)  # Options frame stays fixed
        self.root.grid_rowconfigure(0, weight=1)  # Main frame will expand vertically

        # Main frame (left side)
        self.main_frame = ctk.CTkFrame(self.root)  # Use ctk.CTkFrame directly
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        # Options frame (right side)
        self.options_frame = ctk.CTkFrame(self.root)  # Use ctk.CTkFrame directly
        self.options_frame.grid(row=0, column=1, sticky="ns")  # Stick to top and bottom

        self.logo_label = ctk.CTkLabel(self.options_frame,
                                       text="logo",
                                       )
        self.logo_label.pack(pady=(10, 150))
        # Adding buttons to options_frame
        self.button1 = ctk.CTkButton(self.options_frame,
                                     text="חדש מטופל",
                                     width=200,
                                     height=40,
                                     command=self.show_frame_1)
        self.button1.pack(pady=10)
        self.button2 = ctk.CTkButton(self.options_frame,
                                     text="מטופל חיפוש",
                                     width=200,
                                     height=40,
                                     command=self.show_frame_2)
        self.button2.pack(pady=10)

        # Frames to be displayed in main_frame
        self.new_form_frame = ctk.CTkFrame(self.main_frame, fg_color="lightblue")

        # Configure grid columns and rows to expand equally
        self.new_form_frame.grid_columnconfigure(0, weight=1, )  # Make column 0 expand
        self.new_form_frame.grid_columnconfigure(1, weight=1, )  # Make column 1 expand
        self.new_form_frame.grid_rowconfigure(0, weight=1, )  # Make row 0 expand

        self.new_form_frame.grid_rowconfigure(6, weight=1, )  # Make row 6 expand

        # Frame 1 widgets
        self.f_name_label = ctk.CTkLabel(
            self.new_form_frame,
            text="שם פרטי",
            font=hebrew_font,
            anchor="e"  # Right align the text
        )
        self.f_name_label.grid(row=1, column=1, padx=padX_size, pady=padY_size, sticky=sticky_label)

        self.f_name_entry = ctk.CTkEntry(
            self.new_form_frame,
            font=hebrew_font,
            width=250,
            justify='right'
        )
        self.f_name_entry.grid(row=1, column=0, padx=padX_size, pady=padY_size, sticky=sticky_entry)

        # Last Name
        self.l_name_label = ctk.CTkLabel(
            self.new_form_frame,
            text="שם משפחה",
            font=hebrew_font,
            anchor="e"
        )
        self.l_name_label.grid(row=2, column=1, padx=padX_size, pady=padY_size, sticky=sticky_label)

        self.l_name_entry = ctk.CTkEntry(
            self.new_form_frame,
            font=hebrew_font,
            width=250,
            justify='right'
        )
        self.l_name_entry.grid(row=2, column=0, padx=padX_size, pady=padY_size,
                               sticky=sticky_entry)  # align the entry to the right

        # ID input
        self.id_label = ctk.CTkLabel(
            self.new_form_frame,
            text="תעודת זהות",
            font=hebrew_font,
            anchor=sticky_label
        )
        self.id_label.grid(row=3, column=1, padx=padX_size, pady=padY_size, sticky=sticky_label)

        self.id_entry = ctk.CTkEntry(
            self.new_form_frame,
            font=hebrew_font,
            width=250,
            justify='right'
        )
        self.id_entry.grid(row=3, column=0, padx=padX_size, pady=padY_size, sticky=sticky_entry)

        # Phone input
        self.phone_label = ctk.CTkLabel(
            self.new_form_frame,
            text="טלפון",
            font=hebrew_font,
            anchor="e"
        )
        self.phone_label.grid(row=4, column=1, padx=padX_size, pady=padY_size, sticky=sticky_label)

        self.phone_entry = ctk.CTkEntry(
            self.new_form_frame,
            font=hebrew_font,
            width=250,
            justify='right'
        )
        self.phone_entry.grid(row=4, column=0, padx=padX_size, pady=padY_size,
                              sticky=sticky_entry)  # align the entry to the right

        # Age Input
        self.birth_date_label = ctk.CTkLabel(
            self.new_form_frame,
            text="תאריך לידה",
            font=hebrew_font,
            anchor="e"
        )
        self.birth_date_label.grid(row=5, column=1, padx=padX_size, pady=padY_size, sticky=sticky_label)

        # Add a DateEntry widget (calendar)
        self.calendar = DateEntry(
            self.new_form_frame,
            date_pattern='dd/mm/yyyy',
            width=24,  # Increase the width to make it bigger
            background="darkblue",
            foreground="white",
            font=("Arial", 16)  # Adjust the font size to make the text inside the widget bigger
        )
        self.calendar.grid(row=5, column=0, padx=padX_size, pady=padY_size, sticky=sticky_entry)
        # Submit Button
        self.create_button = ctk.CTkButton(self.new_form_frame,
                                           text="WORD קובץ צור  ",
                                           width=250,
                                           height=100,
                                           command=self.collect_data)
        self.create_button.grid(row=6, column=0, padx=padX_size, pady=padY_size, sticky=sticky_entry)

        self.frame_2 = ctk.CTkFrame(self.main_frame, fg_color="lightgreen")  # Use ctk.CTkFrame directly
        self.frame_2_label = ctk.CTkLabel(self.frame_2, text="This is Frame 2", font=("Arial", 20))
        self.frame_2_label.pack(pady=50)

        # Show the first frame by default
        self.show_frame_1()

    def show_frame_1(self):
        """Display Frame 1 and hide Frame 2"""
        self.new_form_frame.pack(fill="both", expand=True)
        self.frame_2.pack_forget()

    def show_frame_2(self):
        """Display Frame 2 and hide Frame 1"""
        self.frame_2.pack(fill="both", expand=True)
        self.new_form_frame.pack_forget()

    def calculate_age(self, birthdate_str):

        try:
            # Parse the birthdate string
            birthdate = datetime.strptime(birthdate_str, '%d/%m/%Y')
        except ValueError:
            return None

        # Calculate the current age
        current_date = datetime.today()
        age = current_date.year - birthdate.year

        # Adjust for birthday not yet occurring this year
        if current_date.month < birthdate.month or (
                current_date.month == birthdate.month and current_date.day < birthdate.day):
            age -= 1

        return age

    def collect_data(self):
        first_name = self.f_name_entry.get()
        last_name = self.l_name_entry.get()
        ID = self.id_entry.get()
        birth_date = self.calendar.get()
        phone = self.phone_entry.get()
        check_birth_date = birth_date

        if not first_name or not last_name or not birth_date or not ID or not phone:
            messagebox.showwarning("שגיאת קלט", " ! אנא מלא את כל השדות")
            return

            # Check if the birthdate format is correct
        try:

            # Assuming the expected format is 'dd/mm/yyyy'
            check_birth_date = datetime.strptime(check_birth_date, '%d/%m/%Y')

        except ValueError:
            messagebox.showerror("שגיאת קלט", "!תאריך הלידה חייב להיות בפורמט נכון: dd/mm/yyyy")
            return
        try:
            birth_date = str(birth_date)
        except ValueError:
            messagebox.showerror("שגיאת קלט", "!הגיל חייב להיות מספר")
            return

        # Get the current date in the desired format (e.g., dd-mm-yyyy)
        current_date = datetime.now().strftime('%d-%m-%Y')  # Use hyphens instead of slashes
        age = self.calculate_age(birth_date)

        docx = create_docx(first_name, last_name, ID, age, current_date, phone)
        if db.check_patient_id_exists(ID):
            db.insert_visit_record(ID, current_date, docx)
        else:
            db.insert_patient_record(first_name, last_name, ID, birth_date, phone)
            db.insert_visit_record(ID, current_date, docx)

        # Clear all entry widgets
        self.f_name_entry.delete(0, tk.END)
        self.l_name_entry.delete(0, tk.END)
        self.id_entry.delete(0, tk.END)
        self.phone_entry.delete(0, tk.END)


def main():
    # Call the function to create the tables
    db.create_tables()
    root = ctk.CTk()  # create CTk window like you do with the Tk window
    PatientForm(root)
    root.mainloop()


if __name__ == "__main__":
    main()