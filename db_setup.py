import sqlite3


def create_tables():
    # Connect to SQLite database (or create it if it doesn't exist)
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Enable foreign key constraint enforcement (important for SQLite)
    cursor.execute("PRAGMA foreign_keys = ON")

    # Check if the 'patients' table already exists
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='patients'")
    if cursor.fetchone() is None:
        # If the table doesn't exist, create it
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS patients (
            patient_id TEXT PRIMARY KEY,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            age TEXT NOT NULL
        )
        """)
        print("Table 'patients' created successfully!")
    else:
        print("Table 'patients' already exists.")

    # Check if the 'visits' table already exists
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='visits'")
    if cursor.fetchone() is None:
        # If the table doesn't exist, create it
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS visits (
            visit_id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_id TEXT NOT NULL,
            visit_date TEXT NOT NULL,
            docx_path TEXT NOT NULL,
            FOREIGN KEY (patient_id) REFERENCES patients(patient_id)
        )
        """)
        print("Table 'visits' created successfully!")
    else:
        print("Table 'visits' already exists.")

    # Commit the changes and close the connection
    conn.commit()
    conn.close()


def fetch_data():
    """Fetch data from the SQLite database."""
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Query to fetch patient and visit details (without docx_path)
    query = """
        SELECT v.visit_date, p.age,  p.first_name,p.last_name, p.patient_id
        FROM patients p
        LEFT JOIN visits v ON p.patient_id = v.patient_id
    """
    cursor.execute(query)
    rows = cursor.fetchall()

    # Print rows for debugging
    for row in rows:
        print(row)  # Check the data returned

    conn.close()
    return rows


def insert_visit_record(patient_id, time, docx_path):
    # Connect to the SQLite database
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Insert the patient record into the patients table
    cursor.execute("""
        INSERT INTO visits (patient_id, visit_date, docx_path)
        VALUES (?, ?, ?)
        """, (patient_id, docx_path, time))

    # Commit the changes and close the connection
    conn.commit()
    print("Visit record inserted successfully!")
    conn.close()


def insert_patient_record(first_name, last_name, patient_id, age, time, docx_path):
    """
    Inserts a new patient record into the SQLite database.

    Parameters:
        first_name (str): The patient's first name.
        last_name (str): The patient's last name.
        patient_id (str): The patient's unique ID.
        age (str): The patient's age (as a string).
        time (str): The time of the record (e.g., a timestamp or date).
        docx_path (str): The path to the patient's document file.
    """
    # Connect to the SQLite database
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Insert the patient record into the patients table
    cursor.execute("""
    INSERT INTO patients (patient_id, first_name, last_name, age)
    VALUES (?, ?, ?, ?)
    """, (patient_id, first_name, last_name, age))

    # Commit the changes and close the connection
    conn.commit()
    print("Patient record inserted successfully!")
    conn.close()
    insert_visit_record(patient_id, time, docx_path)


def get_docx_path(patient_id):
    # Connect to SQLite database
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Query to get the docx path for the given patient_id
    query = """
    SELECT docx_path
    FROM visits
    WHERE patient_id = ?
    """
    cursor.execute(query, (patient_id,))

    # Fetch the result
    result = cursor.fetchone()

    # Close the connection
    conn.close()

    if result:
        return result[0]  # Return the file path
    else:
        return None  # Return None if no record is found


def search_patients(search_term):
    """
    Search patients in the database based on a search term.

    :param search_term: String to search for in patient records
    :return: List of matching patient records
    """
    # Connect to the database
    conn = sqlite3.connect("patients.db")
    cursor = conn.cursor()

    # Create a search query that checks multiple columns
    query = """
        SELECT  v.visit_date, p.age, p.last_name, p.first_name, p.patient_id
        FROM patients p
        LEFT JOIN visits v ON p.patient_id = v.patient_id
        WHERE 
            LOWER(p.first_name) LIKE ? OR 
            LOWER(p.last_name) LIKE ? OR 
            LOWER(p.age) LIKE ? OR 
            LOWER(v.visit_date) LIKE ? OR
            LOWER(p.patient_id) LIKE ?
        """

    # Use % wildcards for partial matching
    search_param = f'%{search_term.lower()}%'

    # Execute the query
    cursor.execute(query, (search_param, search_param, search_param, search_param, search_param))

    # Fetch and process results
    results = cursor.fetchall()

    conn.close()

    return results
