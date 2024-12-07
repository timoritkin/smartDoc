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
