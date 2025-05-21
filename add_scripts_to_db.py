import os
import sqlite3

DB_PATH = "scripts.db"
SCRIPTS_FOLDER = "powershell_scripts"

def add_scripts_to_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Create table if it doesn't exist
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS scripts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        description TEXT,
        script_content TEXT,
        type TEXT
    )
    """)
    
    for filename in os.listdir(SCRIPTS_FOLDER):
        if filename.endswith(".ps1"):
            path = os.path.join(SCRIPTS_FOLDER, filename)
            with open(path, "r", encoding="utf-8") as f:
                script_content = f.read()
            name = os.path.splitext(filename)[0]
            description = f"Script imported from {filename}"
            type_ = "powershell"
            
            # Check for duplicates
            cursor.execute("SELECT COUNT(*) FROM scripts WHERE name = ?", (name,))
            if cursor.fetchone()[0] == 0:
                cursor.execute(
                    "INSERT INTO scripts (name, description, script_content, type) VALUES (?, ?, ?, ?)",
                    (name, description, script_content, type_)
                )
                print(f"Added {filename} to database.")
            else:
                print(f"Skipped {filename} as it already exists in the database.")
    
    conn.commit()
    conn.close()

if __name__ == "__main__":
    add_scripts_to_db()
    print("All scripts added.")