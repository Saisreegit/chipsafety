import mysql.connector
import os
from dotenv import load_dotenv

load_dotenv()

def get_connection():
    return mysql.connector.connect(
        host=os.getenv("DB_HOST"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        database=os.getenv("DB_NAME"),
        port=int(os.getenv("DB_PORT", 3306))
    )

def insert_excel_data(df):
    conn = get_connection()
    cursor = conn.cursor()

    # Create table if not exists
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS excel_data (
            id INT AUTO_INCREMENT PRIMARY KEY,
            attribute VARCHAR(255),
            value TEXT
        )
    """)

    # Optional: clear old values
    cursor.execute("DELETE FROM excel_data")

    for index, row in df.iterrows():
        attribute = str(row[0])  # First column (read-only)
        value = str(row[1])      # Second column (editable)
        cursor.execute("INSERT INTO excel_data (attribute, value) VALUES (%s, %s)", (attribute, value))

    conn.commit()
    cursor.close()
    conn.close()
