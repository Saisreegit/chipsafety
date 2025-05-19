import mysql.connector
import os
from dotenv import load_dotenv
from werkzeug.security import generate_password_hash

# Load environment variables from .env file
load_dotenv()

def get_db_connection():
    return mysql.connector.connect(
        host=os.getenv("DB_HOST"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        database=os.getenv("DB_NAME"),
        port=int(os.getenv("DB_PORT", 3306))
    )

def insert_excel_data(df):
    """
    Create table if not exists and insert rows from the Excel DataFrame (df)
    into the database table called 'excel_data'.
    """
    conn = get_db_connection()
    cursor = conn.cursor()

    # Create table if not exists
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS excel_data (
            id INT AUTO_INCREMENT PRIMARY KEY,
            attribute VARCHAR(255),
            value TEXT
        )
    """)
    conn.commit()

    # Insert each row from DataFrame into the table
    for index, row in df.iterrows():
        cursor.execute(
            "INSERT INTO excel_data (attribute, value) VALUES (%s, %s)",
            (row['attribute'], row['value'])
        )

    conn.commit()
    cursor.close()
    conn.close()

def insert_user(username, raw_password):
    """
    Insert a new user into the 'users' table with a hashed password.
    """
    conn = get_db_connection()
    cursor = conn.cursor()

    hashed_password = generate_password_hash(raw_password)

    try:
        cursor.execute(
            "INSERT INTO users (username, password) VALUES (%s, %s)",
            (username, hashed_password)
        )
        conn.commit()
        print(f"User '{username}' inserted successfully.")
    except mysql.connector.Error as err:
        print(f"Error inserting user: {err}")
    finally:
        cursor.close()
        conn.close()
