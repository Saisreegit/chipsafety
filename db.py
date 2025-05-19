import mysql.connector
import os
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

def get_db_connection():
    """Establish a connection to the MySQL database."""
    connection = mysql.connector.connect(
        host=os.getenv("DB_HOST"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        database=os.getenv("DB_NAME")
    )
    return connection

def insert_excel_data(sheet_name, row_data):
    """
    Insert data from the Excel file into the corresponding MySQL table.

    Args:
        sheet_name (str): Name of the sheet (which should match the table name in DB).
        row_data (tuple): Data from the Excel sheet to insert into the table.
    """
    connection = get_db_connection()
    cursor = connection.cursor()
    
    # Generate placeholders based on the number of data columns
    placeholders = ', '.join(['%s'] * len(row_data))
    query = f"INSERT INTO `{sheet_name}` VALUES ({placeholders})"

    try:
        cursor.execute(query, row_data)
        connection.commit()
        print(f"Data inserted into {sheet_name} successfully!")
    except mysql.connector.Error as e:
        print(f"Error inserting data into {sheet_name}: {e}")
    finally:
        cursor.close()
        connection.close()

