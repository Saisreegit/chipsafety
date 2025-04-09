# db.py
import mysql.connector

def get_connection():
    return mysql.connector.connect(
        host="your-rds-endpoint.amazonaws.com",
        user="your-username",
        password="your-password",
        database="excel_logs"
    )
