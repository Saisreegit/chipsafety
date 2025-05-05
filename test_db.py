import os
from dotenv import load_dotenv
import mysql.connector

# Load .env file
load_dotenv()

try:
    # Connect to RDS
    conn = mysql.connector.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASS"),
        database=os.getenv("DB_NAME")
    )
    
    cursor = conn.cursor()
    cursor.execute("SELECT NOW();")
    result = cursor.fetchone()
    print("✅ Connected to database successfully. Server time is:", result[0])

    cursor.close()
    conn.close()
except mysql.connector.Error as err:
    print("❌ Database connection failed:", err)
