from db import get_db_connection

try:
    conn = get_db_connection()
    print("✅ Database connection successful!")
    conn.close()
except Exception as e:
    print(f"❌ Database connection failed: {e}")
