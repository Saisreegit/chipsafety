from db import get_db_connection

try:
    conn = get_db_connection()
    print("✅ Database connection successful!")
    conn.close()
except Exception as e:
<<<<<<< HEAD
    print("❌ Database connection failed:", e)
=======
    print(f"❌ Database connection failed: {e}")
>>>>>>> fccd4c5 (Connected EC2 to RDS and tested DB connection)
