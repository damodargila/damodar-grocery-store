import sqlite3

conn = sqlite3.connect("grocery.db")
cursor = conn.cursor()

try:
    cursor.execute("ALTER TABLE products ADD COLUMN discount INTEGER DEFAULT 10")
    print("discount column added")
except:
    print("discount column already exists")

try:
    cursor.execute("ALTER TABLE orders ADD COLUMN status TEXT DEFAULT 'Pending'")
    print("status column added")
except:
    print("status column already exists")

conn.commit()
conn.close()

print("Database updated successfully")