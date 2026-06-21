import sqlite3

conn = sqlite3.connect("grocery.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS coupons(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE,
    discount INTEGER
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS reviews(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER,
    customer_name TEXT,
    rating INTEGER,
    comment TEXT
)
""")

try:
    cursor.execute("INSERT INTO coupons(code, discount) VALUES('SAVE10', 10)")
except:
    pass

try:
    cursor.execute("INSERT INTO coupons(code, discount) VALUES('SAVE20', 20)")
except:
    pass

conn.commit()
conn.close()

print("Coupon and Review tables created successfully")