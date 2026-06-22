import sqlite3

conn = sqlite3.connect("grocery.db")
cursor = conn.cursor()

columns = [
    ("description", "TEXT"),
    ("image2", "TEXT"),
    ("image3", "TEXT"),
    ("image4", "TEXT"),
    ("image5", "TEXT")
]

for name, ctype in columns:
    try:
        cursor.execute(f"ALTER TABLE products ADD COLUMN {name} {ctype}")
        print(name, "added")
    except:
        print(name, "already exists")

conn.commit()
conn.close()

print("Product table updated")