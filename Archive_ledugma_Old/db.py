import sqlite3
db_path = r'C:\Users\Zvi\.gemini\antigravity\conversations\9a1929d9-eeae-443b-95e3-8045d9c9ea22.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()
print('Tables:', tables)
for table in tables:
    cursor.execute(f"PRAGMA table_info({table[0]});")
    print(f'Schema for {table[0]}:', cursor.fetchall())
    
    if table[0] == 'conversation' or table[0] == 'metadata':
        try:
            cursor.execute(f"SELECT * FROM {table[0]}")
            print(f"Data in {table[0]}:", cursor.fetchall())
        except Exception as e:
            print("Error", e)
