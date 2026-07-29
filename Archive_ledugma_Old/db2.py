import sqlite3
db_path = r'C:\Users\Zvi\.gemini\antigravity\conversations\9a1929d9-eeae-443b-95e3-8045d9c9ea22.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()
cursor.execute("SELECT data FROM trajectory_metadata_blob")
blob = cursor.fetchone()[0]
print("Blob size:", len(blob))
print(blob)
