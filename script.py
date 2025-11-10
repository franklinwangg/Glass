import sqlite3, csv, datetime, os, shutil, pandas as pd, openpyxl, sys

# === PARAMETERS ===
# Usage: python script.py <username> <month> <day> <year>
if len(sys.argv) != 5:
    print("Usage: python script.py <username> <month> <day> <year>")
    sys.exit(1)

username = sys.argv[1]
month = int(sys.argv[2])
day = int(sys.argv[3])
year = int(sys.argv[4])

# === CONFIG ===
EDGE_HISTORY = fr"C:\Users\{username}\AppData\Local\Microsoft\Edge\User Data\Default\History"

# Output folder
output_dir = os.path.join(os.getcwd(), "output")
os.makedirs(output_dir, exist_ok=True)

# Target date
target_date = datetime.datetime(year, month, day)
file_name = f"{month}-{day}-{str(year)[-2:]}.csv"
OUT_PATH = os.path.join(output_dir, file_name)

# === AUTOMATICALLY COPY HISTORY ===
temp_copy = "History_copy"
try:
    shutil.copy2(EDGE_HISTORY, temp_copy)
    print(f"📂 Copied browser history from {EDGE_HISTORY}")
except Exception as e:
    print(f"⚠️ Could not copy Edge history: {e}")
    sys.exit(1)

# === CONNECT TO COPIED DB ===
conn = sqlite3.connect(temp_copy)
cur = conn.cursor()

# === TIME RANGE (UTC) ===
start = datetime.datetime(year, month, day, 0, 0, 0, tzinfo=datetime.timezone.utc)
end = datetime.datetime(year, month, day, 23, 59, 59, tzinfo=datetime.timezone.utc)

# === QUERY ===
query = """
SELECT
    title,
    datetime((last_visit_time/1000000) - 11644473600, 'unixepoch') AS visited
FROM urls
WHERE ((last_visit_time/1000000 - 11644473600) BETWEEN ? AND ?)
ORDER BY last_visit_time DESC;
"""
cur.execute(query, (int(start.timestamp()), int(end.timestamp())))
rows = cur.fetchall()
conn.close()
os.remove(temp_copy)

# === CLEAN + FORMAT ===
cleaned_rows = []
for title, visited in rows:
    if not title or title.strip() == "":
        title = "(No title)"
    try:
        dt = datetime.datetime.strptime(visited, "%Y-%m-%d %H:%M:%S")
        visited_local = dt.strftime("%I:%M:%S %p")  # Convert to AM/PM
    except Exception:
        visited_local = visited
    cleaned_rows.append([title, visited_local])

# === WRITE CSV ===
with open(OUT_PATH, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Title", "Visited (Local Time)"])
    writer.writerows(cleaned_rows)

print(f"✅ Exported {len(cleaned_rows)} visits to {OUT_PATH}")