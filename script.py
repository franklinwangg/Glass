import sqlite3, csv, datetime, os, shutil
import pandas as pd
import openpyxl

# === CONFIG ===
DB_PATH = r"C:\Users\frank\Desktop\History"
target_date = datetime.datetime(2025, 11, 9)

# Output folder and file name (e.g., output/11-9-25.csv)
output_dir = os.path.join(os.getcwd(), "output")
os.makedirs(output_dir, exist_ok=True)
file_name = f"{target_date.month}-{target_date.day}-{str(target_date.year)[-2:]}.csv"
OUT_PATH = os.path.join(output_dir, file_name)

# === COPY AND CONNECT ===
temp_copy = "History_copy"
shutil.copy2(DB_PATH, temp_copy)

conn = sqlite3.connect(temp_copy)
cur = conn.cursor()

# === TIME RANGE (UTC) ===
start = datetime.datetime(target_date.year, target_date.month, target_date.day, 0, 0, 0, tzinfo=datetime.timezone.utc)
end = datetime.datetime(target_date.year, target_date.month, target_date.day, 23, 59, 59, tzinfo=datetime.timezone.utc)

print(f"Filtering visits between:\n  {start}  and  {end}")

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

# === CLEAN + FORMAT RESULTS ===
cleaned_rows = []
for title, visited in rows:
    # Handle missing titles
    if not title or title.strip() == "":
        title = "(No title)"
    # Convert to AM/PM format
    try:
        dt = datetime.datetime.strptime(visited, "%Y-%m-%d %H:%M:%S")
        visited_local = dt.strftime("%I:%M:%S %p")  # e.g. "09:22:27 PM"
    except Exception:
        visited_local = visited
    cleaned_rows.append([title, visited_local])

# === WRITE CSV ===
with open(OUT_PATH, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Title", "Visited (Local Time)"])
    writer.writerows(cleaned_rows)

print(f"✅ Exported {len(cleaned_rows)} visits to {OUT_PATH}")