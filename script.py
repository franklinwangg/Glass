import sqlite3, csv, datetime, os, shutil, pandas as pd, openpyxl

# === CONFIG ===
# Your Edge History source file (the original one)
EDGE_HISTORY = os.path.expanduser(
    r"C:\Users\frank\AppData\Local\Microsoft\Edge\User Data\Default\History"
)

# Output folder
output_dir = os.path.join(os.getcwd(), "output")
os.makedirs(output_dir, exist_ok=True)

# Target date
target_date = datetime.datetime(2025, 11, 9)
file_name = f"{target_date.month}-{target_date.day}-{str(target_date.year)[-2:]}.csv"
OUT_PATH = os.path.join(output_dir, file_name)

# === AUTOMATICALLY COPY HISTORY ===
temp_copy = "History_copy"
try:
    shutil.copy2(EDGE_HISTORY, temp_copy)
    print(f"📂 Copied browser history from {EDGE_HISTORY}")
except Exception as e:
    print(f"⚠️ Could not copy Edge history: {e}")
    exit()

# === CONNECT TO COPIED DB ===
conn = sqlite3.connect(temp_copy)
cur = conn.cursor()

# === TIME RANGE (UTC) ===
start = datetime.datetime(target_date.year, target_date.month, target_date.day, 0, 0, 0, tzinfo=datetime.timezone.utc)
end = datetime.datetime(target_date.year, target_date.month, target_date.day, 23, 59, 59, tzinfo=datetime.timezone.utc)

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

# === WRITE EXCEL ===
df = pd.read_csv(OUT_PATH)
excel_path = OUT_PATH.replace(".csv", ".xlsx")
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Browsing History")

print(f"📊 Created Excel table: {excel_path}")
