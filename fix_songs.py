import os

SONGS_DIR = r"E:\YARG\Setlists\Songs"

has_ini = 0
has_chart_only = 0
has_neither = 0
failed_ini = []

for root, dirs, files in os.walk(SONGS_DIR):
    lower = [f.lower() for f in files]
    if 'song.ini' in lower:
        has_ini += 1
    elif any(f.endswith('.chart') or f.endswith('.mid') for f in lower):
        has_chart_only += 1
        failed_ini.append(root)
    elif any(f.endswith('.ogg') or f.endswith('.mp3') for f in lower):
        has_neither += 1

print(f"Has song.ini:       {has_ini}")
print(f"Chart/mid only:     {has_chart_only}")
print(f"No metadata found:  {has_neither}")
print(f"\nTotal detected:     {has_ini + has_chart_only + has_neither}")
print(f"\nFirst 10 chart-only folders:")
for p in failed_ini[:10]:
    print(" ", p)