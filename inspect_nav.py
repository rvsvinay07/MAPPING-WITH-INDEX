import os

downloads_dir = r"c:\Users\Anjani\Downloads"
filepath = os.path.join(downloads_dir, "AlphaNiftyoriginal.html")

with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
    lines = f.readlines()

# Search for nav or header or logo
print("Searching for header/nav/logo in AlphaNiftyoriginal.html:")
for i, line in enumerate(lines):
    if any(x in line.lower() for x in ["nav", "header", "logo", "brand", "menu"]):
        print(f"Line {i+1}: {line.strip()[:100]}")
