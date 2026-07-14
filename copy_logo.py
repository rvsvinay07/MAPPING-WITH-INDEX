import shutil
import os

src = r"c:\Users\Anjani\Downloads\AlphaNifty_HD_Logo-removebg-preview.avif"
dest = r"c:\Users\Anjani\Downloads\MATCH BY INDEX\logo.avif"

try:
    if os.path.exists(src):
        shutil.copy(src, dest)
        print("Success: Copied AlphaNifty logo to workspace as logo.avif")
    else:
        print("Source logo file not found in Downloads")
except Exception as e:
    print(f"Error copying file: {e}")
