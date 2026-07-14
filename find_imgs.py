import os
import re

downloads_dir = r"c:\Users\Anjani\Downloads"
with open(os.path.join(downloads_dir, "AlphaNiftyoriginal.html"), "r", encoding="utf-8", errors="ignore") as f:
    content = f.read()

# Find all image tags
img_tags = re.findall(r"<img[^>]*src=[\"']([^\"']*)[\"'][^>]*>", content, re.IGNORECASE)
print("Image tags in AlphaNiftyoriginal.html:")
for img in img_tags:
    print(img)
