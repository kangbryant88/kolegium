import os
import glob
import re

template_dir = r"c:\Users\Kangb\OneDrive\Documentos\RoboClass_Planner\templates"
files = glob.glob(os.path.join(template_dir, "*.html"))

def remove_classes(match):
    # match is the whole tag string: e.g. <div class="text-dark mb-2">
    tag_string = match.group(0)
    
    # We want to replace occurrences of \btext-dark\b and \btext-muted\b with empty string
    tag_string = re.sub(r'\btext-dark\b', '', tag_string)
    tag_string = re.sub(r'\btext-muted\b', '', tag_string)
    
    # Also for base.html, we remove bg-light and bg-white
    # We will do this globally inside tags, or we can do it specifically in base.html below
    return tag_string

for file_path in files:
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
    original = content
    
    # regex to match any HTML tag and apply the removal inside it
    # <tagname ... class="..." ...>
    content = re.sub(r'<[^>]+>', remove_classes, content)

    # specifically for base.html, remove bg-light and bg-white
    if os.path.basename(file_path) == "base.html":
        content = re.sub(r'\bbg-light\b', '', content)
        content = re.sub(r'\bbg-white\b', '', content)

    if content != original:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Modified: {os.path.basename(file_path)}")

print("Done.")
