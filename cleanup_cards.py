import os
import glob
import re

template_dir = r"c:\Users\Kangb\OneDrive\Documentos\RoboClass_Planner\templates"
files = glob.glob(os.path.join(template_dir, "*.html"))

def clean_card_classes(match):
    tag_string = match.group(0)
    # Only clean if it's a card or main container
    # But user specifically said "Tarjetas en gestion_personal.html, estadistica.html y el dashboard"
    # and "Tarjetas usen EXCLUSIVAMENTE <div class='card'>"
    
    # Let's remove the bad classes from the whole tag string if it's a div
    tag_string = re.sub(r'\bbg-dark\b', '', tag_string)
    tag_string = re.sub(r'\bbg-light\b', '', tag_string)
    tag_string = re.sub(r'\bbg-white\b', '', tag_string)
    # Only remove text-white and text-dark if we're dealing with a card to avoid breaking badges/buttons
    if 'card' in tag_string or 'content-area' in tag_string or 'main' in tag_string or '<body' in tag_string:
        tag_string = re.sub(r'\btext-white\b', '', tag_string)
        tag_string = re.sub(r'\btext-dark\b', '', tag_string)
        
    # Clean up multiple spaces in class
    tag_string = re.sub(r'\s+', ' ', tag_string)
    tag_string = tag_string.replace('class=" "', 'class=""')
    
    return tag_string

for file_path in files:
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
    original = content
    
    # We apply this to ALL div, main, body, and section tags
    content = re.sub(r'<(div|main|body|section)[^>]*class="[^"]*"[^>]*>', clean_card_classes, content)

    if content != original:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Modified: {os.path.basename(file_path)}")

print("Done.")
