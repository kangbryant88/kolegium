import os
import glob
import re

template_dir = r"c:\Users\Kangb\OneDrive\Documentos\RoboClass_Planner\templates"
files = glob.glob(os.path.join(template_dir, "*.html"))

for file_path in files:
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Find and Replace 1: class="card bg-dark text-white" or class="card bg-dark" -> class="card"
    new_content = content.replace('class="card bg-dark text-white"', 'class="card"')
    new_content = new_content.replace("class='card bg-dark text-white'", "class='card'")
    new_content = new_content.replace('class="card bg-dark"', 'class="card"')
    new_content = new_content.replace("class='card bg-dark'", "class='card'")

    # Find and Replace 2: remove text-white inside <p ...> or <span ...>
    # Using regex to find tags <p ...> or <span ...> that contain text-white and remove the word text-white
    # This regex is simple and works for standard cases
    def remove_text_white(match):
        tag_content = match.group(0)
        # remove text-white
        return tag_content.replace('text-white', '')

    new_content = re.sub(r'<(p|span)\b[^>]*>', remove_text_white, new_content)

    # Also clean up multiple spaces in class attributes that might be left
    new_content = new_content.replace('  ', ' ')

    if new_content != content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Modified: {os.path.basename(file_path)}")

print("Done.")
