import os
import glob
import json
import shutil
import urllib.parse

history_dir = r"C:\Users\Kangb\AppData\Roaming\Code\User\History"
project_dir = r"c:/Users/Kangb/OneDrive/Documentos/RoboClass_Planner/templates"
# lowercase for case-insensitive match
project_dir_lower = project_dir.lower()

# For each directory in history
recovered_count = 0

for d in os.listdir(history_dir):
    d_path = os.path.join(history_dir, d)
    if os.path.isdir(d_path):
        entries_file = os.path.join(d_path, 'entries.json')
        if os.path.exists(entries_file):
            try:
                with open(entries_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                resource = urllib.parse.unquote(data.get('resource', ''))
                # convert file:///c:/... to c:/...
                if resource.startswith('file:///'):
                    resource = resource[8:]
                
                if resource.lower().startswith(project_dir_lower) and resource.lower().endswith('.html'):
                    # Found a match!
                    target_file = os.path.join(r"c:\Users\Kangb\OneDrive\Documentos\RoboClass_Planner\templates", os.path.basename(resource))
                    
                    # Find the latest file in the entries
                    entries = data.get('entries', [])
                    if entries:
                        latest_entry = entries[-1]
                        source_id = latest_entry.get('id')
                        source_file = os.path.join(d_path, source_id)
                        
                        if os.path.exists(source_file):
                            shutil.copy2(source_file, target_file)
                            print(f"Recovered: {os.path.basename(resource)}")
                            recovered_count += 1
            except Exception as e:
                print(f"Error processing {entries_file}: {e}")

print(f"Total files recovered: {recovered_count}")
