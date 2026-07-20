import os
import glob
import json
import shutil
import urllib.parse

history_dir = r"C:\Users\Kangb\AppData\Roaming\Code\User\History"
project_dir_lower = "roboclass_planner"

recovered_count = 0

for d in os.listdir(history_dir):
    d_path = os.path.join(history_dir, d)
    if os.path.isdir(d_path):
        entries_file = os.path.join(d_path, 'entries.json')
        if os.path.exists(entries_file):
            try:
                with open(entries_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                resource = urllib.parse.unquote(data.get('resource', '')).lower()
                
                if project_dir_lower in resource and resource.endswith('.html'):
                    basename = os.path.basename(resource)
                    target_file = os.path.join(r"c:\Users\Kangb\OneDrive\Documentos\RoboClass_Planner\templates", basename)
                    
                    entries = data.get('entries', [])
                    if entries:
                        latest_entry = entries[-1]
                        source_id = latest_entry.get('id')
                        source_file = os.path.join(d_path, source_id)
                        
                        if os.path.exists(source_file):
                            shutil.copy2(source_file, target_file)
                            print(f"Recovered loosely: {basename}")
                            recovered_count += 1
            except Exception as e:
                print(f"Error processing {entries_file}: {e}")

print(f"Total files recovered: {recovered_count}")
