import os
import json
import re

logs = [
    r"C:\Users\Kangb\.gemini\antigravity\brain\28f7bb1e-6008-4196-9332-56921d4efcea\.system_generated\logs\overview.txt",
    r"C:\Users\Kangb\.gemini\antigravity\brain\6fd31cfd-971f-4f55-a000-f4e7cb3cb749\.system_generated\logs\overview.txt",
    r"C:\Users\Kangb\.gemini\antigravity\brain\d6b0cd5a-2c85-477c-9cb5-4eaebe6c0a5e\.system_generated\logs\overview.txt"
]

targets = {
    'landing.html': '',
    'mi_perfil.html': '',
    'enlace_expirado.html': '',
    'ver_expediente_temporal.html': ''
}

def extract_file_content(log_path):
    with open(log_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # We will search for Tool Call: default_api:write_to_file
    # and extract the CodeContent and TargetFile using regex
    
    # Simple JSON extraction regex (since the arguments are printed as JSON in the overview)
    matches = re.finditer(r'\{[^{}]*"TargetFile":\s*"[^"]*(landing\.html|mi_perfil\.html|enlace_expirado\.html|ver_expediente_temporal\.html)"[^{}]*"CodeContent":\s*"([^"]*)"', content)
    
    for match in matches:
        filename = match.group(1)
        code = match.group(2)
        # Handle JSON escaped newlines and quotes
        code = code.replace('\\n', '\n').replace('\\"', '"').replace('\\\\', '\\')
        if len(code) > len(targets[filename]):
            targets[filename] = code

    # It might also be from view_file output.
    # We can search for File Path: `...landing.html` followed by lines
    lines = content.split('\n')
    current_file = None
    current_code = []
    capture = False
    for line in lines:
        if 'File Path: `file:///' in line and any(t in line for t in targets.keys()):
            for t in targets.keys():
                if t in line:
                    current_file = t
                    current_code = []
                    capture = False
                    break
        elif current_file and 'The following code has been modified to include a line number' in line:
            capture = True
        elif current_file and capture:
            if line.startswith('The above content does NOT show the entire file contents') or 'Tool Call:' in line or line == '':
                if len('\n'.join(current_code)) > len(targets[current_file]):
                    targets[current_file] = '\n'.join(current_code)
                capture = False
                current_file = None
            else:
                # Remove line numbers like "1: " or "100: "
                clean_line = re.sub(r'^\d+:\s', '', line)
                current_code.append(clean_line)

for log in logs:
    extract_file_content(log)

out_dir = r"c:\Users\Kangb\OneDrive\Documentos\RoboClass_Planner\templates"
for t, code in targets.items():
    if code:
        with open(os.path.join(out_dir, t), 'w', encoding='utf-8') as f:
            f.write(code)
        print(f"Recovered {t} from logs ({len(code)} bytes).")
    else:
        print(f"Could not find full code for {t}.")
