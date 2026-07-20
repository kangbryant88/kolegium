import os
import glob

app_data = r'C:\Users\Kangb\.gemini\antigravity\brain'
logs = glob.glob(os.path.join(app_data, '*', '.system_generated', 'logs', 'overview.txt'))

targets = ['landing.html', 'mi_perfil.html', 'enlace_expirado.html', 'ver_expediente_temporal.html']

for log in logs:
    with open(log, 'r', encoding='utf-8') as f:
        content = f.read()
        for t in targets:
            if t in content:
                print(f"Found {t} in {log}")
