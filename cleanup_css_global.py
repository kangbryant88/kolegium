import os
import glob
import re

template_dir = r"c:\Users\Kangb\OneDrive\Documentos\RoboClass_Planner\templates"
files = glob.glob(os.path.join(template_dir, "*.html"))

# These files extend base.html and need dark mode compatibility
# Exclude standalone pages that don't use data-bs-theme: login.html, registro.html, landing.html
exclude = {'login.html', 'registro.html', 'landing.html', 'enlace_expirado.html',
           'inscripcion_inicial.html', 'ver_expediente_temporal.html', 'reporte.html'}

for file_path in files:
    basename = os.path.basename(file_path)
    if basename in exclude:
        continue
        
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original = content
    
    # Only replace inside <style> blocks
    def fix_style_block(match):
        style_content = match.group(0)
        # Replace background: #fff; with var(--surface-container-lowest)
        style_content = re.sub(
            r'background:\s*#fff\s*;',
            'background: var(--surface-container-lowest);',
            style_content
        )
        # Replace background: #ffffff; with var(--surface-container-lowest)
        style_content = re.sub(
            r'background:\s*#ffffff\s*;',
            'background: var(--surface-container-lowest);',
            style_content
        )
        # Replace background-color: #fff;
        style_content = re.sub(
            r'background-color:\s*#fff\s*;',
            'background-color: var(--surface-container-lowest);',
            style_content
        )
        # Replace hardcoded border colors #f2f4f6 with var
        style_content = re.sub(
            r'border[^:]*:\s*1px solid #f2f4f6\s*;',
            lambda m: m.group(0).replace('#f2f4f6', 'var(--surface-container-high)'),
            style_content
        )
        # Replace color: #191c1e (hardcoded on-surface)
        style_content = re.sub(
            r'color:\s*#191c1e\s*;',
            'color: var(--on-surface);',
            style_content
        )
        return style_content
    
    content = re.sub(r'<style>.*?</style>', fix_style_block, content, flags=re.DOTALL)
    
    if content != original:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Modified: {basename}")

print("Done.")
