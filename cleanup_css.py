import os

files = [
    'templates/gestion_personal.html',
    'templates/mi_perfil.html',
    'templates/portal_trabajador.html'
]

for file_path in files:
    if not os.path.exists(file_path):
        continue
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Textos hardcodeados
    content = content.replace('text-dark', '')
    content = content.replace('text-white', '')
    
    # Fondos hardcodeados
    content = content.replace('bg-white', 'bg-body')
    content = content.replace('bg-light', 'bg-body-tertiary')
    
    # Fix possible double spaces left by removing classes
    content = content.replace('  ', ' ')
    content = content.replace('class=" ', 'class="')
    content = content.replace(' "', '"')
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)

print('Done.')
