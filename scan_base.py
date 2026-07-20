with open('templates/base.html', 'r', encoding='utf-8') as f:
    lines = f.readlines()
targets = ['current_user', 'dashboard', 'EduPlanner OS', 'usuario.role', 'usuario.nombre_completo']
for i, line in enumerate(lines, 1):
    for t in targets:
        if t in line:
            print(f'Line {i}: {line.rstrip()}')
