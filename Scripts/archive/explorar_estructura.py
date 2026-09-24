#!/usr/bin/env python3
"""
Script de exploración - Analiza la estructura actual de la carpeta
Muestra cómo quedó todo después de los cambios
"""

import os
from pathlib import Path
from datetime import datetime

def explore_folder(root_path=".", level=0, max_level=3):
    """Explora recursivamente la carpeta y muestra la estructura"""
    
    path = Path(root_path)
    
    if level == 0:
        print("="*100)
        print("🔍 EXPLORACIÓN DE ESTRUCTURA")
        print("="*100)
        print(f"\n📍 Carpeta raíz: {os.path.abspath(root_path)}\n")
    
    # Limitar profundidad
    if level > max_level:
        return
    
    try:
        items = sorted(path.iterdir())
    except PermissionError:
        return
    
    # Separar carpetas y archivos
    folders = [item for item in items if item.is_dir() and not item.name.startswith('.')]
    files = [item for item in items if item.is_file() and not item.name.startswith('.')]
    
    # Mostrar carpetas
    for i, folder in enumerate(folders):
        is_last_folder = (i == len(folders) - 1) and len(files) == 0
        prefix = "└── " if is_last_folder else "├── "
        print(f"{'  ' * level}{prefix}📁 {folder.name}/")
        
        # Explorar subcarpeta
        explore_folder(folder, level + 1, max_level)
    
    # Mostrar archivos
    for i, file in enumerate(files):
        is_last = i == len(files) - 1
        prefix = "└── " if is_last else "├── "
        
        # Información del archivo
        size_kb = file.stat().st_size / 1024
        modified = datetime.fromtimestamp(file.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        
        # Icono según tipo
        if file.suffix == ".docx":
            icon = "📄"
        elif file.suffix == ".py":
            icon = "🐍"
        elif file.suffix == ".js":
            icon = "📜"
        else:
            icon = "📦"
        
        size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
        print(f"{'  ' * level}{prefix}{icon} {file.name} ({size_str}) - {modified}")

def count_files():
    """Cuenta archivos por tipo"""
    
    print("\n\n" + "="*100)
    print("📊 CONTADOR DE ARCHIVOS")
    print("="*100 + "\n")
    
    docx_files = list(Path(".").rglob("*.docx"))
    py_files = list(Path(".").rglob("*.py"))
    js_files = list(Path(".").rglob("*.js"))
    other_files = []
    
    print(f"📄 DOCX (Recetas):           {len(docx_files)} archivos")
    for f in sorted(docx_files):
        if f.parent.name in [".", "Coleccions"]:
            print(f"   └─ {f.relative_to('.')}")
    
    print(f"\n🐍 Python Scripts:           {len(py_files)} archivos")
    for f in sorted(py_files):
        print(f"   └─ {f.relative_to('.')}")
    
    print(f"\n📜 Google Apps Scripts:      {len(js_files)} archivos")
    for f in sorted(js_files):
        print(f"   └─ {f.relative_to('.')}")
    
    print(f"\n{'='*100}")
    print(f"🔢 TOTAL: {len(docx_files) + len(py_files) + len(js_files)} archivos")

def categorize_docx():
    """Categoriza los DOCX"""
    
    print("\n\n" + "="*100)
    print("📋 CATEGORIZACIÓN DE RECETAS DOCX")
    print("="*100 + "\n")
    
    docx_files = list(Path(".").rglob("*.docx"))
    
    categories = {
        'Ramen': [],
        'L\'Àvia': [],
        'Asiáticas': [],
        'Originales': [],
        'Otros': []
    }
    
    for f in sorted(docx_files):
        filename = f.name.lower()
        path_str = str(f).lower()
        
        if 'ramen' in filename:
            categories['Ramen'].append(str(f.relative_to('.')))
        elif any(x in filename for x in ['avia', 'canelons', 'peus', 'callos', 'sípia', 'bunyols', 'patates', 'rostit', 'arròs', 'mongetes', 'macarrons', 'bacallà', 'crema', 'pastís de llimona', 'pasta']):
            categories['L\'Àvia'].append(str(f.relative_to('.')))
        elif any(x in filename for x in ['pad', 'thai', 'curry', 'chop_suey', 'salsa_', 'pescado', 'carne']):
            categories['Asiáticas'].append(str(f.relative_to('.')))
        elif any(x in filename for x in ['pizza', 'pastís del', 'mandonguilles xai', 'estofat', 'amanida', 'masa']):
            categories['Originales'].append(str(f.relative_to('.')))
        else:
            categories['Otros'].append(str(f.relative_to('.')))
    
    for category, files in categories.items():
        if files:
            print(f"📌 {category}: {len(files)} archivos")
            for f in files:
                print(f"   └─ {f}")
            print()

def main():
    explore_folder()
    count_files()
    categorize_docx()
    
    print("\n" + "="*100)
    print("💡 ANÁLISIS")
    print("="*100)
    print("""
Ahora que ves la estructura actual, podemos:

1. ✓ LIMPIAR manualmente los archivos duplicados/antiguos
2. ✓ REORGANIZAR en carpetas de forma más lógica
3. ✓ ELIMINAR los DOCX que no necesites

¿Cuál es el problema específico que ves?
- ¿Archivos duplicados?
- ¿Clasificación incorrecta?
- ¿Carpetas mal creadas?
- ¿Documentos antiguos sin eliminar?
""")

if __name__ == "__main__":
    main()