#!/usr/bin/env python3
"""
Script para listar y analizar todos los archivos de recetas
Muestra la estructura actual y propone una mejor organización
"""

import os
from pathlib import Path
from datetime import datetime

def analyze_folder(folder_path="."):
    """Analiza la carpeta y lista los archivos"""
    
    # Detectar tipos de archivos
    docx_files = []
    gdoc_files = []  # Por si hay info sobre Google Docs
    
    files = list(Path(folder_path).glob("*.docx"))
    
    print("="*80)
    print("📂 ANÁLISIS DE LA CARPETA DE RECETAS")
    print("="*80)
    print(f"\n📍 Carpeta: {os.path.abspath(folder_path)}")
    print(f"📊 Total de archivos DOCX: {len(files)}\n")
    
    if not files:
        print("❌ No se encontraron archivos DOCX")
        return
    
    # Categorizar archivos
    categories = {
        'Receptas Individuales (l\'àvia)': [],
        'Receptas Ramen': [],
        'Receptas Asiáticas': [],
        'Receptas Originales (Google Docs convertidos)': [],
        'Otros': []
    }
    
    for file in sorted(files):
        filename = file.name
        size = file.stat().st_size / 1024  # KB
        modified = datetime.fromtimestamp(file.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        
        file_info = {
            'name': filename,
            'size_kb': size,
            'modified': modified
        }
        
        # Categorizar
        if 'Recepta' in filename and 'l\'àvia' not in filename and 'Ramen' not in filename:
            if any(x in filename for x in ['Pizza', 'Pad', 'Amanida', 'Pastís', 'Mandonguilles', 'Estofat']):
                categories['Receptas Originales (Google Docs convertidos)'].append(file_info)
            else:
                categories['Otros'].append(file_info)
        elif 'Recepta' in filename and any(x in filename for x in ['Canelons', 'Peus', 'Callos', 'Sípia', 'Bunyols', 'Patates', 'Rostit', 'Arròs', 'Mongetes', 'Macarrons', 'Bacallà', 'Crema', 'Pastís de llimona', 'Pasta']):
            categories['Receptas Individuales (l\'àvia)'].append(file_info)
        elif 'Ramen' in filename:
            categories['Receptas Ramen'].append(file_info)
        elif any(x in filename for x in ['Thai', 'curry', 'chop_suey', 'Salsa']):
            categories['Receptas Asiáticas'].append(file_info)
        else:
            categories['Otros'].append(file_info)
    
    # Mostrar por categoría
    total = 0
    for category, items in categories.items():
        if items:
            print(f"\n{'='*80}")
            print(f"📌 {category} ({len(items)} archivos)")
            print(f"{'='*80}")
            
            for i, file_info in enumerate(items, 1):
                size_str = f"{file_info['size_kb']:.1f} KB" if file_info['size_kb'] < 1024 else f"{file_info['size_kb']/1024:.1f} MB"
                print(f"\n{i}. {file_info['name']}")
                print(f"   📏 Tamaño: {size_str}")
                print(f"   🕐 Modificado: {file_info['modified']}")
            
            total += len(items)
    
    # Resumen
    print(f"\n{'='*80}")
    print("📊 RESUMEN")
    print(f"{'='*80}")
    print(f"\n✓ Receptas Individuales (l'àvia):        {len(categories['Receptas Individuales (l\'àvia)'])} archivos")
    print(f"✓ Receptas Ramen:                        {len(categories['Receptas Ramen'])} archivos")
    print(f"✓ Receptas Asiáticas:                    {len(categories['Receptas Asiáticas'])} archivos")
    print(f"✓ Receptas Originales (Google Docs):    {len(categories['Receptas Originales (Google Docs convertidos)'])} archivos")
    print(f"✓ Otros:                                 {len(categories['Otros'])} archivos")
    print(f"\n🔢 TOTAL:                               {total} archivos")
    
    # Propuesta de organización
    print(f"\n{'='*80}")
    print("💡 PROPUESTA DE ORGANIZACIÓN")
    print(f"{'='*80}")
    print("""
Opción 1: Agrupar en CARPETAS
├── 📁 Caldos
│   ├── Recepta Ramen - Caldos.docx
│   └── (futuros caldos)
├── 📁 Asiáticas
│   ├── Recepta Pad Kra-Prao Adaptat.docx
│   ├── Amanida de cabdells amb gambes vermelles asiàtica.docx
│   ├── Pollo_con_curry_tailandes.docx
│   └── ...
├── 📁 L'Àvia (Tradicional)
│   ├── Recepta Canelons de Nadal x90.docx
│   ├── Recepta Peus de porc.docx
│   └── ...
├── 📁 Complementos
│   └── Recepta Ramen - Complementos.docx
└── 📁 Originales
    ├── Recepta Masa de Pizza Perfecta KitchenAid 2 persones.docx
    ├── Recepta Pastís del Guifré.docx
    └── ...

Opción 2: Renombrar ARCHIVOS con prefijo de categoría
├── [CALDO] Recepta Ramen - Caldos.docx
├── [CALDO] Recepta Tonkotsu - Brou.docx
├── [CALDO] Recepta Yasai Miso - Brou.docx
├── [ASIÀTIC] Recepta Pad Kra-Prao.docx
├── [AVIA] Recepta Canelons de Nadal.docx
├── [COMPLEMENTOS] Recepta Ramen - Complementos.docx
└── ...

Opción 3: Google Drive con Labels/Tags
- Todos en Google Drive
- Usar colores y labels de Google Drive para categorizar
- Usar el buscador simplificado que ya funciona
""")
    
    return categories

if __name__ == "__main__":
    folder = "."  # Carpeta actual, cambiar si es necesario
    analyze_folder(folder)