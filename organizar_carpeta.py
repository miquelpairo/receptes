#!/usr/bin/env python3
"""
Script para limpiar y organizar la carpeta de Receptes
- Crea carpeta Coleccions/ con Las colecciones (RAMEN, l'Àvia)
- Crea carpeta Scripts/ con todos los scripts (.py, .js, etc.)
- Deja limpia la carpeta principal
"""

import os
import shutil
from pathlib import Path

def create_folder_structure():
    """Crea la estructura de carpetas"""
    
    current_dir = Path(".")
    coleccions = current_dir / "Coleccions"
    scripts = current_dir / "Scripts"
    
    # Crear carpetas si no existen
    coleccions.mkdir(exist_ok=True)
    scripts.mkdir(exist_ok=True)
    
    print("="*80)
    print("🧹 LIMPIADOR DE CARPETA RECEPTES")
    print("="*80 + "\n")
    
    # Archivos a mover a Coleccions
    coleccions_items = [
        # RAMEN original y generados
        "Curs de RAMEN.docx",
        "Recepta Ramen - Caldos.docx",
        "Recepta Ramen - Complementos.docx",
        
        # L'Àvia original y generados
        "Còpia de Les receptes de l'àvia.docx",
    ]
    
    # Buscar DOCX de l'Àvia en subcarpetas
    receptes_individuales = current_dir / "Receptes_Individuales"
    if receptes_individuales.exists():
        avia_docx_files = list(receptes_individuales.glob("*.docx"))
        coleccions_items.extend([str(f) for f in avia_docx_files])
    
    # Scripts a mover (por extensión y nombre)
    script_patterns = ["*.py", "*.js", "google_apps_script*", "*script*"]
    script_files = set()
    
    for pattern in script_patterns:
        script_files.update(current_dir.glob(pattern))
    
    # Remover archivos DOCX de los scripts detectados
    script_files = [f for f in script_files if f.suffix != ".docx"]
    
    # Copiar a Coleccions
    print("📂 COLECCIONS")
    print("─" * 80)
    
    if coleccions_items:
        for item_name in coleccions_items:
            item_path = current_dir / item_name
            
            if item_path.exists():
                if item_path.is_file():
                    dest = coleccions / item_path.name
                    try:
                        shutil.copy2(item_path, dest)
                        print(f"✓ Copiado: {item_path.name}")
                    except Exception as e:
                        print(f"⚠ Error copiando {item_path.name}: {e}")
            else:
                print(f"⚠ No encontrado: {item_name}")
        
        # Copiar carpeta Receptes_Individuales completa
        if receptes_individuales.exists():
            dest_individuales = coleccions / "Receptes_Individuales"
            if dest_individuales.exists():
                shutil.rmtree(dest_individuales)
            try:
                shutil.copytree(receptes_individuales, dest_individuales)
                print(f"✓ Carpeta copiada: Receptes_Individuales/")
            except Exception as e:
                print(f"⚠ Error copiando carpeta: {e}")
    
    print("\n📚 SCRIPTS")
    print("─" * 80)
    
    if script_files:
        for script_file in sorted(script_files):
            if script_file.name not in ["__pycache__"]:
                dest = scripts / script_file.name
                try:
                    if script_file.is_file():
                        shutil.copy2(script_file, dest)
                        print(f"✓ Movido: {script_file.name}")
                except Exception as e:
                    print(f"⚠ Error moviendo {script_file.name}: {e}")
    
    # Mostrar resumen
    print("\n" + "="*80)
    print("✅ RESUMEN")
    print("="*80)
    
    coleccions_count = len(list(coleccions.glob("**/*")))
    scripts_count = len(list(scripts.glob("**/*")))
    
    print(f"\n📁 Carpeta 'Coleccions/'")
    print(f"   Archivos/carpetas: {coleccions_count}")
    print(f"   └─ Ramen (Caldos + Complementos)")
    print(f"   └─ Les receptes de l'Àvia (original + individuales)")
    
    print(f"\n📁 Carpeta 'Scripts/'")
    print(f"   Archivos: {scripts_count}")
    print(f"   └─ Todos los scripts Python y Google Apps Script")
    
    print(f"\n{'='*80}")
    print("💡 PRÓXIMOS PASOS:")
    print(f"{'='*80}")
    print("""
1. ✓ Archivos DOCX de recetas finales quedan en carpeta principal
2. ✓ Colecciones (RAMEN + l'Àvia) guardadas en 'Coleccions/'
3. ✓ Scripts organizados en 'Scripts/'

4. Luego: Subir a Google Drive
   - Eliminar: Curs de RAMEN.docx y Còpia de Les receptes de l'àvia.docx
   - Mantener: DOCX individuales en carpeta principal
   - Subir TODO a Google Drive en carpeta Receptes
""")

if __name__ == "__main__":
    try:
        create_folder_structure()
        print("\n✅ ¡Organización completada!")
    except Exception as e:
        print(f"\n❌ Error: {e}")
