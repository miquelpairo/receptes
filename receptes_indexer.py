#!/usr/bin/env python3
"""
Script robusto para indexar recetas DOCX + Google Docs
VERSIÓN MEJORADA: No descarta ninguna receta y procesa todos los archivos individuales.
"""

import json
import re
import os
from pathlib import Path
from docx import Document

print("🍳 Indexador de Receptes (Versió Millorada)\n")

# Intentar importar Google de forma segura
GOOGLE_AVAILABLE = False
try:
    import importlib
    importlib.import_module('google.auth.oauthlib.flow')
    from google.auth.oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    GOOGLE_AVAILABLE = True
    print("✓ Llibreries de Google disponibles\n")
except ImportError as e:
    print(f"⚠️ Llibreries de Google no disponibles ({e})")
    print("  Continuant només amb arxius DOCX locals...\n")

RECIPE_FOLDER_ID = "1DcRhULGxLapaXu6liHuTelbj75Tob450"

def get_drive_service():
    if not GOOGLE_AVAILABLE: return None
    SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("   ⚠️ No trobat: credentials.json\n")
                return None
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)

def list_folder_files(service, folder_id):
    try:
        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            spaces='drive',
            fields='files(id, name, mimeType)',
            pageSize=1000
        ).execute()
        return results.get('files', [])
    except Exception as e:
        print(f"   Error llistant carpeta: {e}")
        return []

def get_google_doc_content(service, doc_id):
    try:
        docs_service = build('docs', 'v1', credentials=service._http.request.credentials)
        doc = docs_service.documents().get(documentId=doc_id).execute()
        text = ''
        if 'body' in doc and 'content' in doc['body']:
            for elem in doc['body']['content']:
                if 'paragraph' in elem:
                    text += elem['paragraph'].get('rawText', '') + '\n'
                elif 'table' in elem:
                    for row in elem['table']['tableRows']:
                        for cell in row['tableCells']:
                            for content in cell['content']:
                                if 'paragraph' in content:
                                    text += content['paragraph'].get('rawText', '') + ' '
                        text += '\n'
        return text
    except Exception as e:
        print(f"   Error llegint Google Doc: {e}")
        return ""

def extract_text_from_docx(filepath):
    try:
        doc = Document(filepath)
        text = ''
        for para in doc.paragraphs:
            text += para.text + '\n'
        return text
    except Exception as e:
        print(f"   Error al llegir DOCX: {e}")
        return ""

def parse_single_recipe(text, filename):
    """Nova versió: Mai falla, indexa el text sencer si no pot separar-ho."""
    if not text or len(text.strip()) < 10:
        return None
    
    # Netejar el nom de l'arxiu per usar-lo com a títol (fem servir flags en lloc de (?i))
    clean_name = re.sub(r'\.docx$', '', filename, flags=re.IGNORECASE)
    clean_name = re.sub(r'^Recepta\s+', '', clean_name, flags=re.IGNORECASE)
    clean_name = re.sub(r'^Receta\s+', '', clean_name, flags=re.IGNORECASE)
    clean_name = clean_name.replace('_', ' ').strip()
    
    recipe = {
        'name': clean_name,
        'source': 'DOCX' if '.docx' in filename.lower() else 'Google Docs',
        'ingredients': [],
        'instructions': '',
        'tags': []
    }
    
    full_text = text.strip()
    
    # Intentar extreure ingredients de manera flexible (Sense (?i) al mig)
    regex_pattern = r'(ingredients?|ingredientes|pel pastís|per la carn)[:\n]+(.*?)((?:\n\s*\n)|(preparació|elaboració|instruccions|pasos|steps|preparación))'
    ingredients_match = re.search(regex_pattern, full_text, flags=re.DOTALL | re.IGNORECASE)
    
    if ingredients_match:
        ing_text = ingredients_match.group(2)
        # Netejar i separar per salts de línia
        ings = [i.strip('-•* ').strip() for i in ing_text.split('\n') if len(i.strip()) > 2]
        recipe['ingredients'] = ings
        recipe['instructions'] = full_text # Guardem tot el text sencer per seguretat
    else:
        # Si no detecta la secció, agafa possibles ingredients de les primeres línies curtes
        lines = [line.strip() for line in full_text.split('\n') if line.strip()]
        potential_ings = []
        for line in lines[1:15]:
            if len(line) < 60 and (re.match(r'^\d', line) or line.startswith('-')):
                potential_ings.append(line.strip('-•* '))
        
        recipe['ingredients'] = potential_ings
        recipe['instructions'] = full_text

    # Tags automàtics millorats
    text_lower = full_text.lower()
    tags = set()
    if any(w in text_lower for w in ['porc', 'vedella', 'pollastre', 'xai', 'carn', 'carne', 'ternera']): tags.add('Carn')
    if any(w in text_lower for w in ['sípia', 'gambes', 'peix', 'bacallà', 'pescado', 'calamar', 'salmon']): tags.add('Peix/Seafood')
    if any(w in text_lower for w in ['forn', 'dessert', 'pastís', 'crema', 'xocolata', 'pastel', 'cheesecake']): tags.add('Postres')
    if any(w in text_lower for w in ['patates', 'pastanagues', 'ceba', 'mongetes', 'verdures']): tags.add('Verdures')
    if any(w in text_lower for w in ['arròs', 'pasta', 'macarrons', 'canelons', 'ramen', 'pizza']): tags.add('Arròs/Pasta')
    if any(w in text_lower for w in ['ramen', 'thai', 'tailandès', 'asiàtic', 'hoisin', 'soja', 'curry']): tags.add('Asiàtic')
    if 'tradicional' in text_lower or 'àvia' in filename.lower() or 'àvia' in text_lower: tags.add('Tradicional')
    
    recipe['tags'] = list(tags) if tags else ['Altres']
    
    return recipe
    
def parse_document(filename, text):
    """Procesa el document tractant-lo per defecte com una recepta única"""
    recipes = []
    
    # Si és el llibre original de l'àvia sencer (no el tallat), mirem de separar-lo
    if "receptes de l'àvia" in filename.lower() and "còpia" not in filename.lower() and len(text) > 2000:
        recipe_blocks = re.split(r'\n\s*#\s+', text)
        for block in recipe_blocks:
            if len(block.strip()) > 30:
                r = parse_single_recipe(block, block.split('\n')[0].strip())
                if r: recipes.append(r)
        return recipes

    # Tota la resta es considera 1 arxiu = 1 recepta
    recipe = parse_single_recipe(text, filename)
    if recipe:
        recipes.append(recipe)
        
    return recipes

# MAIN
all_recipes = []

# 1. DOCX locals (busca també en subcarpetes amb **/*.docx)
print("📁 Buscant DOCX locals...")
docx_files = list(Path('.').glob('**/*.docx'))

if docx_files:
    print(f"   {len(docx_files)} arxiu(s) trobats\n")
    for docx_file in docx_files:
        if '~$' in docx_file.name: continue # Ignora temporals
        print(f"   📄 {docx_file.name}...", end=" ")
        text = extract_text_from_docx(docx_file)
        recipes = parse_document(docx_file.name, text)
        if recipes:
            all_recipes.extend(recipes)
            print(f"✓ ({len(recipes)} receptes)")
        else:
            print("⚠ (Error o buida)")
else:
    print("   (cap)\n")

# 2. Google Docs
if GOOGLE_AVAILABLE:
    print("\n☁️  Connectant a Google Drive...")
    try:
        service = get_drive_service()
        if service:
            google_files = list_folder_files(service, RECIPE_FOLDER_ID)
            docs = [f for f in google_files if 'document' in f.get('mimeType', '') or f['name'].endswith('.docx')]
            
            if docs:
                print(f"   {len(docs)} arxiu(s) a Drive\n")
                for gfile in docs:
                    print(f"   📄 {gfile['name']}...", end=" ")
                    if 'document' in gfile.get('mimeType', ''):
                        text = get_google_doc_content(service, gfile['id'])
                        recipes = parse_document(gfile['name'], text)
                        if recipes:
                            all_recipes.extend(recipes)
                            print(f"✓")
                        else:
                            print("⚠")
                    else:
                        print("⏭️ (És DOCX, ja processat localment)")
            else:
                print("   (cap)\n")
    except Exception as e:
        print(f"   Error: {e}\n")

# Netejar duplicats (basat en el nom exacte)
print(f"\n{'='*50}")
unique_recipes = {}
for r in all_recipes:
    name_key = r['name'].lower().strip()
    # Si ja existeix, ens quedem amb el que tingui instruccions més llargues
    if name_key not in unique_recipes or len(r['instructions']) > len(unique_recipes[name_key]['instructions']):
        unique_recipes[name_key] = r
        
final_recipes = list(unique_recipes.values())

print(f"✅ Total indexades: {len(final_recipes)} receptes úniques (de {len(all_recipes)} analitzades)\n")

with open('receptes_index.json', 'w', encoding='utf-8') as f:
    json.dump({
        'total': len(final_recipes),
        'recipes': final_recipes,
        'generated': True
    }, f, ensure_ascii=False, indent=2)

print("✓ receptes_index.json creat correctament!")