# Receptes

Receptari familiar i personal: una web estàtica (GitHub Pages) amb buscador i filtres per etiquetes, i un backend a Google Apps Script per afegir i editar receptes, també a partir de text o fotos amb Gemini.

🌐 **Web:** https://miquelpairo.github.io/receptes/

## Com funciona

```
┌────────────────────────────┐
│ index.html (GitHub Pages)  │  ← llegeix receptes_index.json
└─────────────┬──────────────┘
              │ POST (acció + contrasenya)
┌─────────────▼──────────────┐
│ Google Apps Script         │  ── Gemini (text / fotos → recepta)
│ apps-script/Code.gs        │
└─────────────┬──────────────┘
              │ GitHub API (commit)
┌─────────────▼──────────────┐      ┌──────────────┐
│ receptes_index.json        │ ───► │ còpia a Drive│
└────────────────────────────┘      └──────────────┘
```

- La web és només lectura: carrega `receptes_index.json` i filtra al navegador.
- Qualsevol canvi (crear, editar, esborrar, etiquetes) passa per l'Apps Script, que valida la contrasenya, fa el commit del JSON a GitHub i en desa una còpia a Drive. Per això l'historial té commits "(admin web)" i "via inbox".
- També hi ha una **inbox a Drive**: els Google Docs amb el nom `RECEPTA…` es processen amb `processInbox()` i s'afegeixen automàticament. Els correctes van a `processats/` i els que fallen a `revisar/`.

## Funcions de la web

- Cerca per text i filtres per etiquetes, agrupats per eixos.
- Vista de targetes o vista compacta.
- **Mode admin** (botó de contrasenya):
  - **admin:** crear, editar, esborrar, canviar etiquetes i afegir etiquetes noves a la taxonomia.
  - **família:** només crear receptes noves.
- **Importar amb Gemini:** enganxar text d'un blog o pujar fins a 4 fotos (també manuscrites) i obtenir la recepta estructurada per revisar-la abans de desar.

## Estructura del repositori

| Ruta | Contingut |
|---|---|
| `index.html` | La web (HTML + CSS + JS en un sol fitxer) |
| `receptes_index.json` | Les dades: totes les receptes i la taxonomia d'etiquetes |
| `apps-script/Code.gs` | Codi del backend (Google Apps Script) |
| `Scripts/` | Scripts Python per importar i indexar els DOCX originals |
| `Coleccions/`, `Receptes_Individuales/`, `Processed/` | Documents Word originals |

## Format de `receptes_index.json`

```json
{
  "total": 120,
  "generated": false,
  "taxonomy": { "Cuina": { "select": "one", "required": true, "tags": ["Català", "…"] }, "…": {} },
  "recipes": [
    {
      "id": "pollastre-teriyaki",
      "name": "Pollastre teriyaki",
      "source": "Web",
      "added": "2026-09-07",
      "tags": ["Asiàtic", "Japonès", "Principal", "Carn"],
      "ingredients": ["…"],
      "instructions": "1. …\n2. …"
    }
  ]
}
```

- `id`: el nom en format *slug*. Ha de ser únic.
- `source`: `Gemini`, `Web`, `DOCX` o `DOCX (recuperada)`.
- Les receptes es desen ordenades alfabèticament i `total` s'actualitza sol.

### Taxonomia

Cada eix defineix si s'hi pot triar una o diverses etiquetes (`select`) i si és obligatori (`required`). L'Apps Script rebutja receptes amb etiquetes fora del vocabulari.

| Eix | Selecció | Obligatori |
|---|---|---|
| Cuina | una | sí |
| Estil asiàtic | diverses | no |
| Tipus de plat | una | sí |
| Ingredient | diverses | no |
| Base | diverses | no |
| Família | diverses | no |

Les etiquetes noves s'afegeixen des de la web en mode admin; no cal editar el JSON a mà.

## Backend (Apps Script)

El codi és a `apps-script/Code.gs`. Està desplegat com a aplicació web i la URL és a `ADMIN_ENDPOINT` dins d'`index.html`.

**Script Properties necessàries** (mai al codi ni al repositori):

| Propietat | Ús |
|---|---|
| `GITHUB_TOKEN` | Token amb permís d'escriptura al contingut d'aquest repositori |
| `GEMINI_API_KEY` | Clau de l'API de Gemini |
| `ADMIN_PASSWORD` | Contrasenya del rol admin |
| `FAMILY_PASSWORD` | Contrasenya del rol família |

**Accions (`doPost`)**: `auth`, `saveRecipe`, `parseText`, `parseImages` (admin i família) i `delete`, `updateTags`, `addTag` (només admin).

**Inbox**: `processInbox()` es pot programar amb un activador de temps.

Després de canviar el codi cal fer **Implementa → Gestiona implementacions → Edita → Versió nova** perquè la URL publicada faci servir la versió nova.

## Scripts Python

Es van fer servir per a la importació inicial des de Word. Requereixen `python-docx`.

- `receptes_indexer.py`: genera un índex JSON a partir dels DOCX i Google Docs.
- `dividir_receptes_avia.py`, `dividir_ramen.py`, `analizar_ramen.py`: divideixen els reculls en receptes individuals.
- `listar_archivos.py`, `explorar_estructura.py`, `organizar_carpeta.py`: utilitats per ordenar la carpeta.

`credentials.json` i `token.json` (accés a Google) estan exclosos al `.gitignore` i no s'han de pujar mai.
