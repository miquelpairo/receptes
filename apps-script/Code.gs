/**
 * Ingesta de receptes: Drive (Google Docs de Gemini) -> GitHub + espill a Drive.
 */

// ============================ CONFIG ============================
const CONFIG = {
  GITHUB_OWNER:  'miquelpairo',
  GITHUB_REPO:   'receptes',
  GITHUB_PATH:   'receptes_index.json',
  GITHUB_BRANCH: 'main',

  INBOX_FOLDER_ID:     '10T1-9JqgjFcZTctI3PuP17uzBAd2F973',
  MIRROR_FOLDER_ID:    '1DcRhULGxLapaXu6liHuTelbj75Tob450',
  PROCESSED_SUBFOLDER: 'processats',
  REVIEW_SUBFOLDER:    'revisar',
  MIRROR_FILENAME:     'receptes_index.json',

  RECIPE_MARKER:       'RECEPTA',
  GEMINI_MODEL:        'gemini-3.6-flash',
};

const FALLBACK_TAXONOMY = {
  "Cuina":         {"select":"one",  "required":true,  "tags":["Català","Asiàtic","Italià","Internacional"]},
  "Estil asiàtic": {"select":"many", "required":false, "tags":["Japonès","Xinès","Thai"]},
  "Tipus de plat": {"select":"one",  "required":true,  "tags":["Principal","Entrant","Postres","Salsa"]},
  "Ingredient":    {"select":"many", "required":false, "tags":["Carn","Peix","Vegetarià"]},
  "Base":          {"select":"many", "required":false, "tags":["Arròs","Pasta"]},
  "Família":       {"select":"many", "required":false, "tags":["Àvia","Erola"]}
};

function taxonomyVocab_(taxonomy) {
  const tax = (taxonomy && Object.keys(taxonomy).length) ? taxonomy : FALLBACK_TAXONOMY;
  const s = new Set();
  Object.values(tax).forEach(ax => (ax.tags || []).forEach(t => s.add(t)));
  return s;
}

function validateTags_(tags, taxonomy) {
  const tax = (taxonomy && Object.keys(taxonomy).length) ? taxonomy : FALLBACK_TAXONOMY;
  const vocab = taxonomyVocab_(tax);
  const bad = tags.filter(t => !vocab.has(t));
  if (bad.length) throw new Error('Tags fora del vocabulari: ' + bad.join(', '));
  Object.keys(tax).forEach(name => {
    const ax = tax[name];
    const n = tags.filter(t => (ax.tags || []).includes(t)).length;
    if (ax.select === 'one' && n > 1)
      throw new Error('Eix "' + name + '": màxim 1 (té ' + n + ').');
    if (ax.required && n < 1)
      throw new Error('Eix "' + name + '": cal com a mínim 1.');
  });
}

const GEMINI_PREAMBLE = [
  'Converteix el text següent en una recepta en format JSON.',
  'Respon NOMÉS amb un objecte JSON amb aquesta estructura exacta:',
  '{"name":"","ingredients":[],"instructions":"","tags":[]}',
  'Escriu-ho en català. Ignora introduccions, històries o comentaris del blog: extreu només la recepta.',
  'instructions = els passos d\'elaboració, amb salts de línia \\n entre passos.',
  'tags: tria NOMÉS d\'aquestes llistes tancades, no n\'inventis cap:',
];

// ===============================================================

function processInbox() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('Ja hi ha una execució en curs. Surto.');
    return;
  }
  try {
    const inbox = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
    moveRecipesToInbox(inbox);
    const processed = getOrCreateSubfolder_(inbox, CONFIG.PROCESSED_SUBFOLDER);
    const review = getOrCreateSubfolder_(inbox, CONFIG.REVIEW_SUBFOLDER);

    const { obj: index, sha } = githubGetIndex_();
    if (!Array.isArray(index.recipes)) index.recipes = [];
    const ids = new Set(index.recipes.map(r => r.id || slugify_(r.name)));

    const docs = inbox.getFilesByType(MimeType.GOOGLE_DOCS);
    const toMerge = [];
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    let dupes = 0, errors = 0;

    while (docs.hasNext()) {
      const file = docs.next();
      try {
        const text = DocumentApp.openById(file.getId()).getBody().getText();
        const rec = parseRecipe_(text);
        validateRecipe_(rec, index.taxonomy);

        rec.id = slugify_(rec.name);
        rec.source = rec.source || 'Gemini';
        rec.added = today;

        if (ids.has(rec.id)) {
          Logger.log('DUPLICADA (id=%s): %s -> moc a processats', rec.id, file.getName());
          file.moveTo(processed);
          dupes++;
          continue;
        }
        ids.add(rec.id);
        toMerge.push({ rec, file });
      } catch (e) {
        Logger.log('ERROR a "%s": %s -> moc a revisar', file.getName(), e.message);
        file.moveTo(review);
        errors++;
      }
    }

    if (toMerge.length === 0) {
      Logger.log('Res nou per fusionar. (duplicades: %s, errors: %s)', dupes, errors);
      return;
    }

    toMerge.forEach(x => index.recipes.push(x.rec));
    index.recipes.sort((a, b) => a.name.localeCompare(b.name, 'ca'));
    index.total = index.recipes.length;
    index.generated = false;

    const jsonStr = JSON.stringify(index, null, 2);
    const msg = `Afegir ${toMerge.length} recepta/es via inbox`;
    githubPutIndex_(jsonStr, sha, msg);
    mirrorToDrive_(jsonStr);

    toMerge.forEach(x => x.file.moveTo(processed));
    Logger.log('OK: %s noves publicades. Total: %s.', toMerge.length, index.total);
  } finally {
    lock.releaseLock();
  }
}

function moveRecipesToInbox(inbox) {
  inbox = inbox || DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
  const marker = CONFIG.RECIPE_MARKER.toUpperCase();
  const it = DriveApp.getRootFolder().getFilesByType(MimeType.GOOGLE_DOCS);
  let moved = 0;
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().toUpperCase().indexOf(marker) === 0) {
      f.moveTo(inbox);
      moved++;
    }
  }
  if (moved) Logger.log('Mogudes %s receptes de l\'arrel a la inbox.', moved);
  return moved;
}

function parseRecipe_(rawText) {
  let t = (rawText || '').trim();
  t = t.replace(/```[a-zA-Z]*\s*/g, '').replace(/```/g, '');
  const first = t.indexOf('{');
  const last = t.lastIndexOf('}');
  if (first === -1 || last === -1 || last < first) {
    throw new Error('No s\'ha trobat cap objecte JSON al document.');
  }
  t = t.substring(first, last + 1)
       .replace(/[“”]/g, '"')
       .replace(/[‘’]/g, "'")
       .replace(/ /g, ' ');

  return JSON.parse(t);
}

function validateRecipe_(rec, taxonomy) {
  if (!rec || typeof rec.name !== 'string' || !rec.name.trim())
    throw new Error('Falta "name".');
  if (!Array.isArray(rec.ingredients))
    throw new Error('"ingredients" ha de ser una llista.');
  if (typeof rec.instructions !== 'string' || !rec.instructions.trim())
    throw new Error('Falta "instructions".');
  if (!Array.isArray(rec.tags) || rec.tags.length === 0)
    throw new Error('Falten "tags".');
  validateTags_(rec.tags, taxonomy);
}

function slugify_(name) {
  return name
    .normalize('NFKD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function githubToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!t) throw new Error('Falta la Script Property GITHUB_TOKEN.');
  return t;
}

function githubHeaders_() {
  return {
    'Authorization': 'Bearer ' + githubToken_(),
    'Accept': 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28',
    'User-Agent': 'receptes-appsscript',
  };
}

function githubContentsUrl_() {
  return 'https://api.github.com/repos/' + CONFIG.GITHUB_OWNER + '/' +
         CONFIG.GITHUB_REPO + '/contents/' + CONFIG.GITHUB_PATH;
}

function githubGetIndex_() {
  const url = githubContentsUrl_() + '?ref=' + CONFIG.GITHUB_BRANCH;
  const res = UrlFetchApp.fetch(url, { method: 'get', headers: githubHeaders_(), muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code !== 200) throw new Error('GitHub GET ' + code + ': ' + res.getContentText());

  const data = JSON.parse(res.getContentText());
  const bytes = Utilities.base64Decode(String(data.content).replace(/\s/g, ''));
  const jsonStr = Utilities.newBlob(bytes).getDataAsString('UTF-8');
  return { obj: JSON.parse(jsonStr), sha: data.sha };
}

function githubPutIndex_(jsonStr, sha, message) {
  const payload = {
    message: message,
    content: Utilities.base64Encode(jsonStr, Utilities.Charset.UTF_8),
    sha: sha,
    branch: CONFIG.GITHUB_BRANCH,
  };
  const res = UrlFetchApp.fetch(githubContentsUrl_(), {
    method: 'put',
    headers: githubHeaders_(),
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code !== 200 && code !== 201)
    throw new Error('GitHub PUT ' + code + ': ' + res.getContentText());
}

function getOrCreateSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function mirrorToDrive_(jsonStr) {
  const folder = DriveApp.getFolderById(CONFIG.MIRROR_FOLDER_ID);
  const it = folder.getFilesByName(CONFIG.MIRROR_FILENAME);
  if (it.hasNext()) {
    it.next().setContent(jsonStr);
  } else {
    folder.createFile(CONFIG.MIRROR_FILENAME, jsonStr, 'application/json');
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'Receptes admin actiu' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function resolveRole_(password) {
  const props = PropertiesService.getScriptProperties();
  const admin  = props.getProperty('ADMIN_PASSWORD');
  const family = props.getProperty('FAMILY_PASSWORD');
  if (admin  && password === admin)  return 'admin';
  if (family && password === family) return 'family';
  return null;
}

const ROLE_ACTIONS = {
  admin:  ['auth', 'delete', 'updateTags', 'saveRecipe', 'parseText', 'parseImages', 'addTag'],
  family: ['auth', 'saveRecipe', 'parseText', 'parseImages'],
};

function doPost(e) {
  const out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  try {
    const body = JSON.parse(e.postData.contents);
    const role = resolveRole_(body.password);
    if (!role) return out.setContent(JSON.stringify({ ok: false, error: 'Contrasenya incorrecta.' }));

    const allowed = ROLE_ACTIONS[role] || [];
    if (allowed.indexOf(body.action) === -1) throw new Error('Acció no permesa per al teu rol.');

    let result;
    if (body.action === 'auth') result = { role: role };
    else if (body.action === 'delete') result = adminDelete_(body.id);
    else if (body.action === 'updateTags') result = adminUpdateTags_(body.id, body.tags);
    else if (body.action === 'saveRecipe') result = adminSaveRecipe_(body.recipe, role);
    else if (body.action === 'parseText') result = adminParseText_(body.text);
    else if (body.action === 'parseImages') result = adminParseImages_(body.images);
    else if (body.action === 'addTag') result = adminAddTag_(body.axis, body.tag);
    else throw new Error('Acció desconeguda: ' + body.action);

    return out.setContent(JSON.stringify(Object.assign({ ok: true }, result)));
  } catch (err) {
    return out.setContent(JSON.stringify({ ok: false, error: String(err && err.message || err) }));
  }
}

function adminDelete_(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { obj, sha } = githubGetIndex_();
    const before = obj.recipes.length;
    obj.recipes = obj.recipes.filter(r => r.id !== id);
    if (obj.recipes.length === before) throw new Error('No existeix cap recepta amb id=' + id);
    obj.recipes.sort((a, b) => a.name.localeCompare(b.name, 'ca'));
    obj.total = obj.recipes.length;
    const jsonStr = JSON.stringify(obj, null, 2);
    githubPutIndex_(jsonStr, sha, 'Esborrar recepta ' + id + ' (admin web)');
    mirrorToDrive_(jsonStr);
    return { total: obj.total };
  } finally {
    lock.releaseLock();
  }
}

function adminUpdateTags_(id, tags) {
  if (!Array.isArray(tags) || tags.length === 0) throw new Error('Cal com a mínim una tag.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { obj, sha } = githubGetIndex_();
    validateTags_(tags, obj.taxonomy);
    const rec = obj.recipes.find(r => r.id === id);
    if (!rec) throw new Error('No existeix cap recepta amb id=' + id);
    rec.tags = tags;
    const jsonStr = JSON.stringify(obj, null, 2);
    githubPutIndex_(jsonStr, sha, 'Actualitzar tags de ' + id + ' (admin web)');
    mirrorToDrive_(jsonStr);
    return { total: obj.total };
  } finally {
    lock.releaseLock();
  }
}

function adminSaveRecipe_(input, role) {
  const rec = {
    name: String(input && input.name || '').trim(),
    ingredients: Array.isArray(input && input.ingredients)
      ? input.ingredients.map(s => String(s).trim()).filter(s => s) : [],
    instructions: String(input && input.instructions || '').trim(),
    tags: Array.isArray(input && input.tags) ? input.tags : [],
  };

  if (role === 'family' && input && input.id) {
    throw new Error('La família només pot crear receptes noves, no editar-ne.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { obj, sha } = githubGetIndex_();
    validateRecipe_(rec, obj.taxonomy);
    const existing = input.id ? obj.recipes.find(r => r.id === input.id) : null;

    if (existing) {
      existing.name = rec.name;
      existing.ingredients = rec.ingredients;
      existing.instructions = rec.instructions;
      existing.tags = rec.tags;
    } else {
      const id = slugify_(rec.name);
      if (obj.recipes.some(r => r.id === id))
        throw new Error('Ja existeix una recepta amb aquest nom (id=' + id + ').');
      obj.recipes.push({
        id: id,
        name: rec.name,
        source: 'Web',
        ingredients: rec.ingredients,
        instructions: rec.instructions,
        tags: rec.tags,
        added: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      });
    }

    obj.recipes.sort((a, b) => a.name.localeCompare(b.name, 'ca'));
    obj.total = obj.recipes.length;
    const jsonStr = JSON.stringify(obj, null, 2);
    const id = existing ? existing.id : slugify_(rec.name);
    githubPutIndex_(jsonStr, sha, (existing ? 'Editar' : 'Crear') + ' recepta ' + id + ' (admin web)');
    mirrorToDrive_(jsonStr);
    return { total: obj.total, id: id };
  } finally {
    lock.releaseLock();
  }
}

const AXIS_HINTS = {
  'Cuina': 'Si cap no encaixa clarament, tria Internacional.',
  'Estil asiàtic': 'Només si la cuina és Asiàtic.',
  'Base': "Només si és l'ingredient principal del plat.",
};

function buildGeminiPrompt_(taxonomy) {
  const tax = (taxonomy && Object.keys(taxonomy).length) ? taxonomy : FALLBACK_TAXONOMY;
  const lines = GEMINI_PREAMBLE.slice();
  Object.keys(tax).forEach(function (name) {
    const ax = tax[name];
    const card = (ax.select === 'one' && ax.required) ? 'EXACTAMENT 1'
               : (ax.select === 'one') ? 'com a màxim 1'
               : 'opcional, 0 o més';
    const hint = AXIS_HINTS[name] ? ' ' + AXIS_HINTS[name] : '';
    lines.push('- ' + name + ' (' + card + '): ' + (ax.tags || []).join(', ') + '.' + hint);
  });
  return lines.join('\n');
}

function currentTaxonomy_() {
  try {
    const { obj } = githubGetIndex_();
    if (obj && obj.taxonomy && Object.keys(obj.taxonomy).length) return obj.taxonomy;
  } catch (e) {
    Logger.log('No he pogut llegir la taxonomia viva; faig servir la de reserva: ' + e.message);
  }
  return FALLBACK_TAXONOMY;
}

/**
 * Crida Gemini 2.5 Flash
 */
function geminiExtractRecipe_(parts, taxonomy) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Falta la Script Property GEMINI_API_KEY.');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
              CONFIG.GEMINI_MODEL + ':generateContent';

  const payload = JSON.stringify({
    contents: [{ parts: parts }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.2 },
  });

  const options = {
    method: 'post',
    headers: { 'x-goog-api-key': apiKey },
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true,
  };

  const maxRetries = 3;
  let res, code;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    res = UrlFetchApp.fetch(url, options);
    code = res.getResponseCode();

    if (code === 200) break;

    if ((code === 429 || code >= 500) && attempt < maxRetries) {
      Logger.log(`Gemini ocupat (${code}). Reintent ${attempt + 1}...`);
      Utilities.sleep(attempt * 2000);
    } else {
      throw new Error(`Gemini ${code}: ` + res.getContentText().slice(0, 200));
    }
  }

  const data = JSON.parse(res.getContentText());
  let outText;
  try {
    outText = data.candidates[0].content.parts[0].text;
  } catch (e) {
    throw new Error('Resposta buida o inesperada de Gemini.');
  }

  let rec;
  try {
    rec = JSON.parse(outText);
  } catch (e) {
    outText = outText.replace(/```[a-zA-Z]*\s*/g, '').replace(/```/g, '').trim();
    rec = JSON.parse(outText);
  }

  const vocab = taxonomyVocab_(taxonomy);
  return {
    recipe: {
      name: String(rec.name || '').trim(),
      ingredients: Array.isArray(rec.ingredients) ? rec.ingredients.map(s => String(s).trim()).filter(s => s) : [],
      instructions: String(rec.instructions || '').trim(),
      tags: Array.isArray(rec.tags) ? rec.tags.filter(t => vocab.has(t)) : [],
    },
  };
}

function adminParseText_(text) {
  text = String(text || '').trim();
  if (!text) throw new Error('Text buit.');
  if (text.length > 20000) text = text.substring(0, 20000);

  const taxonomy = currentTaxonomy_();
  const promptText = buildGeminiPrompt_(taxonomy);
  const parts = [{ text: promptText + '\n\n--- TEXT ---\n' + text }];
  return geminiExtractRecipe_(parts, taxonomy);
}

function adminParseImages_(images) {
  if (!Array.isArray(images) || images.length === 0) throw new Error('Cap imatge rebuda.');
  if (images.length > 4) throw new Error('Màxim 4 imatges.');

  const taxonomy = currentTaxonomy_();
  const promptText = buildGeminiPrompt_(taxonomy) +
    '\n\nLa recepta és a la/les imatge(s) adjunta(es); pot ser manuscrita. Fes-ne la lectura (OCR) i estructura-la.';

  const parts = [{ text: promptText }];
  images.forEach(function (b64) {
    const d = String(b64 || '').replace(/^data:[^,]*,/, '').replace(/\s/g, '');
    if (!d) throw new Error('Imatge buida.');
    parts.push({ inline_data: { mime_type: 'image/jpeg', data: d } });
  });
  return geminiExtractRecipe_(parts, taxonomy);
}

function adminAddTag_(axis, tag) {
  axis = String(axis || '').trim();
  tag  = String(tag  || '').replace(/\s+/g, ' ').trim();
  if (!axis) throw new Error('Falta l\'eix.');
  if (!tag)  throw new Error('Falta el nom de l\'etiqueta.');
  if (tag.length > 40) throw new Error('Nom d\'etiqueta massa llarg.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { obj, sha } = githubGetIndex_();
    if (!obj.taxonomy || !Object.keys(obj.taxonomy).length) {
      obj.taxonomy = JSON.parse(JSON.stringify(FALLBACK_TAXONOMY));
    }
    if (!obj.taxonomy[axis]) throw new Error('Eix desconegut: ' + axis);

    const lower = tag.toLowerCase();
    const dup = Object.values(obj.taxonomy)
      .some(ax => (ax.tags || []).some(t => t.toLowerCase() === lower));
    if (dup) throw new Error('L\'etiqueta "' + tag + '" ja existeix.');

    obj.taxonomy[axis].tags.push(tag);
    const jsonStr = JSON.stringify(obj, null, 2);
    githubPutIndex_(jsonStr, sha, 'Afegir etiqueta "' + tag + '" a ' + axis + ' (admin web)');
    mirrorToDrive_(jsonStr);
    return { axis: axis, tag: tag, taxonomy: obj.taxonomy };
  } finally {
    lock.releaseLock();
  }
}

// ------------------- Funció de diagnòstic directa -------------------

function testGeminiApi() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const genUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + apiKey;
  const payload = {
    contents: [{ parts: [{ text: 'Digues OK' }] }]
  };
  const resGen = UrlFetchApp.fetch(genUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('Codi test generateContent: ' + resGen.getResponseCode());
  Logger.log('Resposta: ' + resGen.getContentText().slice(0, 300));
}
