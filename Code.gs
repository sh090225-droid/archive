/**
 * CHARACTER ARCHIVE 1.0
 * Google Apps Script backend
 *
 * First-time setup:
 * 1) Run setup() once from the Apps Script editor and authorize it.
 * 2) Deploy as Web app. Execute as: Me. Who has access: Only myself.
 *
 * No in-app password. Access is restricted by the Apps Script deployment
 * setting, so do not deploy this version with access set to Anyone.
 */

const APP = {
  PROP_SHEET_ID: 'CHARACTER_ARCHIVE_SPREADSHEET_ID',
  PROP_API_SECRET: 'CHARACTER_ARCHIVE_API_SECRET',
  SHEETS: {
    USERS: ['email', 'name', 'createdAt', 'lastLoginAt'],
    CHARACTERS: [
      'id', 'ownerEmail', 'name', 'visibility', 'folderId', 'tagsJson',
      'favorite', 'imageFileId', 'imageMimeType', 'createdAt', 'updatedAt'
    ],
    FIELDS: ['id', 'characterId', 'ownerEmail', 'label', 'type', 'value', 'sortOrder'],
    FOLDERS: ['id', 'ownerEmail', 'name', 'parentId', 'createdAt']
  },
  MAX_IMAGE_BYTES: 5 * 1024 * 1024
};


/** JSON API for the Cloudflare Pages Function. Never expose the secret in browser code. */
function doPost(e) {
  try {
    const body = JSON.parse(e && e.postData && e.postData.contents || '{}');
    const expected = PropertiesService.getScriptProperties().getProperty(APP.PROP_API_SECRET);
    if (!expected || !body.secret || body.secret !== expected) {
      return jsonOutput_({ error: 'API 인증에 실패했어.' });
    }
    const actions = {
      getInitialData: () => getInitialData(),
      getCharacter: () => getCharacter(body.args[0]),
      getPublicCharacter: () => getPublicCharacter(body.args[0]),
      saveCharacter: () => saveCharacter(body.args[0]),
      uploadCharacterImage: () => uploadCharacterImage(body.args[0], body.args[1], body.args[2], body.args[3]),
      deleteCharacter: () => deleteCharacter(body.args[0]),
      saveFolder: () => saveFolder(body.args[0], body.args[1]),
      deleteFolder: () => deleteFolder(body.args[0])
    };
    if (!Object.prototype.hasOwnProperty.call(actions, body.action)) {
      return jsonOutput_({ error: '허용되지 않은 API야.' });
    }
    return jsonOutput_({ result: actions[body.action]() });
  } catch (err) {
    return jsonOutput_({ error: err && err.message ? err.message : String(err) });
  }
}

function jsonOutput_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.publicCharacterId = (e && e.parameter && e.parameter.character) ? String(e.parameter.character) : '';
  template.webAppUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle('Character Archive')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Run once manually from the editor. Creates the database spreadsheet. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let apiSecret = props.getProperty(APP.PROP_API_SECRET);
  if (!apiSecret) {
    apiSecret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty(APP.PROP_API_SECRET, apiSecret);
  }
  Logger.log('Cloudflare GAS_API_SECRET (copy to Cloudflare Pages secret; keep private): ' + apiSecret);
  let id = props.getProperty(APP.PROP_SHEET_ID);
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Character Archive Database');
    props.setProperty(APP.PROP_SHEET_ID, ss.getId());
  }

  Object.keys(APP.SHEETS).forEach(key => {
    const name = key === 'FIELDS' ? 'CustomFields' : key.charAt(0) + key.slice(1).toLowerCase();
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = APP.SHEETS[key];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    } else {
      const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
      if (headers.some((h, i) => existing[i] !== h)) {
        throw new Error('시트 ' + name + '의 첫 행 구조가 예상과 달라. 기존 데이터를 확인해줘.');
      }
    }
  });

  // Migrate records created by the earlier Google-login version to the single admin owner.
  ['Characters', 'CustomFields', 'Folders'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      const ownerCol = headers.indexOf('ownerEmail') + 1;
      if (ownerCol > 0) sh.getRange(2, ownerCol, sh.getLastRow() - 1, 1).setValue('ADMIN');
    }
  });

  // Remove the default empty sheet if it is not one of our tables.
  ss.getSheets().forEach(sh => {
    if (!Object.values(APP.SHEETS).some(headers => {
      const expected = headers;
      return expected[0] === sh.getRange(1, 1).getValue();
    }) && sh.getName() === 'Sheet1' && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  });

  return {
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    message: '데이터베이스가 준비됐어. 별도의 관리자 암호는 필요하지 않아.',
  };
}

function getInitialData() {
  const owner = 'ADMIN';
  const characters = readObjects_('Characters')
    .filter(r => r.ownerEmail === owner)
    .map(r => ({
      id: r.id, name: r.name, visibility: r.visibility, folderId: r.folderId,
      tags: safeJson_(r.tagsJson, []), favorite: String(r.favorite) === 'true',
      updatedAt: r.updatedAt
    }));
  const folders = readObjects_('Folders')
    .filter(r => r.ownerEmail === owner)
    .map(r => ({ id: r.id, name: r.name, parentId: r.parentId || '' }));
  return { user: { name: '관리자' }, characters, folders };
}

function getCharacter(characterId) {
  const c = findCharacter_(String(characterId));
  if (!c || c.ownerEmail !== 'ADMIN') throw new Error('캐릭터를 찾을 수 없거나 접근 권한이 없어.');
  const fields = readObjects_('CustomFields')
    .filter(f => f.characterId === c.id && f.ownerEmail === 'ADMIN')
    .sort((a, b) => Number(a.sortOrder) - Number(b.sortOrder))
    .map(f => ({ id: f.id, label: f.label, type: f.type, value: f.value }));
  let imageData = '';
  if (c.imageFileId) imageData = getPrivateImageData_(c.imageFileId, 'ADMIN');
  return {
    id: c.id, name: c.name, visibility: c.visibility, folderId: c.folderId || '',
    tags: safeJson_(c.tagsJson, []), favorite: String(c.favorite) === 'true',
    imageFileId: c.imageFileId || '', imageMimeType: c.imageMimeType || '',
    imageData, fields
  };
}

/** Public link endpoint. Returns only public character data. */
function getPublicCharacter(characterId) {
  const c = findCharacter_(String(characterId));
  if (!c || c.visibility !== 'PUBLIC') throw new Error('공개된 캐릭터가 아니거나 존재하지 않아.');
  const fields = readObjects_('CustomFields')
    .filter(f => f.characterId === c.id && f.ownerEmail === c.ownerEmail)
    .sort((a, b) => Number(a.sortOrder) - Number(b.sortOrder))
    .map(f => ({ label: f.label, type: f.type, value: f.value }));
  let imageUrl = '';
  if (c.imageFileId) imageUrl = 'https://drive.google.com/uc?export=view&id=' + encodeURIComponent(c.imageFileId);
  const folder = c.folderId ? findFolder_(c.folderId) : null;
  return {
    id: c.id, name: c.name, tags: safeJson_(c.tagsJson, []),
    folderName: folder ? folder.name : '', imageUrl, fields
  };
}

function saveCharacter(data) {
  data = data || {};
  const name = String(data.name || '').trim();
  if (!name) throw new Error('캐릭터 이름을 입력해줘.');
  if (name.length > 120) throw new Error('캐릭터 이름은 120자 이내로 입력해줘.');

  const id = data.id ? String(data.id) : Utilities.getUuid();
  const existing = findCharacter_(id);
  if (existing && existing.ownerEmail !== 'ADMIN') throw new Error('수정 권한이 없어.');

  const folderId = String(data.folderId || '');
  if (folderId) {
    const folder = findFolder_(folderId);
    if (!folder || folder.ownerEmail !== 'ADMIN') throw new Error('선택한 폴더를 찾을 수 없어.');
  }

  const tags = Array.isArray(data.tags)
    ? data.tags.map(t => String(t).trim()).filter(Boolean).slice(0, 40)
    : [];
  const visibility = data.visibility === 'PUBLIC' ? 'PUBLIC' : 'PRIVATE';
  const now = new Date().toISOString();

  const record = {
    id, ownerEmail: 'ADMIN', name, visibility, folderId,
    tagsJson: JSON.stringify(tags), favorite: !!data.favorite,
    imageFileId: existing ? existing.imageFileId : '',
    imageMimeType: existing ? existing.imageMimeType : '',
    createdAt: existing ? existing.createdAt : now,
    updatedAt: now
  };

  upsertObject_('Characters', 'id', record);

  const fields = Array.isArray(data.fields) ? data.fields : [];
  replaceCharacterFields_(id, 'ADMIN', fields);
  if (record.imageFileId) setImageSharing_(record.imageFileId, visibility);

  return { id, message: '캐릭터를 저장했어.' };
}

function uploadCharacterImage(characterId, fileName, mimeType, base64Data) {
  const c = findCharacter_(String(characterId));
  if (!c || c.ownerEmail !== 'ADMIN') throw new Error('이미지를 추가할 권한이 없어.');

  const allowed = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'];
  if (allowed.indexOf(String(mimeType)) === -1) throw new Error('PNG, JPG, GIF, WEBP 이미지만 지원해.');
  const bytes = Utilities.base64Decode(String(base64Data || ''));
  if (bytes.length > APP.MAX_IMAGE_BYTES) throw new Error('이미지는 5MB 이하로 올려줘.');

  const safeName = String(fileName || 'character-image').replace(/[^\w.\-가-힣 ]/g, '_').slice(0, 100);
  const blob = Utilities.newBlob(bytes, mimeType, safeName);
  const file = DriveApp.createFile(blob);
  setImageSharing_(file.getId(), c.visibility);

  if (c.imageFileId) {
    try { DriveApp.getFileById(c.imageFileId).setTrashed(true); } catch (err) {}
  }
  c.imageFileId = file.getId();
  c.imageMimeType = mimeType;
  c.updatedAt = new Date().toISOString();
  upsertObject_('Characters', 'id', c);
  return { imageFileId: file.getId() };
}

function deleteCharacter(characterId) {
  const c = findCharacter_(String(characterId));
  if (!c || c.ownerEmail !== 'ADMIN') throw new Error('삭제 권한이 없어.');
  if (c.imageFileId) {
    try { DriveApp.getFileById(c.imageFileId).setTrashed(true); } catch (err) {}
  }
  deleteRows_('CustomFields', row => row.characterId === c.id && row.ownerEmail === 'ADMIN');
  deleteRows_('Characters', row => row.id === c.id && row.ownerEmail === 'ADMIN');
  return { message: '캐릭터를 삭제했어.' };
}

function saveFolder(name, parentId) {
  name = String(name || '').trim();
  parentId = String(parentId || '');
  if (!name) throw new Error('폴더 이름을 입력해줘.');
  if (name.length > 80) throw new Error('폴더 이름은 80자 이내로 입력해줘.');
  if (parentId) {
    const parent = findFolder_(parentId);
    if (!parent || parent.ownerEmail !== 'ADMIN') throw new Error('상위 폴더를 찾을 수 없어.');
  }
  const id = Utilities.getUuid();
  appendObject_('Folders', {
    id, ownerEmail: 'ADMIN', name, parentId, createdAt: new Date().toISOString()
  });
  return { id };
}

function deleteFolder(folderId) {
  const folder = findFolder_(String(folderId));
  if (!folder || folder.ownerEmail !== 'ADMIN') throw new Error('폴더를 찾을 수 없어.');
  const children = readObjects_('Folders').filter(f => f.ownerEmail === 'ADMIN' && f.parentId === folder.id);
  if (children.length) throw new Error('하위 폴더를 먼저 이동하거나 삭제해줘.');
  const chars = readObjects_('Characters').filter(c => c.ownerEmail === 'ADMIN' && c.folderId === folder.id);
  if (chars.length) throw new Error('폴더 안의 캐릭터를 먼저 다른 폴더로 옮겨줘.');
  deleteRows_('Folders', f => f.id === folder.id && f.ownerEmail === 'ADMIN');
  return { message: '폴더를 삭제했어.' };
}

/* -------------------- Internal helpers -------------------- */

function db_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP.PROP_SHEET_ID);
  if (!id) throw new Error('먼저 Apps Script 편집기에서 setup()을 실행해줘.');
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  const sh = db_().getSheetByName(name);
  if (!sh) throw new Error('데이터 시트가 없어. setup()을 먼저 실행해줘.');
  return sh;
}

function readObjects_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(v => v !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i]);
    return obj;
  });
}

function upsertObject_(name, key, obj, preserveKeys) {
  const sh = sheet_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const values = sh.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][headers.indexOf(key)]) === String(obj[key])) { rowIndex = i + 1; break; }
  }
  if (rowIndex > 0 && preserveKeys && preserveKeys.length) {
    const existing = {};
    headers.forEach((h, i) => existing[h] = values[rowIndex - 1][i]);
    preserveKeys.forEach(k => { if (existing[k]) obj[k] = existing[k]; });
  }
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  if (rowIndex > 0) sh.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  else sh.appendRow(row);
}

function appendObject_(name, obj) {
  const sh = sheet_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  sh.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
}

function deleteRows_(name, predicate) {
  const sh = sheet_(name);
  const rows = readObjects_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const keyRows = [];
  rows.forEach((obj, i) => { if (predicate(obj)) keyRows.push(i + 2); });
  keyRows.reverse().forEach(r => sh.deleteRow(r));
}

function findCharacter_(id) {
  return readObjects_('Characters').find(r => String(r.id) === String(id)) || null;
}

function findFolder_(id) {
  return readObjects_('Folders').find(r => String(r.id) === String(id)) || null;
}

function replaceCharacterFields_(characterId, email, fields) {
  deleteRows_('CustomFields', f => f.characterId === characterId && f.ownerEmail === email);
  fields.slice(0, 100).forEach((f, i) => {
    const label = String(f.label || '').trim().slice(0, 100);
    const type = ['text', 'textarea', 'number', 'date', 'url', 'image'].includes(f.type) ? f.type : 'text';
    if (!label) return;
    appendObject_('CustomFields', {
      id: Utilities.getUuid(), characterId, ownerEmail: email,
      label, type, value: String(f.value == null ? '' : f.value).slice(0, 10000),
      sortOrder: i
    });
  });
}

function getPrivateImageData_(fileId, email) {
  // The character ownership check happens before this function is called.
  const blob = DriveApp.getFileById(fileId).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function setImageSharing_(fileId, visibility) {
  const file = DriveApp.getFileById(fileId);
  if (visibility === 'PUBLIC') {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } else {
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  }
}

function safeJson_(value, fallback) {
  try { return JSON.parse(String(value || '')); } catch (err) { return fallback; }
}
