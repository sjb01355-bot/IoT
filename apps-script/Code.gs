const DEVICES = ['도어락', '도어벨', '1구 스위치', '2구 스위치', '스피커', '리모트 스위치',
  '미니 스위치', '화재 감지기', '허브', '커튼', 'LED', '추가 조명1', '추가 조명2',
  '추가 리모트', '추가 미니', '추가 화재', '추가 도어벨'];
const CHANNELS = ['전화', '내방', '기타'];
const TABS = {participants:'참여자 기록', logs:'모니터링 기록'};
const HEADERS = ['ID', '엑셀 번호', '기록 데이터', '사진 분할 수'];
const PHOTO_CHUNK = 40000;

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('오아시스 · IoT 보급사업')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function setupOasis() {
  return locked_(() => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) throw new Error('연결된 스프레드시트의 Apps Script 편집기에서 실행해 주세요.');
    const properties = PropertiesService.getScriptProperties();
    const prior = properties.getProperty('OASIS_SPREADSHEET_ID');
    if (prior && prior !== spreadsheet.getId()) throw new Error('이미 다른 스프레드시트가 연결되어 있습니다.');
    Object.values(TABS).forEach(name => {
      const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
      const range = sheet.getRange(1, 1, 1, HEADERS.length);
      if (!sheet.getLastRow()) {
        range.setValues([HEADERS]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(3, 360);
      } else if (range.getValues()[0].join('|') !== HEADERS.join('|')) {
        throw new Error('기존 탭 ‘' + name + '’의 헤더가 다릅니다. 빈 스프레드시트에서 설정해 주세요.');
      }
    });
    properties.setProperty('OASIS_SPREADSHEET_ID', spreadsheet.getId());
  });
}

function locked_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const result = callback();
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function database_() {
  const id = PropertiesService.getScriptProperties().getProperty('OASIS_SPREADSHEET_ID');
  if (!id) throw new Error('앱을 사용하기 전에 Apps Script 편집기에서 setupOasis를 실행해 주세요.');
  return SpreadsheetApp.openById(id);
}

function records_(sheet) {
  const size = sheet.getLastRow() - 1;
  if (size < 1) return [];
  return sheet.getRange(2, 1, size, HEADERS.length).getValues().flatMap(([id, source, data, photoCount], index) =>
    id ? [{id:Number(id), source:String(source || ''), row:index+2, body:JSON.parse(data), photoCount:Number(photoCount)||0}] : []);
}

function state_(spreadsheet) {
  return {
    participants:records_(spreadsheet.getSheetByName(TABS.participants)).map(r=>({...r.body,id:r.id})),
    logs:records_(spreadsheet.getSheetByName(TABS.logs)).map(r=>({...r.body,id:r.id,hasPhoto:r.photoCount>0})),
    devices:DEVICES,
  };
}

function nextId_(table, rows) {
  const properties = PropertiesService.getScriptProperties();
  const key = 'OASIS_SEQUENCE_' + table;
  const current = rows.reduce((max, record)=>Math.max(max, record.id), Number(properties.getProperty(key)) || 0);
  properties.setProperty(key, String(current + 1));
  return current + 1;
}

function write_(sheet, record, photo) {
  const fragments = photo ? photo.match(new RegExp('.{1,' + PHOTO_CHUNK + '}', 'g')) : [];
  const data = JSON.stringify(record.body);
  if (data.length > 45000) throw new Error('상담 내용과 슈퍼비전의 길이를 줄여 주세요.');
  const total = HEADERS.length + Math.max(fragments.length, record.photoCount || 0);
  if (total > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), total - sheet.getMaxColumns());
  if (record.row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), record.row - sheet.getMaxRows());
  const values = [record.id, record.source, data, fragments.length, ...fragments];
  // Prefix every photo fragment so Sheets never interprets its contents as a formula.
  for (let index = HEADERS.length; index < values.length; index++) values[index] = 'b64:' + values[index];
  while (values.length < total) values.push('');
  sheet.getRange(record.row, 1, 1, total).setValues([values]);
}

function boundedText_(value, maximum) {
  const result = String(value ?? '').trim();
  if (result.length > maximum) throw new Error('입력 내용이 너무 깁니다.');
  return result;
}

function participant_(body) {
  const name = boundedText_(body.name, 100);
  if (!name) throw new Error('참여자 이름을 입력해 주세요.');
  if (!['설치 완료','설치 대기'].includes(body.status)) throw new Error('설치 상태를 선택해 주세요.');
  const quantities = Object.entries(body.quantities || {});
  if (quantities.some(([key, value])=>!DEVICES.includes(key) || !Number.isInteger(value) || value<0 || value>9999)) {
    throw new Error('기기 수량은 0~9999 사이의 정수여야 합니다.');
  }
  return {name, status:body.status, quantities:Object.fromEntries(quantities),
    region:boundedText_(body.region,100), phone:boundedText_(body.phone,100),
    address:boundedText_(body.address,300), memo:boundedText_(body.memo,10000)};
}

function monitoring_(body) {
  const date = String(body.date);
  const timestamp = new Date(date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || isNaN(timestamp.getTime()) || timestamp.toISOString().slice(0,10)!==date) {
    throw new Error('상담 날짜를 확인해 주세요.');
  }
  const time = String(body.time || '');
  if (time && !/^([01]\d|2[0-3]):[0-5]\d$/.test(time)) throw new Error('상담 시간을 확인해 주세요.');
  if (!Array.isArray(body.channels) || !body.channels.length || body.channels.some(value=>!CHANNELS.includes(value))) throw new Error('상담 유형을 선택해 주세요.');
  if (!Array.isArray(body.checks) || body.checks.some(value=>!DEVICES.includes(value))) throw new Error('점검 기기 항목을 확인해 주세요.');
  const content = boundedText_(body.content,10000);
  if (!content) throw new Error('상담 내용을 입력해 주세요.');
  return {participant_id:Number(body.participant_id), date, time, channels:[...new Set(body.channels)],
    checks:[...new Set(body.checks)], content, supervision:boundedText_(body.supervision,10000)};
}

function photo_(input) {
  if (!input) return '';
  if (typeof input !== 'string' || input.length > 4000032 ||
      !/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+/]+={0,2}$/.test(input)) {
    throw new Error('3MB 이하의 JPG, PNG, WebP 사진만 첨부할 수 있습니다.');
  }
  const bytes = Utilities.base64Decode(input.split(',')[1]);
  if (bytes.length > 3000000) throw new Error('사진은 3MB 이하로 첨부해 주세요.');
  return input;
}

function readPhoto_(sheet, record) {
  if (!record.photoCount) return '';
  return sheet.getRange(record.row, HEADERS.length+1, 1, record.photoCount)
    .getValues()[0].map(fragment=>String(fragment).slice(4)).join('');
}

function roster_(input) {
  if (typeof input !== 'string' || !input.length || input.length > 8000000) {
    throw new Error('6MB 이하의 엑셀 파일을 선택해 주세요.');
  }
  const bytes = Utilities.base64Decode(input);
  const blob = Utilities.newBlob(bytes,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'oasis-roster.xlsx');
  const converted = Drive.Files.create({name:'오아시스 명단 가져오기 임시 파일',
    mimeType:'application/vnd.google-apps.spreadsheet'}, blob, {fields:'id'});
  try {
    const spreadsheet = SpreadsheetApp.openById(converted.id);
    const sheet = spreadsheet.getSheetByName('설치명단 및 수량') || spreadsheet.getSheets()[0];
    const range = sheet.getRange(1,1,Math.max(2,Math.min(sheet.getLastRow(),5000)),24);
    return parseRoster_(range.getValues(),range.getDisplayValues());
  } finally {
    Drive.Files.remove(converted.id);
  }
}

function parseRoster_(values, display) {
  if (display[1][0].trim() !== '번호' || display[1][3].trim() !== '연락처') {
    throw new Error('오아시스 설치명단 양식인지 확인해 주세요.');
  }
  const result = [];
  const sources = new Set();
  let region = '';
  values.slice(2).forEach((row, index)=>{
    const shown = display[index+2];
    const name = shown[2].trim();
    if (!Number.isInteger(row[0]) || !name) {
      region = '';
      return;
    }
    const source = String(row[0]);
    if (sources.has(source)) throw new Error('엑셀에 중복된 참여자 번호가 있습니다.');
    sources.add(source);
    if (shown[1].trim()) region = shown[1].trim();
    const quantities = Object.fromEntries([...row.slice(6,17), ...row.slice(18,24)]
      .flatMap((quantity,index)=>Number.isInteger(quantity) && quantity>0 ? [[DEVICES[index],quantity]] : []));
    result.push({source, body:participant_({name,region,phone:shown[3],address:shown[4],memo:shown[5],
      status:Object.keys(quantities).length?'설치 완료':'설치 대기',quantities})});
  });
  if (!result.length) throw new Error('가져올 참여자 행이 없습니다.');
  return result;
}

function import_(sheet, roster) {
  const existing = records_(sheet);
  const bySource = new Map(existing.filter(r=>r.source).map(r=>[r.source,r]));
  if (roster.some(r=>bySource.has(r.source) && bySource.get(r.source).body.name!==r.body.name)) {
    throw new Error('기존 명단과 같은 번호에 다른 이름이 있습니다. 엑셀 번호와 이름을 확인해 주세요.');
  }
  const properties = PropertiesService.getScriptProperties();
  const key = 'OASIS_SEQUENCE_participants';
  let lastId = existing.reduce((max, record)=>Math.max(max,record.id), Number(properties.getProperty(key)) || 0);
  const values = sheet.getLastRow()>1 ? sheet.getRange(2,1,sheet.getLastRow()-1,HEADERS.length).getValues() : [];
  roster.forEach(record=>{
    const prior = bySource.get(record.source);
    const row = [prior ? prior.id : ++lastId, record.source, JSON.stringify(record.body), 0];
    if (prior) values[prior.row-2] = row;
    else values.push(row);
  });
  const requiredRows = values.length+1;
  if (requiredRows > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows-sheet.getMaxRows());
  properties.setProperty(key,String(lastId));
  sheet.getRange(2,1,values.length,HEADERS.length).setValues(values);
  return {count:roster.length};
}

function edit_(spreadsheet, table, method, rawId, body) {
  const sheet = spreadsheet.getSheetByName(TABS[table]);
  const rows = records_(sheet);
  const id = Number(rawId);
  const current = rows.find(r=>r.id===id);
  if (method !== 'POST' && !current) throw new Error('항목을 찾을 수 없습니다.');
  if (method === 'DELETE') {
    if (table === 'participants') {
      const logs = spreadsheet.getSheetByName(TABS.logs);
      records_(logs).filter(record=>record.body.participant_id===id).reverse().forEach(record=>logs.deleteRow(record.row));
    }
    sheet.deleteRow(current.row);
    return {id};
  }
  const record = current || {id:nextId_(table,rows),row:sheet.getLastRow()+1,source:'',photoCount:0};
  let photo = '';
  if (table === 'participants') {
    record.body = participant_(body);
  } else {
    record.body = monitoring_(body);
    if (!records_(spreadsheet.getSheetByName(TABS.participants)).some(p=>p.id===record.body.participant_id)) {
      throw new Error('연결할 참여자를 확인해 주세요.');
    }
    photo = body.photo ? photo_(body.photo) : (body.keepPhoto ? readPhoto_(sheet, record) : '');
  }
  write_(sheet,record,photo);
  return {id:record.id};
}

function oasisRequest(path, method, body) {
  if (path === 'import' && method === 'POST') {
    database_();
    const roster = roster_(body);
    return locked_(()=>import_(database_().getSheetByName(TABS.participants),roster));
  }
  return locked_(()=>{
    const database = database_();
    if (path === 'state' && method === 'GET') return state_(database);
    const photo = /^logs\/(\d+)\/photo$/.exec(path);
    if (photo && method === 'GET') {
      const sheet = database.getSheetByName(TABS.logs);
      const record = records_(sheet).find(r=>r.id===Number(photo[1]));
      if (!record) throw new Error('일지를 찾을 수 없습니다.');
      return {photo:readPhoto_(sheet,record)};
    }
    const route = /^(participants|logs)(?:\/(\d+))?$/.exec(path);
    if (route && (method === 'POST' && !route[2] || ['PUT','DELETE'].includes(method) && route[2])) {
      return edit_(database,route[1],method,route[2],body);
    }
    throw new Error('지원하지 않는 요청입니다.');
  });
}
