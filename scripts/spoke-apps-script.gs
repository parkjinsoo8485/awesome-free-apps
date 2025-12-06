// 교과서 선정 – Spoke(학교용) Apps Script 최종본
// 메뉴:
//  1) 허브에서 기준 동기화
//  2) 허브에서 교과서 불러오기
//  3) 【발급】위원별 평가표 만들기
//  4) 【수집】위원 제출본 수집
//  5) PDF 결과지 머지(과목 요약)
//
// SETTINGS 시트(키-값) 필수 항목:
//  - RoundID
//  - CriteriaSetID
//  - HubSheetId(붙여넣기)
//  - TemplateDocId(붙여넣기)
//  - OutputFolderId(붙여넣기)
//
// 스포크 시트 필수 탭/헤더:
//  - MASTER_Evaluators: EvaluatorID | 성명 | 소속 | 이메일 | 이해충돌(Y/N) | EvalSheetId | EvalSheetUrl | 완료체크
//  - ROUND_Subjects:    RoundID | SubjectID
//  - ROUND_Books:       RoundID | SubjectID | BookID | 비고
//  - DATA_Evaluations:  EvalID | Timestamp | RoundID | EvaluatorID | SubjectID | BookID | CriteriaID | 점수(1~5) | 코멘트
//  - CRITERIA_CACHE:    CriteriaID | 가중치
//
// 주소표시줄에서 /d/와 /edit 사이가 스프레드시트 ID입니다.
// 예: https://docs.google.com/spreadsheets/d/1ABCdeFG.../edit
const SPOKE_SHEET_ID = '1UoMr5pejG8ktD2GMY3vN2fk0LsQF1eCDpIEfwwEATgU';

function ss_() { 
  return SpreadsheetApp.openById(SPOKE_SHEET_ID);
}
function fixDataEvaluationsHeader() {
  const sh = ss_().getSheetByName('DATA_Evaluations');
  if (!sh) throw new Error('DATA_Evaluations 시트를 찾지 못했습니다.');
  const HEADER = ['EvalID','Timestamp','RoundID','EvaluatorID','SubjectID','BookID','CriteriaID','점수(1~5)','코멘트'];
  sh.getRange(1,1,1,HEADER.length).setValues([HEADER]); // 정확히 교체
  SpreadsheetApp.getUi().alert('DATA_Evaluations 헤더를 표준으로 교체했습니다.');
}

function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error(`${name} 시트를 찾을 수 없습니다.`);
  return s;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('교과서선정')
    .addItem('허브에서 기준 동기화', 'syncCriteriaFromHub')
    .addItem('허브에서 교과서 불러오기', 'loadBooksFromHub')
    .addSeparator()
    .addItem('【발급】위원별 평가표 만들기', 'issueEvaluatorSheets')
    .addItem('【수집】위원 제출본 수집', 'collectAllEvaluatorSheets')
    .addSeparator()
    .addItem('PDF 결과지 머지(과목 요약)', 'mergeSubjectSummaries')
    .addToUi();
}

/* ------------------ 공통 헬퍼 ------------------ */

function _getSettings_() {
  const sh = ss_().getSheetByName('SETTINGS');
  if (!sh) throw new Error('SETTINGS 시트를 찾을 수 없습니다.');
  const m = {};
  sh.getRange(1,1,sh.getLastRow(),2).getValues().forEach(([k,v])=>{ if(k) m[k]=v; });
  const required = ['RoundID','CriteriaSetID','HubSheetId(붙여넣기)'];
  required.forEach(r=>{ if(!m[r]) throw new Error(`SETTINGS의 ${r} 값이 비어있습니다.`); });
  return m;
}
function _openHub_() {
  const id = _getSettings_()['HubSheetId(붙여넣기)'];
  return SpreadsheetApp.openById(id);
}
function _readAsObjects_(sheetName) {
  const sh = ss_().getSheetByName(sheetName);
  if (!sh) throw new Error(`${sheetName} 시트를 찾을 수 없습니다.`);
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const H = rows[0];
  return rows.slice(1).filter(r=>r.some(x=>x!=='')).map(r=>{
    const o={}; H.forEach((h,i)=>o[h]=r[i]); return o;
  });
}
function _writeFromObjects_(sheetName, rows, headerOrder) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName) || SpreadsheetApp.getActive().insertSheet(sheetName);
  sh.clearContents();
  const header = headerOrder || Object.keys(rows[0] || {});
  sh.getRange(1,1,1,header.length).setValues([header]);
  if (rows.length) {
    const data = rows.map(r => header.map(h => r[h] ?? ''));
    sh.getRange(2,1,data.length, header.length).setValues(data);
  }
}
function _nameMaps_() {
  const hub = _openHub_();
  const toObjs = name => {
    const sh = hub.getSheetByName(name); const rows = sh.getDataRange().getValues(); const H=rows[0];
    return rows.slice(1).filter(r=>r.some(x=>x!=='')).map(r=>{ const o={}; H.forEach((h,i)=>o[h]=r[i]); return o;});
  };
  const subs = toObjs('MASTER_Subjects');
  const pubs = toObjs('MASTER_Publishers');
  const books = toObjs('MASTER_Books');
  return {
    subjectName  : Object.fromEntries(subs.map(o=>[String(o.SubjectID), String(o.과목명)])),
    publisherName: Object.fromEntries(pubs.map(o=>[String(o.PublisherID), String(o.출판사명)])),
    bookName     : Object.fromEntries(books.map(o=>[String(o.BookID), String(o.교과서명)])),
    bookPublisher: Object.fromEntries(books.map(o=>[String(o.BookID), String(o.PublisherID)])),
  };
}
function _criteriaRows_() {
  const setId = _getSettings_()['CriteriaSetID'];
  const hub = _openHub_();
  const sh = hub.getSheetByName('MASTER_Criteria');
  const rows = sh.getDataRange().getValues();
  const H = rows[0], idx = n=>H.indexOf(n);
  return rows.slice(1)
    .filter(r => r[idx('CriteriaSetID')]===setId)
    .map(r => ({CriteriaID:String(r[idx('CriteriaID')]), 순서:+r[idx('순서')], 기준명:String(r[idx('기준명')])}))
    .sort((a,b)=>a.순서-b.순서);
}
function _roundSubjects_() {
  const set = _getSettings_();
  const sh = ss_().getSheetByName('ROUND_Subjects');
  const rows = sh.getDataRange().getValues(); const H=rows[0];
  const ridx=H.indexOf('RoundID'), sidx=H.indexOf('SubjectID');
  return rows.slice(1).filter(r=>r[ridx]===set['RoundID']).map(r=>String(r[sidx]));
}
function _roundBooksBySubject_() {
  const set = _getSettings_();
  const sh = ss_().getSheetByName('ROUND_Books');
  const rows = sh.getDataRange().getValues(); const H=rows[0];
  const ridx=H.indexOf('RoundID'), sidx=H.indexOf('SubjectID'), bidx=H.indexOf('BookID');
  const out={};
  rows.slice(1).filter(r=>r[ridx]===set['RoundID']).forEach(r=>{
    const s=String(r[sidx]), b=String(r[bidx]); (out[s]=out[s]||[]).push(b);
  });
  return out; // {SubjectID:[BookID,...]}
}

/* ------------------ 1) 허브에서 기준 동기화 ------------------ */
function syncCriteriaFromHub() {
  const ui = SpreadsheetApp.getUi();
  try {
    const set = _getSettings_();
    const hub = _openHub_();
    const rows = hub.getSheetByName('MASTER_Criteria').getDataRange().getValues();
    const H = rows[0], idx = n=>H.indexOf(n);
    const filtered = rows.slice(1)
      .filter(r=>r[idx('CriteriaSetID')]===set['CriteriaSetID'])
      .map(r=>({'CriteriaID': String(r[idx('CriteriaID')]), '가중치': Number(r[idx('가중치')])||0}));
    _writeFromObjects_('CRITERIA_CACHE', filtered, ['CriteriaID','가중치']);
    ui.alert(`CRITERIA_CACHE 동기화 완료: ${filtered.length}개 기준`);
  } catch(e) { ui.alert('오류', e.message, ui.ButtonSet.OK); }
}

/* ------------------ 2) 허브에서 교과서 불러오기 ------------------ */
function loadBooksFromHub() {
  const ui = SpreadsheetApp.getUi();
  try {
    const set = _getSettings_();
    const hub = _openHub_();
    const hubBooks = hub.getSheetByName('MASTER_Books').getDataRange().getValues();
    const H = hubBooks[0], idx = n=>H.indexOf(n);
    const all = hubBooks.slice(1).filter(r=>r.some(x=>x!=='')).map(r=>({
      BookID   : String(r[idx('BookID')]),
      SubjectID: String(r[idx('SubjectID')]),
    }));
    const subs = _roundSubjects_();
    const chosen = all.filter(b=>subs.includes(b.SubjectID))
                      .map(b=>({RoundID:set['RoundID'], SubjectID:b.SubjectID, BookID:b.BookID, 비고:''}));
    _writeFromObjects_('ROUND_Books', chosen, ['RoundID','SubjectID','BookID','비고']);
    ui.alert(`ROUND_Books 채움 완료: ${chosen.length}권`);
  } catch(e) { ui.alert('오류', e.message, ui.ButtonSet.OK); }
}

/* ------------------ 3) 【발급】위원별 평가표 만들기 ------------------ */
function issueEvaluatorSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const set = _getSettings_();
  const nm = _nameMaps_();
  const crits = _criteriaRows_();
  const subs = _roundSubjects_();
  const booksBySubj = _roundBooksBySubject_();

  const evalSh = ss.getSheetByName('MASTER_Evaluators');
  const rows = evalSh.getDataRange().getValues(); const H=rows[0];
  const idx = n=>H.indexOf(n);
  const iID=idx('EvaluatorID'), iName=idx('성명'), iMail=idx('이메일'),
        iSheetId=idx('EvalSheetId'), iSheetUrl=idx('EvalSheetUrl');

  let created=0, reused=0;

  for (let r=1; r<rows.length; r++) {
    if (!rows[r][iID]) continue;
    const evaluatorId = String(rows[r][iID]);
    const evaluatorName = String(rows[r][iName]||'');
    const email = String(rows[r][iMail]||'');

    let file, fileId, fileUrl;
    if (rows[r][iSheetId]) {
      try { fileId = String(rows[r][iSheetId]); file = SpreadsheetApp.openById(fileId); fileUrl = file.getUrl(); reused++; }
      catch(e){ rows[r][iSheetId]=''; rows[r][iSheetUrl]=''; }
    }
    if (!file) {
      file = SpreadsheetApp.create(`${set['RoundID']}_평가_${evaluatorName||evaluatorId}`);
      fileId = file.getId(); fileUrl = file.getUrl(); created++;
      try { if (email) DriveApp.getFileById(fileId).addEditor(email); } catch(e){}
    }

    // META
    let meta = file.getSheetByName('META') || file.insertSheet('META');
    meta.clear();
    meta.getRange(1,1,6,2).setValues([
      ['RoundID', set['RoundID']],
      ['EvaluatorID', evaluatorId],
      ['EvaluatorName', evaluatorName],
      ['생성일', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')],
      ['SpokeFileId', ss.getId()],
      ['완료체크(TRUE/FALSE)', '']
    ]);
    meta.setHiddenGridlines(true);

    // 안내/종합의견
    let guide = file.getSheetByName('안내/종합의견') || file.insertSheet('안내/종합의견');
    guide.clear();
    guide.getRange(1,1,8,1).setValues([
      ['① 각 과목 탭에서 기준별로 1~10 점수를 입력하세요(소수 허용).'],
      ['② 빈 셀은 누락으로 표시됩니다.'],
      ['③ 제출 버튼은 없습니다. 중앙에서 수집합니다.'],
      ['④ 필요 시 이 시트에 종합의견을 작성하세요.'],
      [''],
      ['과목 | 교과서 | 종합 의견(자유기술)'],
      ['예) 국어 | 국어 3-1(미래엔) | 활동 구성이 균형 있고…'],
      ['']
    ]);

    // 과목별 탭
    subs.forEach(subj=>{
      const bookIds = booksBySubj[subj]||[];
      const tabName = `${nm.subjectName[subj]||subj}(${subj})`.slice(0,99);
      let sh = file.getSheetByName(tabName);
      if (!sh) sh = file.insertSheet(tabName); else sh.clear();

      const header1 = ['CriteriaID','평가기준명'].concat(bookIds.map(b=>`${nm.bookName[b]||b}\n(${nm.publisherName[nm.bookPublisher[b]]||''})`));
      const header2 = ['(숨김)','(숨김)BookID'].concat(bookIds);
      sh.getRange(1,1,1,header1.length).setValues([header1]).setFontWeight('bold');
      sh.getRange(2,1,1,header2.length).setValues([header2]).setFontColor('#aaaaaa');

      const data = crits.map(c=>[c.CriteriaID, c.기준명].concat(new Array(bookIds.length).fill('')));
      if (data.length) sh.getRange(3,1,data.length,data[0].length).setValues(data);

      sh.setFrozenRows(2); sh.setFrozenColumns(2);
      sh.setColumnWidths(1,1,110); sh.setColumnWidths(2,1,240);
      for (let i=3;i<=2+bookIds.length;i++) sh.setColumnWidth(i,120);

      const rangeScores = sh.getRange(3,3,Math.max(1,crits.length), Math.max(1,bookIds.length));
      const rule = SpreadsheetApp.newDataValidation().requireNumberBetween(1,10).setAllowInvalid(false).build();
      rangeScores.setDataValidation(rule);
      rangeScores.setBackground('#fff8e1'); // 연노랑
      sh.getRange(1,1,crits.length+2,header1.length).setBorder(true,true,true,true,true,true,"#dddddd",SpreadsheetApp.BorderStyle.SOLID);
      sh.getRange(1,1,1,header1.length).setBackground('#f1f3f4');
      sh.getRange(1,1,crits.length+2,2).setBackground('#f8f9fa');
    });

    // 링크 기록
    rows[r][iSheetId]  = fileId;
    rows[r][iSheetUrl] = fileUrl;
  }

  evalSh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  SpreadsheetApp.getUi().alert(`평가표 발급 완료 – 새로 발급: ${created}, 기존 재사용: ${reused}`);
}

/* ------------------ 4) 【수집】위원 제출본 수집 ------------------ */
// 완료자만 반영하려면 true로 변경
const ONLY_COMPLETED = false;

function collectAllEvaluatorSheets() {
  const ss = SpreadsheetApp.getActive();
  const set = _getSettings_();

  // MASTER_Evaluators
  const evalSh = ss.getSheetByName('MASTER_Evaluators');
  const rows = evalSh.getDataRange().getValues(); const H=rows[0], idx=n=>H.indexOf(n);
  const iID=idx('EvaluatorID'), iSheetId=idx('EvalSheetId'), iDone=idx('완료체크');

  // DATA_Evaluations
  const dataSh = ss.getSheetByName('DATA_Evaluations');
  const data = dataSh.getDataRange().getValues();
  const DH = data[0], dIdx=n=>DH.indexOf(n);
  const header = ['EvalID','Timestamp','RoundID','EvaluatorID','SubjectID','BookID','CriteriaID','점수(1~5)','코멘트'];
  if (DH.join('|')!==header.join('|')) throw new Error('DATA_Evaluations 헤더가 요구 형식과 다릅니다.');
  let keep = data.slice(1);
  let appended = 0;
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  for (let r=1; r<rows.length; r++) {
    const eId = String(rows[r][iID]||'');
    const fileId = String(rows[r][iSheetId]||'');
    if (!eId || !fileId) continue;

    // 이번 라운드+해당 위원 기존 데이터 제거
    keep = keep.filter(row => !(row[dIdx('RoundID')]===set['RoundID'] && row[dIdx('EvaluatorID')]===eId));

    let f; try { f = SpreadsheetApp.openById(fileId); } catch(e){ continue; }

    // 완료체크 반영
    let doneFlag = false;
    try {
      const meta = f.getSheetByName('META');
      const v = meta ? meta.getRange('B6').getValue() : '';
      doneFlag = String(v).toUpperCase()==='TRUE';
    } catch(e){}
    if (iDone !== -1) rows[r][iDone] = doneFlag;
    if (ONLY_COMPLETED && !doneFlag) continue;

    // 과목 탭에서 점수 읽기
    const sheets = f.getSheets().filter(s=>/\(.+\)$/.test(s.getName())); // "국어(SUB_...)" 형식
    sheets.forEach(sh=>{
      const lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
      if (lastRow<3 || lastCol<3) return;

      const header2 = sh.getRange(2,1,1,lastCol).getValues()[0]; // 2행: BookID
      const bookIds = header2.slice(2).map(String);
      const body = sh.getRange(3,1,lastRow-2,lastCol).getValues(); // A:CriteriaID, B:기준명, C..점수
      const subjectId = sh.getName().match(/\((.+)\)$/)[1];

      body.forEach(row=>{
        const criteriaId = String(row[0]||'');
        for (let c=0;c<bookIds.length;c++) {
          const score = row[2+c];
          if (typeof score==='number' && !isNaN(score)) {
            const bookId = bookIds[c];
            const evalId = `${set['RoundID']}|${eId}|${subjectId}|${bookId}|${criteriaId}|${now}|${Math.random().toString(36).slice(2,8)}`;
            keep.push([evalId, now, set['RoundID'], eId, subjectId, bookId, criteriaId, score, '']);
            appended++;
          }
        }
      });
    });
  }

  // 결과 반영
  evalSh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  dataSh.clearContents();
  dataSh.getRange(1,1,1,header.length).setValues([header]);
  if (keep.length) dataSh.getRange(2,1,keep.length,header.length).setValues(keep);

  SpreadsheetApp.getUi().alert(`수집 완료: 새로 반영 ${appended}건 (완료체크 갱신됨)`);
}

/* ------------------ 5) PDF 결과지 머지(과목 요약) ------------------ */
function mergeSubjectSummaries() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActive();
    const set = _getSettings_();
    const templateId = set['TemplateDocId(붙여넣기)'];
    const folderId   = set['OutputFolderId(붙여넣기)'];
    if (!templateId || !folderId) throw new Error('SETTINGS에 TemplateDocId / OutputFolderId를 입력하십시오.');

    // 가중치 맵
    const wRows = _readAsObjects_('CRITERIA_CACHE');
    const weight = Object.fromEntries(wRows.map(r=>[String(r['CriteriaID']), Number(r['가중치'])||0]));

    // 집계
    const dataSh = ss.getSheetByName('DATA_Evaluations');
    const rows = dataSh.getDataRange().getValues();
    const H = rows[0], idx=n=>H.indexOf(n);
    const agg = {}; // key: subj|book -> {sum, wsum}
    rows.slice(1).forEach(r=>{
      if (r[idx('RoundID')]!==set['RoundID']) return;
      const subj = String(r[idx('SubjectID')]);
      const book = String(r[idx('BookID')]);
      const cri  = String(r[idx('CriteriaID')]);
      const score= Number(r[idx('점수(1~5)')])||0;
      const w = weight[cri]||0;
      const key = `${subj}|${book}`;
      if (!agg[key]) agg[key]={sum:0,wsum:0};
      agg[key].sum += score*w; agg[key].wsum += w;
    });

    // 과목별 상위 3
    const nm = _nameMaps_();
    const subs = _roundSubjects_();
    const folder = DriveApp.getFolderById(folderId);
    let count = 0;

    subs.forEach(subj=>{
      const arr=[];
      Object.keys(agg).forEach(k=>{
        const [s,b] = k.split('|');
        if (s!==subj) return;
        const {sum,wsum}=agg[k]; if (!wsum) return;
        arr.push({BookID:b, WeightedAvg: sum/wsum});
      });
      arr.sort((a,b)=>b.WeightedAvg-a.WeightedAvg);
      const top = arr.slice(0,3);
      const subjectName = nm.subjectName[subj] || subj;

      const replace = {
        '{{RoundID}}': set['RoundID'],
        '{{SubjectID}}': subj,
        '{{SubjectName}}': subjectName,
        '{{Today}}': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        '{{Top1_Book}}': top[0] ? (nm.bookName[top[0].BookID]||top[0].BookID) : '',
        '{{Top1_Publisher}}': top[0] ? (nm.publisherName[nm.bookPublisher[top[0].BookID]]||'') : '',
        '{{Top1_Score}}': top[0] ? top[0].WeightedAvg.toFixed(3) : '',
        '{{Top2_Book}}': top[1] ? (nm.bookName[top[1].BookID]||top[1].BookID) : '',
        '{{Top2_Publisher}}': top[1] ? (nm.publisherName[nm.bookPublisher[top[1].BookID]]||'') : '',
        '{{Top2_Score}}': top[1] ? top[1].WeightedAvg.toFixed(3) : '',
        '{{Top3_Book}}': top[2] ? (nm.bookName[top[2].BookID]||top[2].BookID) : '',
        '{{Top3_Publisher}}': top[2] ? (nm.publisherName[nm.bookPublisher[top[2].BookID]]||'') : '',
        '{{Top3_Score}}': top[2] ? top[2].WeightedAvg.toFixed(3) : '',
      };

      const copy = DriveApp.getFileById(templateId).makeCopy(`${set['RoundID']}_${subjectName}_요약`, folder);
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();
      Object.keys(replace).forEach(k=>body.replaceText(k, String(replace[k])));
      doc.saveAndClose();

      const pdfBlob = DriveApp.getFileById(copy.getId()).getAs('application/pdf').setName(`${set['RoundID']}_${subjectName}_요약.pdf`);
      folder.createFile(pdfBlob);
      count++;
    });

    ui.alert(`PDF 생성 완료: 과목 요약 ${count}건`);
  } catch(e) { ui.alert('오류', e.message, ui.ButtonSet.OK); }
}
/** ▶ 웹앱 초기 데이터 제공: 라운드/과목/책/기준/기존점수 */
function api_init(rid, eid) {
  const set = _getSettings_();
  const roundId = String((rid||'').toString().trim() || set['RoundID']).trim();
  const evaluatorId = String((eid||'').toString().trim());

  if (!roundId) throw new Error('RoundID가 비어 있습니다. URL 또는 SETTINGS를 확인하세요.');

  // 이번 라운드 과목
  const subjIds = _roundSubjects_(); // ['SUB_KOR_3', ...]
  // 과목별 책
  const booksBySubj = _roundBooksBySubject_(); // {SUB: [B001, ...]}
  // 이름/출판사 맵
  const nm = _nameMaps_();
  // 기준
  const crits = _criteriaRows_(); // [{CriteriaID, 순서, 기준명}...]
  // 가중치
  const wRows = _readAsObjects_('CRITERIA_CACHE');
  const weight = Object.fromEntries(wRows.map(r=>[String(r['CriteriaID']), Number(r['가중치'])||0]));
  crits.forEach(c => c.가중치 = weight[c.CriteriaID] || 0);

  // 기존 입력 복원(해당 evaluatorId만)
  const data = ss_().getSheetByName('DATA_Evaluations').getDataRange().getValues();
  const H = data[0], idx = n => H.indexOf(n);
  const existing = {}; // existing[SubjectID][CriteriaID][BookID] = score
  data.slice(1).forEach(r=>{
    if (String(r[idx('RoundID')]).trim() !== roundId) return;
    if (evaluatorId && String(r[idx('EvaluatorID')]).trim() !== evaluatorId) return;
    const subj = String(r[idx('SubjectID')]).trim();
    const book = String(r[idx('BookID')]).trim();
    const cri  = String(r[idx('CriteriaID')]).trim();
    const score= Number(r[idx('점수(1~5)')]); // 열 이름은 1~5지만 1~10 점수도 숫자면 OK
    if (!existing[subj]) existing[subj]={};
    if (!existing[subj][cri]) existing[subj][cri]={};
    if (!isNaN(score)) existing[subj][cri][book] = score;
  });

  const subjects = subjIds.map(id => ({ id, name: nm.subjectName[id] || id }));
  const books = {};
  subjIds.forEach(s=>{
    books[s] = (booksBySubj[s]||[]).map(bid => ({
      id: bid,
      title: nm.bookName[bid] || bid,
      publisher: nm.publisherName[nm.bookPublisher[bid]] || ''
    }));
  });

  return { roundId, evaluatorId, subjects, books, criteria: crits, existing };
}

/** ▶ 현재 과목 저장(이 위원이 이 과목에 쓴 기존행 삭제 → 새 점수 삽입) */
function api_saveSubject(roundId, evaluatorId, payload) {
  if (!roundId || !evaluatorId) throw new Error('roundId/evaluatorId가 필요합니다.');
  const subjectId = payload && payload.subjectId;
  const rowsByCrit = payload && payload.rows;
  if (!subjectId || !rowsByCrit || !rowsByCrit.length) return { inserted: 0 };

  const sh = ss_().getSheetByName('DATA_Evaluations');
  const values = sh.getDataRange().getValues();
  const H = values[0], idx = n => H.indexOf(n);

  // 1) 기존(같은 Round+Evaluator+Subject) 행 제거
  const keep = values.slice(1).filter(r =>
    !(String(r[idx('RoundID')])===roundId &&
      String(r[idx('EvaluatorID')])===evaluatorId &&
      String(r[idx('SubjectID')])===subjectId)
  );

  // 2) 새 행 만들기
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const header = ['EvalID','Timestamp','RoundID','EvaluatorID','SubjectID','BookID','CriteriaID','점수(1~5)','코멘트'];
  if (H.join('|') !== header.join('|')) throw new Error('DATA_Evaluations 헤더가 요구 형식과 다릅니다.');

  let inserted = 0;
  rowsByCrit.forEach(row=>{
    const cri = row.criteriaId;
    (row.scores||[]).forEach(sc=>{
      const bookId = sc.bookId;
      const score  = Number(sc.score);
      if (!isNaN(score)) {
        const evalId = `${roundId}|${evaluatorId}|${subjectId}|${bookId}|${cri}|${now}|${Math.random().toString(36).slice(2,8)}`;
        keep.push([evalId, now, roundId, evaluatorId, subjectId, bookId, cri, score, '']);
        inserted++;
      }
    });
  });

  // 3) 되쓰기
  sh.clearContents();
  sh.getRange(1,1,1,header.length).setValues([header]);
  if (keep.length) sh.getRange(2,1,keep.length,header.length).setValues(keep);

  return { inserted };
}

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index'); // 파일명: Index (확장자 없이)
  t.params = {
    rid: (e && e.parameter && e.parameter.rid) || '',
    eid: (e && e.parameter && e.parameter.eid) || ''
  };
  return t.evaluate()
           .setTitle('교과서 선정 – 평가 입력')
           .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
