const SCHEMA_VERSION = 'bsc-beauty-co-creation-v5';
const MANAGEMENT_SPREADSHEET_ID = '1qmqs6cBtWZg0YPvWMqkBET61iREmT3PHW6Xsnpwghxk';
const OLD_CONTACT_PROPERTY = 'BSC_GUIDANCE_CONTACT_SHEET_ID';
const SHEET_NAMES = {
  owner: 'サロン・組合員回答',
  practitioner: '外部施術者回答',
  contacts: '案内希望者連絡先'
};
const OWNER_HEADERS = [
  '回答受付ID','回答日時','現在のお立場','店舗の都道府県','市区町村','店舗の形態','所有・賃貸',
  '活用を検討できそうな場所・時間','今後のお店と暮らしで大切にしたいこと','お客様から実際に希望されたサービス',
  '印象に残っているご要望','この構想への関心度','受入れを検討できる施術','その他の施術',
  'お客様への紹介条件','検討できる利用料方式','将来の役割移行','不安・検討しにくい理由',
  '期間限定モニター意向','検討できそうな時期','補足','今後の情報案内希望'
];
const PRACTITIONER_HEADERS = [
  '回答受付ID','回答日時','現在の主な状況','活かしたい技術','資格・修了歴','その他の技術・資格',
  '実務経験','この構想への関心','関心が低い理由','希望する都道府県','希望する市区町村',
  '希望する曜日・時間','施術方法','ご自身のお客様の目安','始めるために大切な条件',
  '検討できる利用料方式','不安なこと','期間限定モニター意向','検討時期','補足','今後の情報案内希望'
];
const CONTACT_HEADERS = ['回答受付ID','回答日時','回答区分','メールアドレス','案内内容','利用目的への同意'];

function doPost(e) {
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw || raw.length > 30000) throw new Error('送信形式');
    const data = JSON.parse(raw);
    validate_(data);
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      writeAnswer_(data, new Date(), false);
      if (data.answers.contact_wish === 'yes') writeContact_(data, new Date());
    } finally {
      lock.releaseLock();
    }
    return page_('ご意見ありがとうございました。現段階では登録・契約は成立していません。', true);
  } catch (error) {
    console.error(error);
    return page_('送信を確認できませんでした。前の画面へ戻って内容を確認してください。', false);
  }
}

function setupManagementSpreadsheet() {
  const ss = SpreadsheetApp.openById(MANAGEMENT_SPREADSHEET_ID);
  const owner = ensureSheet_(ss, SHEET_NAMES.owner, OWNER_HEADERS);
  const practitioner = ensureSheet_(ss, SHEET_NAMES.practitioner, PRACTITIONER_HEADERS);
  const contacts = ensureSheet_(ss, SHEET_NAMES.contacts, CONTACT_HEADERS);
  migrateFormResponses_(owner, practitioner);
  migrateOldContacts_(contacts);
  [owner, practitioner, contacts].forEach((sheet, index) => {
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setBackground('#073b2c').setFontColor('#fff9e9').setFontWeight('bold').setWrap(true);
    sheet.setColumnWidth(1, 190);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidths(3, Math.max(1, sheet.getLastColumn() - 2), 190);
    if (sheet.getFilter()) sheet.getFilter().remove();
    sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), sheet.getLastColumn()).createFilter();
  });
  ss.setActiveSheet(owner);
  return '管理タブを準備しました。既存のフォーム回答と連絡先は元のシートにも残しています。';
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const existing = sheet.getLastRow() ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0] : [];
  if (!existing.length || existing.every(value => value === '')) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  else if (existing.slice(0, headers.length).join('\u001f') !== headers.join('\u001f')) throw new Error('既存タブの列見出しが想定と異なります: ' + name);
  return sheet;
}

function migrateFormResponses_(ownerSheet, practitionerSheet) {
  const source = SpreadsheetApp.openById(MANAGEMENT_SPREADSHEET_ID).getSheetByName('フォームの回答 1');
  if (!source || source.getLastRow() < 2) return;
  const values = source.getDataRange().getValues();
  const headers = values.shift().map(String);
  const jsonColumn = headers.indexOf('回答内容（HTMLフォーム連携用）');
  const dateColumn = headers.indexOf('タイムスタンプ');
  if (jsonColumn < 0) throw new Error('フォーム原票に回答内容JSON列がありません');
  values.forEach((row, index) => {
    try {
      const answers = JSON.parse(row[jsonColumn]);
      if (!answers || !answers.role) return;
      const timestamp = row[dateColumn] || new Date();
      const payload = {
        role: answers.role,
        answers: answers,
        submission_id: 'legacy-' + new Date(timestamp).getTime() + '-' + (index + 2)
      };
      payload.answers.contact_wish = payload.answers.contact_wish || 'unknown';
      writeAnswer_(payload, timestamp, true, ownerSheet, practitionerSheet);
    } catch (error) {
      console.error('原票の回答を読み取れませんでした: ' + error);
    }
  });
}

function migrateOldContacts_(target) {
  const oldId = PropertiesService.getScriptProperties().getProperty(OLD_CONTACT_PROPERTY);
  if (!oldId) return;
  const oldBook = SpreadsheetApp.openById(oldId);
  oldBook.getSheets().forEach(source => {
    const values = source.getDataRange().getValues();
    if (values.length < 2) return;
    const headers = values[0].map(String);
    const col = title => headers.indexOf(title);
    values.slice(1).forEach(row => {
      const submissionId = String(row[col('回答受付ID')] || '');
      const email = String(row[col('メールアドレス')] || '');
      if (!email) return;
      const at = row[col('登録日時')] || new Date();
      const purpose = String(row[col('希望内容')] || '詳しい説明・条件の案内');
      const consent = row[col('利用目的への同意')] === true || String(row[col('利用目的への同意')]).toLowerCase() === 'true';
      const role = String(row[col('回答区分')] || inferLegacyRole_(at) || '');
      const id = submissionId || ('legacy-' + email + '-' + at);
      const targetLastRow = target.getLastRow();
      const match = targetLastRow > 1 ? target.getRange(2, 1, targetLastRow - 1, 1).createTextFinder(id).matchEntireCell(true).findNext() : null;
      if (match) {
        if (role && !target.getRange(match.getRow(), 3).getValue()) target.getRange(match.getRow(), 3).setValue(role);
      } else {
        appendUnique_(target, id, [id, at, role, email, purpose, consent ? '同意あり' : '同意記録なし']);
      }
    });
  });
}

function inferLegacyRole_(timestamp) {
  const source = SpreadsheetApp.openById(MANAGEMENT_SPREADSHEET_ID).getSheetByName('フォームの回答 1');
  if (!source || source.getLastRow() < 2) return '';
  const values = source.getDataRange().getValues();
  const headers = values.shift().map(String);
  const jsonColumn = headers.indexOf('回答内容（HTMLフォーム連携用）');
  const dateColumn = headers.indexOf('タイムスタンプ');
  if (jsonColumn < 0 || dateColumn < 0) return '';
  const minute = Math.floor(new Date(timestamp).getTime() / 60000);
  for (const row of values) {
    if (Math.floor(new Date(row[dateColumn]).getTime() / 60000) !== minute) continue;
    try {
      const answers = JSON.parse(row[jsonColumn]);
      return answers.role === 'owner' ? 'サロン・組合員' : answers.role === 'practitioner' ? '外部施術者' : '';
    } catch (_) {}
  }
  return '';
}

function writeAnswer_(data, timestamp, migrating, ownerSheet, practitionerSheet) {
  const ss = SpreadsheetApp.openById(MANAGEMENT_SPREADSHEET_ID);
  const answers = data.answers;
  const guidance = guidanceLabel_(answers.contact_wish);
  const row = data.role === 'owner' ? ownerRow_(data, timestamp, guidance) : practitionerRow_(data, timestamp, guidance);
  const sheet = data.role === 'owner' ? (ownerSheet || ss.getSheetByName(SHEET_NAMES.owner)) : (practitionerSheet || ss.getSheetByName(SHEET_NAMES.practitioner));
  if (!sheet) throw new Error('管理タブがありません。setupManagementSpreadsheet を実行してください。');
  appendUnique_(sheet, String(data.submission_id), row);
  if (migrating && answers.contact_wish === 'yes' && data.contact_email && data.contact_consent) writeContact_(data, timestamp);
}

function ownerRow_(data, timestamp, guidance) {
  const a = data.answers, d = parseJson_(a.owner_details);
  return [data.submission_id, timestamp, label_(d.position), cell_(d.prefecture), cell_(d.city), label_(d.place_type), label_(d.tenure),
    list_(String(a.space || '').split(',')), list_(d.priorities), list_(d.customer_voice), cell_(d.customer_note), label_(a.interest),
    list_(d.services), cell_(d.other_service), label_(d.referral), list_(d.fee_models), label_(d.succession), list_(d.concerns),
    label_(d.monitor), label_(d.timing), cell_(a.opinion), guidance];
}

function practitionerRow_(data, timestamp, guidance) {
  const a = data.answers, d = parseJson_(a.practitioner_details);
  const schedule = String(a.schedule || '').split(' / ');
  let prefecture = d.prefecture || '', city = d.city || '', days = d.preferred_days || [];
  if (!prefecture && schedule.length) {
    const region = schedule.shift();
    const match = region.match(/^(.*?(?:都|道|府|県|未定|回答しない))(?:\s+(.*))?$/);
    if (match) { prefecture = match[1]; city = city || match[2] || ''; }
  }
  if (!days.length) days = schedule;
  return [data.submission_id, timestamp, label_(a.work_style), label_(a.skills), label_(a.qualifications), cell_(d.other_skill), label_(d.experience),
    label_(a.interest), list_(d.low_interest_reasons), cell_(prefecture), cell_(city), list_(days), label_(d.customer_model || a.use_style),
    cell_(d.customer_count), list_(d.important_conditions), list_(d.fee_models), list_(d.concerns), label_(d.monitor), label_(d.timing), cell_(a.opinion), guidance];
}

function writeContact_(data, timestamp) {
  if (!data.contact_email || !data.contact_consent) throw new Error('案内希望者の同意・メールアドレスがありません');
  const sheet = SpreadsheetApp.openById(MANAGEMENT_SPREADSHEET_ID).getSheetByName(SHEET_NAMES.contacts);
  if (!sheet) throw new Error('連絡先タブがありません。setupManagementSpreadsheet を実行してください。');
  appendUnique_(sheet, String(data.submission_id), [data.submission_id, timestamp, data.role === 'owner' ? 'サロン・組合員' : '外部施術者', data.contact_email, '詳しい説明・条件が決まった際の案内', '同意あり']);
}

function appendUnique_(sheet, id, row) {
  const last = sheet.getLastRow();
  if (last > 1 && sheet.getRange(2, 1, last - 1, 1).createTextFinder(id).matchEntireCell(true).findNext()) return;
  sheet.appendRow(row.map(safeCell_));
}

function validate_(data) {
  if (!data || data.schema_version !== SCHEMA_VERSION || !['owner','practitioner'].includes(data.role)) throw new Error('回答区分');
  if (!/^[0-9a-f-]{36}$/i.test(String(data.submission_id || ''))) throw new Error('受付番号');
  if (!data.answers || typeof data.answers !== 'object' || data.answers.role !== data.role) throw new Error('回答内容');
  if (!['interested','conditional','undecided','less_interested','not_interested'].includes(data.answers.interest)) throw new Error('関心度');
  if (!['yes','no','unknown'].includes(data.answers.contact_wish)) throw new Error('今後の案内希望');
  if (data.answers.contact_wish === 'yes') {
    if (typeof data.contact_email !== 'string' || data.contact_email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.contact_email)) throw new Error('メールアドレス');
    if (data.contact_consent !== true) throw new Error('同意');
  } else if (data.contact_email || data.contact_consent) throw new Error('案内希望のない連絡先');
}

function parseJson_(value) { try { return typeof value === 'string' ? JSON.parse(value || '{}') : (value || {}); } catch (_) { return {}; } }
function list_(value) { return (Array.isArray(value) ? value : (value == null ? [] : [value])).filter(item => item !== '').map(label_).join('、'); }
function cell_(value) { return value == null ? '' : String(value); }
function safeCell_(value) { return typeof value === 'string' && /^[=+@\-]/.test(value) ? "'" + value : value; }
function guidanceLabel_(value) { return ({yes:'案内を希望（連絡先タブに登録）',no:'希望しない',unknown:'未定'})[value] || '未定'; }
function label_(value) {
  if (Array.isArray(value)) return value.map(label_).join('、');
  const map = {
    interested:'とても関心がある',conditional:'やや関心がある',undecided:'どちらともいえない',less_interested:'あまり関心がない',not_interested:'関心がない',
    owner_active:'自分も施術しているオーナー',owner_reducing:'施術を減らしているオーナー',owner_retired:'現在はほとんど施術していないオーナー',successor:'後継者・運営責任者',owner_other:'その他',
    owner:'サロン・組合員',practitioner:'外部施術者',holiday:'定休日',after_hours:'営業時間外',empty_seat:'営業中の空いている席',unused_room:'使っていない部屋・一角',when_present:'自分が店にいる時間に限る',no_space:'現時点では活用が難しい',unknown:'まだ分からない',
    continue:'施術量を減らしてもお店を続けたい',income:'自分の施術以外からも収入を得る道を探したい',customers:'お客様の通い慣れた場所を残したい',new_service:'お客様に新しい美容サービスを届けたい',values:'技術やビーワンへの想いを次につなぎたい',keep:'今の形を続けることを優先したい',none:'今は特にない',
    nail:'ネイル',lash:'まつげ・眉',headspa:'ウェットヘッドスパ・ドライヘッドスパ',wet_headspa:'ウェットヘッドスパ',dry_headspa:'ドライヘッドスパ',facial:'エステ・フェイシャル',hair:'カット・カラーなどのヘア施術',other:'その他',other_service:'その他',
    owner_place_type:'',home:'自宅と併設',independent:'独立した店舗',commercial:'ビル・商業施設内',owner_tenure:'',owned:'自己所有',rented:'賃貸',
    proactive:'自分から案内を検討したい',requested:'お客様から希望があれば案内したい',trust:'人柄や技術を確認してから考えたい',experience:'自分で体験してから考えたい',not_now:'現時点では考えていない',
    commission:'売上に応じた歩合',hourly:'1時間単位の利用料',day:'半日・1日単位の利用料',monthly:'月額の固定利用料',combined:'固定利用料と歩合の組合せ',
    consider:'一緒に働いてから、顧客や店舗の引継ぎも考えたい',partial:'顧客・技術・理念の一部ならつなぎたい',space_only:'場所の活用には関心があるが、継承は考えていない',
    need:'今は必要性を感じない',space:'場所や時間を確保できない',demand:'お客様が利用するか分からない',operations:'予約・案内・管理の手間',cost:'費用を差し引いて収入が残るか',liability:'事故やトラブル時の責任・契約条件',lease:'賃貸契約・設備の条件',
    details:'具体的な説明を聞きたい',conditions:'条件が合えば検討したい',later:'将来なら検討したい',info:'情報収集のみ希望する',no:'現時点では考えていない',soon:'3か月以内',half_year:'半年以内',year:'1年以内',unset:'時期は未定',
    weekday:'平日',weekend:'土日・祝日',salon_holiday:'サロンの定休日',morning:'午前',afternoon:'午後',evening:'夕方以降',
    own:'自分から案内できるお客様を迎えたい',referred:'サロンから紹介されたお客様を担当したい',both:'両方を考えたい',
    short:'短期間・少ない日数から試せる',price:'費用や売上分配が事前に分かる',equipment:'必要な設備・衛生条件が整っている',schedule:'曜日や時間を相談できる',introduction:'店主の紹介や既存のお客様との接点がある',own_customers:'自分のお客様を迎えられる',interior:'内装や設備を自分の基準に合わせられる',
    customers_concern:'お客様を集められるか',qualification:'資格・設備・衛生の条件',work:'今の仕事と両立できるか',
    independent:'独立して活動している',employed:'サロン等に勤務している',side:'副業・兼業を希望している',student:'学習中',
    practitioner_active:'現在施術をしている',
    '1':'関心がない','2':'あまり関心がない','3':'どちらともいえない','4':'やや関心がある','5':'とても関心がある'
  };
  const text = value == null ? '' : String(value);
  return Object.prototype.hasOwnProperty.call(map, text) ? (map[text] || text) : text;
}

function page_(message, success) {
  const title = success ? '回答を受け付けました' : '送信状況を確認してください';
  const safe = message.replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  return HtmlService.createHtmlOutput('<!doctype html><html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>'+title+'</title><style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#003d2b;color:#153a2d;font-family:system-ui,sans-serif}main{width:min(90vw,38rem);padding:2.5rem;border:1px solid #c5a456;border-radius:20px;background:#fff5dc;line-height:1.8}h1{font-family:serif;font-size:1.8rem}</style></head><body><main><h1>'+title+'</h1><p>'+safe+'</p><p>前の画面へ戻るにはブラウザの「戻る」をご利用ください。</p></main></body></html>');
}
