const SCHEMA_VERSION = 'bsc-beauty-co-creation-v4';
const CONTACT_SHEET_PROPERTY = 'BSC_GUIDANCE_CONTACT_SHEET_ID';
const ANSWER_KEYS = {
  owner: ['role','interest','owner_roles','space','customer_needs','owner_details','owner_area','owner_size','opinion','contact_wish'],
  practitioner: ['role','interest','skills','qualifications','work_style','use_style','schedule','conditions','practitioner_details','opinion','contact_wish']
};

function doPost(e) {
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw || raw.length > 30000) throw new Error('送信形式');
    const data = JSON.parse(raw);
    validate_(data);
    const form = FormApp.getActiveForm();
    if (!form) throw new Error('保存先フォームがありません');
    let response = form.createResponse();
    const role = data.role === 'owner' ? 'サロンオーナー・BSC組合員' : '外部施術者';
    const interest = {
      interested: 'とても関心がある',
      conditional: '条件次第',
      undecided: 'まだ判断できない',
      less_interested: 'あまり関心がない',
      not_interested: '関心がない'
    }[data.answers.interest];
    const guidance = {
      yes: '詳しい説明を希望する',
      conditional: '条件が決まったら案内を希望する',
      no: '今回は希望しない'
    }[data.answers.contact_wish];
    response = addChoice_(form, response, 'あなたのお立場を教えてください', role);
    response = addChoice_(form, response, 'この構想にどの程度関心がありますか？', interest);
    response = addParagraph_(form, response, '回答内容（HTMLフォーム連携用）', JSON.stringify(data.answers));
    const area = data.role === 'owner' ? data.answers.owner_area : '';
    if (area) response = addText_(form, response, '店舗または希望する活動地域（都道府県・市区町村まで、任意）', area);
    if (data.answers.opinion) response = addParagraph_(form, response, 'ご期待や気になる点があれば教えてください（任意）', data.answers.opinion);
    response = addChoice_(form, response, '今後の説明やモニター案内について', guidance);
    if (data.answers.contact_wish === 'yes') saveContact_(data);
    response.submit();
    return page_('ご意見ありがとうございました。現段階では登録・契約は成立していません。', true);
  } catch (error) {
    console.error(error);
    return page_('送信を確認できませんでした。前の画面へ戻って内容を確認してください。', false);
  }
}

function validate_(data) {
  if (!data || data.schema_version !== SCHEMA_VERSION || !ANSWER_KEYS[data.role]) throw new Error('回答区分');
  if (!/^[0-9a-f-]{36}$/.test(String(data.submission_id || ''))) throw new Error('受付番号');
  if (!data.answers || typeof data.answers !== 'object' || data.answers.role !== data.role) throw new Error('回答内容');
  if (!['interested','conditional','undecided','less_interested','not_interested'].includes(data.answers.interest)) throw new Error('関心度');
  if (!['yes','conditional','no'].includes(data.answers.contact_wish)) throw new Error('今後の案内希望');
  const email = String(data.contact_email || '').trim();
  if (email.length > 254 || (data.answers.contact_wish === 'yes' && (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || data.contact_consent !== true))) throw new Error('案内用メールアドレス');
  if (data.answers.contact_wish !== 'yes' && (email || data.contact_consent !== false)) throw new Error('案内希望なしの連絡先');
  for (const [key, value] of Object.entries(data.answers)) {
    if (!ANSWER_KEYS[data.role].includes(key)) throw new Error('不明な設問');
    const values = Array.isArray(value) ? value : [value];
    if (values.length > 12 || values.some(item => typeof item !== 'string' || item.length > 2000)) throw new Error('回答の長さ');
  }
}

function addChoice_(form, response, title, value) {
  const item = itemByTitle_(form, title, FormApp.ItemType.MULTIPLE_CHOICE).asMultipleChoiceItem();
  return response.withItemResponse(item.createResponse(value));
}

function addText_(form, response, title, value) {
  const item = itemByTitle_(form, title, FormApp.ItemType.TEXT).asTextItem();
  return response.withItemResponse(item.createResponse(value));
}

function setupContactSpreadsheet() {
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty(CONTACT_SHEET_PROPERTY);
  if (existingId) return SpreadsheetApp.openById(existingId).getUrl();
  const spreadsheet = SpreadsheetApp.create('BSCビューティー共創モデル｜案内希望者連絡先');
  const sheet = spreadsheet.getSheets()[0];
  sheet.setName('案内希望者');
  sheet.appendRow(['登録日時','メールアドレス','希望内容','回答受付ID','利用目的への同意']);
  sheet.setFrozenRows(1);
  properties.setProperty(CONTACT_SHEET_PROPERTY, spreadsheet.getId());
  return spreadsheet.getUrl();
}

function saveContact_(data) {
  const id = PropertiesService.getScriptProperties().getProperty(CONTACT_SHEET_PROPERTY);
  if (!id) throw new Error('案内希望者シートの初期設定が未完了です');
  const sheet = SpreadsheetApp.openById(id).getSheetByName('案内希望者');
  if (!sheet) throw new Error('案内希望者シートがありません');
  const submissionId = data.submission_id;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, 4, lastRow - 1, 1).getDisplayValues().flat();
    if (existingIds.includes(submissionId)) return;
  }
  sheet.appendRow([new Date(), data.contact_email, '詳しい説明・条件案内', submissionId, '同意']);
}

function addParagraph_(form, response, title, value) {
  const item = itemByTitle_(form, title, FormApp.ItemType.PARAGRAPH_TEXT).asParagraphTextItem();
  return response.withItemResponse(item.createResponse(value));
}

function itemByTitle_(form, title, type) {
  const item = form.getItems().find(candidate => candidate.getTitle() === title);
  if (!item || item.getType() !== type) throw new Error('設問の設定不一致: ' + title);
  return item;
}

function page_(message, success) {
  const title = success ? '回答を受け付けました' : '送信状況を確認してください';
  const safe = message.replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  return HtmlService.createHtmlOutput('<!doctype html><html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>'+title+'</title><style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#003d2b;color:#153a2d;font-family:system-ui,sans-serif}main{width:min(90vw,38rem);padding:2.5rem;border:1px solid #c5a456;border-radius:20px;background:#fff5dc;line-height:1.8}h1{font-family:serif;font-size:1.8rem}a{color:#003d2b}</style></head><body><main><h1>'+title+'</h1><p>'+safe+'</p><p>前の画面へ戻るにはブラウザの「戻る」をご利用ください。</p></main></body></html>');
}
