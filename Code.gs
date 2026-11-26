// ╔══════════════════════════════════════════════════════════════════╗
// ║  Muse Lab Studio — Apps Script 後端（完整版）                   ║
// ║                                                                  ║
// ║  試算表需要 4 個工作表分頁：                                     ║
// ║  1. bookings：A時間戳記 B姓名 C電話 D Email E預約日期 F時段 G套數║
// ║  2. blocked ：A類型(date/slot) B日期 C時段                      ║
// ║  3. config  ：A admin_password  B 你的密碼                      ║
// ║  4. email_template：A欄位名 B欄位值                             ║
// ║     （subject / payUrl / body 三列）                            ║
// ╚══════════════════════════════════════════════════════════════════╝

var NOTIFY_EMAIL      = 'muselab.tw@gmail.com';
var SENDER_ALIAS      = 'Muse Lab Studio';
var SHEET_BOOKINGS    = 'bookings';
var SHEET_BLOCKED     = 'blocked';
var SHEET_CONFIG      = 'config';
var SHEET_EMAIL_TPL   = 'email_template';
var SHEET_DEPOSIT_TPL  = 'deposit_template';
var SHEET_CANCEL_TPL    = 'cancel_template';
var SHEET_RESCHEDULE_TPL = 'reschedule_template';
var SPREADSHEET_ID    = '13CGSwtcOleR4KVu7Jx2ZHYEJDEaRvTTBjkeE8YkuQLM';
function SS(){ return SpreadsheetApp.openById(SPREADSHEET_ID); }

// ── Router ────────────────────────────────────────────────────────
function doGet(e) {
  var action   = e.parameter.action  || '';
  var callback = e.parameter.callback || '';
  var result;
  try {
    if      (action === 'getBooked')          result = apiGetBooked();
    else if (action === 'book')               result = apiBook(e.parameter);
    else if (action === 'login')              result = apiLogin(e.parameter);
    else if (action === 'loginGoogle')        result = apiLoginGoogle(e.parameter);
    else if (action === 'getBookings')        result = apiGetBookings(e.parameter.token);
    else if (action === 'getBlocked')         result = apiGetBlocked(e.parameter.token);
    else if (action === 'deleteBooking')      result = apiDeleteBooking(e.parameter.token, e.parameter.id);
    else if (action === 'blockDate')          result = apiBlockDate(e.parameter);
    else if (action === 'blockSlot')          result = apiBlockSlot(e.parameter);
    else if (action === 'unblock')            result = apiUnblock(e.parameter);
    else if (action === 'getEmailTemplate')   result = apiGetEmailTemplate(e.parameter.token);
    else if (action === 'saveEmailTemplate')  result = apiSaveEmailTemplate(e.parameter);
    else if (action === 'confirmDeposit')      result = apiConfirmDeposit(e.parameter);
    else if (action === 'getDepositTemplate')  result = apiGetDepositTemplate(e.parameter.token);
    else if (action === 'saveDepositTemplate')  result = apiSaveDepositTemplate(e.parameter);
    else if (action === 'getCancelTemplate')    result = apiGetCancelTemplate(e.parameter.token);
    else if (action === 'saveCancelTemplate')    result = apiSaveCancelTemplate(e.parameter);
    else if (action === 'getRescheduleTemplate') result = apiGetRescheduleTemplate(e.parameter.token);
    else if (action === 'saveRescheduleTemplate') result = apiSaveRescheduleTemplate(e.parameter);
    else if (action === 'updateBooking')          result = apiUpdateBooking(e.parameter);
    else result = {success:false, message:'未知操作'};
  } catch(err) {
    result = {success:false, message:err.message};
  }
  var json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback+'('+json+')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Auth ──────────────────────────────────────────────────────────
function getPassword() {
  var sheet = SS().getSheetByName(SHEET_CONFIG);
  if (!sheet) return 'admin1234';
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === 'admin_password') return String(data[i][1]);
  }
  return 'admin1234';
}

// 從 config 讀取允許的 email 清單（allowed_emails 列，逗號分隔）
function getAllowedEmails() {
  var sheet = SS().getSheetByName(SHEET_CONFIG);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === 'allowed_emails') {
      return String(data[i][1]).split(',').map(function(e){ return e.trim().toLowerCase(); }).filter(Boolean);
    }
  }
  return [];
}

function makeToken(identifier) {
  // identifier 可以是密碼或 email
  return Utilities.base64Encode(identifier + '_muse_' + new Date().toDateString());
}

function validToken(token) {
  if (!token) return false;
  // 檢查密碼 token
  if (token === makeToken(getPassword())) return true;
  // 檢查 Google email token
  var emails = getAllowedEmails();
  for (var i = 0; i < emails.length; i++) {
    if (token === makeToken(emails[i])) return true;
  }
  return false;
}

// ── Login（密碼）─────────────────────────────────────────────────
function apiLogin(p) {
  if (p.password !== getPassword()) return {success:false, message:'密碼錯誤'};
  return {success:true, token:makeToken(getPassword())};
}

// ── Login（Google）────────────────────────────────────────────────
// 驗證 Google ID token，取出 email，比對白名單
function apiLoginGoogle(p) {
  var credential = p.credential;
  if (!credential) return {success:false, message:'缺少 Google credential'};
  try {
    // Google ID token 是 JWT，直接 base64 decode payload（不需要驗簽，GAS 是 Google 環境）
    var parts = credential.split('.');
    if (parts.length !== 3) return {success:false, message:'無效 token'};
    // base64url → base64
    var payload = parts[1].replace(/-/g,'+').replace(/_/g,'/');
    while (payload.length % 4 !== 0) payload += '=';
    var decoded = JSON.parse(Utilities.newBlob(Utilities.base64Decode(payload)).getDataAsString());
    var email = String(decoded.email || '').toLowerCase();
    var verified = decoded.email_verified;
    if (!email || !verified) return {success:false, message:'Email 未驗證'};
    // 檢查到期時間（exp 是 Unix timestamp，秒）
    var now = Math.floor(Date.now() / 1000);
    if (decoded.exp && decoded.exp < now) return {success:false, message:'Google token 已過期，請重新登入'};
    // 比對白名單
    var allowed = getAllowedEmails();
    if (allowed.indexOf(email) < 0) return {success:false, message:'此帳號無後台權限（'+email+'，白名單：'+allowed.join('|')+'）'};
    // 回傳每日 token
    return {success:true, token:makeToken(email), email:email};
  } catch(err) {
    return {success:false, message:'驗證失敗：'+err.message};
  }
}

// ── Get booked（前台）────────────────────────────────────────────
function apiGetBooked() {
  var ss      = SS();
  var bookSh  = ss.getSheetByName(SHEET_BOOKINGS);
  var blockSh = ss.getSheetByName(SHEET_BLOCKED);
  var booked  = {};
  var bData   = bookSh.getDataRange().getValues();
  for (var i = 1; i < bData.length; i++) {
    var row = bData[i];
    if (!row[4]) continue; // 日期在 E 欄（index 4）
    var ds    = fmtKey(row[4]);
    var slots = String(row[5]).split(',').map(function(s){ return s.trim(); }).filter(Boolean);
    if (!booked[ds]) booked[ds] = [];
    slots.forEach(function(s){ if (booked[ds].indexOf(s)<0) booked[ds].push(s); });
  }
  var blData = blockSh.getDataRange().getValues();
  for (var i = 1; i < blData.length; i++) {
    var row=blData[i], type=String(row[0]), ds=fmtKey(row[1]);
    if (!ds) continue;
    if (!booked[ds]) booked[ds]=[];
    if (type==='date') {
      ['A','B','C','D'].forEach(function(s){ if(booked[ds].indexOf(s)<0) booked[ds].push(s); });
    } else if (type==='slot') {
      var slot=String(row[2]).trim();
      if (booked[ds].indexOf(slot)<0) booked[ds].push(slot);
    }
  }
  return {booked:booked};
}

// ── Book ──────────────────────────────────────────────────────────
function apiBook(p) {
  var name=p.name, phone=p.phone, email=p.email||'', date=p.date, slots=p.slots, slotLabels=p.slotLabels||slots, outfits=p.outfits||'';
  if (!name||!phone||!date||!slots) return {success:false, message:'資料不完整'};
  var ss      = SS();
  var bookSh  = ss.getSheetByName(SHEET_BOOKINGS);
  var blockSh = ss.getSheetByName(SHEET_BLOCKED);
  var slotList = slots.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
  // 檢查整天封鎖
  var blData = blockSh.getDataRange().getValues();
  for (var i=1;i<blData.length;i++) {
    if (String(blData[i][0])==='date' && fmtKey(blData[i][1])===date)
      return {success:false, message:'該日期已關閉預約'};
  }
  // 檢查衝突
  var taken    = getTakenSlots(bookSh, blockSh, date);
  var conflict = slotList.filter(function(s){ return taken.indexOf(s)>-1; });
  if (conflict.length>0) return {success:false, message:'部分時段已被預約，請重新選擇'};
  // 寫入（欄位：時間戳記, 姓名, 電話, Email, 預約日期, 時段, 套數）
  var addons=p.addons||'', note=p.note||'';
  bookSh.appendRow([new Date(), name, phone, email, date, slots, outfits, addons, '', note]);
  // 寄信給管理員
  try { sendAdminMail({name:name, phone:phone, email:email, date:date, slotLabels:slotLabels, outfits:outfits, addons:addons, note:note}); }
  catch(e) { Logger.log('Admin mail error:'+e.message); }
  // 寄確認信給顧客
  if (email) {
    try { sendCustomerMail({name:name, email:email, date:date, slotLabels:slotLabels, outfits:outfits, addons:addons}); }
    catch(e) { Logger.log('Customer mail error:'+e.message); }
  }
  return {success:true};
}

function getTakenSlots(bookSh, blockSh, date) {
  var taken=[];
  var bData=bookSh.getDataRange().getValues();
  for (var i=1;i<bData.length;i++) {
    if (fmtKey(bData[i][4])===date) {
      String(bData[i][5]).split(',').map(function(s){ return s.trim(); }).filter(Boolean)
        .forEach(function(s){ if(taken.indexOf(s)<0) taken.push(s); });
    }
  }
  var blData=blockSh.getDataRange().getValues();
  for (var i=1;i<blData.length;i++) {
    if (fmtKey(blData[i][1])===date && String(blData[i][0])==='slot') {
      var s=String(blData[i][2]).trim();
      if (taken.indexOf(s)<0) taken.push(s);
    }
  }
  return taken;
}

// ── Admin APIs ────────────────────────────────────────────────────
function apiGetBookings(token) {
  if (!validToken(token)) return {success:false, message:'未授權'};
  var sheet=SS().getSheetByName(SHEET_BOOKINGS);
  var data=sheet.getDataRange().getValues();
  var list=[];
  for (var i=1;i<data.length;i++) {
    var row=data[i];
    if (!row[1]&&!row[4]) continue;
    list.push({
      id:        i,
      timestamp: row[0] ? new Date(row[0]).toLocaleString('zh-TW',{timeZone:'Asia/Taipei'}) : '',
      name:      row[1],
      phone:     row[2],
      email:     row[3],
      date:      fmtKey(row[4]),
      slots:     row[5],
      outfits:   row[6],
      addons:    row[7]||'',
      deposit:   row[8]||'',
      note:      row[9]||''
    });
  }
  return {success:true, bookings:list};
}

function apiGetBlocked(token) {
  if (!validToken(token)) return {success:false, message:'未授權'};
  var sheet=SS().getSheetByName(SHEET_BLOCKED);
  var data=sheet.getDataRange().getValues();
  var list=[];
  for (var i=1;i<data.length;i++) {
    if (!data[i][1]) continue;
    list.push({id:i, type:data[i][0], date:fmtKey(data[i][1]), slot:data[i][2]||''});
  }
  return {success:true, blocked:list};
}

function apiDeleteBooking(token, id) {
  if (!validToken(token)) return {success:false, message:'未授權'};
  var sheet=SS().getSheetByName(SHEET_BOOKINGS);
  sheet.deleteRow(parseInt(id)+1);
  return {success:true};
}

function apiBlockDate(p) {
  if (!validToken(p.token)) return {success:false, message:'未授權'};
  var sheet=SS().getSheetByName(SHEET_BLOCKED);
  sheet.appendRow(['date', p.date, '']);
  return {success:true};
}

function apiBlockSlot(p) {
  if (!validToken(p.token)) return {success:false, message:'未授權'};
  var sheet=SS().getSheetByName(SHEET_BLOCKED);
  sheet.appendRow(['slot', p.date, p.slot]);
  return {success:true};
}

function apiUnblock(p) {
  if (!validToken(p.token)) return {success:false, message:'未授權'};
  var sheet=SS().getSheetByName(SHEET_BLOCKED);
  var data=sheet.getDataRange().getValues();
  for (var i=data.length-1;i>=1;i--) {
    var row=data[i];
    if (String(row[0])===p.type && fmtKey(row[1])===p.date) {
      if (p.type==='date' || String(row[2]).trim()===p.slot) {
        sheet.deleteRow(i+1); return {success:true};
      }
    }
  }
  return {success:false, message:'找不到該項目'};
}

// ── Email Template ────────────────────────────────────────────────
function apiGetEmailTemplate(token) {
  if (!validToken(token)) return {success:false, message:'未授權'};
  var sheet=SS().getSheetByName(SHEET_EMAIL_TPL);
  if (!sheet) return {success:true, subject:'', payUrl:'', body:''};
  var data=sheet.getDataRange().getValues();
  var tpl={subject:'', payUrl:'', body:''};
  for (var i=0;i<data.length;i++) {
    var key=String(data[i][0]), val=String(data[i][1]);
    if (key==='subject') tpl.subject=val;
    else if (key==='payUrl') tpl.payUrl=val;
    else if (key==='body') tpl.body=val;
  }
  return {success:true, subject:tpl.subject, payUrl:tpl.payUrl, body:tpl.body};
}

function apiSaveEmailTemplate(p) {
  if (!validToken(p.token)) return {success:false, message:'未授權'};
  var ss=SS();
  var sheet=ss.getSheetByName(SHEET_EMAIL_TPL);
  if (!sheet) { sheet=ss.insertSheet(SHEET_EMAIL_TPL); }
  sheet.clearContents();
  sheet.appendRow(['subject', p.subject||'']);
  sheet.appendRow(['payUrl',  p.payUrl||'']);
  sheet.appendRow(['body',    p.body||'']);
  return {success:true};
}

function getEmailTemplate() {
  var sheet=SS().getSheetByName(SHEET_EMAIL_TPL);
  if (!sheet) return {subject:'【繆思寫真】預約確認通知', payUrl:'', body:''};
  var data=sheet.getDataRange().getValues();
  var tpl={subject:'【繆思寫真】預約確認通知', payUrl:'', body:''};
  for (var i=0;i<data.length;i++) {
    var key=String(data[i][0]), val=String(data[i][1]);
    if (key==='subject') tpl.subject=val;
    else if (key==='payUrl') tpl.payUrl=val;
    else if (key==='body') tpl.body=val;
  }
  return tpl;
}

// ── 寄信給顧客 ────────────────────────────────────────────────────
function sendCustomerMail(p) {
  var tpl = getEmailTemplate();
  var subject = tpl.subject || '【Muse Lab Studio】您的預約確認通知';
  var body = tpl.body;
  var payUrl = tpl.payUrl || '';

  // 預設信件內容
  if (!body) {
    body = '親愛的 {{姓名}} 您好，\n\n感謝您預約 Muse Lab Studio 繆思寫真！\n以下是您的預約資訊：\n\n📅 拍攝日期：{{日期}}\n⏰ 拍攝時段：{{時段}}\n👗 拍攝套數：{{套數}} 套\n\n請在拍攝前完成付款，付款連結如下：\n{{付款連結}}\n\n如有任何問題，歡迎隨時與我們聯繫。\n期待與您相見！\n\nMuse Lab Studio 繆思寫真 敬上';
  }

  // 替換變數
  body = body.replace(/\{\{姓名\}\}/g, p.name)
             .replace(/\{\{日期\}\}/g, p.date)
             .replace(/\{\{時段\}\}/g, p.slotLabels)
             .replace(/\{\{套數\}\}/g, p.outfits)
             .replace(/\{\{加購\}\}/g, p.addons||'無')
             .replace(/\{\{付款連結\}\}/g, payUrl);
  subject = subject.replace(/\{\{姓名\}\}/g, p.name)
                   .replace(/\{\{日期\}\}/g, p.date);

  // HTML 版本
  var htmlBody = '<div style="font-family:sans-serif;max-width:520px;margin:0 auto;">'
    + '<div style="background:#C49ACC;padding:20px 24px;border-radius:10px 10px 0 0;">'
    + '<h2 style="color:#fff;margin:0;font-size:18px;">✦ 預約確認通知</h2>'
    + '<p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px;">Muse Lab Studio 繆思寫真</p>'
    + '</div>'
    + '<div style="border:1.5px solid #DFC8E8;border-top:none;border-radius:0 0 10px 10px;padding:24px;">'
    + '<p style="font-size:14px;line-height:1.8;white-space:pre-wrap;">'+body+'</p>'
    + '<div style="text-align:center;margin-top:20px;"><p style="font-size:13px;color:#7A5E88;margin-bottom:10px;">請掃描以下 QR Code 完成 LINE Pay 付款</p><img src="https://muselab-studio.github.io/muse-booking/linepay-qrcode.png" style="width:200px;height:200px;border-radius:12px;border:1.5px solid #DFC8E8;" alt="LINE Pay QR Code"></div>'
    + (payUrl ? '<div style="text-align:center;margin-top:16px;"><a href="'+payUrl+'" style="display:inline-block;padding:13px 32px;background:#C49ACC;color:#fff;border-radius:40px;font-weight:700;font-size:15px;text-decoration:none;">前往付款 →</a></div>' : '')
    + '<p style="font-size:11px;color:#bbb;margin-top:20px;border-top:1px solid #eee;padding-top:12px;">此信由系統自動發送，請勿直接回覆。</p>'
    + '</div></div>';

  GmailApp.sendEmail(p.email, subject, body, {htmlBody:htmlBody, from:NOTIFY_EMAIL, name:SENDER_ALIAS});
}

// ── 寄信給管理員 ──────────────────────────────────────────────────
function sendAdminMail(p) {
  var subject='【新預約】'+p.name+' · '+p.date+' · '+p.outfits+' 套';
  var html='<div style="font-family:sans-serif;max-width:500px;margin:0 auto;">'
    +'<div style="background:#C49ACC;padding:18px 24px;border-radius:10px 10px 0 0;">'
    +'<h2 style="color:#fff;margin:0;font-size:18px;">✦ 新預約通知</h2>'
    +'<p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px;">Muse Lab Studio 繆思寫真</p>'
    +'</div><div style="border:1.5px solid #DFC8E8;border-top:none;border-radius:0 0 10px 10px;padding:20px 24px;">'
    +'<table style="width:100%;font-size:14px;border-collapse:collapse;">'
    +'<tr><td style="padding:7px 0;color:#7A5E88;width:80px;">姓名</td><td style="font-weight:700;">'+p.name+'</td></tr>'
    +'<tr><td style="padding:7px 0;color:#7A5E88;">電話</td><td>'+p.phone+'</td></tr>'
    +'<tr><td style="padding:7px 0;color:#7A5E88;">Email</td><td>'+p.email+'</td></tr>'
    +'<tr><td style="padding:7px 0;color:#7A5E88;">日期</td><td style="color:#8B5E99;font-weight:700;">'+p.date+'</td></tr>'
    +'<tr><td style="padding:7px 0;color:#7A5E88;">時段</td><td><strong>'+p.slotLabels+'</strong></td></tr>'
    +'<tr><td style="padding:7px 0;color:#7A5E88;">套數</td><td><strong>'+p.outfits+' 套</strong></td></tr>'
    +'<tr><td style="padding:7px 0;color:#7A5E88;">加購</td><td>'+(p.addons||'無')+'</td></tr>'
    +'<tr><td style="padding:7px 0;color:#7A5E88;">備註</td><td>'+(p.note||'—')+'</td></tr>'
    +'</table>'
    +'<p style="font-size:11px;color:#bbb;margin-top:14px;border-top:1px solid #eee;padding-top:12px;">系統自動發送，請勿回覆</p>'
    +'</div></div>';
  GmailApp.sendEmail(NOTIFY_EMAIL, subject, '', {htmlBody:html, name:SENDER_ALIAS});
}

// ── Confirm Deposit ───────────────────────────────────────────────
function apiConfirmDeposit(p) {
  if (!validToken(p.token)) return {success:false, message:'未授權'};
  var id = parseInt(p.id);
  var sheet = SS().getSheetByName(SHEET_BOOKINGS);
  var data  = sheet.getDataRange().getValues();
  var row   = data[id]; // 0-indexed header + rows
  if (!row) return {success:false, message:'找不到該筆預約'};
  // Check already confirmed
  if (String(row[8]) === 'Y') return {success:false, message:'已確認付訂，不可重複操作'};
  // Write Y to column I (index 8)
  sheet.getRange(id + 1, 9).setValue('Y');
  // Send deposit confirmation email to customer
  var email = String(row[3]);
  if (email && email.indexOf('@') > -1) {
    try {
      sendDepositMail({
        name:       String(row[1]),
        email:      email,
        date:       fmtKey(row[4]),
        slotLabels: String(row[5]).split(',').map(function(s){ var m={'A':'10:00','B':'13:00','C':'16:00','D':'19:00'}; return m[s.trim()]||s.trim(); }).join('、'),
        outfits:    String(row[6]),
        addons:     String(row[7]||'')
      });
    } catch(e) { Logger.log('Deposit mail error:'+e.message); }
  }
  return {success:true};
}

// ── Deposit Template APIs ──────────────────────────────────────────
function apiGetDepositTemplate(token) {
  if (!validToken(token)) return {success:false, message:'未授權'};
  var sheet = SS().getSheetByName(SHEET_DEPOSIT_TPL);
  if (!sheet) return {success:true, subject:'', body:''};
  var data = sheet.getDataRange().getValues();
  var tpl  = {subject:'', body:''};
  for (var i=0;i<data.length;i++) {
    if (data[i][0]==='subject') tpl.subject = String(data[i][1]);
    if (data[i][0]==='body')    tpl.body    = String(data[i][1]);
  }
  return {success:true, subject:tpl.subject, body:tpl.body};
}

function apiSaveDepositTemplate(p) {
  if (!validToken(p.token)) return {success:false, message:'未授權'};
  var ss    = SS();
  var sheet = ss.getSheetByName(SHEET_DEPOSIT_TPL);
  if (!sheet) sheet = ss.insertSheet(SHEET_DEPOSIT_TPL);
  sheet.clearContents();
  sheet.appendRow(['subject', p.subject||'']);
  sheet.appendRow(['body',    p.body||'']);
  return {success:true};
}

function getDepositTemplate() {
  var sheet = SS().getSheetByName(SHEET_DEPOSIT_TPL);
  if (!sheet) return {subject:'【Muse Lab Studio】訂金確認通知', body:''};
  var data = sheet.getDataRange().getValues();
  var tpl  = {subject:'【Muse Lab Studio】訂金確認通知', body:''};
  for (var i=0;i<data.length;i++) {
    if (data[i][0]==='subject') tpl.subject = String(data[i][1]);
    if (data[i][0]==='body')    tpl.body    = String(data[i][1]);
  }
  return tpl;
}

function sendDepositMail(p) {
  var tpl    = getDepositTemplate();
  var payUrl = '';
  var subject = tpl.subject || '【Muse Lab Studio】訂金已確認';
  var body    = tpl.body;
  if (!body) {
    body = '親愛的 {{姓名}} 您好，\n\n我們已收到您的訂金，預約正式確認！\n以下是您的預約資訊：\n\n📅 拍攝日期：{{日期}}\n⏰ 拍攝時段：{{時段}}\n👗 拍攝套數：{{套數}} 套\n\n期待與您相見，如有任何問題歡迎聯繫我們。\n\nMuse Lab Studio 繆思寫真 敬上';
  }
  body = replaceVars(body, p.name, p.date, p.slotLabels, p.outfits, p.addons||'無', payUrl);
  subject = subject.split('{{姓名}}').join(p.name).split('{{日期}}').join(p.date);
  var htmlBody = '<div style="font-family:sans-serif;max-width:520px;margin:0 auto;">'
    + '<div style="background:#6BBF8E;padding:20px 24px;border-radius:10px 10px 0 0;">'
    + '<h2 style="color:#fff;margin:0;font-size:18px;">✅ 訂金確認通知</h2>'
    + '<p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:13px;">Muse Lab Studio 繆思寫真</p>'
    + '</div>'
    + '<div style="border:1.5px solid #DFC8E8;border-top:none;border-radius:0 0 10px 10px;padding:24px;">'
    + '<p style="font-size:14px;line-height:1.8;white-space:pre-wrap;">'+body+'</p>'
    + '<p style="font-size:11px;color:#bbb;margin-top:20px;border-top:1px solid #eee;padding-top:12px;">此信由系統自動發送，請勿直接回覆。</p>'
    + '</div></div>';
  GmailApp.sendEmail(p.email, subject, body, {htmlBody:htmlBody, from:NOTIFY_EMAIL, name:SENDER_ALIAS});
}

// ── Update Booking ────────────────────────────────────────────────
function apiUpdateBooking(p) {
  if (!validToken(p.token)) return {success:false, message:'未授權'};
  var id     = parseInt(p.id);
  var sheet  = SS().getSheetByName(SHEET_BOOKINGS);
  var data   = sheet.getDataRange().getValues();
  var row    = data[id];
  if (!row) return {success:false, message:'找不到該筆預約'};

  var name       = p.name       || String(row[1]);
  var phone      = p.phone      || String(row[2]);
  var email      = p.email      || String(row[3]);
  var date       = p.date       || fmtKey(row[4]);
  var slots      = p.slots      || String(row[5]);
  var outfits    = p.outfits    || String(row[6]);
  var addons     = p.addons     !== undefined ? p.addons : String(row[7]||'');
  var deposit    = String(row[8]||'');
  var note       = p.note       !== undefined ? p.note   : String(row[9]||'');

  var slotLabels = slots.split(',').map(function(s){
    var m={'A':'10:00','B':'13:00','C':'16:00','D':'19:00'};
    return m[s.trim()]||s.trim();
  }).join('、');

  // Write updated row
  var rowNum = id + 1;
  sheet.getRange(rowNum, 2).setValue(name);
  sheet.getRange(rowNum, 3).setValue(phone);
  sheet.getRange(rowNum, 4).setValue(email);
  sheet.getRange(rowNum, 5).setValue(date);
  sheet.getRange(rowNum, 6).setValue(slots);
  sheet.getRange(rowNum, 7).setValue(outfits);
  sheet.getRange(rowNum, 8).setValue(addons);
  sheet.getRange(rowNum, 10).setValue(note);

  // Send reschedule notification if email exists
  if (email && email.indexOf('@') > -1) {
    try {
      sendRescheduleMail({
        name: name, email: email, date: date,
        slotLabels: slotLabels, outfits: outfits, addons: addons
      });
    } catch(e) { Logger.log('Reschedule mail error:' + e.message); }
  }
  return {success:true};
}

// ── Reschedule Template APIs ───────────────────────────────────────
function apiGetRescheduleTemplate(token) {
  if (!validToken(token)) return {success:false, message:'未授權'};
  var sheet = SS().getSheetByName(SHEET_RESCHEDULE_TPL);
  if (!sheet) return {success:true, subject:'', body:''};
  var data = sheet.getDataRange().getValues();
  var tpl  = {subject:'', body:''};
  for (var i=0;i<data.length;i++) {
    if (data[i][0]==='subject') tpl.subject = String(data[i][1]);
    if (data[i][0]==='body')    tpl.body    = String(data[i][1]);
  }
  return {success:true, subject:tpl.subject, body:tpl.body};
}

function apiSaveRescheduleTemplate(p) {
  if (!validToken(p.token)) return {success:false, message:'未授權'};
  var ss    = SS();
  var sheet = ss.getSheetByName(SHEET_RESCHEDULE_TPL);
  if (!sheet) sheet = ss.insertSheet(SHEET_RESCHEDULE_TPL);
  sheet.clearContents();
  sheet.appendRow(['subject', p.subject||'']);
  sheet.appendRow(['body',    p.body||'']);
  return {success:true};
}

function getRescheduleTemplate() {
  var sheet = SS().getSheetByName(SHEET_RESCHEDULE_TPL);
  if (!sheet) return {subject:'【Muse Lab Studio】預約更改通知', body:''};
  var data = sheet.getDataRange().getValues();
  var tpl  = {subject:'【Muse Lab Studio】預約更改通知', body:''};
  for (var i=0;i<data.length;i++) {
    if (data[i][0]==='subject') tpl.subject = String(data[i][1]);
    if (data[i][0]==='body')    tpl.body    = String(data[i][1]);
  }
  return tpl;
}

function sendRescheduleMail(p) {
  var tpl     = getRescheduleTemplate();
  var subject = tpl.subject || '【Muse Lab Studio】預約更改通知';
  var body    = tpl.body;
  if (!body) {
    body = '親愛的 {{姓名}} 您好，\n\n您的預約資訊已更新，以下為最新預約內容：\n\n預約日期：{{日期}}\n預約時段：{{時段}}\n拍攝套數：{{套數}} 套\n\n如有任何疑問，請聯繫我們。\n\nMuse Lab Studio 繆思寫真 敬上';
  }
  body = body
    .split('{{姓名}}').join(p.name||'')
    .split('{{日期}}').join(p.date||'')
    .split('{{時段}}').join(p.slotLabels||'')
    .split('{{套數}}').join(p.outfits||'')
    .split('{{加購}}').join(p.addons||'無');
  subject = subject.split('{{姓名}}').join(p.name||'').split('{{日期}}').join(p.date||'');
  var htmlBody = '<div style="font-family:sans-serif;max-width:520px;margin:0 auto;">'
    + '<div style="background:#C49ACC;padding:20px 24px;border-radius:10px 10px 0 0;">'
    + '<h2 style="color:#fff;margin:0;font-size:18px;">預約更改通知</h2>'
    + '<p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:13px;">Muse Lab Studio 繆思寫真</p>'
    + '</div>'
    + '<div style="border:1.5px solid #DFC8E8;border-top:none;border-radius:0 0 10px 10px;padding:24px;">'
    + '<p style="font-size:14px;line-height:1.8;white-space:pre-wrap;">'+body+'</p>'
    + '<p style="font-size:11px;color:#bbb;margin-top:20px;border-top:1px solid #eee;padding-top:12px;">此信由系統自動發送，請勿直接回覆。</p>'
    + '</div></div>';
  GmailApp.sendEmail(p.email, subject, body, {htmlBody:htmlBody, from:NOTIFY_EMAIL, name:SENDER_ALIAS});
}

// ── Cancel Template ───────────────────────────────────────────────
function apiGetCancelTemplate(token) {
  if (!validToken(token)) return {success:false, message:'未授權'};
  var sheet = SS().getSheetByName(SHEET_CANCEL_TPL);
  if (!sheet) return {success:true, subject:'', body:''};
  var data = sheet.getDataRange().getValues();
  var tpl  = {subject:'', body:''};
  for (var i=0;i<data.length;i++) {
    if (data[i][0]==='subject') tpl.subject = String(data[i][1]);
    if (data[i][0]==='body')    tpl.body    = String(data[i][1]);
  }
  return {success:true, subject:tpl.subject, body:tpl.body};
}

function apiSaveCancelTemplate(p) {
  if (!validToken(p.token)) return {success:false, message:'未授權'};
  var ss    = SS();
  var sheet = ss.getSheetByName(SHEET_CANCEL_TPL);
  if (!sheet) sheet = ss.insertSheet(SHEET_CANCEL_TPL);
  sheet.clearContents();
  sheet.appendRow(['subject', p.subject||'']);
  sheet.appendRow(['body',    p.body||'']);
  return {success:true};
}

function getCancelTemplate() {
  var sheet = SS().getSheetByName(SHEET_CANCEL_TPL);
  if (!sheet) return {subject:'【Muse Lab Studio】預約取消通知', body:''};
  var data = sheet.getDataRange().getValues();
  var tpl  = {subject:'【Muse Lab Studio】預約取消通知', body:''};
  for (var i=0;i<data.length;i++) {
    if (data[i][0]==='subject') tpl.subject = String(data[i][1]);
    if (data[i][0]==='body')    tpl.body    = String(data[i][1]);
  }
  return tpl;
}

function sendCancelMail(p) {
  var tpl     = getCancelTemplate();
  var subject = tpl.subject || '【Muse Lab Studio】預約取消通知';
  var body    = tpl.body;
  if (!body) {
    body = '親愛的 {{姓名}} 您好，\n\n很遺憾通知您，您的預約已被取消。\n\n預約日期：{{日期}}\n預約時段：{{時段}}\n拍攝套數：{{套數}} 套\n\n如有任何疑問，請聯繫我們。\n\nMuse Lab Studio 繆思寫真 敬上';
  }
  body = body
    .split('{{姓名}}').join(p.name||'')
    .split('{{日期}}').join(p.date||'')
    .split('{{時段}}').join(p.slotLabels||'')
    .split('{{套數}}').join(p.outfits||'')
    .split('{{加購}}').join(p.addons||'無');
  subject = subject.split('{{姓名}}').join(p.name||'').split('{{日期}}').join(p.date||'');
  var htmlBody = '<div style="font-family:sans-serif;max-width:520px;margin:0 auto;">'
    + '<div style="background:#E87878;padding:20px 24px;border-radius:10px 10px 0 0;">'
    + '<h2 style="color:#fff;margin:0;font-size:18px;">預約取消通知</h2>'
    + '<p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:13px;">Muse Lab Studio 繆思寫真</p>'
    + '</div>'
    + '<div style="border:1.5px solid #DFC8E8;border-top:none;border-radius:0 0 10px 10px;padding:24px;">'
    + '<p style="font-size:14px;line-height:1.8;white-space:pre-wrap;">'+body+'</p>'
    + '<p style="font-size:11px;color:#bbb;margin-top:20px;border-top:1px solid #eee;padding-top:12px;">此信由系統自動發送，請勿直接回覆。</p>'
    + '</div></div>';
  GmailApp.sendEmail(p.email, subject, body, {htmlBody:htmlBody, from:NOTIFY_EMAIL, name:SENDER_ALIAS});
}

// ── Helpers ───────────────────────────────────────────────────────
function replaceVars(body, name, date, slots, outfits, addons, payUrl) {
  return body
    .split('{{姓名}}').join(name||'')
    .split('{{日期}}').join(date||'')
    .split('{{時段}}').join(slots||'')
    .split('{{套數}}').join(outfits||'')
    .split('{{加購}}').join(addons||'無')
    .split('{{付款連結}}').join(payUrl||'');
}
function fmtKey(val) {
  if (!val) return '';
  if (typeof val==='string') return val.substring(0,10);
  var d=new Date(val);
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
