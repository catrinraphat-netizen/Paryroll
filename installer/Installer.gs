/**
 * ตัวติดตั้งระบบเงินเดือน — ดึงโค้ดจาก GitHub มาแทนที่โปรเจกต์นี้ แล้ว deploy เป็นเว็บแอป
 * ใช้ครั้งเดียว: เลือกฟังก์ชัน install → Run → อนุญาตสิทธิ์
 * ต้องเปิด Google Apps Script API ก่อนที่ https://script.google.com/home/usersettings
 */
var SRC = 'https://raw.githubusercontent.com/catrinraphat-netizen/Paryroll/claude/keen-cerf-xkjf4s/';
var FILES = [
  ['appsscript', 'JSON', 'appsscript.json'],
  ['Code', 'SERVER_JS', 'Code.gs'],
  ['TimesheetOcr', 'SERVER_JS', 'TimesheetOcr.gs'],
  ['Index', 'HTML', 'Index.html'],
];

function install() {
  var files = FILES.map(function (f) {
    var r = UrlFetchApp.fetch(SRC + f[2], { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('ดึงไฟล์ ' + f[2] + ' ไม่ได้ (HTTP ' + r.getResponseCode() + ')');
    return { name: f[0], type: f[1], source: r.getContentText('UTF-8') };
  });
  var id = ScriptApp.getScriptId();
  api_('put', 'projects/' + id + '/content', { files: files });
  var ver = api_('post', 'projects/' + id + '/versions', { description: 'ติดตั้งครั้งแรก' });
  var dep = api_('post', 'projects/' + id + '/deployments', {
    versionNumber: ver.versionNumber, manifestFileName: 'appsscript', description: 'ติดตั้งครั้งแรก',
  });
  var url = ((dep.entryPoints || []).filter(function (e) { return e.webApp; })[0] || {}).webApp;
  Logger.log('ติดตั้งเสร็จ · deploymentId = ' + dep.deploymentId);
  Logger.log('ลิงก์เว็บ = ' + (url ? url.url : '(ไม่พบ)'));
  Logger.log('ขั้นต่อไป: รีเฟรชหน้านี้ → เลือกฟังก์ชัน setup → Run → อนุญาตสิทธิ์');
}

function api_(method, path, body) {
  var r = UrlFetchApp.fetch('https://script.googleapis.com/v1/' + path, {
    method: method, contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify(body),
  });
  var code = r.getResponseCode(), text = r.getContentText();
  if (code === 403 && /has not been used|is disabled|SERVICE_DISABLED|User has not enabled/i.test(text)) {
    throw new Error('ยังไม่ได้เปิด Google Apps Script API → เปิดที่ https://script.google.com/home/usersettings แล้ว Run install ใหม่\nคำตอบจาก Google: ' + text.slice(0, 800));
  }
  if (code >= 300) throw new Error('Apps Script API ' + path + ' ตอบ ' + code + ': ' + text.slice(0, 500));
  return JSON.parse(text);
}
