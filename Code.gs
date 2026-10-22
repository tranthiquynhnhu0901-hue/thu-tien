/**
 * TRO VIET MANAGER - GUI EMAIL PDF QUA GOOGLE APPS SCRIPT (MIEN PHI)
 * Quan trong: Chi dan doan Code.gs nay vao https://script.google.com/ (khong dang len GitHub).
 * Web app su dung quyen Gmail cua chu tai khoan, bao ve endpoint bang token 64 ky tu.
 */

function setup() {
  // Run once inside Apps Script editor as the Gmail account that will SEND messages.
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TRO_VIET_TOKEN', token);
  Logger.log('APPS_SCRIPT_TOKEN = "' + token + '"');
  Logger.log('So nguoi nhan email con lai hom nay: ' + MailApp.getRemainingDailyQuota());
  Logger.log('Hay sao chep TOKEN vao Streamlit Secrets. Khong gui TOKEN cho ai.');
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return reply_({ok:true, service:'Tro Viet MailApp', message:'API dang hoat dong. Chi nhan POST co token.'});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const expected = PropertiesService.getScriptProperties().getProperty('TRO_VIET_TOKEN');
    if (!expected || !data.token || data.token !== expected) {
      return reply_({ok:false,error:'Sai khoa bao mat. Kiem tra APPS_SCRIPT_TOKEN.'});
    }
    const to = String(data.to || '').trim();
    if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(to)) {
      return reply_({ok:false,error:'Email nguoi nhan khong hop le.'});
    }
    const subject = String(data.subject || '').slice(0, 200);
    const body = String(data.body || '').slice(0, 15000);
    const id = String(data.request_id || '');
    if (!subject || !body || !/^[a-zA-Z0-9_-]{10,100}$/.test(id)) {
      return reply_({ok:false,error:'Thieu tieu de, noi dung hoac ma yeu cau.'});
    }
    const encoded = String(data.pdf_base64 || '');
    if (encoded.length > 10000000) {
      return reply_({ok:false,error:'File PDF vuot gioi han 7MB.'});
    }

    // Serialise duplicate checks and sends; avoid resending when network response gets lost.
    lock.waitLock(30000);
    locked = true;
    const props = PropertiesService.getScriptProperties();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id)
      .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
    const sentKey = 'sent_' + digest;
    if (props.getProperty(sentKey)) {
      return reply_({ok:true,duplicate:true,message:'Thu nay da gui truoc do. Khong gui lai.'});
    }
    if (MailApp.getRemainingDailyQuota() < 1) {
      return reply_({ok:false,error:'Da het han muc email hom nay. Vui long gui vao ngay mai.'});
    }

    const message = {to:to, subject:subject, body:body, name:'Trọ Việt Manager'};
    if (encoded) {
      const bytes = Utilities.base64Decode(encoded);
      // %PDF- file signature.
      if (bytes.length < 5 || bytes[0] !== 37 || bytes[1] !== 80 ||
          bytes[2] !== 68 || bytes[3] !== 70 || bytes[4] !== 45) {
        return reply_({ok:false,error:'PDF dinh kem khong hop le.'});
      }
      const filename = String(data.file_name || 'Thong_bao_tien_phong.pdf')
        .replace(/[\\/\r\n]/g, '_').slice(0, 100);
      message.attachments = [Utilities.newBlob(bytes, 'application/pdf', filename)];
    }
    MailApp.sendEmail(message);
    props.setProperty(sentKey, new Date().toISOString());
    return reply_({ok:true,duplicate:false,message:'Google da tiep nhan yeu cau gui email.'});
  } catch (err) {
    return reply_({ok:false,error: String(err && err.message ? err.message : err).slice(0,240)});
  } finally {
    if (locked) lock.releaseLock();
  }
}
