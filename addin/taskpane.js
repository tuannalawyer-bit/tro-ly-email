/* Task pane Trợ lý Email — lớp bề mặt Outlook.
 *
 * Trang này được nạp TỪ https://localhost:8765 nên fetch tới backend là
 * same-origin. Token lấy từ thẻ <meta> do backend chèn vào lúc phục vụ tệp.
 */
'use strict';

const BACKEND = 'https://localhost:8765';
const TOKEN = (document.querySelector('meta[name="addin-token"]') || {}).content || '';

const $ = id => document.getElementById(id);

/* ---------------------------------------------------------------- đính kèm */

// MIME suy từ phần mở rộng, KHÔNG dùng a.contentType: Exchange rất hay khai
// application/octet-stream và Gemini từ chối MIME đó.
const SUPPORTED = {
  pdf: 'application/pdf',
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
  gif: 'image/gif', webp: 'image/webp',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};
const MAX_ONE = 8 * 1024 * 1024;
const MAX_TOTAL = 12 * 1024 * 1024;
const MAX_COUNT = 5;
const MIN_IMAGE = 8 * 1024;          // ảnh nhỏ hơn ~8KB gần như chắc chắn là logo chữ ký
const MAX_HTML_BYTES = 30 * 1024;    // displayReplyForm giới hạn ~32KB

const extOf = name => {
  const parts = String(name || '').split('.');
  return parts.length > 1 ? parts.pop().toLowerCase() : '';
};

const isImage = ext => (SUPPORTED[ext] || '').startsWith('image/');

/* ------------------------------------------------------------------ trạng thái */

let currentEmail = { subject: '', sender_name: '', sender_email: '', body: '' };
let currentAttachments = [];
let busy = false;

function setStatus(kind, text) {
  const el = $('status');
  if (!text) { el.hidden = true; return; }
  el.hidden = false;
  el.className = `status ${kind}`;
  el.textContent = text;
}

function setNotes(notes) {
  const el = $('notes');
  if (!notes || !notes.length) { el.hidden = true; return; }
  el.hidden = false;
  el.innerHTML = '<ul>' + notes.map(n => `<li>${escapeHtml(n)}</li>`).join('') + '</ul>';
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"]/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function setBusy(on) {
  busy = on;
  ['generate', 'classify', 'refine', 'reply', 'replyAll', 'copy']
    .forEach(id => { const b = $(id); if (b) b.disabled = on; });
}

/* --------------------------------------------------------------------- API */

async function request(path, payload) {
  const res = await fetch(`${BACKEND}${path}`, {
    method: payload === undefined ? 'GET' : 'POST',
    headers: payload === undefined
      ? { 'X-Addin-Token': TOKEN }
      : { 'Content-Type': 'application/json', 'X-Addin-Token': TOKEN },
    body: payload === undefined ? undefined : JSON.stringify(payload),
  });
  if (res.status === 401) {
    throw new Error('Backend đã khởi động lại. Hãy đóng và mở lại task pane.');
  }
  const data = await res.json();
  if (!data.success) throw new Error(data.error || 'Backend báo lỗi không rõ nguyên nhân.');
  return data.data;
}

const call = (path, payload) => request(path, payload || {});
const get = path => request(path);

/* ---------------------------------------------------------------- mẫu thư */

// Danh sách lấy từ kien_thuc/loai_thu/. Để trống ô chọn thì backend tự khớp theo
// từ khoá trong chủ đề và thân thư — vẫn là hành vi mặc định như trước.
async function loadEmailTypes() {
  const sel = $('emailType');
  try {
    const data = await get('/api/email-types');
    const types = Array.isArray(data) ? data : (data && data.types ? data.types : []);
    if (!types || !types.length) {
      sel.disabled = true;
      $('typeHint').textContent =
        'Chưa có mẫu thư nào. Chạy XUAT_THU.bat rồi để Antigravity viết vào kien_thuc/loai_thu/.';
      return;
    }
    sel.disabled = false;
    sel.innerHTML = '';

    // Mẫu mặc định: Thẩm định mặt bằng & mở điểm mới
    const defaultOpt = document.createElement('option');
    defaultOpt.value = 'tham-dinh-mat-bang';
    defaultOpt.textContent = 'Thẩm định mặt bằng & mở điểm mới (Mặc định)';
    sel.appendChild(defaultOpt);

    const autoOpt = document.createElement('option');
    autoOpt.value = '';
    autoOpt.textContent = '✨ AI tự nhận diện theo nội dung email';
    sel.appendChild(autoOpt);

    types.forEach(t => {
      if (t.name === 'tham-dinh-mat-bang') return;
      const opt = document.createElement('option');
      opt.value = t.name;
      opt.textContent = t.title;
      sel.appendChild(opt);
    });

    sel.value = 'tham-dinh-mat-bang';
    $('typeHint').textContent = `${types.length} mẫu thư sẵn có.`;
  } catch (e) {
    sel.disabled = true;
    $('typeHint').textContent = `Không nạp được danh sách mẫu thư: ${e.message || e}`;
  }
}

function showUsedType(data) {
  if (!data.email_type) return '';
  const label = data.email_type_title || data.email_type;
  return data.email_type_auto ? ` AI đã chọn mẫu "${label}".` : ` Dùng mẫu "${label}".`;
}

/* ------------------------------------------------------------------ Office */

function readBody() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, r => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(r.error);
      resolve(r.value || '');
    });
  });
}

function getAttachmentContent(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(id, r => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(r.error);
      resolve(r.value);
    });
  });
}

async function collectAttachments(notes) {
  const item = Office.context.mailbox.item;
  if (!Office.context.requirements.isSetSupported('Mailbox', '1.8')) {
    notes.push('Outlook này không hỗ trợ đọc nội dung tệp đính kèm (cần Mailbox 1.8).');
    return [];
  }

  const metas = (item.attachments || []).filter(a => {
    const ext = extOf(a.name);
    if (a.attachmentType !== Office.MailboxEnums.AttachmentType.File) {
      notes.push(`Bỏ qua "${a.name}": không phải tệp thường (${a.attachmentType}).`);
      return false;
    }
    if (a.isInline) return false;                        // logo chữ ký, bỏ im lặng
    if (!(ext in SUPPORTED)) {
      notes.push(`Bỏ qua "${a.name}": chưa hỗ trợ định dạng .${ext}.`);
      return false;
    }
    if (a.size > MAX_ONE) {
      notes.push(`Bỏ qua "${a.name}": ${(a.size / 1048576).toFixed(1)} MB, quá lớn.`);
      return false;
    }
    if (isImage(ext) && a.size < MIN_IMAGE) return false;
    return true;
  });

  // PDF và ảnh trước vì giá trị thông tin cao nhất, rồi mới tới bảng tính/văn bản.
  const rank = a => (extOf(a.name) === 'pdf' ? 0 : isImage(extOf(a.name)) ? 1 : 2);
  metas.sort((x, y) => rank(x) - rank(y));

  const out = [];
  let total = 0;
  for (const a of metas) {
    if (out.length >= MAX_COUNT || total + a.size > MAX_TOTAL) {
      notes.push(`Bỏ qua "${a.name}": đã chạm giới hạn tổng dung lượng.`);
      continue;
    }
    try {
      const c = await getAttachmentContent(a.id);
      if (c.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        notes.push(`Bỏ qua "${a.name}": Outlook trả về định dạng ${c.format}.`);
        continue;
      }
      out.push({ name: a.name, content_type: SUPPORTED[extOf(a.name)],
                 size: a.size, data_b64: c.content });
      total += a.size;
    } catch (e) {
      notes.push(`Không đọc được "${a.name}": ${e.message || e}`);
    }
  }
  return out;
}

async function readEmail() {
  const item = Office.context.mailbox.item;
  const from = item.from || item.sender || {};   // from là null trên thư đã gửi/nháp
  currentEmail = {
    subject: item.subject || '',
    sender_name: from.displayName || '',
    sender_email: from.emailAddress || '',
    body: await readBody(),
  };
  $('subject').textContent = currentEmail.subject || '(Không có chủ đề)';
  $('sender').textContent = currentEmail.sender_name
    ? `${currentEmail.sender_name} <${currentEmail.sender_email}>`
    : (currentEmail.sender_email || '');

  const files = (item.attachments || []).filter(a => !a.isInline);
  $('attach').textContent = files.length
    ? `${files.length} tệp đính kèm: ${files.map(a => a.name).join(', ')}`
    : 'Không có tệp đính kèm';
  return currentEmail;
}

function me() {
  const p = Office.context.mailbox.userProfile || {};
  return { name: p.displayName || '', email: p.emailAddress || '' };
}

/* ------------------------------------------------------------ chèn trả lời */

async function insertReply(all = true) {
  const html = $('draft').innerHTML.trim();
  if (!html) { setStatus('error', 'Chưa có nội dung để chèn.'); return; }

  // Giới hạn 32KB
  const bytes = new Blob([html]).size;
  if (bytes > MAX_HTML_BYTES) {
    setStatus('error',
      `Bản nháp dài ${(bytes / 1024).toFixed(1)} KB, vượt giới hạn ` +
      `${MAX_HTML_BYTES / 1024} KB của Outlook. Hãy rút ngắn hoặc dùng nút Sao chép HTML.`);
    return;
  }

  const item = Office.context && Office.context.mailbox && Office.context.mailbox.item;

  // 1. Chế độ Compose (người dùng đang mở sẵn cửa sổ thư nháp/trả lời):
  if (item && item.body && typeof item.body.setAsync === 'function') {
    setStatus('info', 'Đang điền nội dung vào thư nháp…');
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, r => {
      if (r && r.status === Office.AsyncResultStatus.Failed) {
        setStatus('error', `Lỗi điền thư nháp: ${(r.error || {}).message || ''}`);
      } else {
        setStatus('done', 'Đã tự động điền xong nội dung vào thư nháp trong Outlook. Kiểm tra lại rồi bấm Gửi.');
      }
    });
    return;
  }

  // 2. Chế độ Read (đang đọc thư): Tự động mở cửa sổ Trả lời tất cả (Reply All)
  setStatus('info', 'Đang mở cửa sổ Trả lời tất cả trong Outlook…');

  // Ưu tiên 1: Gọi backend COM để mở cửa sổ Reply All trực tiếp trong Outlook Desktop
  try {
    const res = await call('/api/open-reply-all', {
      html_body: html,
      subject: (currentEmail && currentEmail.subject) || (item ? item.subject : ''),
      sender_email: (currentEmail && currentEmail.sender_email) || '',
    });
    if (res && res.success) {
      setStatus('done', 'Đã tự động mở cửa sổ Trả lời tất cả trong Outlook. Kiểm tra lại rồi bấm Gửi.');
      return;
    }
  } catch (e) {
    console.warn("Backend COM open-reply-all failed, fallback to Office JS:", e);
  }

  // Fallback 2: Sử dụng Office JS displayReplyAllForm
  if (item) {
    const fn = all !== false ? 'displayReplyAllForm' : 'displayReplyForm';
    const fnAsync = fn + 'Async';
    try {
      if (typeof item[fnAsync] === 'function' && Office.context.requirements && Office.context.requirements.isSetSupported('Mailbox', '1.9')) {
        item[fnAsync]({ htmlBody: html }, r => {
          if (r && r.status === Office.AsyncResultStatus.Succeeded) {
            setStatus('done', 'Đã tự động mở thư nháp trả lời tất cả trong Outlook. Kiểm tra lại rồi bấm Gửi.');
          } else {
            tryOfficeJsSync();
          }
        });
        return;
      }
    } catch (_) {}
    tryOfficeJsSync();
  } else {
    setStatus('done', 'Đã tạo bản nháp thành công bên dưới.');
  }

  function tryOfficeJsSync() {
    try {
      const fn = all !== false ? 'displayReplyAllForm' : 'displayReplyForm';
      if (item && typeof item[fn] === 'function') {
        try {
          item[fn](html);
          setStatus('done', 'Đã tự động mở thư nháp trả lời tất cả trong Outlook. Kiểm tra lại rồi bấm Gửi.');
          return;
        } catch (_) {
          item[fn]({ htmlBody: html });
          setStatus('done', 'Đã tự động mở thư nháp trả lời tất cả trong Outlook. Kiểm tra lại rồi bấm Gửi.');
          return;
        }
      }
    } catch (e) {
      console.warn("Office JS displayReplyForm error:", e);
    }
    setStatus('done', 'Đã tạo bản nháp thành công bên dưới.');
  }
}


/* ------------------------------------------------------------------ thao tác */

$('generate').onclick = async () => {
  if (busy) return;
  setBusy(true);
  setNotes([]);
  setStatus('info', 'Đang đọc email…');
  try {
    const email = await readEmail();
    const notes = [];
    let attachments = [];
    if ($('useAttach').checked) {
      setStatus('info', 'Đang đọc tệp đính kèm…');
      attachments = await collectAttachments(notes);
    }
    setStatus('info', attachments.length
      ? `Đang soạn thư (Gemini đang đọc ${attachments.length} tệp, có thể mất 1-2 phút)…`
      : 'Đang soạn thư…');

    const data = await call('/api/generate-reply', {
      email, instruction: $('instruction').value, attachments, me: me(),
      email_type: $('emailType').value,
    });

    $('draft').innerHTML = data.html_body || '';
    $('draftBox').hidden = false;
    setNotes(notes.concat(data.notes || []));
    $('draftBox').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    
    // Tự động mở cửa sổ thư nháp Trả lời tất cả trong Outlook
    await insertReply(true);
  } catch (e) {
    setStatus('error', e.message || String(e));
  } finally {
    setBusy(false);
  }
};

$('refine').onclick = async () => {
  if (busy) return;
  const feedback = $('feedback').value.trim();
  if (!feedback) { setStatus('error', 'Hãy nhập yêu cầu chỉnh sửa.'); return; }
  setBusy(true);
  setStatus('info', 'Đang chỉnh sửa bản nháp…');
  try {
    const data = await call('/api/refine-draft', {
      draft_html: $('draft').innerHTML, feedback,
    });
    $('draft').innerHTML = data.html_body || '';
    $('feedback').value = '';
    // Tự động cập nhật mở lại thư nháp trong Outlook
    await insertReply(true);
  } catch (e) {
    setStatus('error', e.message || String(e));
  } finally {
    setBusy(false);
  }
};

$('classify').onclick = async () => {
  if (busy) return;
  setBusy(true);
  setStatus('info', 'Đang phân loại…');
  try {
    const email = await readEmail();
    const d = await call('/api/classify-email', { email });
    const badge = $('badge');
    if (d.priority) {
      badge.hidden = false;
      badge.className = 'badge ' + d.priority.toLowerCase().replace(/\s+/g, '');
      badge.textContent = d.priority;
    }
    setStatus('done',
      `${d.category || 'Không rõ'} · ${d.needs_reply ? 'cần trả lời' : 'không cần trả lời'}` +
      (d.summary ? ` — ${d.summary}` : ''));
  } catch (e) {
    setStatus('error', e.message || String(e));
  } finally {
    setBusy(false);
  }
};

$('reply').onclick = () => insertReply(false);
$('replyAll').onclick = () => insertReply(true);

$('copy').onclick = async () => {
  try {
    await navigator.clipboard.writeText($('draft').innerHTML);
    setStatus('done', 'Đã sao chép HTML vào clipboard.');
  } catch (e) {
    setStatus('error', `Không sao chép được: ${e.message || e}`);
  }
};

/* --------------------------------------------------- học văn phong (nền) */

let stylePolling = null;

function renderStyleInfo(kb) {
  if (!kb) { $('styleInfo').textContent = ''; return; }
  const parts = [];
  parts.push(kb.has_learned
    ? `Đã thống kê ${kb.sample_count} thư` +
      (kb.learned_at ? ` (${new Date(kb.learned_at).toLocaleDateString('vi-VN')})` : '')
    : 'Chưa xuất thư lần nào.');
  parts.push(kb.type_count
    ? `${kb.type_count} loại thư đã có hướng dẫn riêng.`
    : 'Chưa có hướng dẫn theo loại thư.');
  $('styleInfo').textContent = parts.join(' · ');
}

function renderStyleProgress(s) {
  const box = $('styleProgress');
  const bar = $('styleBar');
  box.hidden = false;
  $('stylePhase').textContent = s.total
    ? `${s.phase} ${s.current}/${s.total}`
    : (s.phase || 'Đang xử lý…');
  if (s.total) {
    bar.classList.remove('indeterminate');
    bar.style.width = `${Math.round((s.current / s.total) * 100)}%`;
  } else {
    // Giai đoạn quét Outlook không biết trước tổng số -> chạy thanh vô định.
    bar.classList.add('indeterminate');
    bar.style.width = '';
  }
}

async function pollStyle() {
  try {
    const s = await get('/api/export-emails/status');
    renderStyleProgress(s);
    if (s.running) return;

    clearInterval(stylePolling);
    stylePolling = null;
    $('learnStyle').disabled = false;
    $('styleProgress').hidden = true;

    if (s.error) {
      setStatus('error', `Xuất thư thất bại: ${s.error}`);
    } else if (s.done && s.result) {
      const r = s.result;
      setStatus('done',
        `Đã xuất ${r.unique} thư (quét ${r.total_raw}, ${r.short} thư ngắn) ` +
        `thành ${r.files} tệp trong ${r.dir}. ` +
        `Mở thư mục đó bằng Antigravity và ra lệnh "đọc 00_TONG_QUAN.md rồi làm theo đề bài".`);
      refreshStyleInfo();
    }
  } catch (e) {
    clearInterval(stylePolling);
    stylePolling = null;
    $('learnStyle').disabled = false;
    $('styleProgress').hidden = true;
    setStatus('error', e.message || String(e));
  }
}

async function refreshStyleInfo() {
  try {
    renderStyleInfo(await get('/api/knowledge'));
  } catch {
    $('styleInfo').textContent = 'Không đọc được trạng thái văn phong.';
  }
}

$('learnStyle').onclick = async () => {
  if (stylePolling) return;
  $('learnStyle').disabled = true;
  setStatus('info', 'Đang quét hộp thư. Giữ Outlook mở — việc này có thể mất 10–30 phút.');
  try {
    renderStyleProgress(await call('/api/export-emails',
                                   { deep: $('deepScan').checked }));
    stylePolling = setInterval(pollStyle, 2000);
  } catch (e) {
    $('learnStyle').disabled = false;
    $('styleProgress').hidden = true;
    setStatus('error', e.message || String(e));
  }
};

/* --------------------------------------------------------------- khởi động */

if (window.Office) {
  Office.onReady(() => {
    if (!TOKEN || TOKEN.indexOf('{{') === 0) {
      setStatus('error',
        'Không lấy được mã phiên. Trang này phải được mở qua backend ' +
        '(chạy CHAY_ADDIN_BACKEND.bat), không mở trực tiếp từ đĩa.');
      return;
    }
    readEmail().catch(e =>
      setStatus('error', `Không đọc được email: ${e.message || e}`));

    loadEmailTypes();
    refreshStyleInfo();
    // Task pane có thể bị đóng/mở lại giữa lúc đang học — bắt lại tiến độ nếu còn chạy.
    get('/api/export-emails/status').then(s => {
      if (s.running) {
        $('styleBox').open = true;
        $('learnStyle').disabled = true;
        renderStyleProgress(s);
        stylePolling = setInterval(pollStyle, 2000);
      }
    }).catch(() => {});
  });
}
