const state = { items: [], busy: false };

const fileInput = document.querySelector('#file-input');
const dropzone = document.querySelector('#dropzone');
const fileList = document.querySelector('#file-list');
const schoolInput = document.querySelector('#school');
const yearInput = document.querySelector('#year');
const generateButton = document.querySelector('#generate');
const validation = document.querySelector('#validation');
const toast = document.querySelector('#toast');

yearInput.value = new Date().getFullYear();

function passStats(data) {
  const below = data.dist
    .filter(([label]) => label.includes('以下') && label.includes('60'))
    .reduce((sum, row) => sum + Number(row[1]), 0);
  const passed = Number(data.n) - below;
  return { passed, rate: data.n ? passed / Number(data.n) * 100 : 0 };
}

function showToast(message, isError = false) {
  toast.textContent = message;
  toast.className = `toast show${isError ? ' error' : ''}`;
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => { toast.className = 'toast'; }, 3200);
}

async function readError(response) {
  try { return (await response.json()).error || '处理失败'; }
  catch { return '处理失败，请稍后重试'; }
}

async function uploadFiles(files) {
  const accepted = [...files].filter(file => file.name.toLowerCase().endsWith('.docx'));
  if (accepted.length !== files.length) showToast('已忽略非 .docx 文件', true);
  if (!accepted.length) return;
  const totalBytes = accepted.reduce((sum, file) => sum + file.size, 0);
  if (accepted.length > 8) {
    showToast('一次最多上传 8 份报告', true);
    return;
  }
  if (totalBytes > 4 * 1024 * 1024) {
    showToast('单次上传文件合计不能超过 4MB', true);
    return;
  }
  state.busy = true;
  updateAll();
  const placeholders = accepted.map(file => ({ filename: file.name, loading: true }));
  state.items.push(...placeholders);
  renderFiles();

  const form = new FormData();
  accepted.forEach(file => form.append('files', file));
  try {
    const response = await fetch('/api/parse', { method: 'POST', body: form });
    if (!response.ok) throw new Error(await readError(response));
    const payload = await response.json();
    placeholders.forEach((placeholder, index) => {
      const position = state.items.indexOf(placeholder);
      if (position >= 0) state.items.splice(position, 1, payload.files[index]);
    });
  } catch (error) {
    placeholders.forEach(placeholder => {
      const position = state.items.indexOf(placeholder);
      if (position >= 0) state.items.splice(position, 1, { filename: placeholder.filename, ok: false, error: error.message });
    });
    showToast(error.message, true);
  } finally {
    state.busy = false;
    fileInput.value = '';
    updateAll();
  }
}

function renderFiles() {
  if (!state.items.length) {
    fileList.innerHTML = '<div class="empty-state">尚未添加报告</div>';
    return;
  }
  fileList.innerHTML = state.items.map((item, index) => {
    if (item.loading) return `
      <div class="file-row error-row">
        <div class="file-name"><strong title="${escapeHtml(item.filename)}">${escapeHtml(item.filename)}</strong><span>正在解析数据表…</span></div>
        <div class="error-message">请稍候</div>
      </div>`;
    if (!item.ok) return `
      <div class="file-row error-row">
        <div class="file-name"><strong title="${escapeHtml(item.filename)}">${escapeHtml(item.filename)}</strong><span>解析失败</span></div>
        <div class="error-message">${escapeHtml(item.error)}</div>
        <button class="remove-button" data-remove="${index}" title="移除文件" aria-label="移除文件">×</button>
      </div>`;
    const stats = passStats(item.data);
    return `
      <div class="file-row">
        <div class="file-name"><strong title="${escapeHtml(item.filename)}">${escapeHtml(item.filename)}</strong><span>${item.data.primary.length} 个一级维度 · ${item.data.secondary.length} 个二级维度</span></div>
        <input class="age-input" data-age="${index}" list="age-options" value="${escapeHtml(item.data.name)}" placeholder="填写年龄段" aria-label="年龄段">
        <div class="data-cell"><span>样本</span><strong>${item.data.n}</strong></div>
        <div class="data-cell"><span>平均分</span><strong>${Number(item.data.mean).toFixed(2)}</strong></div>
        <div class="data-cell"><span>达标率</span><strong>${stats.rate.toFixed(2)}%</strong></div>
        <button class="remove-button" data-remove="${index}" title="移除文件" aria-label="移除文件">×</button>
      </div>`;
  }).join('');
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
}

function validItems() { return state.items.filter(item => item.ok); }

function updateSummary() {
  const items = validItems();
  const total = items.reduce((sum, item) => sum + Number(item.data.n), 0);
  const weighted = total ? items.reduce((sum, item) => sum + Number(item.data.mean) * Number(item.data.n), 0) / total : null;
  const passed = items.reduce((sum, item) => sum + passStats(item.data).passed, 0);
  document.querySelector('#metric-groups').textContent = items.length;
  document.querySelector('#metric-samples').textContent = total || '--';
  document.querySelector('#metric-mean').textContent = weighted === null ? '--' : weighted.toFixed(2);
  document.querySelector('#metric-pass').textContent = total ? `${(passed / total * 100).toFixed(2)}%` : '--';
}

function validate() {
  const items = validItems();
  const names = items.map(item => item.data.name.trim());
  let message = '';
  let ready = false;
  if (state.busy) message = '正在解析报告，请稍候';
  else if (items.length < 2) message = '请先上传至少两份有效的年龄段报告';
  else if (names.some(name => !name)) message = '请补充未识别的年龄段名称';
  else if (new Set(names).size !== names.length) message = '年龄段名称不能重复';
  else if (!schoolInput.value.trim()) message = '请填写学校名称';
  else { message = `已就绪，将生成 ${items.length + 1} 个主体部分`; ready = true; }
  validation.textContent = message;
  validation.className = `validation${ready ? ' ready' : ''}`;
  generateButton.disabled = !ready || state.busy;
  return ready;
}

function updateAll() { renderFiles(); updateSummary(); validate(); }

fileList.addEventListener('click', event => {
  const button = event.target.closest('[data-remove]');
  if (!button) return;
  state.items.splice(Number(button.dataset.remove), 1);
  updateAll();
});

fileList.addEventListener('input', event => {
  if (!event.target.matches('[data-age]')) return;
  const item = state.items[Number(event.target.dataset.age)];
  if (item?.ok) item.data.name = event.target.value;
  updateSummary();
  validate();
});

fileInput.addEventListener('change', () => uploadFiles(fileInput.files));
['dragenter', 'dragover'].forEach(name => dropzone.addEventListener(name, event => { event.preventDefault(); dropzone.classList.add('dragging'); }));
['dragleave', 'drop'].forEach(name => dropzone.addEventListener(name, event => { event.preventDefault(); dropzone.classList.remove('dragging'); }));
dropzone.addEventListener('drop', event => uploadFiles(event.dataTransfer.files));
schoolInput.addEventListener('input', validate);
yearInput.addEventListener('input', validate);

generateButton.addEventListener('click', async () => {
  if (!validate()) return;
  state.busy = true;
  generateButton.disabled = true;
  document.querySelector('#button-label').textContent = '正在生成报告…';
  try {
    const response = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ school: schoolInput.value.trim(), year: Number(yearInput.value), datasets: validItems().map(item => item.data) })
    });
    if (!response.ok) throw new Error(await readError(response));
    const blob = await response.blob();
    const disposition = response.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
    const filename = match ? decodeURIComponent(match[1]) : '学校学生素养测评数据分析报告.docx';
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    showToast('报告已生成并开始下载');
  } catch (error) {
    showToast(error.message, true);
  } finally {
    state.busy = false;
    document.querySelector('#button-label').textContent = '生成学校报告';
    validate();
  }
});

updateAll();
