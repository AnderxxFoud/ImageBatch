const MAX_NAMES = 5000;
const CHUNK_SIZE = 50;
let state = {
  excelFile: null,
  imgFile: null,
  names: [],
  sheetNames: [],
  zipBlob: null,
};

const $ = (id) => document.getElementById(id);

const dropExcel   = $('drop-excel');
const dropImg     = $('drop-img');
const inputExcel  = $('input-excel');
const inputImg    = $('input-img');
const excelInfo   = $('excel-info');
const excelName   = $('excel-name');
const clearExcel  = $('clear-excel');
const imgInfo     = $('img-info');
const imgName     = $('img-name');
const imgThumb    = $('img-thumb');
const clearImg    = $('clear-img');
const configRow   = $('config-row');
const colInput    = $('col-input');
const sheetSelect = $('sheet-select');
const namesCount  = $('names-count');
const summary     = $('summary');
const sImg        = $('s-img');
const sCount      = $('s-count');
const sSize       = $('s-size');
const btnGo       = $('btn-go');
const btnLabel    = $('btn-label');
const progressSec = $('progress-section');
const progressBar = $('progress-bar');
const progressCnt = $('progress-count');
const progressSub = $('progress-sub');
const doneSec     = $('done-section');
const doneSub     = $('done-sub');
const btnDownload = $('btn-download');
const btnReset    = $('btn-reset');

function sanitize(name) {
  return String(name).replace(/[\\/?%*:|"<>\n\r]/g, '-').trim();
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function parseExcel(file, col, sheetIndex) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames[sheetIndex];
        if (!sheetName) return reject(new Error('Hoja no encontrada'));
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const colIdx = col.toUpperCase().charCodeAt(0) - 65;
        const names = rows
          .map((r) => r[colIdx])
          .filter((v) => v !== undefined && v !== null && String(v).trim() !== '')
          .map((v) => String(v).trim())
          .slice(0, MAX_NAMES);
        resolve({ names, sheetNames: wb.SheetNames });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('Error al leer el archivo'));
    reader.readAsArrayBuffer(file);
  });
}

function fileToArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = () => reject(new Error('Error leyendo imagen'));
    reader.readAsArrayBuffer(file);
  });
}

function updateExcelUI() {
  const hasExcel = !!state.excelFile;
  dropExcel.hidden = hasExcel;
  excelInfo.hidden = !hasExcel;
  configRow.hidden = !hasExcel;
  if (hasExcel) excelName.textContent = state.excelFile.name;
}

function updateImgUI() {
  const hasImg = !!state.imgFile;
  dropImg.hidden = hasImg;
  imgInfo.hidden = !hasImg;
  if (hasImg) {
    imgName.textContent = state.imgFile.name;
    imgThumb.src = URL.createObjectURL(state.imgFile);
  }
}

function updateNamesDisplay() {
  const n = state.names.length;
  namesCount.textContent = n === 0 ? '—' : n.toLocaleString('es-MX');
  if (n === MAX_NAMES) {
    namesCount.textContent += ` (límite: ${MAX_NAMES.toLocaleString('es-MX')})`;
  }
}

function updateSummary() {
  const ready = state.imgFile && state.names.length > 0;
  summary.hidden = !ready;
  if (!ready) return;

  sImg.textContent = state.imgFile.name;
  sCount.textContent = state.names.length.toLocaleString('es-MX') + ' copias';
  const est = state.imgFile.size * state.names.length;
  sSize.textContent = '~' + formatBytes(est);
}

function updateButton() {
  const ready = state.excelFile && state.imgFile && state.names.length > 0;
  btnGo.disabled = !ready;
  btnLabel.textContent = ready
    ? `Generar ${state.names.length.toLocaleString('es-MX')} imágenes`
    : 'Esperando archivos…';
}

function refresh() {
  updateNamesDisplay();
  updateSummary();
  updateButton();
}

async function loadExcel() {
  if (!state.excelFile) return;
  try {
    const col = colInput.value || 'A';
    const sheetIdx = sheetSelect.selectedIndex >= 0 ? sheetSelect.selectedIndex : 0;
    const result = await parseExcel(state.excelFile, col, sheetIdx);
    state.names = result.names;

    if (result.sheetNames.length && sheetSelect.options.length === 0) {
      result.sheetNames.forEach((n, i) => {
        const opt = document.createElement('option');
        opt.value = i;
        opt.textContent = n;
        sheetSelect.appendChild(opt);
      });
    }
    refresh();
  } catch (err) {
    alert('Error al leer el Excel: ' + err.message);
  }
}

async function generateZip() {
  const imgBuffer = await fileToArrayBuffer(state.imgFile);
  const ext = state.imgFile.name.includes('.')
    ? '.' + state.imgFile.name.split('.').pop()
    : '';

  const zip = new JSZip();
  const folder = zip.folder('imagenes');
  const total = state.names.length;

  progressSec.hidden = false;
  progressBar.style.width = '0%';
  progressCnt.textContent = `0 / ${total.toLocaleString('es-MX')}`;
  progressSub.textContent = 'Preparando archivos…';

  for (let i = 0; i < total; i += CHUNK_SIZE) {
    const chunk = state.names.slice(i, i + CHUNK_SIZE);
    for (const name of chunk) {
      const filename = sanitize(name) + ext;
      folder.file(filename, imgBuffer, { binary: true });
    }
    const done = Math.min(i + CHUNK_SIZE, total);
    const pct = Math.round((done / total) * 100);
    progressBar.style.width = pct + '%';
    progressCnt.textContent = `${done.toLocaleString('es-MX')} / ${total.toLocaleString('es-MX')}`;
    progressSub.textContent = `Empaquetando: ${state.names[Math.min(i, total - 1)]}…`;
    await sleep(0);
  }

  progressSub.textContent = 'Comprimiendo ZIP…';
  await sleep(10);

  const blob = await zip.generateAsync(
    {
      type: 'blob',
      compression: 'STORE',
      streamFiles: true,
    },
    (meta) => {
      const pct = Math.round(meta.percent);
      progressBar.style.width = pct + '%';
      progressSub.textContent = `Comprimiendo… ${pct}%`;
    }
  );

  return blob;
}

btnGo.addEventListener('click', async () => {
  btnGo.disabled = true;
  btnLabel.textContent = 'Procesando…';
  doneSec.hidden = true;

  try {
    const blob = await generateZip();
    state.zipBlob = blob;

    progressSec.hidden = true;
    doneSec.hidden = false;
    doneSub.textContent =
      `${state.names.length.toLocaleString('es-MX')} imágenes · ${formatBytes(blob.size)}`;
  } catch (err) {
    alert('Error al generar el ZIP: ' + err.message);
    btnGo.disabled = false;
    btnLabel.textContent = `Generar ${state.names.length.toLocaleString('es-MX')} imágenes`;
    progressSec.hidden = true;
  }
});

btnDownload.addEventListener('click', () => {
  if (!state.zipBlob) return;
  const url = URL.createObjectURL(state.zipBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'imagenes.zip';
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
});

btnReset.addEventListener('click', () => {
  state = { excelFile: null, imgFile: null, names: [], sheetNames: [], zipBlob: null };
  inputExcel.value = '';
  inputImg.value = '';
  sheetSelect.innerHTML = '';
  imgThumb.src = '';
  doneSec.hidden = true;
  progressSec.hidden = true;
  updateExcelUI();
  updateImgUI();
  refresh();
});

function handleExcelFile(file) {
  if (!file) return;
  const ok = /\.(xlsx|xls)$/i.test(file.name);
  if (!ok) { alert('Por favor sube un archivo .xlsx o .xls'); return; }
  state.excelFile = file;
  sheetSelect.innerHTML = '';
  updateExcelUI();
  loadExcel();
}

function handleImgFile(file) {
  if (!file) return;
  if (!file.type.startsWith('image/')) { alert('Por favor sube un archivo de imagen'); return; }
  state.imgFile = file;
  updateImgUI();
  refresh();
}

dropExcel.addEventListener('click', () => inputExcel.click());
dropExcel.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') inputExcel.click(); });
dropImg.addEventListener('click', () => inputImg.click());
dropImg.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') inputImg.click(); });

inputExcel.addEventListener('change', (e) => handleExcelFile(e.target.files[0]));
inputImg.addEventListener('change', (e) => handleImgFile(e.target.files[0]));

dropExcel.addEventListener('dragover', (e) => { e.preventDefault(); dropExcel.classList.add('over'); });
dropExcel.addEventListener('dragleave', () => dropExcel.classList.remove('over'));
dropExcel.addEventListener('drop', (e) => {
  e.preventDefault();
  dropExcel.classList.remove('over');
  handleExcelFile(e.dataTransfer.files[0]);
});

dropImg.addEventListener('dragover', (e) => { e.preventDefault(); dropImg.classList.add('over'); });
dropImg.addEventListener('dragleave', () => dropImg.classList.remove('over'));
dropImg.addEventListener('drop', (e) => {
  e.preventDefault();
  dropImg.classList.remove('over');
  handleImgFile(e.dataTransfer.files[0]);
});

clearExcel.addEventListener('click', () => {
  state.excelFile = null;
  state.names = [];
  sheetSelect.innerHTML = '';
  inputExcel.value = '';
  updateExcelUI();
  refresh();
});

clearImg.addEventListener('click', () => {
  state.imgFile = null;
  inputImg.value = '';
  URL.revokeObjectURL(imgThumb.src);
  imgThumb.src = '';
  updateImgUI();
  refresh();
});

colInput.addEventListener('input', () => loadExcel());
sheetSelect.addEventListener('change', () => loadExcel());
