"""
Milka Implementation Report Generator
Flask web app: upload WhatsApp ZIP → generate PPTX
Run: python app.py
"""
import os, uuid, threading, shutil
from datetime import datetime, timedelta
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template_string

from processor import process_zip, CHAIN_ORDER

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB

BASE_DIR  = Path(__file__).parent
UPLOAD_DIR = Path(os.environ.get('UPLOAD_DIR', '/tmp/milka_uploads'))
TEMPLATE_PPTX = BASE_DIR / 'Template_Ejemplo.pptx'
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# In-memory job store  {job_id: {status, progress, result, error}}
jobs: dict = {}


# ── HTML ──────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Generador de Informes de Implementación</title>
<style>
  :root {
    --navy: #103B78;
    --navy-dark: #0A2757;
    --navy-light: #2E6BC4;
    --navy-bg: #EEF3FB;
    --red: #CC1A0E;
    --red-light: #E8382B;
    --white: #FFFFFF;
    --gray: #6B7280;
    --success: #10B981;
    --error: #EF4444;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Segoe UI', system-ui, sans-serif;
    background: var(--navy-bg);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
  }
  .card {
    background: var(--white);
    border-radius: 20px;
    box-shadow: 0 8px 40px rgba(16,59,120,0.13);
    padding: 48px;
    width: 100%;
    max-width: 700px;
  }
  .logo { text-align: center; margin-bottom: 32px; }
  .logo img { height: 56px; }
  .logo h1 {
    font-size: 1.8rem;
    font-weight: 800;
    color: var(--navy-dark);
    margin-top: 12px;
  }
  .logo p { color: var(--gray); margin-top: 6px; font-size: 0.95rem; }

  /* Drop zone */
  .drop-zone {
    border: 2.5px dashed var(--navy-light);
    border-radius: 14px;
    padding: 48px 24px;
    text-align: center;
    cursor: pointer;
    transition: all .2s;
    background: var(--navy-bg);
    position: relative;
  }
  .drop-zone:hover, .drop-zone.drag-over {
    border-color: var(--navy);
    background: #D9E6F7;
  }
  .drop-zone input[type=file] {
    position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%;
  }
  .drop-zone .icon { font-size: 3rem; margin-bottom: 12px; }
  .drop-zone h3 { font-size: 1.1rem; color: var(--navy-dark); font-weight: 700; }
  .drop-zone p  { font-size: 0.85rem; color: var(--gray); margin-top: 6px; }
  .drop-zone .file-name {
    margin-top: 14px; font-weight: 700; color: var(--navy);
    background: #D9E6F7; padding: 8px 16px; border-radius: 8px;
    display: inline-block;
  }

  /* Date row */
  .date-row {
    display: flex; gap: 16px; margin: 24px 0;
  }
  .field { flex: 1; }
  label { display: block; font-size: 0.82rem; font-weight: 700;
          color: var(--navy-dark); margin-bottom: 6px; text-transform: uppercase; letter-spacing: .5px; }
  input[type=date] {
    width: 100%; padding: 10px 14px; border: 2px solid #E5E7EB;
    border-radius: 10px; font-size: 0.95rem; outline: none;
    transition: border-color .2s; color: #111;
  }
  input[type=date]:focus { border-color: var(--navy); }

  /* Button */
  .btn {
    width: 100%; padding: 16px; border: none; border-radius: 12px;
    font-size: 1.05rem; font-weight: 800; cursor: pointer;
    background: linear-gradient(135deg, var(--red-light) 0%, var(--red) 100%);
    color: var(--white); letter-spacing: .5px;
    transition: opacity .2s, transform .1s;
  }
  .btn:hover:not(:disabled) { opacity: .92; transform: translateY(-1px); }
  .btn:disabled { opacity: .5; cursor: not-allowed; }

  /* Progress */
  #progress-section { display: none; margin-top: 28px; }
  .progress-bar-wrap {
    background: #E5E7EB; border-radius: 99px; height: 10px; overflow: hidden; margin-bottom: 10px;
  }
  .progress-bar {
    height: 100%; border-radius: 99px;
    background: linear-gradient(90deg, var(--navy-light), var(--navy));
    transition: width .4s ease;
  }
  .progress-label { font-size: 0.88rem; color: var(--gray); text-align: center; }

  /* Results */
  #results-section { display: none; margin-top: 28px; }
  .results-title { font-weight: 800; color: var(--navy-dark); font-size: 1.1rem; margin-bottom: 16px; }
  .chain-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 12px; margin-bottom: 24px; }
  .chain-chip {
    background: var(--navy-bg); border-radius: 12px; padding: 14px;
    text-align: center; border: 2px solid #C0D4EE;
  }
  .chain-chip .name { font-size: 0.78rem; font-weight: 700; color: var(--gray); text-transform: uppercase; letter-spacing: .5px; }
  .chain-chip .count { font-size: 2rem; font-weight: 800; color: var(--navy); line-height: 1.1; margin-top: 4px; }
  .chain-chip .label { font-size: 0.72rem; color: var(--gray); }

  .btn-download {
    width: 100%; padding: 16px; border: none; border-radius: 12px;
    font-size: 1.05rem; font-weight: 800; cursor: pointer;
    background: linear-gradient(135deg, var(--success) 0%, #059669 100%);
    color: var(--white); letter-spacing: .5px; display: flex; align-items: center;
    justify-content: center; gap: 10px; text-decoration: none;
    transition: opacity .2s;
  }
  .btn-download:hover { opacity: .9; }

  .error-box {
    background: #FEF2F2; border: 2px solid #FECACA; border-radius: 12px;
    padding: 16px; color: #B91C1C; font-size: 0.9rem; margin-top: 20px;
    display: none;
  }

  .steps {
    display: flex; gap: 8px; margin-bottom: 28px;
    font-size: 0.78rem; color: var(--gray);
    justify-content: center;
  }
  .steps span { background: #D9E6F7; border-radius: 99px; padding: 4px 12px; color: var(--navy); font-weight: 600; }

  /* Template upload */
  .template-zone {
    border: 2.5px dashed var(--navy-light);
    border-radius: 14px;
    padding: 32px 24px;
    text-align: center;
    cursor: pointer;
    transition: all .2s;
    background: var(--navy-bg);
    position: relative;
    margin: 0 0 24px 0;
  }
  .template-zone:hover, .template-zone.drag-over {
    border-color: var(--navy);
    background: #D9E6F7;
  }
  .template-zone input[type=file] {
    position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%;
  }
  .template-zone .icon { font-size: 2rem; margin-bottom: 8px; }
  .template-zone h3 { font-size: 1rem; color: var(--navy-dark); font-weight: 700; }
  .template-zone p { font-size: 0.85rem; color: var(--gray); margin-top: 4px; }
  .template-zone .file-name {
    margin-top: 12px; font-weight: 700; color: var(--navy);
    background: #D9E6F7; padding: 6px 14px; border-radius: 8px;
    display: inline-flex; align-items: center; gap: 8px;
  }
  .template-clear {
    background: none; border: none; cursor: pointer;
    color: #9CA3AF; font-size: 0.9rem; padding: 0 2px;
  }
  .template-clear:hover { color: var(--error); }

  footer {
    margin-top: 32px; text-align: center;
    font-size: 0.8rem; color: var(--gray);
  }
</style>
</head>
<body>

<div class="card">
  <div class="logo">
    <svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="923" height="216" viewBox="0 0 923 216" style="height:56px;width:auto;display:inline-block"><path fill="#103B78" d="M626.624 44.7331C661.733 43.311 691.343 70.6452 692.748 105.775C694.152 140.905 666.82 170.519 631.709 171.907C596.622 173.295 567.049 145.969 565.646 110.863C564.242 75.7565 591.538 46.1543 626.624 44.7331ZM633.01 149.971C655.905 147.841 672.781 127.609 670.782 104.69C668.782 81.77 648.658 64.7697 625.74 66.6408C602.639 68.5268 585.482 88.86 587.497 111.963C589.513 135.067 609.931 152.118 633.01 149.971Z"/><defs><linearGradient id="so_g0" gradientUnits="userSpaceOnUse" x1="565.00726" y1="1.6763315" x2="455.34055" y2="130.89998"><stop offset="0" stop-color="#FF1D0E"/><stop offset="1" stop-color="#CA0713"/></linearGradient></defs><path fill="url(#so_g0)" d="M565.284 0.711721L566.181 0.484396C567.557 2.47589 566.909 12.7871 566.65 15.5765C563.407 50.4314 544.113 81.9252 518.006 104.255C517.762 120.388 516.417 138.848 515.582 155.067C512.772 159.231 506.268 166.345 502.705 170.709C495.523 165.513 495.52 160.008 493.11 151.972C490.382 142.876 488.225 133.507 485.22 124.505C478.09 128.556 462.807 131.793 454.353 132.689C453.592 121.919 454.552 111.096 457.197 100.628C443.661 98.6192 427.128 98.0194 413.989 94.8222C412.763 94.5239 409.998 92.1233 408.883 91.1745C411.503 87.6041 418.767 79.3423 421.722 76.0817C426.993 74.4964 435.146 72.9919 440.688 71.8043C451.146 69.5316 461.617 67.314 472.099 65.1518C488.836 33.6757 529.691 2.79826 565.284 0.711721Z"/><path fill="#103B78" d="M111.378 46.6599C131.671 46.2639 152.169 46.8886 172.481 46.6136C178.066 46.538 184.139 46.4804 189.68 46.832C190.036 52.3076 189.908 60.2994 189.706 65.7985L154.655 65.7814C151.43 65.7766 132.381 65.066 131.698 67.5132C129.723 74.5868 130.639 90.3056 130.703 97.9896C140.191 97.6628 150.677 98.0075 160.258 97.8817C165.859 97.8081 176.988 97.6095 182.051 98.233C182.144 104.364 182.135 110.497 182.025 116.628C176.147 117.015 168.644 116.821 162.66 116.821L130.655 116.842L130.666 125.859C130.689 130.175 130.153 143.153 131.028 148.132C131.225 149.253 132.011 149.792 133.211 150.096C142.494 152.447 186.175 148.485 191.587 151.515C192.812 153.677 193.074 167.451 191.563 169.217C188.56 170.69 167.869 170.043 163.096 170.043L108.146 170.036C107.632 160.853 107.936 149.142 107.94 139.749L107.914 88.3506L107.906 61.5397C107.906 58.7723 107.77 49.8974 108.268 47.5723C109.429 46.5034 109.487 46.819 111.378 46.6599Z"/><path fill="#103B78" d="M706.074 46.644C712.415 46.6127 719.21 46.4692 725.5 46.7891C726.157 50.9568 725.941 59.8524 725.937 64.3414L725.887 100.333C725.874 109.092 725.464 117.505 726.84 126.26C729.844 145.384 746.252 152.994 764.279 149.651C790.624 144.766 787.585 118.751 787.597 97.9845L787.617 60.0902C787.607 57.4261 787.113 49.2259 788.323 47.4271C791.306 46.0203 805.803 46.5726 809.663 46.8337C810.121 53.4974 809.894 61.9612 809.88 68.7683L809.905 101.166C809.923 117.682 810.73 133.756 801.73 148.443C795.499 158.611 784.644 165.708 773.078 168.19C743.907 174.452 712.844 165.41 705.351 133.448C704.315 129.028 703.687 124.523 703.474 119.988C702.914 108.013 703.367 91.8886 703.38 79.5714L703.407 57.6102C703.406 54.5879 702.881 50.7229 703.659 47.8339C703.919 46.8718 705.183 46.8378 706.074 46.644Z"/><path fill="#103B78" d="M48.2514 46.6381C64.4507 46.311 76.9433 49.3852 91.0061 57.0824C89.1594 62.3462 86.1601 69.1482 84.0122 74.4435C74.5874 68.7575 63.8905 65.5215 52.8959 65.0302C41.9158 64.6046 25.51 67.0907 24.8861 81.1794C24.5059 92.9587 39.019 95.6894 47.7963 97.9624C64.7071 102.314 90.0555 105.854 95.0723 125.936C97.7298 137.422 94.6095 150.092 85.7123 158.104C73.4695 169.13 57.1626 170.327 41.5159 169.711C26.3279 168.845 12.6317 164.62 0.232724 155.781C2.6923 149.988 5.24384 144.235 7.88638 138.523C21.5768 147.923 34.5663 152.134 51.3336 151.543C60.4641 151.221 75.0413 147.078 74.1913 135.494C73.059 120.063 46.2821 119.238 35.1549 115.353C26.7878 112.432 18.6199 110.54 11.7739 104.391C-0.36436 94.147 0.477533 72.119 10.752 60.7861C20.6234 49.8976 34.1963 47.5091 48.2514 46.6381Z"/><path fill="#103B78" d="M303.198 46.6792C309.491 46.4929 317.76 46.5176 323.938 46.781C325.127 78.9931 322.89 115.076 324.584 147.609C324.608 148.054 325.055 149.323 325.242 149.811C327.386 151.382 378.725 150.485 386.162 150.786C386.544 154.384 386.881 165.485 385.99 169.206C385.853 169.775 385.118 169.851 384.546 170.01C357.767 170.431 328.196 170.351 301.41 169.968L301.402 169.885C300.636 160.554 300.716 49.4289 301.463 47.7588C301.826 46.9452 302.401 46.9518 303.198 46.6792Z"/><path fill="#103B78" d="M206.678 46.6522L226.684 46.6315C227.593 70.2758 226.294 95.6659 226.82 119.423C226.966 125.983 226.25 143.416 227.597 149.06C230.801 152.917 279.079 149.026 288.052 151.024C289.531 152.194 289.422 153.353 289.344 155.209C289.151 159.761 290.2 164.814 288.961 169.195C288.785 169.816 287.747 169.883 287.198 170.011C260.084 170.45 231.425 170.116 204.244 170.005C203.994 164.011 204.05 157.407 204.158 151.401C204.771 117.1 203.29 82.247 204.256 48.006C205.033 46.7425 204.877 47.095 206.678 46.6522Z"/><path fill="#103B78" d="M819.391 47.1445C825.827 46.4019 845.64 46.8983 852.71 46.8991L923 46.9314L923 65.9093C916.259 66.2267 888.659 65.4003 884.467 66.8339C882.407 69.8725 883.082 91.1724 883.083 95.7887L883.105 142.609C883.109 151.033 883.384 161.38 882.897 169.676C875.925 170.089 867.468 169.856 860.379 169.827C859.719 146.776 860.847 123.242 860.358 100.132C860.165 90.9964 861.553 75.0826 858.844 67.0362C858.349 65.5654 823.086 66.0299 819.44 65.9018C819.189 59.9301 819.358 53.1798 819.391 47.1445Z"/><defs><linearGradient id="so_g1" gradientUnits="userSpaceOnUse" x1="448.19553" y1="139.27383" x2="421.69492" y2="170.39836"><stop offset="0" stop-color="red"/><stop offset="1" stop-color="#F50"/></linearGradient></defs><path fill="url(#so_g1)" d="M411.617 163.947C416.488 155.003 423.17 138.468 432.735 135.643C437.312 134.307 442.229 134.822 446.43 137.078C450.234 139.161 453.14 142.572 454.341 146.765C455.593 151.203 455.05 155.957 452.828 159.998C449.619 165.763 437.168 174.337 431.385 178.87C427.656 174.22 417.839 168.446 413.785 163.229C413.58 162.964 412.008 163.777 411.617 163.947Z"/><defs><linearGradient id="so_g2" gradientUnits="userSpaceOnUse" x1="422.44202" y1="169.54677" x2="386.57278" y2="215.85545"><stop offset="0" stop-color="#FF6201"/><stop offset="1" stop-color="#FF8000"/></linearGradient></defs><path fill="url(#so_g2)" d="M411.617 163.947C412.008 163.777 413.58 162.964 413.785 163.229C417.839 168.446 427.656 174.22 431.385 178.87L384.081 216L383.05 216C383.916 213.56 390.324 202.337 391.838 199.604L411.617 163.947Z"/></svg>
    <h1>Generador de Informes de Implementación</h1>
    <p>Genera tu presentación a partir del chat de WhatsApp</p>
  </div>

  <div class="steps">
    <span>① Sube el ZIP de WhatsApp</span>
    <span>② Selecciona el período</span>
    <span>③ Descarga el PPTX</span>
  </div>

  <form id="upload-form">
    <div class="drop-zone" id="drop-zone">
      <input type="file" id="zip-input" accept=".zip" />
      <div class="icon">📦</div>
      <h3>Arrastra el ZIP de WhatsApp aquí</h3>
      <p>O haz clic para seleccionar el archivo</p>
      <div class="file-name" id="file-name" style="display:none"></div>
    </div>

    <div class="date-row">
      <div class="field">
        <label>Desde</label>
        <input type="date" id="start-date" />
      </div>
      <div class="field">
        <label>Hasta</label>
        <input type="date" id="end-date" />
      </div>
    </div>

    <div class="template-zone" id="template-zone">
      <input type="file" id="template-input" accept=".pptx" />
      <div class="icon">📊</div>
      <h3>Arrastra el Template PPTX aquí</h3>
      <p>O haz clic para seleccionar · Si no subes uno se usará el template por defecto</p>
      <div class="file-name" id="template-name" style="display:none">
        <span id="template-name-text"></span>
        <button type="button" class="template-clear" id="template-clear" onclick="clearTemplate(event)">✕</button>
      </div>
    </div>

    <button class="btn" type="submit" id="submit-btn" disabled>
      Generar presentación
    </button>
  </form>

  <div id="progress-section">
    <div class="progress-bar-wrap">
      <div class="progress-bar" id="progress-bar" style="width:0%"></div>
    </div>
    <div class="progress-label" id="progress-label">Procesando…</div>
  </div>

  <div class="error-box" id="error-box"></div>

  <div id="results-section">
    <div class="results-title" id="results-title"></div>
    <div class="chain-grid" id="chain-grid"></div>
    <a class="btn-download" id="download-btn" href="#" target="_blank">
      ⬇️ Descargar presentación PPTX
    </a>
  </div>
</div>

<footer>Sell-Out · Generador de Informes de Implementación</footer>

<script>
// Default dates: last 14 days
const today = new Date();
const twoWeeksAgo = new Date(today);
twoWeeksAgo.setDate(today.getDate() - 14);
const fmt = d => d.toISOString().split('T')[0];
document.getElementById('end-date').value = fmt(today);
document.getElementById('start-date').value = fmt(twoWeeksAgo);

// Drop zone
const dropZone = document.getElementById('drop-zone');
const zipInput = document.getElementById('zip-input');
const fileNameEl = document.getElementById('file-name');
const submitBtn = document.getElementById('submit-btn');
let selectedFile = null;

// Template
const templateInput = document.getElementById('template-input');
const templateNameEl = document.getElementById('template-name');
const templateNameText = document.getElementById('template-name-text');
const templateZone = document.getElementById('template-zone');
let selectedTemplate = null;

function setTemplate(file) {
  if (!file || !file.name.endsWith('.pptx')) {
    alert('Por favor selecciona un archivo .pptx');
    return;
  }
  selectedTemplate = file;
  templateNameText.textContent = '📄 ' + file.name;
  templateNameEl.style.display = 'inline-flex';
}

function clearTemplate(e) {
  e.stopPropagation();
  selectedTemplate = null;
  templateInput.value = '';
  templateNameEl.style.display = 'none';
  templateNameText.textContent = '';
}

templateInput.addEventListener('change', e => setTemplate(e.target.files[0]));
templateZone.addEventListener('dragover', e => { e.preventDefault(); templateZone.classList.add('drag-over'); });
templateZone.addEventListener('dragleave', () => templateZone.classList.remove('drag-over'));
templateZone.addEventListener('drop', e => {
  e.preventDefault();
  templateZone.classList.remove('drag-over');
  setTemplate(e.dataTransfer.files[0]);
});

function onFileSelected(file) {
  if (!file || !file.name.endsWith('.zip')) {
    alert('Por favor selecciona un archivo .zip');
    return;
  }
  selectedFile = file;
  fileNameEl.textContent = '📁 ' + file.name + ' (' + (file.size / 1024 / 1024).toFixed(1) + ' MB)';
  fileNameEl.style.display = 'inline-block';
  submitBtn.disabled = false;
}

zipInput.addEventListener('change', e => onFileSelected(e.target.files[0]));
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  onFileSelected(e.dataTransfer.files[0]);
});

// Submit
document.getElementById('upload-form').addEventListener('submit', async e => {
  e.preventDefault();
  if (!selectedFile) return;

  const startDate = document.getElementById('start-date').value;
  const endDate   = document.getElementById('end-date').value;

  // Show progress
  document.getElementById('progress-section').style.display = 'block';
  document.getElementById('results-section').style.display = 'none';
  document.getElementById('error-box').style.display = 'none';
  submitBtn.disabled = true;
  setProgress(5, 'Subiendo archivo…');

  // Upload
  const formData = new FormData();
  formData.append('file', selectedFile);
  formData.append('start_date', startDate);
  formData.append('end_date', endDate);
  if (selectedTemplate) formData.append('template', selectedTemplate);

  let jobId;
  try {
    const res = await fetch('/upload', { method: 'POST', body: formData });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Error al subir');
    jobId = data.job_id;
  } catch (err) {
    showError(err.message);
    return;
  }

  // Poll status
  await pollJob(jobId);
});

async function pollJob(jobId) {
  const LABELS = [
    [10, 'Extrayendo archivos del ZIP…'],
    [30, 'Leyendo mensajes del chat…'],
    [50, 'Identificando tiendas por cadena…'],
    [70, 'Generando presentación PowerPoint…'],
    [90, 'Insertando fotos en los slides…'],
  ];
  let labelIdx = 0;
  const interval = setInterval(async () => {
    try {
      const res = await fetch(`/status/${jobId}`);
      const data = await res.json();

      if (data.status === 'processing') {
        if (labelIdx < LABELS.length) {
          setProgress(LABELS[labelIdx][0], LABELS[labelIdx][1]);
          labelIdx++;
        }
      } else if (data.status === 'done') {
        clearInterval(interval);
        setProgress(100, '¡Listo!');
        setTimeout(() => showResults(data.result, jobId), 400);
      } else if (data.status === 'error') {
        clearInterval(interval);
        showError(data.error);
      }
    } catch (err) {
      clearInterval(interval);
      showError('Error de conexión al verificar estado');
    }
  }, 1200);
}

function setProgress(pct, label) {
  document.getElementById('progress-bar').style.width = pct + '%';
  document.getElementById('progress-label').textContent = label;
}

function showError(msg) {
  document.getElementById('progress-section').style.display = 'none';
  const box = document.getElementById('error-box');
  box.textContent = '⚠️ ' + msg;
  box.style.display = 'block';
  submitBtn.disabled = false;
}

function showResults(result, jobId) {
  document.getElementById('progress-section').style.display = 'none';
  const sec = document.getElementById('results-section');
  sec.style.display = 'block';
  submitBtn.disabled = false;

  document.getElementById('results-title').textContent =
    `✅ Presentación generada: ${result.total_slides} slides · ${result.total_stores} tiendas`;

  const grid = document.getElementById('chain-grid');
  grid.innerHTML = '';

  // Merge y renombrar cadenas para el resumen
  const MERGE = {
    'UNIMARC':      'SMU',
    'SANTA ISABEL': 'SISA',
    'HIPER':        'HIPER LIDER',
    'EXPRESS':      'EXPRESS LIDER',
  };
  const merged = {};
  Object.entries(result.summary).forEach(([chain, count]) => {
    const display = MERGE[chain] || chain;
    merged[display] = (merged[display] || 0) + count;
  });

  const ORDER = ['SISA', 'JUMBO', 'HIPER LIDER', 'EXPRESS LIDER', 'TOTTUS', 'SMU'];
  ORDER.forEach(chain => {
    const count = merged[chain];
    if (!count) return;
    const chip = document.createElement('div');
    chip.className = 'chain-chip';
    chip.innerHTML = `<div class="name">${chain}</div>
                      <div class="count">${count}</div>
                      <div class="label">tiendas</div>`;
    grid.appendChild(chip);
  });

  document.getElementById('download-btn').href = `/download/${jobId}`;
}
</script>
</body>
</html>"""


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/health')
def health():
    return jsonify(status='ok')


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify(error='No se recibió ningún archivo'), 400
    f = request.files['file']
    if not f.filename.endswith('.zip'):
        return jsonify(error='Solo se aceptan archivos .zip'), 400

    start_str = request.form.get('start_date', '')
    end_str   = request.form.get('end_date', '')
    try:
        start_date = datetime.strptime(start_str, '%Y-%m-%d')
        end_date   = datetime.strptime(end_str,   '%Y-%m-%d') + timedelta(days=1)
    except ValueError:
        return jsonify(error='Fechas inválidas'), 400

    job_id = str(uuid.uuid4())[:8]
    job_dir = UPLOAD_DIR / job_id
    job_dir.mkdir(parents=True)

    zip_path    = job_dir / 'chat.zip'
    photos_dir  = str(job_dir / 'media')
    output_path = str(job_dir / 'Implementacion_Milka_Easter.pptx')

    f.save(str(zip_path))

    # Template: usa el personalizado si se subió, si no el default
    template_file = request.files.get('template')
    if template_file and template_file.filename.endswith('.pptx'):
        template_path = str(job_dir / 'template.pptx')
        template_file.save(template_path)
    else:
        template_path = str(TEMPLATE_PPTX)

    jobs[job_id] = {'status': 'processing', 'result': None, 'error': None}

    def worker():
        try:
            result = process_zip(
                str(zip_path), photos_dir, start_date, end_date,
                template_path, output_path
            )
            jobs[job_id]['result'] = {
                'total_slides': result['total_slides'],
                'total_stores': len(result['stores']),
                'summary': result['summary'],
            }
            jobs[job_id]['status'] = 'done'
        except Exception as e:
            jobs[job_id]['status'] = 'error'
            jobs[job_id]['error']  = str(e)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify(job_id=job_id)


@app.route('/status/<job_id>')
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify(error='Job no encontrado'), 404
    return jsonify(
        status=job['status'],
        result=job.get('result'),
        error=job.get('error'),
    )


@app.route('/download/<job_id>')
def download(job_id):
    output = UPLOAD_DIR / job_id / 'Implementacion_Milka_Easter.pptx'
    if not output.exists():
        return 'Archivo no encontrado', 404
    return send_file(
        str(output),
        as_attachment=True,
        download_name='Implementacion_Milka_Easter.pptx',
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    print('\n🚀  Sell-Out · Generador de Informes de Implementación')
    print(f'   Abre tu navegador en:  http://localhost:{port}\n')
    app.run(debug=False, host='0.0.0.0', port=port, threaded=True)
