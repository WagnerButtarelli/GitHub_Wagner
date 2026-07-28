// ===== CONFIG =====
const GAS_URL = 'https://script.google.com/macros/s/AKfycbwHPhiemtPgDMhRBl8TekWNhoM7ATvXdl5iLxZ0qDQodsAx6IytV3E8DhpeS3MAawuV/exec';
const DB_NAME = 'coletaCampoDB';
const STORE = 'registros';

// ===== REGISTRO DO SERVICE WORKER =====
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').then(reg => {
    navigator.serviceWorker.addEventListener('message', ev => {
      if (ev.data && ev.data.tipo === 'TRIGGER_SYNC') sincronizar();
    });
  });
}

// ===== INDEXEDDB =====
function abrirDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: 'id' });
    };
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = e => reject(e.target.error);
  });
}
async function salvarRegistro(reg) {
  const db = await abrirDB();
  return new Promise((res, rej) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(reg);
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
}
async function listarRegistros() {
  const db = await abrirDB();
  return new Promise((res, rej) => {
    const tx = db.transaction(STORE, 'readonly');
    const req = tx.objectStore(STORE).getAll();
    req.onsuccess = () => res(req.result);
    req.onerror = () => rej(req.error);
  });
}

// ===== CAPTURA (foto + GPS) - funciona 100% offline =====
document.getElementById('btnFoto').addEventListener('click', () => document.getElementById('inputFoto').click());

document.getElementById('inputFoto').addEventListener('change', e => {
  const arquivo = e.target.files[0];
  if (!arquivo) return;

  navigator.geolocation.getCurrentPosition(pos => {
    const registro = {
      id: 'reg_' + Date.now(),
      blob: arquivo,                 // guarda o Blob direto, sem base64
      mimeType: arquivo.type,
      nomeArquivo: 'foto_' + Date.now() + '.jpg',
      lat: pos.coords.latitude,
      lng: pos.coords.longitude,
      timestamp: new Date().toISOString(),
      sincronizado: false,
      tentativas: 0
    };
    salvarRegistro(registro).then(() => {
      renderizarLista();
      if (navigator.onLine) sincronizar();
    });
  }, err => alert('GPS indisponível: ' + err.message), { enableHighAccuracy: true, timeout: 15000 });
});

// ===== SYNC =====
function blobParaBase64(blob) {
  return new Promise((res, rej) => {
    const leitor = new FileReader();
    leitor.onload = () => res(leitor.result.split(',')[1]);
    leitor.onerror = rej;
    leitor.readAsDataURL(blob);
  });
}

async function sincronizar() {
  if (!navigator.onLine || GAS_URL.includes('COLE_AQUI')) return;
  const registros = await listarRegistros();
  const pendentes = registros.filter(r => !r.sincronizado);

  for (const r of pendentes) {
    try {
      const base64 = await blobParaBase64(r.blob);
      const payload = {
        id: r.id, foto: base64, mimeType: r.mimeType, nomeArquivo: r.nomeArquivo,
        lat: r.lat, lng: r.lng, timestamp: r.timestamp
      };
      // text/plain evita preflight CORS (GAS não trata bem OPTIONS)
      const resp = await fetch(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      });
      const json = await resp.json();
      if (json.sucesso) {
        r.sincronizado = true;
        await salvarRegistro(r);
      } else {
        r.tentativas++; await salvarRegistro(r);
      }
    } catch (e) {
      r.tentativas++; await salvarRegistro(r); // sem rede de verdade ou instável -> tenta de novo depois
    }
  }
  renderizarLista();
}

// Tenta registrar Background Sync (bônus, best-effort)
async function registrarBackgroundSync() {
  if ('serviceWorker' in navigator && 'SyncManager' in window) {
    try {
      const reg = await navigator.serviceWorker.ready;
      await reg.sync.register('sync-registros');
    } catch (e) { /* iOS/Safari cai aqui - ignora */ }
  }
}

document.getElementById('btnSync').addEventListener('click', sincronizar);
window.addEventListener('online', () => { atualizarStatus(); sincronizar(); registrarBackgroundSync(); });
window.addEventListener('offline', atualizarStatus);
setInterval(() => { if (navigator.onLine) sincronizar(); }, 30000); // rede de campo (3G instável) - reforço além do evento 'online'

// ===== UI =====
function atualizarStatus() {
  const el = document.getElementById('status');
  el.textContent = navigator.onLine ? '🟢 Online' : '🔴 Offline - salvando localmente';
  el.className = navigator.onLine ? 'online' : 'offline';
}

async function renderizarLista() {
  const registros = await listarRegistros();
  const pendentes = registros.filter(r => !r.sincronizado);
  document.getElementById('contador').textContent = pendentes.length;

  document.getElementById('lista').innerHTML = registros.slice().reverse().map(r => {
    const classe = r.sincronizado ? 'sincronizado' : (r.tentativas > 3 ? 'erro' : 'pendente');
    const icone = r.sincronizado ? '✅' : (r.tentativas > 3 ? '⚠️' : '⏳');
    return `<div class="item ${classe}">${icone} ${r.nomeArquivo}<br>
      Lat: ${r.lat.toFixed(6)} | Lng: ${r.lng.toFixed(6)}<br>
      ${new Date(r.timestamp).toLocaleString('pt-BR')}${r.tentativas > 3 ? ' - ' + r.tentativas + ' tentativas' : ''}</div>`;
  }).join('');
}

atualizarStatus();
renderizarLista();
registrarBackgroundSync();
