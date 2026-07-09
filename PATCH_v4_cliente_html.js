// ============================================================
//  PATCH CLIENTE v4.0 — Fila durável + envio em LOTE
//  ------------------------------------------------------------
//  Cole este bloco no HTML SUBSTITUINDO as funções:
//    postRequest, enviarParaSheets, reenvioAutomaticoSilencioso
//  e ADICIONE o handler `beforeunload` no final.
// ============================================================

const LOTE_MAX      = 25;     // scans por request
const LOTE_INTERVAL = 1500;   // ms — janela de acumulação
const BACKOFF_BASE  = 1500;   // ms
const BACKOFF_MAX   = 30000;  // ms
const LS_FILA       = 'freq_fila_pendente_v4';

let filaEnvio      = [];      // registros aguardando envio
let filaTimer      = null;
let tentativas     = 0;
let enviando       = false;

// Carrega fila persistida ao iniciar
try {
  const raw = localStorage.getItem(LS_FILA);
  if (raw) filaEnvio = JSON.parse(raw) || [];
} catch (_) { filaEnvio = []; }

function persistirFila() {
  try { localStorage.setItem(LS_FILA, JSON.stringify(filaEnvio)); } catch (_) {}
}

// Substitui enviarParaSheets(registro) — agora só ENFILEIRA.
async function enviarParaSheets(registro) {
  registro.syncStatus = 'pendente';
  salvarRegistrosLocal();
  renderizarLog();

  // Deduplicação client-side por matricula+data
  const chave = registro.matricula + '|' + registro.data;
  if (!filaEnvio.some(r => (r.matricula + '|' + r.data) === chave)) {
    filaEnvio.push(registro);
    persistirFila();
  }
  agendarFlushLote();
}

function agendarFlushLote() {
  if (filaTimer) return;
  // Se acumular o suficiente, envia já; senão espera a janela.
  if (filaEnvio.length >= LOTE_MAX) return flushLote();
  filaTimer = setTimeout(() => { filaTimer = null; flushLote(); }, LOTE_INTERVAL);
}

async function flushLote() {
  if (enviando) return;
  if (filaTimer) { clearTimeout(filaTimer); filaTimer = null; }
  if (filaEnvio.length === 0) return;

  const url = getUrl();
  if (!urlValida(url)) return;

  const lote = filaEnvio.slice(0, LOTE_MAX);
  enviando = true;
  setSyncStatus('syncing');

  try {
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'registrarPresencaLote', registros: lote }),
    });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const data = await resp.json();
    if (!data || data.status !== 'ok') throw new Error(data && data.message || 'Resposta inválida');

    // Marca cada registro do lote como sincronizado
    const chaves = new Set(lote.map(r => r.matricula + '|' + r.data));
    for (const r of registrosHoje) {
      if (chaves.has(r.matricula + '|' + r.data)) r.syncStatus = 'ok';
    }
    filaEnvio = filaEnvio.slice(lote.length);
    tentativas = 0;
    persistirFila();
    salvarRegistrosLocal();
    renderizarLog();
    setSyncStatus('synced');
  } catch (err) {
    console.warn('Falha lote:', err.message);
    tentativas++;
    setSyncStatus('fail');
    // Backoff exponencial com jitter
    const delay = Math.min(BACKOFF_BASE * Math.pow(2, tentativas - 1), BACKOFF_MAX) + Math.random() * 800;
    setTimeout(() => { enviando = false; if (filaEnvio.length) flushLote(); }, delay);
    return;
  } finally {
    enviando = false;
  }

  // Se ainda sobrou fila, agenda novo ciclo
  if (filaEnvio.length > 0) agendarFlushLote();
}

// Reenvio automático — apenas garante que a fila é processada.
async function reenvioAutomaticoSilencioso() {
  const pendentes = registrosHoje.filter(r => r.syncStatus !== 'ok');
  for (const r of pendentes) {
    const k = r.matricula + '|' + r.data;
    if (!filaEnvio.some(x => (x.matricula + '|' + x.data) === k)) {
      filaEnvio.push(r);
    }
  }
  persistirFila();
  if (filaEnvio.length) agendarFlushLote();
}
setInterval(reenvioAutomaticoSilencioso, 8000);

// Ao fechar/atualizar aba: tentar enviar o que restar via sendBeacon
// (não bloqueia navegação e sobrevive ao unload).
window.addEventListener('beforeunload', () => {
  if (!filaEnvio.length) return;
  const url = getUrl();
  if (!urlValida(url)) return;
  const blob = new Blob(
    [JSON.stringify({ action: 'registrarPresencaLote', registros: filaEnvio.slice(0, 50) })],
    { type: 'text/plain;charset=utf-8' }
  );
  navigator.sendBeacon(url, blob);
});
