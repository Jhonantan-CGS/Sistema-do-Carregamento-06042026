const safeStorage = {
  getItem(key, fallback = null) {
    try {
      const value = localStorage.getItem(key);
      return value === null ? fallback : value;
    } catch (_) {
      return fallback;
    }
  },
  setItem(key, value) {
    try { localStorage.setItem(key, value); return true; } catch (_) { return false; }
  },
  removeItem(key) {
    try { localStorage.removeItem(key); return true; } catch (_) { return false; }
  }
};

const safeJsonParse = (raw, fallback = {}) => {
  try {
    const parsed = JSON.parse(raw);
    return parsed === null ? fallback : parsed;
  } catch (_) { return fallback; }
};

const reportClientError = (message) => {
  const dbg = globalThis.__debugEngineRef;
  if (dbg && typeof dbg.log === 'function') dbg.log(message, "error");
  else console.error(message);
};

window.onerror = function(msg, url, lineNo) {
  reportClientError(`Erro Linha ${lineNo}: ${msg}`);
  return false;
};
window.onunhandledrejection = function(event) {
  reportClientError(`Promessa Rejeitada: ${event.reason}`);
};

const config = {
  app: {
    version: '20260406.02',
    buildTag: '20260409r07',
    versionStorageKey: 'cysyAppVersion'
  },
  base: {
    empresa: 'Cysy Mineração',
    endereco: 'Rod. SC 100 - Km 26 (Estrada Geral), Bairro Riacho dos Franciscos, Jaguaruna - SC, CEP 88715-000',
    enderecoCurto: 'Rod. SC 100 / Rodovia Claudino Abel Botega, Jaguaruna - SC',
    atividade: 'Extração mineral para corretivos, adubos e fertilizantes',
    cep: '88715-000',
    lat: -28.6131180,
    lon: -48.9518536,
    coordLabel: "28° 36' 47\" S, 48° 57' 07\" O",
    areaPolygon: [
      [-28.6127389, -48.9530531],
      [-28.6134971, -48.9518058],
      [-28.6143777, -48.9500088],
      [-28.6119336, -48.9524710]
    ]
  },
  api: {
    scriptUrl: 'https://script.google.com/macros/s/AKfycbwPCWcAgb90Esm31atFZNAkHf0QYa0Qeg9QVMjj6v4kuzS8yq2NF43xAdo2IKOgUM8Z/exec',
    sheetId: '1bF_9A1P12OAITTcYskNsNRJ8ti_PXjNC3ti7ij0TVw0',
    estoqueSheetId: '1ubMI0M0znO3DTSx4w3dIM8AnQZ5KaJc22ZMyC_oU63c',
    historicoSheetId: '1KAOCFX7_raD0YMhA3yPh8HjS1fF9TlJWp9TYGACCMCI',
    rncSheetId: '1lyWfTaAkrgMWB0JrzbiPt9RIz7jB1WdCe2A9PwnRivY',
    key: safeStorage.getItem('cysyGoogleApiKey', 'AIzaSyAdzZznAPDRlGLqbT-8AAVSKslZYfr7Jc0'),
    range: 'Carregamento!A:Q',
    timeoutMs: 15000
  },
  checklist: [
    {id:'carroceria', lbl:'Baú ou carroceria limpa, livre de resíduos?', type:'yn'},
    {id:'umidade', lbl:'Ausência de umidade que comprometa o produto?', type:'yn'},
    {id:'lona', lbl:'Lona ou cobertura em bom estado?', type:'yn'},
    {id:'estrutura', lbl:'Laterais e piso em boas condições estruturais?', type:'yn'},
    {id:'indicacao', lbl:'Carregado conforme instrução do motorista?', type:'yn'},
    {id:'excesso', lbl:'A carga foi realizada com excesso de peso?', type:'yn_r'},
    {id:'poeira', lbl:'Produto carregado com excesso de poeira?', type:'yn_r'},
    {id:'tamanho', lbl:'O tamanho da carroceria adequado?', type:'yn'},
    {id:'forracao', lbl:'Há necessidade de forração extra?', type:'yn_r'},
    {id:'obs_geral', lbl:'Observações gerais de qualidade:', type:'text'}
  ]
};

const parseBRNumber = (str) => {
  if (typeof str === 'number') return str;
  if (!str) return 0;
  let s = String(str).replace(/[^\d,\.-]/g, '').trim();
  if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
  else if (s.lastIndexOf('.') > s.lastIndexOf(',')) s = s.replace(/,/g, '');
  else if (s.includes(',')) s = s.replace(',', '.');
  return parseFloat(s) || 0;
};

const normalizeName = (s) => String(s || '').toUpperCase().replace(/\s+/g, ' ').trim();
const formatTons = (val) => { const n = parseBRNumber(val); return (isNaN(n) ? 0 : n).toFixed(3); };
const escapeHTML = (str) => String(str || '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&#39;');

const stableHash = (value) => {
  const str = String(value || '');
  let hash = 2166136261;
  for (let i = 0; i < str.length; i++) {
    hash ^= str.charCodeAt(i);
    hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
};

const randomId = (prefix = 'id') => {
  try {
    if (window.crypto?.randomUUID) return `${prefix}_${window.crypto.randomUUID()}`;
  } catch (_) {}
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2, 10)}`;
};

const parseSheetDate = (value) => {
  if (!value && value !== 0) return null;
  if (value instanceof Date && !isNaN(value)) {
    const d = new Date(value);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  if (typeof value === 'number') {
    const tzOffset = new Date().getTimezoneOffset() * 60000;
    const date = new Date((value - 25569) * 86400 * 1000 + tzOffset);
    if (!isNaN(date)) {
      date.setHours(0, 0, 0, 0);
      return date;
    }
  }

  const str = String(value).trim();
  const br = /(\d{1,2})\/(\d{1,2})\/(\d{4})/;
  const iso = /^(\d{4})-(\d{2})-(\d{2})/;
  const brDash = /(\d{1,2})-(\d{1,2})-(\d{4})/;

  let match = str.match(br);
  if (match) {
    const d = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
    d.setHours(0, 0, 0, 0);
    return isNaN(d) ? null : d;
  }

  match = str.match(brDash);
  if (match) {
    const d = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
    d.setHours(0, 0, 0, 0);
    return isNaN(d) ? null : d;
  }

  match = str.match(iso);
  if (match) {
    const d = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    d.setHours(0, 0, 0, 0);
    return isNaN(d) ? null : d;
  }

  const nativeDate = new Date(str);
  if (!isNaN(nativeDate)) {
    nativeDate.setHours(0, 0, 0, 0);
    return nativeDate;
  }
  
  match = str.match(/(\d{1,2})[\s\/](\d{1,2})[\s\/](\d{4})/);
  if (match) {
    const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
    date.setHours(0, 0, 0, 0);
    return isNaN(date) ? null : date;
  }

  return null;
};

const isSameDay = (a, b) => {
  if (!a || !b) return false;
  const da = new Date(a); const db = new Date(b);
  da.setHours(0,0,0,0); db.setHours(0,0,0,0);
  return da.getTime() === db.getTime();
};

const plateRules = {
  emptySentinel: 'SEM PLACA',
  patterns: [
    { key: 'intl', regex: /^[A-Z]{3}[0-9]{3}$/ }, // CAL042
    { key: 'mercosul', regex: /^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$/ }, // ABC1D23
    { key: 'legacy', regex: /^[A-Z]{3}[0-9]{4}$/ }, // ABC1234
    { key: 'fleet', regex: /^[A-Z]{4}[0-9]{3}$/ } // ABCD123
  ],
  normalize(raw = '') {
    return String(raw || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  },
  isEmpty(raw = '') {
    return /\bSEM\W*PLACAS?\b/.test(String(raw || '').toUpperCase());
  },
  isValid(raw = '') {
    const normalized = this.normalize(raw);
    if (!normalized) return false;
    return this.patterns.some(({ regex }) => regex.test(normalized));
  },
  extract(text = '') {
    const source = String(text || '').toUpperCase();
    if (this.isEmpty(source)) return this.emptySentinel;
    const candidates = source.match(/[A-Z0-9-]{6,8}/g) || [];
    for (const candidate of candidates) {
      const normalized = this.normalize(candidate);
      if (this.isValid(normalized)) return normalized;
    }
    return '';
  }
};

const extractPlaca = (obs) => plateRules.extract(obs);

const extrairHorarioObs = (obs) => {
  const text = String(obs || '');
  const match = text.match(/\b([01]?\d|2[0-3]):([0-5]\d)\b/);
  return match ? `${match[1].padStart(2, '0')}:${match[2]}` : '';
};

const debugEngine = {
  logs: [],
  isOpen: false,
  maxLogs: 300,
  verboseMode: false,
  formatLogHtml(entry) {
    const clr = entry.type === 'error' ? '#EF4444' : entry.type === 'success' ? '#10B981' : entry.type === 'warn' ? '#F59E0B' : '#CBD5E1';
    return `<div style="padding-bottom:6px; border-bottom:1px solid rgba(255,255,255,0.1);"><span style="color:#94A3B8;">[${entry.time}]</span> <span style="color:${clr}">${escapeHTML(entry.msg)}</span></div>`;
  },
  renderLogs() {
    const consoleDiv = document.getElementById('dbgLogConsole');
    if (!consoleDiv) return;
    consoleDiv.innerHTML = this.logs.map((entry) => this.formatLogHtml(entry)).join('');
    consoleDiv.scrollTop = consoleDiv.scrollHeight;
  },
  log(msg, type = 'info') {
    if (type === 'info' && !this.verboseMode && !this.isOpen) return;
    const time = new Date().toLocaleTimeString();
    this.logs.push({ time, msg, type });
    if (this.logs.length > this.maxLogs) this.logs = this.logs.slice(this.logs.length - this.maxLogs);
    if (this.isOpen) this.renderLogs();
  },
  clearLogs() {
    this.logs = [];
    const consoleDiv = document.getElementById('dbgLogConsole');
    if (consoleDiv) consoleDiv.innerHTML = '';
  },
  updateAPIStatus(connected, detail = '') {
    const el = document.getElementById('dbgStatusAPI');
    if (el) el.textContent = connected ? `🟢 API Conectada ${detail}` : `🔴 API Falha ${detail}`;
  },
  updateNetStatus(online) {
    const el = document.getElementById('dbgStatusNet');
    if (el) el.textContent = online ? '🟢 Online' : '🔴 Offline';
  },
  updateSWStatus(active) {
    const el = document.getElementById('dbgStatusSW');
    if (el) el.textContent = active ? '🟢 SW Ativo' : '⚪ SW Inativo';
  },
  togglePanel() {
    this.isOpen = !this.isOpen;
    const p = document.getElementById('debugPanel');
    if (p) p.style.right = this.isOpen ? '0' : '-300px';
    if (this.isOpen) this.renderLogs();
  },
  copiarLogs() {
    const texto = this.logs.map(l => `[${l.time}] [${l.type.toUpperCase()}] ${l.msg}`).join('\n');
    navigator.clipboard.writeText(texto).then(() => {
      alert("Logs copiados! Envie para o desenvolvedor.");
    }).catch(() => {
      alert("Não foi possível copiar automaticamente. Copie manualmente no painel.");
    });
  },
  async runFullTest() {
    this.log("=== INICIANDO TESTE COMPLETO ===", "info");
    
    this.log("--- Teste 1: Verificando LocalStorage ---", "info");
    const user = safeStorage.getItem('cysyUser', '');
    this.log(`Usuário logado: ${user || 'NENHUM'}`, user ? "success" : "warn");
    
    this.log("--- Teste 2: Testando parseSheetDate ---", "info");
    const testDates = [
      "20/03/2026",
      "2026-03-20",
      "20-03-2026",
      new Date(),
      45678,
      "20 de Março de 2026",
      "03/20/2026"
    ];
    testDates.forEach(d => {
      const result = parseSheetDate(d);
      this.log(`parseSheetDate("${String(d).substring(0,20)}") => ${result ? result.toLocaleDateString('pt-BR') : 'NULL'}`, result ? "success" : "error");
    });
    
    this.log("--- Teste 3: Verificando dados do parse ---", "info");
    if (appController.lastParsedData && appController.lastParsedData.length > 0) {
      this.log(`Dados parseados: ${appController.lastParsedData.length} registros`, "success");
      const sample = appController.lastParsedData[0];
      this.log(`Amostra - Cliente: ${sample.cliente}`, "info");
      this.log(`Amostra - dataRaw: "${sample.dataRaw}"`, "info");
      this.log(`Amostra - parseSheetDate: ${parseSheetDate(sample.dataRaw) ? parseSheetDate(sample.dataRaw).toLocaleDateString('pt-BR') : 'NULL'}`, "info");
    } else {
      this.log("Nenhum dado parseado encontrado!", "warn");
    }
    
    this.log("--- Teste 4: Verificando estoque ---", "info");
    if (uiBuilder.lotesFlat && uiBuilder.lotesFlat.length > 0) {
      this.log(`Lotes carregados: ${uiBuilder.lotesFlat.length}`, "success");
    } else {
      this.log("Nenhum lote encontrado!", "warn");
    }
    
    this.log("--- Teste 5: Navegador ---", "info");
    this.log(`User Agent: ${navigator.userAgent.substring(0, 50)}...`, "info");
    this.log(`Online: ${navigator.onLine}`, "info");
    this.log(`Protocol: ${window.location.protocol}`, "info");
    
    this.log("=== TESTE COMPLETO FINALIZADO ===", "success");
  }
};

globalThis.__debugEngineRef = debugEngine;

const soundManager = {
  ctx: null,
  enabled: true,
  init() {
    try {
      this.ctx = new (window.AudioContext || window.webkitAudioContext)();
      this.enabled = safeStorage.getItem('cysySoundEnabled', 'true') !== 'false';
      const btn = document.getElementById('soundToggleBtn');
      if (btn) btn.textContent = this.enabled ? 'Som: ON' : 'Som: OFF';
    } catch(e) { this.enabled = false; }
  },
  resume() {
    if (this.ctx && this.ctx.state === 'suspended') this.ctx.resume();
  },
  play(type) {
    if (!this.enabled || !this.ctx) return;
    this.resume();
    const ctx = this.ctx;
    const now = ctx.currentTime;
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    
    const sounds = {
      click: { freq: 800, dur: 0.05, type: 'sine', vol: 0.1 },
      tab: { freqs: [460, 620], dur: 0.09, type: 'triangle', vol: 0.075 },
      success: { freq: 880, dur: 0.15, type: 'sine', vol: 0.15 },
      error: { freq: 220, dur: 0.3, type: 'sawtooth', vol: 0.1 },
      notify: { freq: 600, dur: 0.1, type: 'triangle', vol: 0.12 },
      swoosh: { freqs: [320, 480, 720], dur: 0.08, type: 'triangle', vol: 0.055 },
      liberar: { freqs: [523, 659, 784], dur: 0.12, type: 'sine', vol: 0.12 },
      alert: { freqs: [400, 300], dur: 0.15, type: 'square', vol: 0.08 },
      criticalAlert: { freqs: [880, 660, 880, 520], dur: 0.14, type: 'triangle', vol: 0.11 }
    };
    
    const s = sounds[type] || sounds.click;
    if (s.freqs) {
      s.freqs.forEach((f, i) => {
        const o = ctx.createOscillator();
        const g = ctx.createGain();
        o.connect(g);
        g.connect(ctx.destination);
        o.type = s.type;
        o.frequency.value = f;
        g.gain.setValueAtTime(s.vol, now + i * 0.08);
        g.gain.exponentialRampToValueAtTime(0.001, now + i * 0.08 + s.dur);
        o.start(now + i * 0.08);
        o.stop(now + i * 0.08 + s.dur + 0.01);
      });
    } else {
      osc.type = s.type;
      osc.frequency.value = s.freq;
      gain.gain.setValueAtTime(s.vol, now);
      gain.gain.exponentialRampToValueAtTime(0.001, now + s.dur);
      osc.start(now);
      osc.stop(now + s.dur + 0.01);
    }
  },
  toggle() {
    this.enabled = !this.enabled;
    safeStorage.setItem('cysySoundEnabled', String(this.enabled));
    return this.enabled;
  }
};

const toastManager = {
  container: null,
  init() {
    if (this.container) return;
    this.container = document.createElement('div');
    this.container.id = 'toastContainer';
    this.container.style.cssText = 'position:fixed;bottom:88px;left:50%;transform:translateX(-50%);z-index:100000;display:flex;flex-direction:column;align-items:center;gap:10px;pointer-events:none;max-width:min(92vw,460px);width:100%;padding:0 12px;';
    document.body.appendChild(this.container);
  },
  show(msg, type = 'info', duration = 3500) {
    this.init();
    const icons = { success: '✓', error: '✕', warning: '⚠', info: 'ℹ' };
    const colors = {
      success: { bg: '#ECFDF5', border: '#10B981', text: '#047857', iconBg: '#D1FAE5' },
      error: { bg: '#FEF2F2', border: '#EF4444', text: '#B91C1C', iconBg: '#FEE2E2' },
      warning: { bg: '#FFFBEB', border: '#F59E0B', text: '#B45309', iconBg: '#FEF3C7' },
      info: { bg: '#EFF6FF', border: '#3B82F6', text: '#1D4ED8', iconBg: '#DBEAFE' }
    };
    const c = colors[type] || colors.info;
    const toast = document.createElement('div');
    toast.style.cssText = `background:${c.bg};border:1px solid ${c.border};border-left:5px solid ${c.border};border-radius:14px;padding:12px 14px;display:flex;align-items:flex-start;gap:12px;animation:slideUp 0.2s ease;box-shadow:0 10px 22px rgba(15,23,42,0.12);pointer-events:auto;width:100%;`;
    toast.innerHTML = `<span style="flex:0 0 auto;width:32px;height:32px;border-radius:10px;background:${c.iconBg};display:inline-flex;align-items:center;justify-content:center;font-size:17px;color:${c.text};font-weight:900;">${icons[type] || '•'}</span><span style="color:#0F172A;font-size:14px;line-height:1.45;font-weight:800;">${escapeHTML(msg)}</span>`;
    this.container.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(8px)';
      toast.style.transition = 'opacity 0.2s ease, transform 0.2s ease';
      setTimeout(() => toast.remove(), 220);
    }, duration);
  }
};

const envManager = {
  init() {
    setInterval(() => {
      const now = new Date();
      const cw = document.getElementById('clockWidget');
      if (cw) cw.innerText = now.toLocaleTimeString('pt-BR', {hour:'2-digit', minute:'2-digit', second:'2-digit'});
    }, 1000);
    this.fetchWeather();
    setInterval(() => this.fetchWeather(), 3600000);
  },
  async fetchWeather() {
    try {
      const url = `https://api.open-meteo.com/v1/forecast?latitude=${encodeURIComponent(config.base.lat)}&longitude=${encodeURIComponent(config.base.lon)}&current=temperature_2m,precipitation,wind_speed_10m&daily=temperature_2m_max,precipitation_probability_max,wind_speed_10m_max,uv_index_max&timezone=America%2FSao_Paulo`;
      const res = await fetch(url);
      const data = await res.json();

      const cur = data.current;
      const daily = data.daily;
      const tempNow = Math.round(cur.temperature_2m);
      const windNow = Math.round(cur.wind_speed_10m);
      const tempTom = Math.round(daily.temperature_2m_max[1]);
      const rainTom = daily.precipitation_probability_max[1];
      const uvTom = daily.uv_index_max[1] || '--';

      const ids = [
        ['wTempNow', `${tempNow}°C`],
        ['wWindNow', `${windNow} km/h`],
        ['wTempTom', `${tempTom}°C`],
        ['wRainTom', `${rainTom}%`],
        ['wUvTom', uvTom]
      ];
      ids.forEach(([id, value]) => {
        const el = document.getElementById(id);
        if (el) el.innerText = value;
      });

      const alertBanner = document.getElementById('weatherAlertBanner');
      if (!alertBanner) return;
      const todayKey = new Date().toISOString().slice(0, 10);
      const canRegisterPriority = Boolean(safeStorage.getItem('cysyUser', '').trim()) && window.priorityAlertManager;

      if (windNow > 25 || daily.wind_speed_10m_max[0] > 35 || daily.wind_speed_10m_max[1] > 35) {
        alertBanner.style.background = 'rgba(220, 38, 38, 0.2)';
        alertBanner.style.color = '#FCA5A5';
        alertBanner.innerHTML = '⚠️ VENTO FORTE: Risco de Lonas Voando!';
        if (canRegisterPriority) {
          window.priorityAlertManager.registerAlert({
            id: `weather_wind_${todayKey}`,
            title: '⚠️ ALERTA DE VENTO FORTE',
            body: 'Reforce lonas e proteções. Há risco operacional com vento forte na área da Cysy.',
            severity: 'danger',
            actionKey: 'tab_op',
            actionLabel: '📋 Abrir liberação',
            source: 'clima'
          });
        }
      } else if (cur.precipitation > 0 || daily.precipitation_probability_max[0] > 60 || daily.precipitation_probability_max[1] > 60) {
        alertBanner.style.background = 'rgba(220, 38, 38, 0.2)';
        alertBanner.style.color = '#FCA5A5';
        alertBanner.innerHTML = '🌧️ ALERTA DE CHUVA: Proteja os materiais expostos.';
        if (canRegisterPriority) {
          window.priorityAlertManager.registerAlert({
            id: `weather_rain_${todayKey}`,
            title: '🌧️ ALERTA DE CHUVA',
            body: 'Proteja os materiais expostos e confirme a revisão operacional da área.',
            severity: 'danger',
            actionKey: 'tab_op',
            actionLabel: '📋 Abrir liberação',
            source: 'clima'
          });
        }
      } else if (tempNow > 32 || daily.temperature_2m_max[0] > 33 || daily.temperature_2m_max[1] > 33 || Number(uvTom) >= 8) {
        alertBanner.style.background = 'rgba(220, 38, 38, 0.2)';
        alertBanner.style.color = '#FCA5A5';
        alertBanner.innerHTML = '☀️ SOL EXTREMO: Atenção aos materiais sensíveis.';
        if (canRegisterPriority) {
          window.priorityAlertManager.registerAlert({
            id: `weather_sun_${todayKey}`,
            title: '☀️ ALERTA DE SOL EXTREMO',
            body: 'Monitore os materiais sensíveis. O risco de ressecamento e aquecimento está elevado.',
            severity: 'warning',
            actionKey: 'tab_op',
            actionLabel: '📋 Abrir liberação',
            source: 'clima'
          });
        }
      } else {
        alertBanner.style.background = 'rgba(16, 185, 129, 0.2)';
        alertBanner.style.color = '#34D399';
        alertBanner.innerHTML = '✅ Clima atual favorável para a operação.';
      }
    } catch(e) {}
  }
};

const wakeLockManager = {
  prefKey: 'cysyWakeLockPreferenceV1',
  supportToastKey: 'cysyWakeLockSupportToastV1',
  manualHintKey: 'cysyWakeLockManualHintV1',
  lock: null,
  initialized: false,
  requestedFromFirstAccess: false,

  isSupported() {
    return Boolean(navigator.wakeLock && typeof navigator.wakeLock.request === 'function');
  },

  isEnabled() {
    return safeStorage.getItem(this.prefKey, 'true') !== 'false';
  },

  async init() {
    if (this.initialized) return;
    this.initialized = true;

    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') {
        this.ensureActive(false);
      }
    });
    window.addEventListener('focus', () => this.ensureActive(false));
    window.addEventListener('pageshow', () => this.ensureActive(false));
  },

  getManualHintMessage() {
    const ua = navigator.userAgent || '';
    if (/Android/i.test(ua)) {
      return 'Para máxima estabilidade, remova a otimização de bateria do navegador/PWA nas configurações do Android.';
    }
    return 'Se o dispositivo continuar suspendendo a tela, desative restrições de energia para o navegador/PWA.';
  },

  async enableOnFirstAccess(fromUserGesture = false) {
    if (!this.requestedFromFirstAccess) {
      this.requestedFromFirstAccess = true;
      safeStorage.setItem(this.prefKey, 'true');
    }
    return this.ensureActive(fromUserGesture);
  },

  async ensureActive(fromUserGesture = false) {
    if (!this.isEnabled()) return false;
    if (!this.isSupported()) {
      if (fromUserGesture && safeStorage.getItem(this.supportToastKey, 'false') !== 'true') {
        safeStorage.setItem(this.supportToastKey, 'true');
        toastManager.show('Este navegador não suporta manter a tela ativa automaticamente.', 'warning', 5200);
      }
      return false;
    }
    if (document.visibilityState === 'hidden') return false;
    if (this.lock && !this.lock.released) return true;

    try {
      this.lock = await navigator.wakeLock.request('screen');
      this.lock.addEventListener('release', () => {
        this.lock = null;
        if (document.visibilityState === 'visible' && this.isEnabled()) {
          setTimeout(() => this.ensureActive(false), 300);
        }
      });
      if (fromUserGesture) {
        toastManager.show('Tela ativa protegida para operação em campo.', 'success', 3800);
      }
      return true;
    } catch (err) {
      debugEngine.log(`Wake Lock indisponível: ${err.message}`, 'warn');
      if (fromUserGesture && safeStorage.getItem(this.manualHintKey, 'false') !== 'true') {
        safeStorage.setItem(this.manualHintKey, 'true');
        toastManager.show(this.getManualHintMessage(), 'warning', 6500);
      }
      return false;
    }
  },

  async release() {
    try {
      await this.lock?.release?.();
    } catch (_) {}
    this.lock = null;
  }
};

const permissionManager = {
  requestKey: 'cysyPermissionBootstrapV1',
  async requestAll(fromUserGesture = false) {
    const alreadyRequested = safeStorage.getItem(this.requestKey, 'false') === 'true';
    if (alreadyRequested && !fromUserGesture) return;
    const resumo = { geo: 'nao_suportado', notif: 'nao_suportado', camera: 'nao_suportado' };

    if ("Notification" in window) {
      try {
        if (Notification.permission === 'granted') resumo.notif = 'granted';
        else if (Notification.permission === 'denied') resumo.notif = 'denied';
        else resumo.notif = fromUserGesture ? await Notification.requestPermission() : 'pendente_gesto';
      } catch (_) { resumo.notif = 'erro'; }
    }

    if (navigator.geolocation) {
      if (fromUserGesture) {
        resumo.geo = await new Promise((resolve) => {
          navigator.geolocation.getCurrentPosition(
            (pos) => {
              safeStorage.setItem('cysyLastDeviceGeo', JSON.stringify({
                lat: Number(pos.coords.latitude.toFixed(6)),
                lon: Number(pos.coords.longitude.toFixed(6)),
                accuracy: Math.round(pos.coords.accuracy || 0),
                capturedAt: new Date().toISOString()
              }));
              resolve('granted');
            },
            (err) => {
              if (err && err.code === 1) resolve('denied');
              else resolve('erro');
            },
            { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
          );
        });
      } else {
        resumo.geo = safeStorage.getItem('cysyLastDeviceGeo', '') ? 'cache_local' : 'pendente_gesto';
      }
    }

    if (navigator.mediaDevices?.getUserMedia && window.isSecureContext) {
      try {
        if (fromUserGesture) {
          const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' }, audio: false });
          stream.getTracks().forEach(t => t.stop());
          resumo.camera = 'granted';
        } else {
          resumo.camera = 'pendente_gesto';
        }
      } catch (_) { resumo.camera = 'denied'; }
    }

    if (navigator.storage?.persist) {
      try { await navigator.storage.persist(); } catch (_) {}
    }

    try {
      await wakeLockManager.enableOnFirstAccess(fromUserGesture);
    } catch (_) {}

    safeStorage.setItem(this.requestKey, 'true');
    safeStorage.setItem('cysyPermissionLastSummary', JSON.stringify({ ...resumo, at: new Date().toISOString() }));
    backupManager.addEntry('PERMISSOES_INICIAIS', resumo);
  },

  isSafari() {
    const ua = navigator.userAgent || '';
    return /Safari/i.test(ua) && !/Chrome|CriOS|Chromium|Edg|OPR/i.test(ua);
  },

  getGpsErrorMessage(err) {
    const secure = window.isSecureContext && location.protocol === 'https:';
    const msgBase = "Erro ao obter GPS.";
    if (!secure) {
      return `${msgBase} Abra em HTTPS para habilitar localização precisa.`;
    }
    if (!err) {
      return `${msgBase} Autorize a localização no navegador/PWA e tente novamente.`;
    }
    if (err.code === 1) {
      if (this.isSafari()) {
        return `${msgBase} No Safari/iOS, ative Localização para este site (Ajustes > Safari > Localização) e permita também no modo PWA.`;
      }
      return `${msgBase} Permissão de localização negada. Libere GPS para este site e tente novamente.`;
    }
    if (err.code === 2) return `${msgBase} Sinal indisponível. Vá para área aberta e tente novamente.`;
    if (err.code === 3) return `${msgBase} Tempo de resposta excedido. Verifique GPS e conectividade.`;
    return `${msgBase} Verifique permissões do navegador/PWA e se o acesso está em HTTPS.`;
  }
};
