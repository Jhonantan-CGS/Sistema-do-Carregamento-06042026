const loginManager = {
  isLoggedIn() {
    return Boolean(safeStorage.getItem('cysyUser', '').trim());
  },
  init() {
    const user = safeStorage.getItem('cysyUser', '');
    if (!user) document.getElementById('loginOverlay').style.display = 'flex';
    else this.setFields(user);
  },
  async entrar() {
    soundManager.play('click');
    const name = document.getElementById('loginNameInput').value.trim();
    if (name.split(/\s+/).length < 2) {
      soundManager.play('error');
      toastManager.show('Digite Nome e Sobrenome para rastreabilidade.', 'warning');
      return;
    }
    safeStorage.setItem('cysyUser', name);
    backupManager.addEntry('LOGIN', { usuario: name, dataHora: new Date().toLocaleString('pt-BR') });
    soundManager.play('success');
    toastManager.show(`Bem-vindo(a), ${name.split(' ')[0]}!`, 'success');
    document.getElementById('loginOverlay').style.display = 'none';
    this.setFields(name);

    await permissionManager.requestAll(true).catch(() => {});
    try { priorityAlertManager.init(); } catch (_) {}
    try { await envManager.fetchWeather(); } catch (_) {}
    setTimeout(() => {
      try { priorityAlertManager.remindActiveAlerts({ force: true }); } catch (_) {}
    }, 1200);
  },
  logout() {
    if (!confirm("Deseja realmente sair? Dados não salvos serão perdidos.")) return;
    soundManager.play('click');
    backupManager.addEntry('LOGOUT', { usuario: safeStorage.getItem('cysyUser', ''), dataHora: new Date().toLocaleString('pt-BR') });
    try { wakeLockManager.release(); } catch (_) {}
    safeStorage.removeItem('cysyUser');
    safeStorage.removeItem('cysyFaturamentoConfirmacoes');
    safeStorage.removeItem('cysyFaturamentoEvidencias');
    location.reload();
  },
  setFields(name) {
    const op = document.getElementById('selAcomp');
    const perda = document.getElementById('perdaAcomp');
    const userNameEl = document.getElementById('userLoggedName');
    if (op) op.value = name;
    if (perda) perda.value = name;
    if (userNameEl) {
      userNameEl.innerText = `👤 ${name}`;
      userNameEl.style.display = 'block';
    }
  },
  isSuperAdmin() {
    const user = safeStorage.getItem('cysyUser', '').toLowerCase().trim();
    return user === 'jhonantan.goncalves';
  }
};

const dbManager = {
  dbName: 'CysyDBv8',
  storeName: 'uploads',
  sentIndexKey: 'cysySyncSentIndexV2',
  maxOfflineOperacoes: 10,
  db: null,
  memoryFallback: [],
  sentIndex: {},
  loadSentIndex() {
    this.sentIndex = safeJsonParse(safeStorage.getItem(this.sentIndexKey, '{}'), {});
    if (!this.sentIndex || typeof this.sentIndex !== 'object') this.sentIndex = {};
  },
  persistSentIndex() {
    const entries = Object.entries(this.sentIndex || {}).sort((a, b) => {
      const ad = new Date(a[1]?.sentAt || 0).getTime();
      const bd = new Date(b[1]?.sentAt || 0).getTime();
      return bd - ad;
    }).slice(0, 12000);
    this.sentIndex = Object.fromEntries(entries);
    safeStorage.setItem(this.sentIndexKey, JSON.stringify(this.sentIndex));
  },
  countPayloadImages(payload = {}) {
    const fotos = payload?.fotos || {};
    let total = 0;
    ['frente', 'traseira', 'assoalho', 'imagemCargaPronta'].forEach((k) => {
      const val = fotos[k] || payload[k] || '';
      if (typeof val === 'string' && val.startsWith('data:image/')) total++;
    });
    return total;
  },
  buildFingerprint(type, payload = {}) {
    const normalizedType = String(type || payload.tipo || '').toUpperCase();

    // Baixa de perdas precisa considerar os itens informados; antes
    // gerávamos sempre o mesmo hash, fazendo qualquer nova baixa ser
    // tratada como duplicada. Incluímos acompanhandte e cada linha de perda
    // na assinatura para permitir envios distintos.
    if (normalizedType === 'BAIXA_PERDA') {
      const perdasSig = Array.isArray(payload.perdas)
        ? payload.perdas.map((p) => ({
            p: normalizeName(p.produto || ''),
            l: String(p.lote || '').toUpperCase(),
            q: Number(p.qtd || 0),
            t: normalizeName(p.tipo || ''),
            o: String(p.obs || '').trim()
          }))
        : [];
      return stableHash(
        JSON.stringify({
          tipo: normalizedType,
          acomp: normalizeName(payload.acomp || ''),
          perdas: perdasSig
        })
      );
    }

    const base = {
      tipo: normalizedType,
      placa: normalizeName(payload.placa || ''),
      pedido: normalizeName(payload.pedido || ''),
      cliente: normalizeName(payload.cliente || ''),
      status: normalizeName(payload.status || ''),
      lotePrincipal: normalizeName(payload.lote || ''),
      linhas: Array.isArray(payload.linhasOperacao) ? payload.linhasOperacao.map((l) => ({
        p: normalizeName(l.produto || ''),
        q: Number(l.quantidade || 0),
        lo: normalizeName(l.lote || ''),
        pe: Number(l.perda || 0)
      })) : [],
      checklist: payload.checklist || {},
      fotosMeta: {
        frente: Boolean(payload?.fotos?.frente),
        traseira: Boolean(payload?.fotos?.traseira),
        assoalho: Boolean(payload?.fotos?.assoalho),
        imgCarga: Boolean(payload.imagemCargaPronta)
      }
    };
    return stableHash(JSON.stringify(base));
  },
  buildMeta(type, payload = {}, options = {}) {
    const nowIso = new Date().toISOString();
    const localId = options.localId || payload.localId || randomId('loc');
    const fingerprint = options.fingerprint || payload.fingerprint || this.buildFingerprint(type, payload);
    const syncId = options.syncId || payload.syncId || `${String(type || 'REG').toUpperCase()}_${Date.now()}_${stableHash(`${localId}|${fingerprint}|${Math.random()}`)}`;
    return {
      localId,
      syncId,
      fingerprint,
      createdAt: options.createdAt || payload.createdAt || nowIso,
      imageCount: Number(options.imageCount ?? payload.imageCount ?? this.countPayloadImages(payload) ?? 0)
    };
  },
  normalizeRecord(type, payload = {}, options = {}) {
    const meta = this.buildMeta(type, payload, options);
    const nowIso = new Date().toISOString();
    const desiredStatus = String(options.status || '').toUpperCase();
    const status = ['AGUARDANDO_INTERNET', 'ENVIANDO', 'ERRO', 'ENVIADO'].includes(desiredStatus)
      ? desiredStatus
      : (navigator.onLine ? 'ERRO' : 'AGUARDANDO_INTERNET');
    return {
      type: String(type || payload.tipo || '').toUpperCase(),
      payload: { ...payload, syncId: meta.syncId, localId: meta.localId, fingerprint: meta.fingerprint },
      localId: meta.localId,
      syncId: meta.syncId,
      fingerprint: meta.fingerprint,
      status,
      createdAt: meta.createdAt,
      updatedAt: nowIso,
      lastAttemptAt: options.lastAttemptAt || '',
      attemptCount: Number(options.attemptCount || 0),
      lastError: String(options.lastError || ''),
      sentAt: options.sentAt || '',
      lockVersion: Number(options.lockVersion || 0),
      imageCount: Number(meta.imageCount || 0)
    };
  },
  isSent(syncId, fingerprint = '') {
    if (!syncId) return false;
    const row = this.sentIndex[syncId];
    if (!row) return false;
    if (!fingerprint) return true;
    return String(row.fingerprint || '') === String(fingerprint || '');
  },
  isFingerprintSent(fingerprint = '') {
    const target = String(fingerprint || '');
    if (!target) return false;
    return Object.values(this.sentIndex || {}).some((row) => String(row?.fingerprint || '') === target);
  },
  markSent(syncId, fingerprint, type, payload = {}) {
    if (!syncId) return;
    this.sentIndex[syncId] = {
      fingerprint: String(fingerprint || ''),
      type: String(type || '').toUpperCase(),
      sentAt: new Date().toISOString(),
      placa: normalizeName(payload.placa || ''),
      pedido: normalizeName(payload.pedido || ''),
      localId: String(payload.localId || '')
    };
    this.persistSentIndex();
  },
  async _dbGetAll() {
    if (!this.db) return [];
    return new Promise((resolve) => {
      try {
        const req = this.db.transaction(this.storeName, 'readonly').objectStore(this.storeName).getAll();
        req.onsuccess = (e) => resolve(e.target.result || []);
        req.onerror = () => resolve([]);
      } catch (_) {
        resolve([]);
      }
    });
  },
  async _dbAdd(record) {
    if (!this.db) return null;
    return new Promise((resolve) => {
      try {
        const tx = this.db.transaction(this.storeName, 'readwrite');
        const req = tx.objectStore(this.storeName).add(record);
        req.onsuccess = (e) => resolve(e.target.result);
        req.onerror = () => resolve(null);
      } catch (_) {
        resolve(null);
      }
    });
  },
  async _dbPut(record) {
    if (!this.db) return false;
    return new Promise((resolve) => {
      try {
        const tx = this.db.transaction(this.storeName, 'readwrite');
        tx.objectStore(this.storeName).put(record);
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => resolve(false);
      } catch (_) {
        resolve(false);
      }
    });
  },
  async _dbDelete(id) {
    if (!this.db) return;
    return new Promise((resolve) => {
      try {
        const tx = this.db.transaction(this.storeName, 'readwrite');
        tx.objectStore(this.storeName).delete(id);
        tx.oncomplete = () => resolve();
        tx.onerror = () => resolve();
      } catch (_) {
        resolve();
      }
    });
  },
  sortQueue(list = []) {
    return [...list].sort((a, b) => {
      const ad = new Date(a.createdAt || a.timestamp || 0).getTime();
      const bd = new Date(b.createdAt || b.timestamp || 0).getTime();
      return ad - bd;
    });
  },
  sanitizeRecord(row) {
    if (!row || typeof row !== 'object') return null;
    const rec = {
      ...row,
      type: String(row.type || '').toUpperCase(),
      status: String(row.status || '').toUpperCase()
    };
    if (!rec.payload || typeof rec.payload !== 'object') rec.payload = {};
    if (!rec.syncId) rec.syncId = rec.payload.syncId || `${rec.type || 'REG'}_${rec.id || Date.now()}_${stableHash(JSON.stringify(rec.payload))}`;
    if (!rec.localId) rec.localId = rec.payload.localId || randomId('loc');
    if (!rec.fingerprint) rec.fingerprint = rec.payload.fingerprint || this.buildFingerprint(rec.type, rec.payload);
    if (!['AGUARDANDO_INTERNET', 'ENVIANDO', 'ERRO', 'ENVIADO'].includes(rec.status)) rec.status = 'AGUARDANDO_INTERNET';
    if (!rec.createdAt) rec.createdAt = new Date(rec.timestamp || Date.now()).toISOString();
    if (!rec.updatedAt) rec.updatedAt = rec.createdAt;
    if (!rec.imageCount && rec.imageCount !== 0) rec.imageCount = this.countPayloadImages(rec.payload);
    rec.attemptCount = Number(rec.attemptCount || 0);
    rec.lockVersion = Number(rec.lockVersion || 0);
    rec.lastAttemptAt = rec.lastAttemptAt || '';
    rec.lastError = rec.lastError || '';
    rec.sentAt = rec.sentAt || '';
    rec.payload.syncId = rec.syncId;
    rec.payload.localId = rec.localId;
    rec.payload.fingerprint = rec.fingerprint;
    return rec;
  },
  async init() {
    this.loadSentIndex();
    return new Promise((resolve) => {
      if (!('indexedDB' in window)) {
        this.db = null;
        debugEngine.log("IndexedDB indisponível. Usando fila em memória.", "warn");
        return resolve();
      }
      let settled = false;
      const done = () => { if (!settled) { settled = true; resolve(); } };
      try {
        const req = indexedDB.open(this.dbName, 2);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          let store = null;
          if (!db.objectStoreNames.contains(this.storeName)) {
            store = db.createObjectStore(this.storeName, { keyPath: 'id', autoIncrement: true });
          } else {
            store = req.transaction.objectStore(this.storeName);
          }
          if (store && !store.indexNames.contains('syncId')) store.createIndex('syncId', 'syncId', { unique: false });
          if (store && !store.indexNames.contains('status')) store.createIndex('status', 'status', { unique: false });
          if (store && !store.indexNames.contains('type')) store.createIndex('type', 'type', { unique: false });
          if (store && !store.indexNames.contains('createdAt')) store.createIndex('createdAt', 'createdAt', { unique: false });
        };
        req.onsuccess = async (e) => {
          this.db = e.target.result;
          await this.migrateLegacyRecords();
          done();
        };
        req.onerror = () => {
          this.db = null;
          debugEngine.log("Falha ao abrir IndexedDB. Fila offline em memória.", "warn");
          done();
        };
        req.onblocked = () => {
          this.db = null;
          debugEngine.log("IndexedDB bloqueado por outra aba. Fila offline em memória.", "warn");
          done();
        };
      } catch (err) {
        this.db = null;
        debugEngine.log(`Erro IndexedDB: ${err.message}`, "warn");
        done();
      }
    });
  },
  async migrateLegacyRecords() {
    const all = await this._dbGetAll();
    let changed = false;
    for (const row of all) {
      const clean = this.sanitizeRecord(row);
      if (!clean) continue;
      const mustUpdate =
        clean.syncId !== row.syncId ||
        clean.localId !== row.localId ||
        clean.fingerprint !== row.fingerprint ||
        clean.status !== row.status ||
        clean.imageCount !== row.imageCount ||
        clean.type !== row.type ||
        !row.createdAt;
      if (mustUpdate) {
        clean.id = row.id;
        await this._dbPut(clean);
        changed = true;
      }
      if (clean.status === 'ENVIADO') this.markSent(clean.syncId, clean.fingerprint, clean.type, clean.payload);
    }
    if (changed) debugEngine.log("Fila offline migrada para modelo transacional.", "info");
  },
  async getAllRecords(options = {}) {
    const includeSent = Boolean(options.includeSent);
    const raw = this.db ? await this._dbGetAll() : [];
    const fromDb = raw.map((r) => this.sanitizeRecord(r)).filter(Boolean);
    const fromMem = this.memoryFallback.map((r) => this.sanitizeRecord(r)).filter(Boolean);
    const merged = this.sortQueue([...fromDb, ...fromMem]);
    if (includeSent) return merged;
    return merged.filter((item) => item.status !== 'ENVIADO');
  },
  async getAllPending() {
    return this.getAllRecords({ includeSent: false });
  },
  async findBySyncId(syncId) {
    if (!syncId) return null;
    const all = await this.getAllRecords({ includeSent: true });
    return all.find((item) => item.syncId === syncId) || null;
  },
  async findOpenDuplicate(syncId, fingerprint) {
    const all = await this.getAllRecords({ includeSent: false });
    return all.find((item) =>
      item.syncId === syncId ||
      (fingerprint && item.fingerprint === fingerprint && item.status !== 'ENVIADO')
    ) || null;
  },
  async countPendingOperacoes() {
    const all = await this.getAllRecords({ includeSent: false });
    return all.filter((item) => item.type === 'OPERACAO').length;
  },
  async savePending(type, payload, options = {}) {
    const normalizedType = String(type || payload?.tipo || '').toUpperCase();
    const rec = this.normalizeRecord(normalizedType, payload, options);

    if (this.isSent(rec.syncId, rec.fingerprint) || this.isFingerprintSent(rec.fingerprint)) {
      return { ok: false, reason: 'JA_ENVIADO', record: rec };
    }

    const dup = await this.findOpenDuplicate(rec.syncId, rec.fingerprint);
    if (dup) {
      return { ok: false, reason: 'DUPLICADO_LOCAL', record: dup };
    }

    if (normalizedType === 'OPERACAO') {
      const pendingOps = await this.countPendingOperacoes();
      if (pendingOps >= this.maxOfflineOperacoes) {
        return {
          ok: false,
          reason: 'OFFLINE_LIMIT',
          limit: this.maxOfflineOperacoes,
          queueType: 'OPERACAO'
        };
      }
    }

    if (!this.db) {
      const memRecord = { ...rec, id: `mem_${Date.now()}_${Math.random()}` };
      this.memoryFallback.push(memRecord);
      return { ok: true, record: memRecord };
    }

    const id = await this._dbAdd(rec);
    if (id === null) {
      const memRecord = { ...rec, id: `mem_${Date.now()}_${Math.random()}` };
      this.memoryFallback.push(memRecord);
      return { ok: true, record: memRecord, fallback: true };
    }
    return { ok: true, record: { ...rec, id } };
  },
  async updateRecord(id, updates = {}) {
    this.memoryFallback = this.memoryFallback.map((item) => {
      if (item.id !== id) return item;
      return this.sanitizeRecord({ ...item, ...updates, updatedAt: new Date().toISOString() });
    });
    if (!this.db || String(id).startsWith('mem_')) return true;
    const all = await this._dbGetAll();
    const existing = all.find((row) => row.id === id);
    if (!existing) return false;
    const next = this.sanitizeRecord({ ...existing, ...updates, updatedAt: new Date().toISOString() });
    if (!next) return false;
    next.id = id;
    const ok = await this._dbPut(next);
    if (ok && next.status === 'ENVIADO') this.markSent(next.syncId, next.fingerprint, next.type, next.payload);
    return ok;
  },
  async markAsSending(id) {
    const row = (await this.getAllRecords({ includeSent: true })).find((item) => item.id === id);
    if (!row || row.status === 'ENVIADO') return false;
    const nextAttempt = Number(row.attemptCount || 0) + 1;
    return this.updateRecord(id, {
      status: 'ENVIANDO',
      attemptCount: nextAttempt,
      lastAttemptAt: new Date().toISOString(),
      lastError: '',
      lockVersion: Number(row.lockVersion || 0) + 1
    });
  },
  async markAsError(id, errorMessage = '') {
    return this.updateRecord(id, {
      status: navigator.onLine ? 'ERRO' : 'AGUARDANDO_INTERNET',
      lastError: String(errorMessage || '').slice(0, 300),
      updatedAt: new Date().toISOString()
    });
  },
  async markAsSent(id, responseMeta = {}) {
    const row = (await this.getAllRecords({ includeSent: true })).find((item) => item.id === id);
    if (!row) return false;
    const sentAt = new Date().toISOString();
    const ok = await this.updateRecord(id, {
      status: 'ENVIADO',
      sentAt,
      updatedAt: sentAt,
      lastError: '',
      responseMeta
    });
    this.markSent(row.syncId, row.fingerprint, row.type, row.payload);
    return ok;
  },
  async markWaitingForConnection() {
    const all = await this.getAllRecords({ includeSent: false });
    await Promise.all(all.map((item) => {
      if (item.status === 'ENVIADO') return Promise.resolve();
      return this.updateRecord(item.id, { status: 'AGUARDANDO_INTERNET' });
    }));
  },
  async deletePending(id) {
    this.memoryFallback = this.memoryFallback.filter(item => item.id !== id);
    if (!this.db || String(id).startsWith('mem_')) return;
    await this._dbDelete(id);
  },
  async clearAll() {
    this.memoryFallback = [];
    this.sentIndex = {};
    safeStorage.removeItem(this.sentIndexKey);
    if (!this.db) return;
    return new Promise((resolve) => {
      try {
        const tx = this.db.transaction(this.storeName, 'readwrite');
        tx.objectStore(this.storeName).clear();
        tx.oncomplete = () => resolve();
        tx.onerror = () => resolve();
      } catch (_) {
        resolve();
      }
    });
  },
  async getSyncSnapshot() {
    const all = await this.getAllRecords({ includeSent: true });
    const waiting = all.filter((item) => item.status === 'AGUARDANDO_INTERNET');
    const sending = all.filter((item) => item.status === 'ENVIANDO');
    const errors = all.filter((item) => item.status === 'ERRO');
    const sent = all.filter((item) => item.status === 'ENVIADO').slice(-40).reverse();
    const pendentesOperacao = all.filter((item) => item.type === 'OPERACAO' && item.status !== 'ENVIADO').length;
    return {
      waiting,
      sending,
      errors,
      sent,
      pendentesOperacao,
      totalPendentes: waiting.length + sending.length + errors.length
    };
  },
  async replaceQueue(records = []) {
    const clean = Array.isArray(records) ? records.map((r) => this.sanitizeRecord(r)).filter(Boolean) : [];
    await this.clearAll();
    this.loadSentIndex();
    for (const item of clean) {
      if (item.status === 'ENVIADO') this.markSent(item.syncId, item.fingerprint, item.type, item.payload);
      if (!this.db) {
        this.memoryFallback.push({ ...item, id: item.id || `mem_${Date.now()}_${Math.random()}` });
      } else {
        const copy = { ...item };
        delete copy.id;
        await this._dbAdd(copy);
      }
    }
    return clean.length;
  }
};

const backupManager = {
  dbName: 'CysyBackupDBv1',
  storeName: 'logs',
  db: null,
  memoryLogs: [],
  async init() {
    return new Promise((resolve) => {
      if (!('indexedDB' in window)) {
        this.db = null;
        debugEngine.log("Backup em IndexedDB indisponível. Usando memória.", "warn");
        return resolve();
      }
      let settled = false;
      const done = () => { if (!settled) { settled = true; resolve(); } };
      try {
        const req = indexedDB.open(this.dbName, 1);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains(this.storeName)) {
            const store = db.createObjectStore(this.storeName, { keyPath: 'id', autoIncrement: true });
            store.createIndex('timestamp', 'timestamp', { unique: false });
          }
        };
        req.onsuccess = (e) => { this.db = e.target.result; done(); };
        req.onerror = () => { this.db = null; done(); };
        req.onblocked = () => { this.db = null; done(); };
      } catch (_) {
        this.db = null;
        done();
      }
    });
  },
  async addEntry(type, data = {}) {
    const entry = { type, timestamp: new Date().toISOString(), data };
    if (!this.db) {
      this.memoryLogs.push(entry);
      return;
    }
    return new Promise((resolve) => {
      const tx = this.db.transaction(this.storeName, 'readwrite');
      tx.objectStore(this.storeName).add(entry);
      tx.oncomplete = () => resolve();
      tx.onerror = () => {
        this.memoryLogs.push(entry);
        resolve();
      };
    });
  },
  async getAll() {
    if (!this.db) return [...this.memoryLogs];
    return new Promise((resolve) => {
      const req = this.db.transaction(this.storeName, 'readonly').objectStore(this.storeName).getAll();
      req.onsuccess = (e) => {
        const all = [...(e.target.result || []), ...this.memoryLogs];
        all.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
        resolve(all);
      };
      req.onerror = () => resolve([...this.memoryLogs]);
    });
  },
  async clearAll() {
    this.memoryLogs = [];
    if (!this.db) return;
    return new Promise((resolve) => {
      const tx = this.db.transaction(this.storeName, 'readwrite');
      tx.objectStore(this.storeName).clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    });
  },
  async replaceLogs(entries = []) {
    const cleanEntries = (Array.isArray(entries) ? entries : [])
      .filter(e => e && typeof e === 'object' && typeof e.type === 'string')
      .slice(-20000)
      .map(e => ({
        type: String(e.type).slice(0, 120),
        timestamp: e.timestamp && !Number.isNaN(new Date(e.timestamp).getTime()) ? String(e.timestamp) : new Date().toISOString(),
        data: e.data && typeof e.data === 'object' ? e.data : {}
      }));
    await this.clearAll();
    if (!this.db) {
      this.memoryLogs = cleanEntries;
      return cleanEntries.length;
    }
    return new Promise((resolve) => {
      const tx = this.db.transaction(this.storeName, 'readwrite');
      const store = tx.objectStore(this.storeName);
      cleanEntries.forEach(item => store.add(item));
      tx.oncomplete = () => resolve(cleanEntries.length);
      tx.onerror = () => resolve(cleanEntries.length);
    });
  },
  triggerImportBackup() {
    const input = document.getElementById('restoreBackupInput');
    if (!input) return;
    input.value = '';
    input.click();
  },
  async importBackupFromInput(event) {
    const file = event?.target?.files?.[0];
    if (!file) return;
    if (file.size > 25 * 1024 * 1024) {
      toastManager.show('Backup excede 25MB e foi bloqueado por segurança.', 'error');
      return;
    }
    try {
      const raw = await file.text();
      const parsed = safeJsonParse(raw, null);
      if (!parsed || typeof parsed !== 'object' || !Array.isArray(parsed.logs)) {
        throw new Error('Formato inválido: logs ausentes.');
      }
      const confirma = confirm(`Restaurar backup com ${parsed.logs.length} log(s)? Os dados locais atuais serão sobrescritos.`);
      if (!confirma) return;
      const datasets = parsed.datasets || {};
      const storageSnapshot = datasets.storageSnapshot || {};
      if (storageSnapshot && typeof storageSnapshot === 'object') {
        Object.keys(storageSnapshot).forEach((k) => {
          if (!k.startsWith('cysy')) return;
          const v = storageSnapshot[k];
          if (typeof v === 'string') safeStorage.setItem(k, v);
        });
      }
      if (Array.isArray(datasets.faturamentoConfirmacoes)) {
        safeStorage.setItem('cysyFaturamentoConfirmacoes', JSON.stringify(datasets.faturamentoConfirmacoes));
      }
      if (datasets.faturamentoEvidencias && typeof datasets.faturamentoEvidencias === 'object') {
        safeStorage.setItem('cysyFaturamentoEvidencias', JSON.stringify(datasets.faturamentoEvidencias));
      }
      if (datasets.permissoes && typeof datasets.permissoes === 'object') {
        safeStorage.setItem('cysyPermissionLastSummary', JSON.stringify(datasets.permissoes));
      }
      if (Array.isArray(datasets.queueRegistros)) {
        await dbManager.replaceQueue(datasets.queueRegistros);
      }
      if (datasets.sentIndex && typeof datasets.sentIndex === 'object') {
        const merged = { ...(dbManager.sentIndex || {}), ...datasets.sentIndex };
        safeStorage.setItem(dbManager.sentIndexKey, JSON.stringify(merged));
        dbManager.loadSentIndex();
      }
      const restoredCount = await this.replaceLogs(parsed.logs);
      await this.addEntry('BACKUP_RESTAURADO', {
        dataHora: new Date().toLocaleString('pt-BR'),
        arquivo: file.name,
        logsRestaurados: restoredCount,
        filaRestaurada: Array.isArray(datasets.queueRegistros) ? datasets.queueRegistros.length : 0
      });
      if (appController && typeof appController.hydrateCachedState === 'function') {
        appController.hydrateCachedState();
      }
      toastManager.show(`Backup restaurado com sucesso (${restoredCount} logs).`, 'success', 4500);
      syncManager.refreshSyncView();
      appController.handleRefresh(true);
    } catch (err) {
      toastManager.show(`Falha ao restaurar backup: ${err.message}`, 'error', 5000);
    }
  },
  async exportAllLogs() {
    const logs = await this.getAll();
    let pending = [];
    let queueFull = [];
    try {
      pending = await dbManager.getAllPending();
      queueFull = await dbManager.getAllRecords({ includeSent: true });
    } catch (_) {}
    const storageSnapshot = {};
    try {
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (!k || !k.startsWith('cysy')) continue;
        storageSnapshot[k] = localStorage.getItem(k);
      }
    } catch (_) {}
    const payload = {
      exportedAt: new Date().toISOString(),
      total: logs.length,
      snapshot: {
        usuarioAtual: safeStorage.getItem('cysyUser', ''),
        pendenciasOffline: (pending || []).length,
        versaoApp: config.app.version
      },
      datasets: {
        faturamentoConfirmacoes: safeJsonParse(safeStorage.getItem('cysyFaturamentoConfirmacoes', '[]'), []),
        faturamentoEvidencias: safeJsonParse(safeStorage.getItem('cysyFaturamentoEvidencias', '{}'), {}),
        permissoes: safeJsonParse(safeStorage.getItem('cysyPermissionLastSummary', '{}'), {}),
        queueRegistros: queueFull,
        sentIndex: dbManager.sentIndex || {},
        storageSnapshot
      },
      logs
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
    a.download = `cysy-backup-${stamp}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    toastManager.show(`Backup exportado (${logs.length} registros).`, 'success');
  }
};

const syncManager = {
  isOnline: navigator.onLine,
  isSyncing: false,
  lastSyncAt: '',
  activeSyncIds: new Set(),
  async init() {
    window.addEventListener('online', async () => {
      this.isOnline = true;
      toastManager.show('Conexão restabelecida. Sincronização automática iniciada.', 'info');
      await this.updateStatus(true, { autoProcess: false });
      await this.processQueue();
    });
    window.addEventListener('offline', async () => {
      this.isOnline = false;
      await dbManager.markWaitingForConnection();
      await this.updateStatus(false);
      await this.refreshSyncView();
    });
    await dbManager.init().catch(() => {});
    this.lastSyncAt = safeStorage.getItem('cysyLastSyncAt', '');
    await this.updateStatus(navigator.onLine, { autoProcess: false });
    await this.refreshSyncView();
    setInterval(() => this.processQueue(), 20000);
    setInterval(() => this.refreshSyncView(), 8000);
  },
  async updateStatus(online, options = {}) {
    this.isOnline = online;
    debugEngine.updateNetStatus(online);
    if (!online) {
      await dbManager.markWaitingForConnection();
    }
    let count = 0;
    try {
      const pending = await dbManager.getAllPending();
      count = pending ? pending.length : 0;
    } catch(e) {}
    const badge = document.getElementById('netStatusBadge');
    if (!badge) return;

    if (this.isSyncing) {
      badge.innerHTML = '🔄 Sincronizando...';
      badge.className = 'status-badge syncing';
    } else if (this.isOnline) {
      badge.innerHTML = `🟢 ONLINE ${count > 0 ? `| Fila: ${count}` : ''}`;
      badge.className = 'status-badge online';
      if (count > 0 && options.autoProcess !== false) this.processQueue();
    } else {
      badge.innerHTML = `🔴 OFFLINE ${count > 0 ? `| Fila: ${count}` : ''}`;
      badge.className = 'status-badge offline';
    }
    await this.refreshSyncView();
  },
  async refreshSyncView() {
    try {
      const snapshot = await dbManager.getSyncSnapshot();
      if (uiBuilder && typeof uiBuilder.renderSyncCenter === 'function') {
        uiBuilder.renderSyncCenter({
          ...snapshot,
          online: this.isOnline,
          syncing: this.isSyncing,
          lastSyncAt: this.lastSyncAt
        });
      }
    } catch (_) {}
  },
  async processQueue(options = {}) {
    if (!this.isOnline || this.isSyncing) return;
    let pending = [];
    try {
      pending = await dbManager.getAllRecords({ includeSent: false });
    } catch (_) {}
    if (!pending || pending.length === 0) return;
    const onlyIds = Array.isArray(options.onlyIds) ? new Set(options.onlyIds.map(String)) : null;
    const queue = pending
      .filter((item) => ['AGUARDANDO_INTERNET', 'ERRO', 'ENVIANDO'].includes(String(item.status || '').toUpperCase()))
      .filter((item) => !onlyIds || onlyIds.has(String(item.id)))
      .sort((a, b) => new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime());
    if (queue.length === 0) return;

    this.isSyncing = true;
    await this.updateStatus(true, { autoProcess: false });
    let successCount = 0;
    try {
      for (const item of queue) {
        if (!navigator.onLine) break;
        if (this.activeSyncIds.has(item.syncId)) continue;
        this.activeSyncIds.add(item.syncId);
        try {
          const current = await dbManager.findBySyncId(item.syncId);
          if (!current || current.status === 'ENVIADO') continue;
          if (dbManager.isSent(current.syncId, current.fingerprint) || dbManager.isFingerprintSent(current.fingerprint)) {
            await dbManager.markAsSent(current.id, { deduplicado: true, origem: 'sentIndex' });
            continue;
          }
          const locked = await dbManager.markAsSending(current.id);
          if (!locked) continue;
          await this.refreshSyncView();

          const resp = await apiService.sendDataToAppScript(current.payload);
          if (resp && resp.success) {
            await dbManager.markAsSent(current.id, { sucesso: true });
            backupManager.addEntry('SYNC_ENVIADO', {
              syncId: current.syncId,
              tipo: current.type,
              dataHora: new Date().toLocaleString('pt-BR')
            });
            successCount++;
          } else {
            throw new Error(resp?.message || 'Resposta inválida do endpoint.');
          }
        } catch(e) {
          const msg = e?.message || 'Falha ao sincronizar';
          await dbManager.markAsError(item.id, msg);
          backupManager.addEntry('SYNC_ERRO', {
            syncId: item.syncId,
            tipo: item.type,
            erro: msg,
            dataHora: new Date().toLocaleString('pt-BR')
          });
        } finally {
          this.activeSyncIds.delete(item.syncId);
          await this.refreshSyncView();
        }
      }
    } finally {
      this.isSyncing = false;
    }
    if (successCount > 0) {
      this.lastSyncAt = new Date().toISOString();
      safeStorage.setItem('cysyLastSyncAt', this.lastSyncAt);
      uiBuilder.updateGlobalUpdateTimestamp();
      appController.handleRefresh();
      toastManager.show(`${successCount} item(ns) sincronizado(s) com sucesso.`, 'success');
    }
    this.isOnline = navigator.onLine;
    await this.updateStatus(this.isOnline, { autoProcess: false });
  },
  async retryItem(idOrSyncId, buttonEl = null) {
    if (buttonEl) {
      buttonEl.disabled = true;
      buttonEl.style.opacity = '0.65';
    }
    try {
      const all = await dbManager.getAllRecords({ includeSent: false });
      const row = all.find((item) => String(item.id) === String(idOrSyncId) || String(item.syncId) === String(idOrSyncId));
      if (!row) {
        toastManager.show('Item não encontrado para reprocessamento.', 'warning');
        return;
      }
      if (dbManager.isSent(row.syncId, row.fingerprint) || dbManager.isFingerprintSent(row.fingerprint)) {
        await dbManager.markAsSent(row.id, { deduplicado: true, origem: 'retry' });
        toastManager.show('Esse registro já foi enviado anteriormente.', 'info');
        await this.refreshSyncView();
        return;
      }
      await dbManager.updateRecord(row.id, {
        status: this.isOnline ? 'ERRO' : 'AGUARDANDO_INTERNET',
        lastError: ''
      });
      if (this.isOnline) {
        await this.processQueue({ onlyIds: [row.id] });
      } else {
        toastManager.show('Sem internet no momento. Item mantido na fila.', 'warning');
        await this.refreshSyncView();
      }
    } finally {
      if (buttonEl) {
        buttonEl.disabled = false;
        buttonEl.style.opacity = '1';
      }
    }
  },
  async retryAllErrors() {
    const snapshot = await dbManager.getSyncSnapshot();
    const ids = (snapshot.errors || []).map((row) => row.id);
    if (ids.length === 0) {
      toastManager.show('Não há itens com erro para reenviar.', 'info');
      return;
    }
    if (!this.isOnline) {
      toastManager.show('Sem internet. Conecte-se para tentar novamente.', 'warning');
      return;
    }
    await this.processQueue({ onlyIds: ids });
  }
};

const swManager = {
  async init() {
    if (!('serviceWorker' in navigator)) {
      debugEngine.updateSWStatus(false);
      debugEngine.log("Service Worker não suportado neste navegador.", "warn");
      return;
    }
    try {
      await cacheJanitor.firstBootClear();
      const buildTag = String(config.app.buildTag || '20260406r02');
      const reg = await navigator.serviceWorker.register(`./service-worker.js?v=${buildTag}`, { scope: './' });
      if (navigator.storage?.persist) {
        try {
          const alreadyPersistent = await navigator.storage.persisted?.();
          if (!alreadyPersistent) await navigator.storage.persist();
        } catch (_) {}
      }
      debugEngine.updateSWStatus(true);
      debugEngine.log(`Service Worker ativo (${reg.scope}).`, "success");
    } catch (err) {
      debugEngine.updateSWStatus(false);
      debugEngine.log(`Falha ao registrar Service Worker: ${err.message}`, "error");
    }
  }
};

const installManager = {
  deferredPrompt: null,
  manualHintShown: false,
  isStandalone() {
    return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone;
  },
  getContext() {
    const ua = navigator.userAgent || '';
    const isIOS = /iPad|iPhone|iPod/.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    const isSafari = isIOS && /Safari/.test(ua) && !/CriOS|FxiOS|EdgiOS|OPiOS/.test(ua);
    const isIOSNonSafari = isIOS && !isSafari;
    const isSamsungBrowser = /SamsungBrowser/.test(ua);
    const isEdge = /Edg\//.test(ua);
    const isAndroid = /Android/.test(ua);
    const isChromeLike = /Chrome|CriOS/.test(ua) && !isEdge && !/OPR|SamsungBrowser/.test(ua);
    return {
      isSecure: Boolean(window.isSecureContext),
      isIOS,
      isSafari,
      isIOSNonSafari,
      isSamsungBrowser,
      isEdge,
      isAndroid,
      isChromeLike
    };
  },
  getManualInstallMessage() {
    const context = this.getContext();
    if (!context.isSecure) return 'Para instalar como aplicativo real no smartphone, abra este sistema pela URL HTTPS publicada. Arquivo local ou HTTP não habilitam instalação PWA.';
    if (context.isIOSNonSafari) return 'No iPhone/iPad, abra este sistema no Safari e use Compartilhar > "Adicionar à Tela de Início". Chrome e Edge no iOS não instalam PWA diretamente.';
    if (context.isSafari) return 'No Safari do iPhone/iPad: toque em Compartilhar e escolha "Adicionar à Tela de Início".';
    if (context.isAndroid && context.isChromeLike) return 'No Chrome Android: use o botão de instalar. Se o prompt não aparecer, abra o menu ⋮ e toque em "Instalar aplicativo".';
    if (context.isSamsungBrowser) return 'No Samsung Internet: abra o menu e use a opção de instalar ou adicionar à tela inicial.';
    if (context.isEdge) return 'No Microsoft Edge: abra o menu e vá em Aplicativos > Instalar este site como aplicativo.';
    if (context.isChromeLike) return 'No Chrome: use o menu do navegador e procure a opção de instalar este aplicativo.';
    return 'A instalação automática não ficou disponível neste navegador. Use o menu do navegador para instalar ou adicionar à tela inicial.';
  },
  updateButtonState() {
    const btn = document.getElementById('installAppBtn');
    if (!btn) return;
    const context = this.getContext();
    if (this.isStandalone()) {
      btn.style.display = 'none';
      btn.classList.remove('install-ready');
      return;
    }
    btn.style.display = 'inline-flex';
    btn.classList.toggle('install-ready', Boolean(this.deferredPrompt) && context.isSecure);
    if (!context.isSecure) {
      btn.textContent = '🔒 Instalar via HTTPS';
      btn.title = 'Abra o sistema pela URL HTTPS publicada para instalar como app.';
      return;
    }
    if (this.deferredPrompt) {
      btn.textContent = '📲 Instalar app';
      btn.title = 'Instalar aplicativo no dispositivo';
      return;
    }
    if (context.isIOSNonSafari) {
      btn.textContent = '🧭 Abrir no Safari';
      btn.title = 'No iPhone/iPad, a instalação deve ser feita pelo Safari.';
      return;
    }
    if (context.isSafari) {
      btn.textContent = '📲 Instalar no Safari';
      btn.title = 'Mostrar os passos para adicionar o app à Tela de Início.';
      return;
    }
    if (context.isAndroid && context.isChromeLike) {
      btn.textContent = '📲 Instalar app';
      btn.title = 'Se o prompt não abrir, use o menu do navegador para instalar.';
      return;
    }
    btn.textContent = '📲 Como instalar';
    btn.title = 'Ver instruções de instalação para este navegador.';
  },
  init() {
    const btn = document.getElementById('installAppBtn');
    if (!btn) return;
    this.updateButtonState();
    setTimeout(() => {
      if (!this.manualHintShown && !this.deferredPrompt && !this.isStandalone()) {
        this.manualHintShown = true;
        const type = window.isSecureContext ? 'info' : 'warning';
        toastManager.show(this.getManualInstallMessage(), type, 5200);
      }
    }, 1800);
    window.addEventListener('beforeinstallprompt', (event) => {
      event.preventDefault();
      this.deferredPrompt = event;
      this.updateButtonState();
      toastManager.show('Aplicativo disponível para instalação.', 'success');
    });
    window.addEventListener('appinstalled', () => {
      this.deferredPrompt = null;
      this.updateButtonState();
      toastManager.show('Aplicativo instalado com sucesso.', 'success');
    });
    window.addEventListener('pageshow', () => this.updateButtonState());
    window.addEventListener('visibilitychange', () => {
      if (!document.hidden) this.updateButtonState();
    });
  },
  async installApp() {
    const btn = document.getElementById('installAppBtn');
    const context = this.getContext();
    if (this.isStandalone()) {
      toastManager.show('O aplicativo já está instalado neste dispositivo.', 'info');
      this.updateButtonState();
      return;
    }
    if (!context.isSecure) {
      toastManager.show(this.getManualInstallMessage(), 'warning', 7000);
      return;
    }
    if (!this.deferredPrompt) {
      toastManager.show(this.getManualInstallMessage(), 'info', 6500);
      return;
    }
    this.deferredPrompt.prompt();
    const choice = await this.deferredPrompt.userChoice;
    if (choice?.outcome === 'accepted') {
      toastManager.show('Instalação iniciada.', 'success');
      if (btn) btn.style.display = 'none';
    } else {
      toastManager.show('Instalação cancelada pelo usuário.', 'warning');
    }
    this.deferredPrompt = null;
    this.updateButtonState();
  }
};

const versionManager = {
    cleanupWaveKey: 'cysyLegacyCleanupWave2026033011',
  async limparCachesRuntime() {
    try {
      if ('caches' in window) {
        const keys = await caches.keys();
        const buildTag = String(config.app.buildTag || '').trim();
        const expectedPrefixes = [
          buildTag ? `cysy-log360-v${buildTag}` : ''
        ].filter(Boolean);
        await Promise.all(
          keys
            .filter((key) =>
              key.startsWith('cysy-') &&
              !expectedPrefixes.some((prefix) => key === prefix || key.startsWith(prefix))
            )
            .map((key) => caches.delete(key))
        );
      }
    } catch (_) {}
  },
  async limparArtefatosLegados() {
    if (safeStorage.getItem(this.cleanupWaveKey, '')) return;
    try {
      const dbAtivo = new Set([String(dbManager?.dbName || ''), String(backupManager?.dbName || '')]);
      const legados = ['CysyDBv5', 'CysyDBv6', 'CysyDBv7', 'CysyDBv9', 'CysyBackupDBv0', 'CysyBackupDBv2']
        .filter((name) => name && !dbAtivo.has(name));
      const deleteDb = (name) => new Promise((resolve) => {
        try {
          if (!('indexedDB' in window)) return resolve();
          const req = indexedDB.deleteDatabase(name);
          req.onsuccess = () => resolve();
          req.onerror = () => resolve();
          req.onblocked = () => resolve();
        } catch (_) { resolve(); }
      });
      await Promise.all(legados.map(deleteDb));

      // Remove aliases antigos sem tocar nos dados atuais.
      safeStorage.removeItem('cysySyncSentIndex');
      safeStorage.removeItem('cysySyncSentIndexV1');
      safeStorage.removeItem('cysyLocations');
      safeStorage.removeItem('cysyLocationLogV1');
      safeStorage.removeItem('cysyMapSyncState');
      safeStorage.removeItem('cysyMapViewState');
      safeStorage.removeItem('cysyStockAlertLastAt');

      safeStorage.setItem(this.cleanupWaveKey, new Date().toISOString());
      backupManager.addEntry('LIMPEZA_LEGADO', {
        dataHora: new Date().toLocaleString('pt-BR'),
        dbRemovidos: legados.length
      });
    } catch (_) {}
  },

  async ensureVersionRuntime() {
    const current = String(config.app.version || '');
    const key = String(config.app.versionStorageKey || 'cysyAppVersion');
    const old = safeStorage.getItem(key, '');

    await this.limparArtefatosLegados();
    if (old === current) return;

    await this.limparCachesRuntime();
    safeStorage.setItem(key, current);
    safeStorage.setItem(`${key}:updatedAt`, new Date().toISOString());
    backupManager.addEntry('VERSAO_CACHE_LIMPO', { versaoAnterior: old || 'N/A', versaoAtual: current });
    toastManager.show(`Cache técnico atualizado para versão ${current} sem perda de dados.`, 'info', 4200);
  }
};

const cacheJanitor = {
  lastRunKey: 'cysyCacheJanitorLastRun',
  intervalMs: 12 * 60 * 60 * 1000,
  firstBootKey: 'cysyFirstBootCacheClearV3',
  preserveStorageKeys: new Set([
    'cysyGoogleApiKey'
  ]),
  transientStorageKeys: new Set([
    'cysyAppCacheV1',
    'cysyCacheJanitorLastRun',
    'cysyLastSyncAt',
    'cysyPermissionBootstrapV1',
    'cysyPermissionLastSummary',
    'cysyPriorityAlertsV2'
  ]),

  getAppScopePath() {
    try {
      return new URL('./', window.location.href).pathname;
    } catch (_) {
      return '/';
    }
  },

  async getScopedRegistrations() {
    if (!('serviceWorker' in navigator)) return [];
    try {
      const scopePath = this.getAppScopePath();
      const registrations = await navigator.serviceWorker.getRegistrations();
      return registrations.filter((registration) => {
        try {
          return new URL(registration.scope).pathname.startsWith(scopePath);
        } catch (_) {
          return false;
        }
      });
    } catch (_) {
      return [];
    }
  },

  async clearManagedCaches(forceAllManaged = false) {
    if (!('caches' in window)) return;
    const keepCaches = new Set([
      `cysy-log360-v${config.app.buildTag}`
    ].filter(Boolean));
    try {
      const keys = await caches.keys();
      await Promise.allSettled(
        keys
          .filter((key) => key.startsWith('cysy-') && (forceAllManaged || !keepCaches.has(key)))
          .map((key) => caches.delete(key))
      );
    } catch (_) {}
  },

  clearManagedLocalStorage(includePersistent = false) {
    try {
      const removableKeys = new Set([
        ...this.transientStorageKeys,
        ...(includePersistent ? [...this.preserveStorageKeys] : [])
      ]);
      removableKeys.forEach((key) => {
        if (key.startsWith('cysy')) safeStorage.removeItem(key);
      });
    } catch (_) {}
  },

  clearTransientStorageArtifacts() {
    this.clearManagedLocalStorage(false);
    try { sessionStorage.clear(); } catch (_) {}
  },

  shouldRun(now = Date.now()) {
    const last = Number(safeStorage.getItem(this.lastRunKey, '0')) || 0;
    return !last || (now - last) >= this.intervalMs;
  },

  async run(reason = 'startup') {
    if (!('caches' in window)) return;
    const now = Date.now();
    if (reason !== 'force' && !this.shouldRun(now)) return;

    try {
      await this.clearManagedCaches(false);
      safeStorage.setItem(this.lastRunKey, String(now));
      backupManager.addEntry('LIMPEZA_CACHE_AUTOMATICA', {
        motivo: reason,
        quando: new Date(now).toISOString()
      });
    } catch (_) {}
  },

  async firstBootClear() {
    if (safeStorage.getItem(this.firstBootKey, '') === 'done') return;
    try {
      const registrations = await this.getScopedRegistrations();
      await Promise.allSettled(registrations.map((registration) => registration.unregister()));
      await this.clearManagedCaches(true);
      this.clearTransientStorageArtifacts();
    } catch (_) {}
    safeStorage.setItem(this.firstBootKey, 'done');
  },

  schedule() {
    this.run('startup');
    setInterval(() => this.run('interval'), this.intervalMs);
  }
};

const priorityAlertManager = {
  storageKey: 'cysyPriorityAlertsV2',
  reminderMs: 2 * 60 * 1000,
  cleanupAfterMs: 14 * 24 * 60 * 60 * 1000,
  alerts: {},
  currentOverlayAlertId: '',
  lastOverlayReminderAt: 0,
  overlayDismissedIds: {},
  initialized: false,
  severityRank: { danger: 3, warning: 2, info: 1, success: 0 },

  init() {
    if (this.initialized) return;
    this.initialized = true;
    this.load();
    this.cleanup();
    this.renderBar();
    setInterval(() => this.remindActiveAlerts(), 30000);
    window.addEventListener('visibilitychange', () => {
      if (!document.hidden) this.remindActiveAlerts();
    });
    window.addEventListener('focus', () => this.remindActiveAlerts());
  },

  load() {
    const parsed = safeJsonParse(safeStorage.getItem(this.storageKey, '{}'), {});
    this.alerts = parsed && typeof parsed === 'object' ? parsed : {};
    this.removeLegacyMappingAlerts();
  },

  persist() {
    safeStorage.setItem(this.storageKey, JSON.stringify(this.alerts));
  },

  cleanup() {
    const now = Date.now();
    let changed = false;
    Object.entries(this.alerts || {}).forEach(([id, alert]) => {
      if (this.isLegacyMappingAlert(alert, id)) {
        delete this.alerts[id];
        changed = true;
        return;
      }
      const ackAt = new Date(alert?.acknowledgedAt || 0).getTime();
      if (ackAt && (now - ackAt) > this.cleanupAfterMs) {
        delete this.alerts[id];
        changed = true;
      }
    });
    if (changed) this.persist();
  },

  normalizeSeverity(severity = 'info') {
    const target = String(severity || '').toLowerCase();
    return this.severityRank[target] === undefined ? 'info' : target;
  },

  getIcon(alert = {}) {
    if (alert.icon) return alert.icon;
    const severity = this.normalizeSeverity(alert.severity);
    if (severity === 'danger') return '🚨';
    if (severity === 'warning') return '⚠️';
    if (severity === 'success') return '✅';
    return 'ℹ️';
  },

  getCurrentTabId() {
    const activeTab = document.querySelector('.tab-content.active');
    if (!activeTab?.id) return '';
    return String(activeTab.id).replace(/^tab-/, '').trim().toLowerCase();
  },

  shouldSuppressOverlay() {
    const tabId = this.getCurrentTabId();
    return tabId === 'op' || tabId === 'perdas';
  },

  resetOverlayDismissals() {
    this.overlayDismissedIds = {};
  },

  getActionLabel(actionKey = '') {
    const action = String(actionKey || '');
    if (action === 'tab_op') return '📋 Abrir liberação';
    return '';
  },

  sortAlerts(list = []) {
    return [...list].sort((a, b) => {
      const severityDiff = (this.severityRank[this.normalizeSeverity(b.severity)] || 0) - (this.severityRank[this.normalizeSeverity(a.severity)] || 0);
      if (severityDiff !== 0) return severityDiff;
      return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
    });
  },

  getActiveAlerts() {
    if (!loginManager.isLoggedIn()) return [];
    return this.sortAlerts(Object.values(this.alerts || {}).filter((alert) => alert && !alert.acknowledgedAt));
  },

  getTopAlert() {
    return this.getActiveAlerts()[0] || null;
  },

  isLegacyMappingAlert(alert = {}, fallbackId = '') {
    const blob = [
      String(alert?.id || fallbackId || '').toLowerCase(),
      String(alert?.title || '').toLowerCase(),
      String(alert?.body || '').toLowerCase(),
      String(alert?.actionKey || '').toLowerCase(),
      String(alert?.actionLabel || '').toLowerCase(),
      String(alert?.source || '').toLowerCase()
    ].join(' | ');
    const legacyTokens = [
      'localização mapeada',
      'localizacao mapeada',
      'sem localização',
      'sem localizacao',
      'mapeamento',
      'mapa dos lotes',
      'revisar o mapa',
      'abrir mapa',
      'tab_mapa',
      'mapa_missing_locations',
      'estoque_pendente_'
    ];
    return legacyTokens.some((token) => blob.includes(token));
  },

  removeLegacyMappingAlerts() {
    let changed = false;
    Object.entries(this.alerts || {}).forEach(([id, alert]) => {
      if (!this.isLegacyMappingAlert(alert, id)) return;
      delete this.alerts[id];
      changed = true;
    });
    if (changed) this.persist();
  },

  registerAlert(payload = {}) {
    if (!loginManager.isLoggedIn()) return null;
    const id = String(payload.id || '').trim();
    if (!id) return null;
    if (this.isLegacyMappingAlert(payload, id)) return null;
    const now = new Date().toISOString();
    const signature = stableHash(JSON.stringify({
      title: payload.title || '',
      body: payload.body || '',
      severity: this.normalizeSeverity(payload.severity),
      actionKey: payload.actionKey || '',
      actionPayload: payload.actionPayload || null,
      source: payload.source || ''
    }));
    const previous = this.alerts[id];
    const isNewCycle = !previous || previous.signature !== signature;
    const next = {
      id,
      title: String(payload.title || 'Alerta operacional'),
      body: String(payload.body || ''),
      severity: this.normalizeSeverity(payload.severity || 'warning'),
      icon: String(payload.icon || ''),
      actionKey: String(payload.actionKey || ''),
      actionLabel: String(payload.actionLabel || this.getActionLabel(payload.actionKey)),
      actionPayload: payload.actionPayload || null,
      source: String(payload.source || 'sistema'),
      signature,
      createdAt: isNewCycle ? now : (previous?.createdAt || now),
      updatedAt: now,
      acknowledgedAt: isNewCycle ? '' : (previous?.acknowledgedAt || ''),
      lastNotifiedAt: isNewCycle ? '' : (previous?.lastNotifiedAt || ''),
      lastSoundAt: isNewCycle ? '' : (previous?.lastSoundAt || ''),
      meta: payload.meta || previous?.meta || {}
    };
    this.alerts[id] = next;
    this.persist();
    this.renderBar();
    this.remindActiveAlerts({ force: isNewCycle, preferredId: id });
    return next;
  },

  updateNotificationTimestamp(id) {
    if (!this.alerts[id]) return;
    this.alerts[id].lastNotifiedAt = new Date().toISOString();
    this.persist();
    this.renderBar();
  },

  updateSoundTimestamp(id) {
    if (!this.alerts[id]) return;
    this.alerts[id].lastSoundAt = new Date().toISOString();
    this.persist();
  },

  playAlertSound(alert, force = false) {
    if (!alert || alert.acknowledgedAt || !soundManager?.enabled) return;
    const lastAt = new Date(alert.lastSoundAt || 0).getTime();
    if (!force && lastAt && (Date.now() - lastAt) < this.reminderMs) return;
    const severity = this.normalizeSeverity(alert.severity);
    if (severity === 'danger') soundManager.play('criticalAlert');
    else if (severity === 'warning') soundManager.play('alert');
    else soundManager.play('notify');
    this.updateSoundTimestamp(alert.id);
  },

  async showBrowserNotification(alert, force = false) {
    if (!alert || alert.acknowledgedAt || this.isLegacyMappingAlert(alert)) return;
    if (!("Notification" in window) || Notification.permission !== 'granted') return;
    const lastAt = new Date(alert.lastNotifiedAt || 0).getTime();
    if (!force && lastAt && (Date.now() - lastAt) < this.reminderMs) return;

    const options = {
      body: alert.body,
      icon: './assets/icons/icon-192.png',
      badge: './assets/icons/icon-192.png',
      tag: alert.id,
      renotify: true,
      requireInteraction: true,
      vibrate: [220, 120, 220],
      data: {
        alertId: alert.id,
        actionKey: alert.actionKey || ''
      }
    };
    try {
      const reg = await navigator.serviceWorker.getRegistration();
      if (reg?.showNotification) {
        await reg.showNotification(alert.title, options);
      } else {
        new Notification(alert.title, options);
      }
      this.updateNotificationTimestamp(alert.id);
    } catch (_) {
      try {
        new Notification(alert.title, options);
        this.updateNotificationTimestamp(alert.id);
      } catch (_) {}
    }
  },

  buildMeta(alert = {}) {
    const lines = [];
    lines.push(`Origem: ${alert.source || 'sistema'}`);
    lines.push(`Criado em: ${new Date(alert.createdAt || Date.now()).toLocaleString('pt-BR')}`);
    if (alert.lastNotifiedAt) {
      lines.push(`Último lembrete: ${new Date(alert.lastNotifiedAt).toLocaleString('pt-BR')}`);
    } else {
      lines.push('Último lembrete: aguardando primeira notificação');
    }
    if (alert.actionLabel) lines.push(`Ação sugerida: ${alert.actionLabel}`);
    return lines.join('\n');
  },

  applyOverlaySeverity(alert = {}) {
    const box = document.getElementById('alertGlobalBox');
    const title = document.getElementById('globalAlertTitle');
    if (!box || !title) return;
    const severity = this.normalizeSeverity(alert.severity);
    if (severity === 'danger') {
      box.style.borderColor = '#F87171';
      box.style.background = '#FFF7ED';
      title.style.color = '#7F1D1D';
      return;
    }
    if (severity === 'warning') {
      box.style.borderColor = '#F59E0B';
      box.style.background = '#FFFBEB';
      title.style.color = '#92400E';
      return;
    }
    box.style.borderColor = '#60A5FA';
    box.style.background = '#EFF6FF';
    title.style.color = '#1E3A8A';
  },

  renderBar() {
    const bar = document.getElementById('priorityAlertBar');
    const chip = document.getElementById('priorityAlertChip');
    const title = document.getElementById('priorityAlertTitleBar');
    const body = document.getElementById('priorityAlertBodyBar');
    const counter = document.getElementById('headerAlertCounter');
    if (!bar || !chip || !title || !body) return;

    const active = this.getActiveAlerts();
    const top = active[0];
    if (!top) {
      bar.classList.remove('active');
      bar.style.display = 'none';
      if (counter) {
        counter.classList.remove('show', 'critical', 'warning', 'info');
        counter.style.display = 'none';
        counter.textContent = '🚨 0 alertas';
      }
      return;
    }

    bar.style.display = '';
    bar.classList.add('active');
    bar.dataset.severity = this.normalizeSeverity(top.severity);
    chip.textContent = `${this.getIcon(top)} Prioridade máxima • ${active.length} pendente(s)`;
    title.textContent = top.title;
    body.textContent = top.body;
    if (counter) {
      const severity = this.normalizeSeverity(top.severity);
      counter.style.display = '';
      counter.classList.add('show');
      counter.classList.remove('critical', 'warning', 'info');
      counter.classList.add(severity === 'danger' ? 'critical' : severity === 'warning' ? 'warning' : 'info');
      counter.textContent = `${this.getIcon(top)} ${active.length} alerta${active.length > 1 ? 's' : ''}`;
      counter.title = top.title;
    }
  },

  openTopAlert() {
    this.openAlert('');
  },

  openAlert(alertId = '') {
    const overlay = document.getElementById('globalAlertOverlay');
    const stack = document.getElementById('globalAlertStack');
    if (!overlay || !stack) return;
    const active = this.getActiveAlerts();
    const targetId = String(alertId || '').trim();
    const visibleAlerts = active.filter((alert) => {
      if (alert.acknowledgedAt) return false;
      if (targetId) return alert.id === targetId;
      return !this.overlayDismissedIds[alert.id];
    });
    if (visibleAlerts.length === 0) {
      this.hideOverlay();
      return;
    }

    this.currentOverlayAlertId = visibleAlerts[0]?.id || '';
    stack.innerHTML = visibleAlerts.map((alert, index) => {
      const severity = this.normalizeSeverity(alert.severity);
      const actionLabel = alert.actionLabel
        ? `<button class="btn-danger" type="button" onclick="priorityAlertManager.handleAction('${escapeHTML(alert.id)}')" style="height:56px;">${escapeHTML(alert.actionLabel)}</button>`
        : '';
      return `<article class="priority-alert-card" data-severity="${escapeHTML(severity)}" data-alert-id="${escapeHTML(alert.id)}">
        <div class="priority-alert-card-head">
          <div class="priority-alert-card-title">
            <span class="priority-alert-card-subtitle">${escapeHTML(this.getIcon(alert))} alerta prioritário ${visibleAlerts.length > 1 ? `• ${index + 1}/${visibleAlerts.length}` : ''}</span>
            <h3>${escapeHTML(alert.title)}</h3>
          </div>
          <button type="button" class="priority-alert-card-close" onclick="priorityAlertManager.dismissOverlayAlert('${escapeHTML(alert.id)}')" aria-label="Fechar alerta sem confirmar">×</button>
        </div>
        <div class="priority-alert-card-body">${escapeHTML(alert.body)}</div>
        <div class="priority-modal-meta">${escapeHTML(this.buildMeta(alert))}</div>
        <div class="priority-modal-actions">
          ${actionLabel}
          <button class="btn-success" type="button" onclick="priorityAlertManager.acknowledgeAlert('${escapeHTML(alert.id)}')" style="height:56px;">✅ Confirmar visualização</button>
          <button class="priority-alert-btn ghost" type="button" onclick="priorityAlertManager.dismissOverlayAlert('${escapeHTML(alert.id)}')">Lembrar depois</button>
        </div>
      </article>`;
    }).join('');
    overlay.style.display = 'flex';
  },

  hideOverlay() {
    const overlay = document.getElementById('globalAlertOverlay');
    const stack = document.getElementById('globalAlertStack');
    if (overlay) overlay.style.display = 'none';
    if (stack) stack.innerHTML = '';
    this.currentOverlayAlertId = '';
  },

  dismissOverlayAlert(alertId = '') {
    const target = alertId ? this.alerts[alertId] : this.getTopAlert();
    if (!target) return;
    this.overlayDismissedIds[target.id] = Date.now();
    const remaining = this.getActiveAlerts().filter((alert) => !this.overlayDismissedIds[alert.id]);
    if (remaining.length === 0) {
      this.hideOverlay();
      return;
    }
    this.openAlert('');
  },

  acknowledgeAlert(alertId = '') {
    const alert = alertId ? this.alerts[alertId] : this.getTopAlert();
    if (!alert) return;
    this.alerts[alert.id] = {
      ...alert,
      acknowledgedAt: new Date().toISOString()
    };
    delete this.overlayDismissedIds[alert.id];
    this.persist();
    const remaining = this.getActiveAlerts().filter((item) => item.id !== alert.id && !this.overlayDismissedIds[item.id]);
    if (remaining.length > 0) this.openAlert('');
    else this.hideOverlay();
    this.renderBar();
    backupManager.addEntry('ALERTA_CONFIRMADO', {
      alertaId: alert.id,
      titulo: alert.title,
      dataHora: new Date().toLocaleString('pt-BR')
    });
    toastManager.show(`Alerta confirmado: ${alert.title}`, 'success', 4200);
  },

  acknowledgeTopAlert() {
    this.acknowledgeAlert('');
  },

  acknowledgeFromOverlay() {
    this.acknowledgeAlert(this.currentOverlayAlertId);
  },

  handleAction(alert = null) {
    const target = typeof alert === 'string'
      ? this.alerts[alert]
      : (alert || (this.currentOverlayAlertId ? this.alerts[this.currentOverlayAlertId] : this.getTopAlert()));
    if (!target) return;
    if (target.actionKey === 'tab_op') {
      delete this.overlayDismissedIds[target.id];
      this.hideOverlay();
      uiBuilder.switchTab(null, 'op');
      return;
    }
    if (target.actionKey === 'tab_perdas') {
      delete this.overlayDismissedIds[target.id];
      this.hideOverlay();
      uiBuilder.switchTab(null, 'perdas');
    }
  },

  handleActionFromOverlay() {
    this.handleAction();
  },

  remindActiveAlerts(options = {}) {
    if (!loginManager.isLoggedIn()) {
      this.resetOverlayDismissals();
      this.hideOverlay();
      this.renderBar();
      return;
    }
    const active = this.getActiveAlerts();
    if (active.length === 0) {
      this.resetOverlayDismissals();
      this.hideOverlay();
      this.renderBar();
      return;
    }
    const preferred = options.preferredId ? this.alerts[options.preferredId] : null;
    const top = preferred && !preferred.acknowledgedAt ? preferred : active[0];
    this.renderBar();

    const shouldRepeatOverlay = options.force || !this.lastOverlayReminderAt || (Date.now() - this.lastOverlayReminderAt) >= this.reminderMs;
    if (shouldRepeatOverlay) {
      this.resetOverlayDismissals();
      if (!document.hidden && !this.shouldSuppressOverlay()) {
        this.openAlert('');
        this.lastOverlayReminderAt = Date.now();
      } else {
        this.hideOverlay();
      }
    } else if (this.shouldSuppressOverlay()) {
      this.hideOverlay();
    }
    this.playAlertSound(top, Boolean(options.force));
    this.showBrowserNotification(top, Boolean(options.force));
  }
};

window.priorityAlertManager = priorityAlertManager;

const alertManager = {
  notifiedPlates: new Set(),
  init() {
    setInterval(() => this.checkTimes(), 30000);
  },
  checkTimes() {
    if (!loginManager.isLoggedIn()) return;
    const obs = document.getElementById('inpObs');
    const placa = document.getElementById('inpPlaca');
    if (!obs || !placa || !obs.value || !placa.value) return;
    const match = /\b([0-1]?[0-9]|2[0-3]):([0-5][0-9])\b/.exec(obs.value);
    if (!match) return;

    const now = new Date();
    const currTotal = now.getHours() * 60 + now.getMinutes();
    const schedTotal = parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
    if (currTotal >= (schedTotal - 15) && currTotal <= (schedTotal + 15) && !this.notifiedPlates.has(placa.value)) {
      this.notifiedPlates.add(placa.value);
      const dayKey = now.toISOString().slice(0, 10);
      priorityAlertManager.registerAlert({
        id: `carregamento_${stableHash(`${placa.value}|${match[0]}|${dayKey}`)}`,
        title: '🕒 HORA DE CARREGAMENTO',
        body: `A placa ${placa.value} está programada para agora (${match[0]}).`,
        severity: 'warning',
        actionKey: 'tab_op',
        actionLabel: '📋 Abrir liberação',
        source: 'agendamento'
      });
    }
  },
  resetStockAlertWindow() {
    toastManager.show('Os alertas legados foram removidos deste app.', 'info', 3200);
  },
  showNotification(title, body) {
    priorityAlertManager.registerAlert({
      id: `alerta_manual_${stableHash(`${title}|${body}`)}`,
      title,
      body,
      severity: 'info',
      source: 'manual'
    });
  }
};

const apiService = {
  globalLocSheetExists: null,
  globalLocSheetCheckedAt: 0,
  ensureSheetsConfig() {
    if (!config.api.key || !config.api.sheetId) {
      debugEngine.log("Configuração da API do Google ausente.", "error");
      throw new Error("API_KEY_NAO_CONFIGURADA");
    }
  },
  async fetchSheetData() {
    this.ensureSheetsConfig();
    if (!navigator.onLine) {
      debugEngine.log("Falha ao buscar Carregamentos: OFFLINE", "error");
      throw new Error("OFFLINE");
    }
    debugEngine.log("Buscando dados da aba de Carregamentos...", "info");
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${config.api.sheetId}/values/${encodeURIComponent(config.api.range)}?key=${config.api.key}&valueRenderOption=UNFORMATTED_VALUE`;
    const fetchPromise = fetch(url);
    const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout de 15s excedido")), config.api.timeoutMs));

    try {
      const res = await Promise.race([fetchPromise, timeoutPromise]);
      if (!res.ok) {
        let errMsg = `HTTP ${res.status}`;
        try {
          const errObj = await res.json();
          errMsg += ` - ${errObj.error.message}`;
        } catch(e){}
        debugEngine.log(`Erro API Carregamentos: ${errMsg}`, "error");
        debugEngine.updateAPIStatus(false, `(Erro ${res.status})`);
        throw new Error(errMsg);
      }
      const json = await res.json();
      debugEngine.log(`Carregamentos obtidos com sucesso (${json.values ? json.values.length : 0} linhas).`, "success");
      debugEngine.updateAPIStatus(true, "(200 OK)");
      
      if (json.values && json.values.length > 0) {
        debugEngine.log("=== DEBUG: Amostra das primeiras 3 linhas ===", "info");
        for (let i = 0; i < Math.min(3, json.values.length); i++) {
          const row = json.values[i];
          debugEngine.log(`Linha ${i}: ${JSON.stringify(row).substring(0, 200)}`, "info");
        }
        if (json.values[2]) {
          debugEngine.log(`Coluna A (dataRaw): "${json.values[2][0]}" (tipo: ${typeof json.values[2][0]})`, "info");
        }
      }
      
      return json.values || [];
    } catch(e) {
      debugEngine.log(`Exceção no fetchSheetData: ${e.message}`, "error");
      debugEngine.updateAPIStatus(false, "(Falha)");
      throw e;
    }
  },

  async fetchEstoqueData() {
    this.ensureSheetsConfig();
    if (!navigator.onLine) return [];
    debugEngine.log("Buscando dados de Estoque...", "info");
    try {
      const resInfo = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${config.api.estoqueSheetId}?key=${config.api.key}&fields=sheets.properties.title`);
      const infoJson = await resInfo.json();
      if (!infoJson.sheets) throw new Error("Aba de estoque não encontrada");

      const visible = infoJson.sheets;
      const now = new Date();
      const meses = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
      const t = visible.find(s => s.properties.title.toUpperCase() === `${meses[now.getMonth()]}/${now.getFullYear()}`) || visible[0];
      if (!t) throw new Error("Nenhuma aba visível no estoque");

      const resData = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${config.api.estoqueSheetId}/values/'${encodeURIComponent(t.properties.title)}'!A:F?key=${config.api.key}&valueRenderOption=UNFORMATTED_VALUE`);
      const dataJson = await resData.json();
      debugEngine.log("Estoque sincronizado.", "success");
      return dataJson.values || [];
    } catch (e) {
      debugEngine.log(`Erro ao buscar Estoque: ${e.message}`, "warn");
      return [];
    }
  },

  async fetchHistoricoData() {
    this.ensureSheetsConfig();
    if (!navigator.onLine) return [];
    debugEngine.log("Buscando Histórico de Cargas...", "info");
    try {
      const resInfo = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${config.api.historicoSheetId}?key=${config.api.key}&fields=sheets.properties.title`);
      const infoJson = await resInfo.json();
      if (!infoJson.sheets) return [];
      const targetSheet = infoJson.sheets[0];
      if (!targetSheet) return [];
      const resData = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${config.api.historicoSheetId}/values/'${encodeURIComponent(targetSheet.properties.title)}'!A:Z?key=${config.api.key}&valueRenderOption=FORMATTED_VALUE`);
      const json = await resData.json();
      return json.values || [];
    } catch (e) {
      debugEngine.log(`Erro ao buscar Histórico: ${e.message}`, "warn");
      return [];
    }
  },

  async fetchRncData() {
    this.ensureSheetsConfig();
    if (!navigator.onLine) return [];
    debugEngine.log("Buscando Reclamações (RNC)...", "info");
    try {
      const resInfo = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${config.api.rncSheetId}?key=${config.api.key}&fields=sheets.properties.title`);
      const infoJson = await resInfo.json();
      if (!infoJson.sheets) return [];
      const targetSheet = infoJson.sheets.find(s => s.properties.title.toUpperCase().includes('RNC')) || infoJson.sheets[0];
      if (!targetSheet) return [];
      const resData = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${config.api.rncSheetId}/values/'${encodeURIComponent(targetSheet.properties.title)}'!A:N?key=${config.api.key}&valueRenderOption=UNFORMATTED_VALUE`);
      const json = await resData.json();
      return json.values || [];
    } catch (e) {
      debugEngine.log(`Erro ao buscar RNC: ${e.message}`, "warn");
      return [];
    }
  },

  async sendDataToAppScript(payload) {
    if (!navigator.onLine) throw new Error("OFFLINE");
    debugEngine.log(`Enviando POST payload (${payload.tipo})...`, "info");
    const res = await fetch(config.api.scriptUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });
    if (!res.ok) throw new Error(`Falha HTTP ${res.status}`);
    const text = await res.text();
    let parsed = null;
    try {
      parsed = JSON.parse(text);
    } catch (_) {
      throw new Error("Resposta inválida do Apps Script.");
    }
    if (!parsed || !parsed.success) throw new Error(parsed?.message || "Apps Script retornou falha.");
    return parsed;
  }
};





