const mapPersistentStore = {
  db: null,
  dbName: 'CysyMapStoreV1',
  dirName: 'cysy-data',
  fileName: 'map-lotes-snapshot-v2.json',
  backendKey: 'cysyMapPersistentBackendV2',
  schemaKey: 'cysyMapPersistentSchemaV2',
  schemaVersion: 2,

  normalizeSnapshotPayload(payload = {}) {
    const locations = Array.isArray(payload.locations) ? payload.locations : [];
    const logs = Array.isArray(payload.logs) ? payload.logs : [];
    return {
      locations: locations.filter((row) => row && row.lote).map((row) => ({ lote: String(row.lote), raw: String(row.raw || '') })),
      logs: logs.filter((row) => row && row.lote).map((row) => ({ lote: String(row.lote), entry: row.entry || {} }))
    };
  },

  async init() {
    await this.initIndexedDb();
    if (this.supportsOpfs()) {
      await this.getOpfsDirectory(true);
    }
    return true;
  },

  async initIndexedDb() {
    if (this.db) return this.db;
    if (!('indexedDB' in window)) return null;
    return new Promise((resolve) => {
      try {
        const req = indexedDB.open(this.dbName, 1);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains('locations')) db.createObjectStore('locations', { keyPath: 'lote' });
          if (!db.objectStoreNames.contains('logs')) db.createObjectStore('logs', { keyPath: 'lote' });
        };
        req.onsuccess = (e) => { this.db = e.target.result; resolve(this.db); };
        req.onerror = () => resolve(null);
        req.onblocked = () => resolve(null);
      } catch (_) { resolve(null); }
    });
  },

  supportsOpfs() {
    return typeof navigator !== 'undefined' &&
      Boolean(navigator.storage && typeof navigator.storage.getDirectory === 'function');
  },

  async getOpfsDirectory(create = true) {
    if (!this.supportsOpfs()) return null;
    try {
      const root = await navigator.storage.getDirectory();
      return await root.getDirectoryHandle(this.dirName, { create });
    } catch (_) {
      return null;
    }
  },

  async loadFromOpfs() {
    const dir = await this.getOpfsDirectory(false);
    if (!dir) return null;
    try {
      const fileHandle = await dir.getFileHandle(this.fileName);
      const file = await fileHandle.getFile();
      const text = await file.text();
      if (!text) return null;
      const parsed = safeJsonParse(text, null);
      if (!parsed || typeof parsed !== 'object') return null;
      const normalized = this.normalizeSnapshotPayload(parsed);
      safeStorage.setItem(this.backendKey, 'opfs');
      safeStorage.setItem(this.schemaKey, String(this.schemaVersion));
      return { ...normalized, source: 'opfs' };
    } catch (_) {
      return null;
    }
  },

  async saveToOpfs(locDB = {}, locationLog = {}) {
    const dir = await this.getOpfsDirectory(true);
    if (!dir) return false;
    try {
      const fileHandle = await dir.getFileHandle(this.fileName, { create: true });
      const writable = await fileHandle.createWritable();
      const payload = {
        schemaVersion: this.schemaVersion,
        updatedAt: new Date().toISOString(),
        locations: Object.entries(locDB || {}).map(([lote, raw]) => ({ lote, raw: String(raw || '') })),
        logs: Object.entries(locationLog || {}).map(([lote, entry]) => ({ lote, entry }))
      };
      await writable.write(JSON.stringify(payload));
      await writable.close();
      safeStorage.setItem(this.backendKey, 'opfs');
      safeStorage.setItem(this.schemaKey, String(this.schemaVersion));
      return true;
    } catch (_) {
      return false;
    }
  },

  async loadFromIndexedDb() {
    const db = await this.initIndexedDb();
    if (!db) return null;
    const tx = db.transaction(['locations', 'logs'], 'readonly');
    const locReq = tx.objectStore('locations').getAll();
    const logReq = tx.objectStore('logs').getAll();
    return new Promise((resolve) => {
      tx.oncomplete = () => {
        const normalized = this.normalizeSnapshotPayload({
          locations: locReq.result || [],
          logs: logReq.result || []
        });
        resolve({ ...normalized, source: 'indexeddb' });
      };
      tx.onerror = () => resolve(null);
      tx.onabort = () => resolve(null);
    });
  },

  async saveToIndexedDb(locDB = {}, locationLog = {}) {
    const db = await this.initIndexedDb();
    if (!db) return false;
    return new Promise((resolve) => {
      try {
        const tx = db.transaction(['locations', 'logs'], 'readwrite');
        const locStore = tx.objectStore('locations');
        const logStore = tx.objectStore('logs');
        locStore.clear();
        logStore.clear();
        Object.entries(locDB || {}).forEach(([lote, raw]) => {
          locStore.put({ lote, raw: String(raw || '') });
        });
        Object.entries(locationLog || {}).forEach(([lote, entry]) => {
          logStore.put({ lote, entry });
        });
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => resolve(false);
        tx.onabort = () => resolve(false);
      } catch (_) { resolve(false); }
    });
  },

  async loadSnapshot() {
    const opfs = await this.loadFromOpfs();
    if (opfs) return opfs;
    const indexed = await this.loadFromIndexedDb();
    if (indexed) {
      safeStorage.setItem(this.backendKey, 'indexeddb');
      safeStorage.setItem(this.schemaKey, String(this.schemaVersion));
      return indexed;
    }
    return { locations: [], logs: [], source: 'none' };
  },

  async saveSnapshot(locDB = {}, locationLog = {}) {
    const opfsSaved = await this.saveToOpfs(locDB, locationLog);
    const indexedSaved = await this.saveToIndexedDb(locDB, locationLog);
    return Boolean(opfsSaved || indexedSaved);
  }
};

const mapController = {
  locDB: safeJsonParse(safeStorage.getItem('cysyLocations', '{}'), {}),
  locationLogKey: 'cysyLocationLogV1',
  locationLog: safeJsonParse(safeStorage.getItem('cysyLocationLogV1', '{}'), {}),
  pendingSharedConfirmationKey: 'cysyMapPendingSharedConfirmationsV1',
  pendingSharedConfirmations: safeJsonParse(safeStorage.getItem('cysyMapPendingSharedConfirmationsV1', '{}'), {}),
  registrationTargetKey: 'cysyMapRegistrationTargetV1',
  selectedRegistrationLote: safeStorage.getItem('cysyMapRegistrationTargetV1', ''),
  manuallyDeletedLotesKey: 'cysyMapDeletedLotesV1',
  manuallyDeletedLotes: safeJsonParse(safeStorage.getItem('cysyMapDeletedLotesV1', '[]'), []),
  mapResetVersionKey: 'cysyMapResetVersion',
  mapResetVersionValue: '2026-04-02-reset-lotes-v2',
  map: null,
  satelliteLayer: null,
  streetLayer: null,
  activeBaseLayer: null,
  markers: {},
  resizeObs: null,
  userMarker: null,
  companyMarker: null,
  companyAreaLayer: null,
  isFullscreen: false,
  currentBounds: null,
  resizeBound: false,
  sharedSyncTimer: null,
  sharedSyncInFlight: false,
  lastSharedSyncAt: '',
  lastSharedSyncError: '',
  lastSharedFetchAt: 0,
  lastSharedRows: [],
  tileFailureCount: 0,
  usingFallbackLayer: false,
  connectionLayers: [],
  lastRenderedConnectionsKey: '',
  decorationLayers: [],
  selectedLote: safeStorage.getItem('cysySelectedMapLote', ''),
  visualMode: safeStorage.getItem('cysyMapVisualMode', 'clean'),
  lastRenderedMapped: [],
  lastRenderedPending: [],
  lastRenderedProximity: { edges: [], averageNearestDistance: 0, threshold: 0 },
  statusFilter: safeStorage.getItem('cysyMapStatusFilter', 'all'),
  sheetTab: safeStorage.getItem('cysyMapSheetTab', 'lotes'),
  sheetExpanded: (() => {
    const mobile = typeof window !== 'undefined' && window.matchMedia && window.matchMedia('(max-width: 768px)').matches;
    if (mobile) return true;
    const saved = safeStorage.getItem('cysyMapSheetExpanded', '');
    if (saved === 'true') return true;
    if (saved === 'false') return false;
    return !mobile;
  })(),
  mapStageExpanded: (() => {
    const mobile = typeof window !== 'undefined' && window.matchMedia && window.matchMedia('(max-width: 768px)').matches;
    if (mobile) return false;
    const saved = safeStorage.getItem('cysyMapStageExpanded', '');
    if (saved === 'true') return true;
    if (saved === 'false') return false;
    return !mobile;
  })(),
  offlineTileCacheName: 'cysy-map-tiles-v20260406r02',
  offlineMapStateKey: 'cysyOfflineMapStateV2',
  offlineSaveTimer: null,
  offlineSaveInFlight: false,
  layoutRefreshRaf: 0,
  pendingRenderTimer: 0,
  searchRefreshTimer: 0,
  lastRenderAt: 0,
  sharedConfirmationTimeoutMs: 20000,
  sharedConfirmationIntervalMs: 1400,
  advancedActionsExpandedKey: 'cysyMapAdvancedActionsExpanded',
  advancedActionsExpanded: (() => {
    const saved = safeStorage.getItem('cysyMapAdvancedActionsExpanded', '');
    return saved === 'true';
  })(),
  overlaysExpanded: (() => {
    const saved = safeStorage.getItem('cysyMapOverlaysExpanded', '');
    if (saved === 'true') return true;
    if (saved === 'false') return false;
    return false;
  })(),
  currentBaseLayerKind: 'street',

  getMapContainer() {
    return document.getElementById('patioMapContainer');
  },

  isCompactMobile() {
    return typeof window !== 'undefined' &&
      window.matchMedia &&
      (window.matchMedia('(max-width: 768px)').matches || window.matchMedia('(pointer: coarse)').matches);
  },

  getMapDock() {
    return document.getElementById('mapDock');
  },

  getFullscreenOverlay() {
    return document.getElementById('mapFullscreenOverlay');
  },

  getFullscreenMount() {
    return document.getElementById('mapFullscreenMount');
  },

  getCompanyAreaCoords() {
    return Array.isArray(config.base.areaPolygon) ? config.base.areaPolygon.filter((row) => Array.isArray(row) && row.length === 2) : [];
  },

  getBaseFocusBounds() {
    const areaCoords = this.getCompanyAreaCoords();
    return areaCoords.length >= 3 ? L.latLngBounds(areaCoords) : null;
  },

  normalizeSharedCoordinate(value) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return null;
    if (Math.abs(numeric) > 180) return numeric / 1000000;
    return numeric;
  },

  normalizeSharedUpdatedAt(value) {
    if (value === undefined || value === null || value === '') return '';
    const numeric = Number(value);
    if (Number.isFinite(numeric) && numeric > 30000 && numeric < 90000) {
      const excelEpoch = Date.UTC(1899, 11, 30);
      const millis = Math.round(numeric * 86400000);
      return new Date(excelEpoch + millis).toLocaleString('pt-BR');
    }
    return String(value);
  },

  getSortableTimestamp(value) {
    if (!value && value !== 0) return 0;
    const numeric = Number(value);
    if (Number.isFinite(numeric) && numeric > 30000 && numeric < 90000) {
      const excelEpoch = Date.UTC(1899, 11, 30);
      return excelEpoch + Math.round(numeric * 86400000);
    }
    const parsed = Date.parse(String(value));
    if (Number.isFinite(parsed)) return parsed;
    const brMatch = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:,?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(String(value).trim());
    if (brMatch) {
      const [, dd, mm, yyyy, hh = '0', min = '0', ss = '0'] = brMatch;
      return new Date(Number(yyyy), Number(mm) - 1, Number(dd), Number(hh), Number(min), Number(ss)).getTime();
    }
    return 0;
  },

  getVisibleLotes() {
    return [...(uiBuilder.lotesFlat || [])].filter((item) => item && !this.shouldRemoveLoteFromMap(item.lote)).sort((a, b) =>
      String(a.lote || '').localeCompare(String(b.lote || ''), 'pt-BR')
    );
  },

  getAllLotes() {
    return [...(uiBuilder.allLotesFlat || uiBuilder.lotesFlat || [])].filter(Boolean).sort((a, b) =>
      String(a.lote || '').localeCompare(String(b.lote || ''), 'pt-BR')
    );
  },

  getNormalizedLoteKey(lote) {
    return normalizeName(String(lote || '').trim());
  },

  shouldIgnoreLote(lote) {
    const key = this.getNormalizedLoteKey(lote);
    return key.startsWith('TESTE_') || key.startsWith('TESTE ');
  },

  getManualDeletedSet() {
    const source = Array.isArray(this.manuallyDeletedLotes) ? this.manuallyDeletedLotes : [];
    return new Set(source.map((item) => this.getNormalizedLoteKey(item)).filter(Boolean));
  },

  isManuallyDeletedLote(lote) {
    if (!lote) return false;
    return this.getManualDeletedSet().has(this.getNormalizedLoteKey(lote));
  },

  markLoteAsManuallyDeleted(lote) {
    const key = this.getNormalizedLoteKey(lote);
    if (!key) return;
    const next = this.getManualDeletedSet();
    next.add(key);
    this.manuallyDeletedLotes = [...next];
  },

  clearManualDeletedMark(lote) {
    const key = this.getNormalizedLoteKey(lote);
    if (!key) return;
    const next = this.getManualDeletedSet();
    if (!next.delete(key)) return;
    this.manuallyDeletedLotes = [...next];
  },

  getPendingSharedConfirmation(lote) {
    const key = this.getNormalizedLoteKey(lote);
    if (!key) return null;
    const source = this.pendingSharedConfirmations || {};
    return source[key] || null;
  },

  setPendingSharedConfirmation(lote, lat, lon, description = '', updatedAt = '', updatedBy = '') {
    const key = this.getNormalizedLoteKey(lote);
    const numericLat = Number(lat);
    const numericLon = Number(lon);
    if (!key || !Number.isFinite(numericLat) || !Number.isFinite(numericLon)) return;
    this.pendingSharedConfirmations = {
      ...(this.pendingSharedConfirmations || {}),
      [key]: {
        lote: String(lote || '').trim(),
        lat: Number(numericLat.toFixed(6)),
        lon: Number(numericLon.toFixed(6)),
        description: String(description || '').trim(),
        updatedAt: String(updatedAt || '').trim(),
        updatedBy: String(updatedBy || '').trim()
      }
    };
  },

  clearPendingSharedConfirmation(lote) {
    const key = this.getNormalizedLoteKey(lote);
    if (!key || !this.pendingSharedConfirmations || !this.pendingSharedConfirmations[key]) return;
    delete this.pendingSharedConfirmations[key];
  },

  shouldPreservePendingLocation(lote, lat, lon, toleranceMeters = 6) {
    const pending = this.getPendingSharedConfirmation(lote);
    const sharedLat = Number(lat);
    const sharedLon = Number(lon);
    if (!pending) return false;
    if (!Number.isFinite(sharedLat) || !Number.isFinite(sharedLon)) return true;
    const distance = this.calculateDistanceMeters(pending.lat, pending.lon, sharedLat, sharedLon);
    if (distance <= Math.max(2, Number(toleranceMeters || 6))) {
      this.clearPendingSharedConfirmation(lote);
      return false;
    }
    return true;
  },

  shouldRemoveLoteFromMap(lote) {
    if (!lote || this.shouldIgnoreLote(lote) || this.isManuallyDeletedLote(lote)) return true;
    const info = this.getLoteInfo(lote);
    if (!info) return false;
    return Number(info.saldo || 0) <= 0;
  },

  purgeUnmappableLocationEntries() {
    let changed = false;
    Object.keys(this.locDB || {}).forEach((lote) => {
      if (!this.shouldRemoveLoteFromMap(lote)) return;
      delete this.locDB[lote];
      this.clearPendingSharedConfirmation(lote);
      changed = true;
    });
    Object.keys(this.locationLog || {}).forEach((lote) => {
      if (!this.shouldRemoveLoteFromMap(lote)) return;
      delete this.locationLog[lote];
      this.clearPendingSharedConfirmation(lote);
      changed = true;
    });
    if (changed) this.persistLocationState();
  },

  applyRequestedMapReset() {
    const appliedVersion = safeStorage.getItem(this.mapResetVersionKey, '');
    if (appliedVersion === this.mapResetVersionValue) return;
    const knownRegisteredLotes = new Set([
      ...Object.keys(this.locDB || {}),
      ...Object.keys(this.locationLog || {})
    ].map((item) => this.getNormalizedLoteKey(item)).filter(Boolean));
    this.locDB = {};
    this.locationLog = {};
    this.pendingSharedConfirmations = {};
    this.manuallyDeletedLotes = [...knownRegisteredLotes];
    safeStorage.setItem(this.mapResetVersionKey, this.mapResetVersionValue);
    this.persistLocationState();
  },

  async clearMapDataDeep() {
    const knownRegisteredLotes = new Set([
      ...Object.keys(this.locDB || {}),
      ...Object.keys(this.locationLog || {}),
      ...(Array.isArray(this.lastRenderedMapped) ? this.lastRenderedMapped.map((item) => item?.lote) : []),
      ...(Array.isArray(this.lastSharedRows) ? this.lastSharedRows.map((row) => row?.[0]) : [])
    ].map((item) => this.getNormalizedLoteKey(item)).filter(Boolean));

    this.locDB = {};
    this.locationLog = {};
    this.pendingSharedConfirmations = {};
    this.manuallyDeletedLotes = [...knownRegisteredLotes];
    this.selectedLote = '';
    this.selectedRegistrationLote = '';
    safeStorage.removeItem('cysySelectedMapLote');
    safeStorage.removeItem(this.registrationTargetKey);
    safeStorage.removeItem(this.offlineMapStateKey);
    this.persistLocationState();

    try {
      const pending = await dbManager.getAllPending();
      const filtered = (pending || []).filter((item) => String(item.type || '').toUpperCase() !== 'LOCALIZACAO');
      await dbManager.replaceQueue(filtered);
    } catch (_) {}

    try {
      if ('caches' in window && this.offlineTileCacheName) {
        await caches.delete(this.offlineTileCacheName);
      }
    } catch (_) {}

    this.lastRenderedMapped = [];
    this.lastRenderedProximity = { edges: [], averageNearestDistance: 0, threshold: 0 };
    this.lastSharedRows = [];
    this.renderPins(true);
    this.refreshMapLayout();
  },

  purgeIgnoredLocationEntries() {
    let changed = false;
    Object.keys(this.locDB || {}).forEach((lote) => {
      if (!this.shouldIgnoreLote(lote)) return;
      delete this.locDB[lote];
      this.clearPendingSharedConfirmation(lote);
      changed = true;
    });
    Object.keys(this.locationLog || {}).forEach((lote) => {
      if (!this.shouldIgnoreLote(lote)) return;
      delete this.locationLog[lote];
      this.clearPendingSharedConfirmation(lote);
      changed = true;
    });
    if (changed) this.persistLocationState();
  },

  clearLotesCadastrados() {
    const countLoc = Object.keys(this.locDB || {}).length;
    const countLog = Object.keys(this.locationLog || {}).length;
    if (countLoc === 0 && countLog === 0) {
      toastManager.show('Nenhum lote cadastrado para limpar.', 'info');
      return;
    }
    const confirmed = window.confirm(`Tem certeza que deseja limpar todos os lotes cadastrados no mapa?\n\nIsso removerá:\n• ${countLoc} localização(ões)\n• ${countLog} log(s) de localização`);
    if (!confirmed) return;
    this.locDB = {};
    this.locationLog = {};
    this.pendingSharedConfirmations = {};
    this.manuallyDeletedLotes = [];
    this.selectedRegistrationLote = '';
    safeStorage.removeItem(this.registrationTargetKey);
    this.persistLocationState();
    this.renderPins();
    toastManager.show(`Lotes limpos: ${countLoc} localizações e ${countLog} logs removidos.`, 'success');
  },

  async clearMapDataWithConfirmation() {
    const countLoc = Object.keys(this.locDB || {}).length;
    const countLog = Object.keys(this.locationLog || {}).length;
    const countShared = Array.isArray(this.lastSharedRows) ? Math.max(0, this.lastSharedRows.length - 1) : 0;
    const confirmed = window.confirm(
      `Limpeza profunda do mapa?\n\nIsso irá:\n• apagar coordenadas locais\n• apagar log local do mapa\n• limpar fila offline de localizações\n• impedir que os lotes já registrados reapareçam automaticamente\n\nRegistros atuais detectados:\n• ${countLoc} localização(ões) locais\n• ${countLog} log(s) locais\n• ${countShared} linha(s) compartilhadas em memória\n\nDeseja continuar?`
    );
    if (!confirmed) return;
    await this.clearMapDataDeep();
    toastManager.show('Limpeza profunda do mapa concluída. Você já pode refazer os cadastros.', 'success', 5200);
  },

  deleteSingleLote(lote) {
    if (!lote) return false;
    const loteName = String(lote).trim();
    const logEntryRaw = this.locationLog?.[loteName];
    const logEntry = logEntryRaw ? this.normalizeLocationLogEntry(logEntryRaw) : null;
    const produto = logEntry?.product || '';
    const saldo = (logEntry?.saldo > 0) ? (formatTons(logEntry.saldo) + ' t') : 'sem saldo registrado';
    const hasLoc = Boolean(this.locDB?.[loteName]);
    const hasLog = Boolean(this.locationLog?.[loteName]);
    if (!hasLoc && !hasLog) {
      toastManager.show('Este lote não possui dados de localização registrados.', 'info');
      return false;
    }
    const step1 = window.confirm(
      '\u26A0\uFE0F AVISO DE RISCO \u2014 Exclus\u00e3o de localiza\u00e7\u00e3o de lote\n\n' +
      'Lote: ' + loteName + '\n' +
      'Produto: ' + (produto || 'N\u00e3o informado') + '\n' +
      'Saldo: ' + saldo + '\n\n' +
      'Esta a\u00e7\u00e3o remover\u00e1 as coordenadas GPS deste lote do mapa.\n' +
      'O lote continuar\u00e1 existindo no estoque, mas desaparecer\u00e1 do mapa at\u00e9 ser regeorreferenciado.\n\n' +
      'Deseja continuar?'
    );
    if (!step1) {
      toastManager.show('Exclus\u00e3o cancelada.', 'info');
      return false;
    }
    const step2 = window.confirm(
      '\uD83D\uDD34 CONFIRMA\u00c7\u00c3O FINAL\n\n' +
      'Voc\u00ea est\u00e1 prestes a apagar definitivamente a localiza\u00e7\u00e3o do lote:\n\n' +
      '"' + loteName + '"\n\n' +
      'Esta a\u00e7\u00e3o n\u00e3o pode ser desfeita.\n' +
      'Clique em OK para confirmar a exclus\u00e3o.'
    );
    if (!step2) {
      toastManager.show('Exclus\u00e3o cancelada.', 'info');
      return false;
    }
    let removed = false;
    if (hasLoc) { delete this.locDB[loteName]; removed = true; }
    if (hasLog) { delete this.locationLog[loteName]; removed = true; }
    if (removed) {
      this.clearPendingSharedConfirmation(loteName);
      this.markLoteAsManuallyDeleted(loteName);
      if (this.getNormalizedLoteKey(this.selectedRegistrationLote) === this.getNormalizedLoteKey(loteName)) {
        this.selectedRegistrationLote = '';
        safeStorage.removeItem(this.registrationTargetKey);
      }
      this.persistLocationState();
      this.renderPins();
      this.purgeQueuedLocationEntriesForLote(loteName).catch(() => {});
      toastManager.show('Localiza\u00e7\u00e3o do lote "' + loteName + '" removida com sucesso.', 'success');
      return true;
    }
    return false;
  },

  async purgeQueuedLocationEntriesForLote(lote) {
    const loteKey = this.getNormalizedLoteKey(lote);
    if (!loteKey) return 0;
    let removed = 0;
    try {
      const records = await dbManager.getAllRecords({ includeSent: false });
      for (const row of records || []) {
        if (String(row?.type || '').toUpperCase() !== 'LOCALIZACAO') continue;
        if (this.getNormalizedLoteKey(row?.payload?.lote) !== loteKey) continue;
        await dbManager.deletePending(row.id);
        removed += 1;
      }
      if (removed > 0) syncManager.refreshSyncView();
    } catch (_) {}
    return removed;
  },

  persistLocationState() {
    mapPersistentStore.saveSnapshot(this.locDB || {}, this.locationLog || {}).catch(() => {});
    safeStorage.setItem('cysyLocations', JSON.stringify(this.locDB || {}));
    safeStorage.setItem(this.locationLogKey, JSON.stringify(this.locationLog || {}));
    safeStorage.setItem(this.pendingSharedConfirmationKey, JSON.stringify(this.pendingSharedConfirmations || {}));
    safeStorage.setItem(this.manuallyDeletedLotesKey, JSON.stringify(this.manuallyDeletedLotes || []));
  },

  async hydrateFromPersistent() {
    try {
      const snap = await mapPersistentStore.loadSnapshot();
      const hasPersistent = (snap.locations && snap.locations.length) || (snap.logs && snap.logs.length);
      if (hasPersistent) {
        const locObj = {};
        (snap.locations || []).forEach((row) => { if (row && row.lote) locObj[row.lote] = row.raw; });
        const logObj = {};
        (snap.logs || []).forEach((row) => { if (row && row.lote) logObj[row.lote] = row.entry; });
        this.locDB = locObj;
        this.locationLog = logObj;
        safeStorage.setItem('cysyLocations', JSON.stringify(this.locDB || {}));
        safeStorage.setItem(this.locationLogKey, JSON.stringify(this.locationLog || {}));
      } else if (Object.keys(this.locDB || {}).length > 0 || Object.keys(this.locationLog || {}).length > 0) {
        // Migra do localStorage para o armazenamento persistente.
        await mapPersistentStore.saveSnapshot(this.locDB || {}, this.locationLog || {});
      }
    } catch (_) {}
  },

  getLoteInfo(lote) {
    const all = this.getAllLotes();
    return all.find((item) => String(item.lote || '').trim() === String(lote || '').trim()) || null;
  },

  getLocationLogEntryByLote(lote) {
    const loteName = String(lote || '').trim();
    if (!loteName) return null;
    const direct = this.locationLog?.[loteName];
    if (direct) return this.normalizeLocationLogEntry(direct);
    const loteKey = this.getNormalizedLoteKey(loteName);
    const match = Object.values(this.locationLog || {}).find((row) =>
      this.getNormalizedLoteKey(row?.lote) === loteKey
    );
    return match ? this.normalizeLocationLogEntry(match) : null;
  },

  formatLocationMoment(value) {
    if (!value && value !== 0) return '';
    const parsed = this.getSortableTimestamp(value);
    if (parsed > 0) return new Date(parsed).toLocaleString('pt-BR');
    return String(value || '');
  },

  getLocationSyncMeta(entry = {}) {
    const pendingConfirmation = entry?.lote ? this.getPendingSharedConfirmation(entry.lote) : null;
    let status = String(entry.syncStatus || '').trim().toLowerCase();
    if (!status) {
      if (pendingConfirmation) status = 'aguardando_confirmacao_global';
      else if (String(entry.source || '').trim().toLowerCase() === 'compartilhado') status = 'sincronizado';
    }
    const statusMap = {
      capturando: {
        code: 'capturando',
        label: 'Capturando GPS',
        className: 'map-sync-pill map-sync-pill--capturing',
        summary: 'O app está obtendo a coordenada mais precisa disponível.'
      },
      enviando: {
        code: 'enviando',
        label: 'Enviando',
        className: 'map-sync-pill map-sync-pill--sending',
        summary: 'O lote já foi capturado localmente e está sendo enviado ao endpoint.'
      },
      aguardando_confirmacao_global: {
        code: 'aguardando_confirmacao_global',
        label: 'Aguardando confirmação global',
        className: 'map-sync-pill map-sync-pill--waiting',
        summary: 'O envio foi aceito, mas o mapa compartilhado ainda não confirmou o lote para todos.'
      },
      sincronizado: {
        code: 'sincronizado',
        label: 'Sincronizado para todos',
        className: 'map-sync-pill map-sync-pill--ok',
        summary: 'O lote já apareceu no mapa compartilhado e foi confirmado na leitura global.'
      },
      pendente_com_falha: {
        code: 'pendente_com_falha',
        label: 'Pendente com falha',
        className: 'map-sync-pill map-sync-pill--error',
        summary: 'O lote ficou pendente e precisa de nova confirmação da sincronização compartilhada.'
      }
    };
    return statusMap[status] || {
      code: 'sincronizado',
      label: 'Sincronizado para todos',
      className: 'map-sync-pill map-sync-pill--ok',
      summary: 'O lote está pronto no mapa compartilhado.'
    };
  },

  normalizeLocationLogEntry(entry = {}) {
    const lote = String(entry.lote || '').trim();
    if (!lote || this.shouldRemoveLoteFromMap(lote)) return null;
    const lat = entry.lat !== undefined && entry.lat !== null && entry.lat !== '' ? Number(entry.lat) : null;
    const lon = entry.lon !== undefined && entry.lon !== null && entry.lon !== '' ? Number(entry.lon) : null;
    const info = this.getLoteInfo(lote);
    return {
      lote,
      lat: Number.isFinite(lat) ? lat : null,
      lon: Number.isFinite(lon) ? lon : null,
      description: String(entry.description || entry.desc || this.extrairDescricao(this.locDB[lote]) || '').trim(),
      updatedBy: String(entry.updatedBy || entry.usuario || '').trim(),
      updatedAt: String(entry.updatedAt || entry.dataHora || entry.timestamp || '').trim(),
      source: String(entry.source || 'local').trim(),
      product: String(entry.product || info?.produto || '').trim(),
      saldo: Number(entry.saldo ?? info?.saldo ?? 0),
      syncStatus: String(entry.syncStatus || '').trim(),
      syncMessage: String(entry.syncMessage || '').trim(),
      lastAttemptAt: String(entry.lastAttemptAt || '').trim(),
      confirmationSource: String(entry.confirmationSource || '').trim(),
      syncId: String(entry.syncId || '').trim(),
      localId: String(entry.localId || '').trim(),
      fingerprint: String(entry.fingerprint || '').trim()
    };
  },

  upsertLocationLogEntry(entry = {}) {
    const normalized = this.normalizeLocationLogEntry(entry);
    if (!normalized) return;
    const prev = this.locationLog[normalized.lote] || {};
    this.locationLog[normalized.lote] = {
      ...prev,
      ...normalized,
      product: normalized.product || prev.product || '',
      saldo: Number.isFinite(normalized.saldo) ? normalized.saldo : Number(prev.saldo || 0),
      syncStatus: normalized.syncStatus || prev.syncStatus || '',
      syncMessage: normalized.syncMessage || prev.syncMessage || '',
      lastAttemptAt: normalized.lastAttemptAt || prev.lastAttemptAt || '',
      confirmationSource: normalized.confirmationSource || prev.confirmationSource || '',
      syncId: normalized.syncId || prev.syncId || '',
      localId: normalized.localId || prev.localId || '',
      fingerprint: normalized.fingerprint || prev.fingerprint || ''
    };
  },

  setLocationSyncState(lote, patch = {}, options = {}) {
    const loteName = String(lote || '').trim();
    if (!loteName) return null;
    const info = this.getLoteInfo(loteName);
    const existing = this.getLocationLogEntryByLote(loteName) || {
      lote: loteName,
      lat: null,
      lon: null,
      description: '',
      updatedBy: '',
      updatedAt: '',
      source: 'local',
      product: String(info?.produto || '').trim(),
      saldo: Number(info?.saldo ?? 0),
      syncStatus: '',
      syncMessage: '',
      lastAttemptAt: '',
      confirmationSource: '',
      syncId: '',
      localId: '',
      fingerprint: ''
    };
    this.upsertLocationLogEntry({
      ...existing,
      ...patch,
      lote: loteName
    });
    if (options.persist !== false) this.persistLocationState();
    if (options.render === 'full') this.renderPins(true);
    else if (options.render === 'panel') this.renderPanelsFromSnapshot(Boolean(options.refreshSearch));
    return this.getLocationLogEntryByLote(loteName);
  },

  rebuildLocationLogFromLocDb() {
    let changed = false;
    Object.keys(this.locDB || {}).forEach((lote) => {
      if (this.shouldRemoveLoteFromMap(lote)) return;
      const parsed = this.parseLocationEntry(this.locDB[lote]);
      if (!parsed) return;
      const existing = this.locationLog[lote] || {};
      const before = JSON.stringify(existing || {});
      this.upsertLocationLogEntry({
        lote,
        lat: parsed.lat,
        lon: parsed.lon,
        description: parsed.description,
        updatedAt: existing.updatedAt || '',
        updatedBy: existing.updatedBy || '',
        source: existing.source || 'cache_local',
        product: existing.product || '',
        saldo: existing.saldo,
        syncStatus: existing.syncStatus || '',
        syncMessage: existing.syncMessage || '',
        lastAttemptAt: existing.lastAttemptAt || '',
        confirmationSource: existing.confirmationSource || '',
        syncId: existing.syncId || '',
        localId: existing.localId || '',
        fingerprint: existing.fingerprint || ''
      });
      if (!changed && before !== JSON.stringify(this.locationLog[lote] || {})) changed = true;
    });
    if (changed) this.persistLocationState();
  },

  getLocationLogEntries(searchRaw = '') {
    const term = String(searchRaw || '').trim().toUpperCase();
    const entries = Object.values(this.locationLog || {})
      .map((row) => this.normalizeLocationLogEntry(row))
      .filter(Boolean)
      .sort((a, b) => {
        const ad = this.getSortableTimestamp(a.updatedAt || 0);
        const bd = this.getSortableTimestamp(b.updatedAt || 0);
        return bd - ad;
      });
    if (!term) return entries;
    return entries.filter((item) => {
      const hay = [
        item.lote,
        item.product,
        item.description,
        item.updatedBy,
        item.updatedAt
      ].join(' ').toUpperCase();
      return hay.includes(term);
    });
  },

  renderLocationLog(searchRaw = '') {
    const container = document.getElementById('mapLocationLogList');
    const countEl = document.getElementById('mapLocationLogCount');
    if (!container) return;
    const entries = this.getLocationLogEntries(searchRaw);
    if (countEl) countEl.textContent = String(entries.length);
    container.innerHTML = entries.length > 0
      ? entries.map((item) => {
          const loteToken = encodeURIComponent(item.lote);
          const hasCoords = Number.isFinite(item.lat) && Number.isFinite(item.lon);
          const subtitle = item.product || 'Lote registrado sem produto ativo no estoque atual';
          const updatedMeta = item.updatedAt ? `Atualizado em ${escapeHTML(item.updatedAt)}` : 'Data de atualização indisponível';
          const byMeta = item.updatedBy ? ` por ${escapeHTML(item.updatedBy)}` : '';
          const saldoLabel = item.saldo > 0 ? `${formatTons(item.saldo)} t` : 'LOG';
          const syncMeta = this.getLocationSyncMeta(item);
          const isRegistrationTarget = this.getNormalizedLoteKey(this.selectedRegistrationLote) === this.getNormalizedLoteKey(item.lote);
          return `<article class="map-lote-card">
            <div class="map-lote-top">
              <div>
                <div class="map-lote-code">${escapeHTML(item.lote)}</div>
                <div class="map-lote-product">${escapeHTML(subtitle)}</div>
              </div>
              <div style="display:flex;flex-direction:column;align-items:flex-end;gap:6px;">
                <div class="map-lote-saldo">${escapeHTML(saldoLabel)}</div>
                <button type="button"
                  onclick="mapController.deleteSingleLote(decodeURIComponent('${loteToken}'))"
                  title="Apagar localização deste lote do mapa"
                  style="background:rgba(220,38,38,0.1);border:1px solid rgba(220,38,38,0.35);color:#DC2626;border-radius:8px;padding:4px 10px;font-size:11px;font-weight:800;cursor:pointer;min-height:28px;box-shadow:none;white-space:nowrap;"
                  onmouseover="this.style.background='rgba(220,38,38,0.22)'"
                  onmouseout="this.style.background='rgba(220,38,38,0.1)'">🗑️ Apagar</button>
              </div>
            </div>
            <div class="map-lote-location">
              ${hasCoords ? `Coordenadas: ${Number(item.lat).toFixed(6)}, ${Number(item.lon).toFixed(6)}<br>` : ''}
              ${escapeHTML(item.description || 'Sem descrição complementar')}<br>
              <span class="${syncMeta.className}">${escapeHTML(syncMeta.label)}</span><br>
              <span style="color:#475569;">${updatedMeta}${byMeta}</span>
            </div>
            <div class="map-lote-actions">
              <button type="button" class="${isRegistrationTarget ? 'btn-success' : ''}" onclick="mapController.setRegistrationTarget(decodeURIComponent('${loteToken}'), { refreshSearch: true })">${isRegistrationTarget ? '✅ Lote selecionado' : 'Selecionar lote'}</button>
              ${hasCoords
                ? `<button type="button" onclick="mapController.selectLote(decodeURIComponent('${loteToken}'), { center: true })">Focar lote</button>
                   <button type="button" onclick="mapController.confirmNavigationToLote(decodeURIComponent('${loteToken}'))">Ir até o local</button>`
                : `<button type="button" onclick="mapController.capturarGPSEditar(decodeURIComponent('${loteToken}'))">Capturar GPS</button>`
              }
            </div>
          </article>`;
        }).join('')
      : `<div class="map-empty-card">Nenhuma localização registrada ainda. Os lotes salvos aparecerão aqui para todos os usuários.</div>`;
  },

  getLocationString(lat, lon, desc) {
    return (lat && lon) ? `📍 Lat: ${lat}, Lon: ${lon} \n${desc}` : String(desc || '');
  },

  parseLocationEntry(locRaw) {
    const raw = String(locRaw || '').trim();
    if (!raw) return null;
    const match = /Lat:\s*([-0-9.]+),\s*Lon:\s*([-0-9.]+)/i.exec(raw);
    if (!match) return null;
    const lat = Number.parseFloat(match[1]);
    const lon = Number.parseFloat(match[2]);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    return {
      lat,
      lon,
      description: this.extrairDescricao(raw) || 'Localização georreferenciada'
    };
  },

  getMarkerLabel(lote, fallbackIndex = 1) {
    const raw = String(lote || '').trim();
    const compact = raw.replace(/\s+/g, '');
    const digits = (compact.match(/\d+/g) || []).join('');
    const label = digits ? digits.slice(-4) : compact.slice(-4);
    return String(label || fallbackIndex).toUpperCase();
  },

  getVisualModes() {
    return ['clean', 'informative', 'detailed'];
  },

  getVisualModeLabel(mode) {
    const labels = {
      clean: 'Modo limpo',
      informative: 'Modo informativo',
      detailed: 'Modo detalhado'
    };
    return labels[String(mode || '').toLowerCase()] || 'Modo limpo';
  },

  setVisualMode(mode = 'clean', silent = false) {
    const normalized = this.getVisualModes().includes(mode) ? mode : 'clean';
    this.visualMode = normalized;
    safeStorage.setItem('cysyMapVisualMode', normalized);
    this.updateVisualModeControls();
    this.renderPins();
    if (!silent) {
      toastManager.show(`${this.getVisualModeLabel(normalized)} ativado.`, 'info');
    }
  },

  updateVisualModeControls() {
    this.getVisualModes().forEach((mode) => {
      const button = document.getElementById(`mapMode-${mode}`);
      if (!button) return;
      const active = this.visualMode === mode;
      button.classList.toggle('active', active);
      button.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
  },

  toggleMapOverlays(forceState = null) {
    this.overlaysExpanded = typeof forceState === 'boolean' ? forceState : !this.overlaysExpanded;
    safeStorage.setItem('cysyMapOverlaysExpanded', String(this.overlaysExpanded));
    this.updateOverlayUi();
    this.refreshMapLayout();
  },

  updateOverlayUi() {
    const canvas = document.querySelector('.map-ops-canvas');
    const toggle = document.getElementById('mapOverlayToggle');
    if (canvas) {
      canvas.classList.toggle('map-overlays-collapsed', !this.overlaysExpanded);
      canvas.classList.toggle('map-overlays-expanded', this.overlaysExpanded);
    }
    if (toggle) {
      toggle.textContent = this.overlaysExpanded ? 'Ocultar controles' : 'Mostrar controles';
      toggle.setAttribute('aria-pressed', this.overlaysExpanded ? 'true' : 'false');
    }
  },

  updateAdvancedActionsUi() {
    const panel = document.getElementById('mapAdvancedActions');
    const button = document.getElementById('mapAdvancedActionsToggle');
    if (panel) panel.hidden = !this.advancedActionsExpanded;
    if (button) {
      button.textContent = this.advancedActionsExpanded ? 'Menos opções' : 'Mais opções';
      button.setAttribute('aria-expanded', this.advancedActionsExpanded ? 'true' : 'false');
    }
  },

  toggleAdvancedActions(forceState = null) {
    this.advancedActionsExpanded = typeof forceState === 'boolean' ? forceState : !this.advancedActionsExpanded;
    safeStorage.setItem(this.advancedActionsExpandedKey, String(this.advancedActionsExpanded));
    this.updateAdvancedActionsUi();
    if (this.mapStageExpanded) this.refreshMapLayout();
  },

  getRegistrationTargetInfo(loteOverride = '') {
    const loteName = String(loteOverride || this.selectedRegistrationLote || '').trim();
    if (!loteName) return null;
    const info = this.getAllLotes().find((item) => this.getNormalizedLoteKey(item.lote) === this.getNormalizedLoteKey(loteName));
    if (!info || this.shouldRemoveLoteFromMap(info.lote)) return null;
    const logEntry = this.getLocationLogEntryByLote(info.lote);
    const renderedPoint = this.getRenderablePointByLote(info.lote, this.lastRenderedMapped);
    const parsedLoc = renderedPoint || this.parseLocationEntry(this.locDB?.[info.lote]);
    const syncMeta = this.getLocationSyncMeta(logEntry || { lote: info.lote, source: parsedLoc ? 'compartilhado' : 'local' });
    return {
      lote: info.lote,
      produto: info.produto || '',
      saldo: Number(info.saldo || 0),
      hasCoordinates: Boolean(parsedLoc && Number.isFinite(parsedLoc.lat) && Number.isFinite(parsedLoc.lon)),
      lat: parsedLoc?.lat ?? null,
      lon: parsedLoc?.lon ?? null,
      syncMeta,
      syncMessage: String(logEntry?.syncMessage || syncMeta.summary || '').trim(),
      updatedAt: this.formatLocationMoment(logEntry?.updatedAt || ''),
      updatedBy: String(logEntry?.updatedBy || '').trim()
    };
  },

  setRegistrationTarget(lote, options = {}) {
    const target = this.getRegistrationTargetInfo(lote);
    if (!target) {
      toastManager.show('Selecione um lote ativo do estoque para registrar o GPS.', 'warning');
      return null;
    }
    this.selectedRegistrationLote = target.lote;
    safeStorage.setItem(this.registrationTargetKey, target.lote);
    const input = document.getElementById('inpBuscaMapa');
    if (input && options.populateSearch !== false) input.value = target.lote;
    this.updateRegistrationUi();
    if (options.refreshSearch) this.buscarNaLista();
    if (!options.silent) {
      toastManager.show(`Lote ${target.lote} pronto para o registro do GPS.`, 'info', 2600);
    }
    return target;
  },

  clearRegistrationTarget(silent = false) {
    this.selectedRegistrationLote = '';
    safeStorage.removeItem(this.registrationTargetKey);
    this.updateRegistrationUi();
    this.buscarNaLista();
    if (!silent) toastManager.show('Seleção do lote para GPS foi limpa.', 'info');
  },

  updateRegistrationUi() {
    const label = document.getElementById('mapRegistrationTargetLabel');
    const meta = document.getElementById('mapRegistrationTargetMeta');
    const syncBadge = document.getElementById('mapRegistrationSyncBadge');
    const syncCopy = document.getElementById('mapRegistrationSyncCopy');
    const actionBtn = document.getElementById('mapRegisterSelectedBtn');
    const clearBtn = document.getElementById('mapRegistrationClearBtn');
    const target = this.getRegistrationTargetInfo();
    if (!target && this.selectedRegistrationLote) {
      this.selectedRegistrationLote = '';
      safeStorage.removeItem(this.registrationTargetKey);
    }
    if (label) {
      label.textContent = target
        ? `${target.lote} • ${target.produto || 'Produto não informado'}`
        : 'Nenhum lote selecionado';
    }
    if (meta) {
      meta.textContent = target
        ? `Saldo disponível: ${formatTons(target.saldo)} t${target.hasCoordinates ? ' • Já possui GPS salvo' : ' • Ainda sem GPS válido'}`
        : '1. Busque um lote  2. Toque em “Selecionar lote”  3. Use o botão principal para registrar o GPS.';
    }
    if (syncBadge) {
      const syncMeta = target ? target.syncMeta : null;
      syncBadge.className = syncMeta ? syncMeta.className : 'map-sync-pill map-sync-pill--idle';
      syncBadge.textContent = syncMeta ? syncMeta.label : 'Aguardando lote';
    }
    if (syncCopy) {
      syncCopy.textContent = target
        ? (target.syncMessage || target.syncMeta.summary)
        : 'O registro só será concluído com sucesso depois que o mapa compartilhado confirmar este lote para todos os usuários.';
    }
    if (actionBtn) {
      actionBtn.disabled = !target;
      actionBtn.textContent = target ? `📍 Registrar GPS do lote ${target.lote}` : '📍 Registrar GPS do lote selecionado';
    }
    if (clearBtn) clearBtn.hidden = !target;
  },

  handleRegisterSelectedLote() {
    const target = this.getRegistrationTargetInfo();
    if (!target) {
      toastManager.show('Busque e selecione um lote antes de registrar o GPS.', 'warning');
      return;
    }
    this.capturarGPSEditar(target.lote);
  },

  renderPanelsFromSnapshot(refreshSearch = false) {
    const mapped = Array.isArray(this.lastRenderedMapped) ? this.lastRenderedMapped : [];
    const pending = Array.isArray(this.lastRenderedPending) ? this.lastRenderedPending : [];
    const proximity = this.lastRenderedProximity && Array.isArray(this.lastRenderedProximity.edges)
      ? this.lastRenderedProximity
      : this.getProximityWeb(mapped);
    this.renderPanels(mapped, pending, proximity);
    if (this.sheetTab === 'tech') {
      this.renderLocationLog(document.getElementById('inpBuscaMapaLog')?.value || '');
    }
    if (refreshSearch) this.buscarNaLista();
    else this.updateRegistrationUi();
  },

  syncUiState() {
    this.updateVisualModeControls();
    this.updateOverlayUi();
    this.updateAdvancedActionsUi();
    this.updateMapStageUi();
    this.updateSheetUi();
    this.updateStatusFilterControls();
    this.updateRegistrationUi();
  },

  toggleMapStage(forceState = null) {
    const nextState = typeof forceState === 'boolean' ? forceState : !this.mapStageExpanded;
    this.mapStageExpanded = nextState;
    safeStorage.setItem('cysyMapStageExpanded', String(this.mapStageExpanded));
    if (!nextState && this.isFullscreen) {
      this.toggleFullScreen(false);
    }
    if (!nextState) {
      this.clearMapDynamicLayers();
    }
    this.updateMapStageUi();
    this.refreshMapLayout();
    if (nextState) {
      setTimeout(() => {
        this.initLeaflet();
        this.renderPins(true);
        this.refreshMapLayout();
        this.refreshSharedLocations(true);
      }, 120);
    }
  },

  updateMapStageUi() {
    const stage = document.getElementById('mapOpsStage');
    const toggle = document.getElementById('mapStageToggle');
    const label = document.getElementById('mapStageToggleLabel');
    const hint = document.getElementById('mapStageToggleHint');
    if (stage) {
      stage.classList.toggle('map-stage-collapsed', !this.mapStageExpanded);
      stage.classList.toggle('map-stage-expanded', this.mapStageExpanded);
    }
    if (toggle) {
      toggle.setAttribute('aria-expanded', this.mapStageExpanded ? 'true' : 'false');
    }
    if (label) {
      label.textContent = this.mapStageExpanded ? 'Recolher mapa' : 'Expandir mapa';
    }
    if (hint) {
      hint.textContent = this.mapStageExpanded
        ? 'Toque para recolher o mapa e priorizar a lista de lotes.'
        : 'Mapa recolhido para caber melhor em tela pequena. Toque para abrir.';
    }
  },

  getSheetTabs() {
    return ['lotes', 'tech'];
  },

  switchSheetTab(tab = 'lotes', silent = false) {
    const normalized = this.getSheetTabs().includes(tab) ? tab : 'lotes';
    this.sheetTab = normalized;
    safeStorage.setItem('cysyMapSheetTab', normalized);
    this.updateSheetUi();
    if (normalized === 'tech') {
      this.renderLocationLog(document.getElementById('inpBuscaMapaLog')?.value || '');
    }
    if (!silent) {
      toastManager.show(normalized === 'lotes' ? 'Painel de lotes ativado.' : 'Dados técnicos do mapa ativados.', 'info');
    }
  },

  toggleSheet(forceState = null) {
    this.sheetExpanded = true;
    safeStorage.setItem('cysyMapSheetExpanded', String(this.sheetExpanded));
    this.updateSheetUi();
    this.refreshMapLayout();
    setTimeout(() => this.renderPanelsFromSnapshot(true), 120);
  },

  updateSheetUi() {
    const sheet = document.getElementById('mapOpsSheet');
    const toggleLabel = document.getElementById('mapSheetToggleLabel');
    if (sheet) {
      sheet.classList.toggle('collapsed', !this.sheetExpanded);
      sheet.classList.toggle('expanded', this.sheetExpanded);
    }
    if (toggleLabel) toggleLabel.textContent = this.sheetExpanded ? 'Recolher painel' : 'Expandir painel';

    this.getSheetTabs().forEach((tab) => {
      const button = document.getElementById(`mapSheetTab-${tab}`);
      const panel = document.getElementById(`mapSheetPanel-${tab}`);
      const active = this.sheetTab === tab;
      if (button) {
        button.classList.toggle('active', active);
        button.setAttribute('aria-selected', active ? 'true' : 'false');
      }
      if (panel) {
        panel.hidden = !active || !this.sheetExpanded;
        panel.classList.toggle('map-sheet-panel-active', active && this.sheetExpanded);
      }
    });
  },

  setStatusFilter(filter = 'all', silent = false) {
    const allowed = ['all', 'ok', 'review', 'missing'];
    const normalized = allowed.includes(filter) ? filter : 'all';
    this.statusFilter = normalized;
    safeStorage.setItem('cysyMapStatusFilter', normalized);
    this.updateStatusFilterControls();
    this.renderPanelsFromSnapshot(true);
    if (!silent) {
      const labels = {
        all: 'Exibindo todos os lotes.',
        ok: 'Filtrando lotes posicionados e atualizados.',
        review: 'Filtrando lotes que precisam de revisão.',
        missing: 'Filtrando lotes pendentes de GPS.'
      };
      toastManager.show(labels[normalized] || 'Filtro aplicado.', 'info');
    }
  },

  updateStatusFilterControls() {
    ['all', 'ok', 'review', 'missing'].forEach((filter) => {
      const button = document.getElementById(`mapFilter-${filter}`);
      if (!button) return;
      const active = this.statusFilter === filter;
      button.classList.toggle('active', active);
      button.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
  },

  handleSearchInput() {
    if (this.searchRefreshTimer) clearTimeout(this.searchRefreshTimer);
    this.searchRefreshTimer = setTimeout(() => {
      this.searchRefreshTimer = 0;
      this.buscarNaLista();
      this.renderPanelsFromSnapshot(false);
    }, this.isCompactMobile() ? 140 : 90);
  },

  focusSelectedOrBase() {
    if (this.selectedLote) {
      this.selectLote(this.selectedLote, { center: true, mapped: this.lastRenderedMapped, zoom: 19 });
      return;
    }
    this.centralizarEmpresa(true);
  },

  getMappedFreshnessStatus(item = {}) {
    const hasCoords = Number.isFinite(Number(item.lat)) && Number.isFinite(Number(item.lon));
    if (!hasCoords) {
      return {
        code: 'missing',
        label: 'Sem coordenadas',
        summary: 'Este lote ainda não possui coordenadas válidas no sistema.',
        badgeClass: 'map-lote-state map-lote-state--missing'
      };
    }
    const updatedAtMs = this.getSortableTimestamp(item.updatedAt || '');
    const ageHours = updatedAtMs > 0 ? ((Date.now() - updatedAtMs) / 3600000) : Number.POSITIVE_INFINITY;
    if (Number.isFinite(ageHours) && ageHours <= 72) {
      return {
        code: 'ok',
        label: 'Posicionado e atualizado',
        summary: 'Coordenadas recentes e prontas para operação.',
        badgeClass: 'map-lote-state map-lote-state--ok'
      };
    }
    return {
      code: 'review',
      label: 'GPS desatualizado',
      summary: 'Recomenda-se revisar o ponto GPS deste lote.',
      badgeClass: 'map-lote-state map-lote-state--review'
    };
  },

  getCombinedPanelItems(mapped = [], pending = []) {
    const mappedItems = (mapped || []).map((item) => ({
      ...item,
      panelStatus: this.getMappedFreshnessStatus(item),
      hasCoordinates: true
    }));
    const pendingItems = (pending || []).map((item) => ({
      ...item,
      panelStatus: this.getMappedFreshnessStatus({ ...item, lat: null, lon: null }),
      hasCoordinates: false
    }));
    return [...mappedItems, ...pendingItems];
  },

  applySheetFilters(items = []) {
    const searchTerm = String(document.getElementById('inpBuscaMapa')?.value || '').trim().toUpperCase();
    const order = { ok: 0, review: 1, missing: 2 };
    return (items || []).filter((item) => {
      if (this.statusFilter !== 'all' && item.panelStatus?.code !== this.statusFilter) return false;
      if (!searchTerm) return true;
      const hay = [
        item.lote,
        item.produto,
        item.description,
        item.updatedBy
      ].join(' ').toUpperCase();
      return hay.includes(searchTerm);
    }).sort((a, b) => {
      const aSelected = this.getNormalizedLoteKey(a.lote) === this.getNormalizedLoteKey(this.selectedLote) ? 1 : 0;
      const bSelected = this.getNormalizedLoteKey(b.lote) === this.getNormalizedLoteKey(this.selectedLote) ? 1 : 0;
      return (bSelected - aSelected) ||
        ((order[a.panelStatus?.code] ?? 9) - (order[b.panelStatus?.code] ?? 9)) ||
        String(a.lote || '').localeCompare(String(b.lote || ''), 'pt-BR');
    });
  },

  getPanelRenderWindow(items = []) {
    const normalized = Array.isArray(items) ? items : [];
    if (!this.isCompactMobile()) {
      return { items: normalized, hiddenCount: 0, truncated: false };
    }
    const searchTerm = String(document.getElementById('inpBuscaMapa')?.value || '').trim();
    if (searchTerm || this.statusFilter !== 'all') {
      return { items: normalized, hiddenCount: 0, truncated: false };
    }
    const limit = 24;
    return {
      items: normalized.slice(0, limit),
      hiddenCount: Math.max(0, normalized.length - limit),
      truncated: normalized.length > limit
    };
  },

  calculatePolygonPerimeter(coords = []) {
    const normalized = Array.isArray(coords) ? coords.filter((row) => Array.isArray(row) && row.length === 2) : [];
    if (normalized.length < 2) return 0;
    let total = 0;
    const toRad = (deg) => (deg * Math.PI) / 180;
    for (let i = 0; i < normalized.length; i++) {
      const [lat1, lon1] = normalized[i];
      const [lat2, lon2] = normalized[(i + 1) % normalized.length];
      const dLat = toRad(lat2 - lat1);
      const dLon = toRad(lon2 - lon1);
      const a = Math.sin(dLat / 2) ** 2 +
        Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
      total += 6371000 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    }
    return total;
  },

  calculatePolygonArea(coords = []) {
    const normalized = Array.isArray(coords) ? coords.filter((row) => Array.isArray(row) && row.length === 2) : [];
    if (normalized.length < 3) return 0;
    const avgLat = normalized.reduce((sum, row) => sum + Number(row[0] || 0), 0) / normalized.length;
    const metersPerDegLat = 111320;
    const metersPerDegLon = Math.cos((avgLat * Math.PI) / 180) * 111320;
    const points = normalized.map(([lat, lon]) => ({
      x: Number(lon) * metersPerDegLon,
      y: Number(lat) * metersPerDegLat
    }));
    let area = 0;
    for (let i = 0; i < points.length; i++) {
      const current = points[i];
      const next = points[(i + 1) % points.length];
      area += (current.x * next.y) - (next.x * current.y);
    }
    return Math.abs(area / 2);
  },

  formatMeters(value = 0) {
    const meters = Number(value || 0);
    if (!Number.isFinite(meters) || meters <= 0) return 'n/d';
    return meters >= 1000 ? `${(meters / 1000).toFixed(2)} km` : `${Math.round(meters)} m`;
  },

  renderTechnicalPanel(mapped = [], pending = [], proximity = null) {
    const baseContainer = document.getElementById('mapTechnicalBase');
    const syncContainer = document.getElementById('mapTechnicalSync');
    if (!baseContainer && !syncContainer) return;

    const areaCoords = this.getCompanyAreaCoords();
    const areaValue = this.calculatePolygonArea(areaCoords);
    const perimeterValue = this.calculatePolygonPerimeter(areaCoords);
    const bounds = this.getBaseFocusBounds();
    const proximityData = proximity || this.getProximityWeb(mapped);
    const avgDistance = Number(proximityData?.averageNearestDistance || 0);
    const currentLayerLabel = {
      street: 'Mapa padrão',
      satellite: 'Satélite',
      hybrid: 'Híbrido'
    }[this.currentBaseLayerKind] || 'Mapa padrão';

    if (baseContainer) {
      baseContainer.innerHTML = `
        <div class="map-tech-list">
          <div class="map-tech-item">
            <span>Base</span>
            <strong>${escapeHTML(config.base.empresa)}<br>${escapeHTML(config.base.enderecoCurto)}</strong>
          </div>
          <div class="map-tech-item">
            <span>Área do terreno</span>
            <strong>${areaValue > 0 ? `${Math.round(areaValue).toLocaleString('pt-BR')} m²` : 'n/d'}</strong>
          </div>
          <div class="map-tech-item">
            <span>Perímetro estimado</span>
            <strong>${this.formatMeters(perimeterValue)}</strong>
          </div>
          <div class="map-tech-item">
            <span>Centro operacional</span>
            <strong>${config.base.lat.toFixed(6)}, ${config.base.lon.toFixed(6)}</strong>
          </div>
          <div class="map-tech-item">
            <span>Bounds da área</span>
            <strong>${bounds ? `${bounds.getSouthWest().lat.toFixed(5)}, ${bounds.getSouthWest().lng.toFixed(5)} ↔ ${bounds.getNorthEast().lat.toFixed(5)}, ${bounds.getNorthEast().lng.toFixed(5)}` : 'n/d'}</strong>
          </div>
          <div class="map-tech-item">
            <span>Distância média entre vizinhos</span>
            <strong>${this.formatMeters(avgDistance)}</strong>
          </div>
        </div>`;
    }

    if (syncContainer) {
      syncContainer.innerHTML = `
        <div class="map-tech-list">
          <div class="map-tech-item">
            <span>Camada ativa</span>
            <strong>${currentLayerLabel}</strong>
          </div>
          <div class="map-tech-item">
            <span>Sincronização compartilhada</span>
            <strong>${this.lastSharedSyncAt ? new Date(this.lastSharedSyncAt).toLocaleString('pt-BR') : 'Ainda não sincronizado nesta sessão'}</strong>
          </div>
          <div class="map-tech-item">
            <span>Cache offline</span>
            <strong>${navigator.onLine ? 'Preparado para reuso offline dos tiles já visitados' : 'Usando cache local e fila offline do mapa'}</strong>
          </div>
          <div class="map-tech-item">
            <span>Lotes no mapa</span>
            <strong>${mapped.length} com coordenadas • ${pending.length} pendentes</strong>
          </div>
        </div>`;
    }
    const btnLimpar = document.getElementById('btnLimparLotesMapa');
    if (btnLimpar) {
      btnLimpar.hidden = !loginManager.isSuperAdmin();
    }
  },

  clearDecorationLayers() {
    (this.decorationLayers || []).forEach((layer) => {
      try { this.map?.removeLayer(layer); } catch (_) {}
    });
    this.decorationLayers = [];
  },

  clearMapDynamicLayers() {
    Object.values(this.markers || {}).forEach((marker) => {
      try { this.map?.removeLayer(marker); } catch (_) {}
    });
    this.markers = {};
    this.clearDecorationLayers();
    this.clearConnectionLayers();
  },

  getRenderablePointByLote(lote, sourceList = null) {
    const key = this.getNormalizedLoteKey(lote);
    const list = Array.isArray(sourceList) ? sourceList : this.getRenderableMappedPoints();
    return list.find((item) => this.getNormalizedLoteKey(item.lote) === key) || null;
  },

  getSelectedMappedPoint(sourceList = null) {
    if (!this.selectedLote) return null;
    return this.getRenderablePointByLote(this.selectedLote, sourceList);
  },

  clearSelectedLote(silent = false) {
    this.selectedLote = '';
    safeStorage.removeItem('cysySelectedMapLote');
    this.renderPins();
    if (!silent) toastManager.show('Foco do lote limpo.', 'info');
  },

  selectLote(lote, options = {}) {
    const mapped = Array.isArray(options.mapped) ? options.mapped : this.getRenderableMappedPoints();
    const item = this.getRenderablePointByLote(lote, mapped);
    if (!item) {
      toastManager.show('Este lote ainda não possui localização disponível no mapa.', 'warning');
      return null;
    }

    this.selectedLote = item.lote;
    safeStorage.setItem('cysySelectedMapLote', item.lote);
    if (options.center !== false) {
      this.initLeaflet();
      const zoom = Number(options.zoom || (this.map?.getZoom?.() || 19));
      try {
        this.map?.setView([item.lat, item.lon], Math.max(18, zoom), { animate: true });
      } catch (_) {}
    }
    this.renderPins();
    return item;
  },

  confirmNavigationToLote(lote) {
    const item = this.getRenderablePointByLote(lote, this.lastRenderedMapped.length > 0 ? this.lastRenderedMapped : this.getRenderableMappedPoints());
    if (!item) {
      toastManager.show('Este lote ainda não possui coordenadas prontas para navegação.', 'warning');
      return;
    }
    this.selectLote(item.lote, { center: true, mapped: this.lastRenderedMapped, zoom: 19 });
    const confirmed = window.confirm(`Quer ir até o lote ${item.lote}?\n\nMaterial: ${item.produto || 'Não informado'}\nQuantidade: ${item.active ? `${formatTons(item.saldo)} t` : 'Registro do log compartilhado'}\n\nSe confirmar, o app abrirá a navegação a pé para este lote.`);
    if (!confirmed) return;
    this.openNavigation(item.lat, item.lon);
  },

  getLiveLayoutConfig(pointsCount = 0) {
    const zoom = Number(this.map?.getZoom?.() || 18);
    const dense = pointsCount >= 8;
    if (this.visualMode === 'detailed') {
      return {
        badgeWidth: zoom >= 18 ? 76 : 70,
        badgeHeight: 38,
        pointRadius: 7,
        minGap: dense ? 60 : 66,
        radiusStep: dense ? 22 : 24,
        showLeaders: true
      };
    }
    if (this.visualMode === 'informative') {
      return {
        badgeWidth: zoom >= 18 ? 66 : 62,
        badgeHeight: 34,
        pointRadius: 6,
        minGap: dense ? 54 : 58,
        radiusStep: 20,
        showLeaders: true
      };
    }
    return {
      badgeWidth: 56,
      badgeHeight: 30,
      pointRadius: 5,
      minGap: dense ? 48 : 52,
      radiusStep: 18,
      showLeaders: dense || zoom >= 17
    };
  },

  getMarkerPalette() {
    return ['#2563EB', '#DC2626', '#059669', '#7C3AED', '#EA580C', '#0891B2', '#BE123C', '#4F46E5', '#0F766E', '#CA8A04'];
  },

  getStatusColor(statusCode = 'ok', fallback = '#2563EB') {
    const palette = {
      ok: '#16A34A',
      review: '#D97706',
      missing: '#64748B'
    };
    return palette[statusCode] || fallback;
  },

  getCollisionArea(layoutBounds, padding = 24) {
    return {
      minX: Number(layoutBounds?.minX ?? 0) + padding,
      minY: Number(layoutBounds?.minY ?? 0) + padding,
      maxX: Number(layoutBounds?.maxX ?? 0) - padding,
      maxY: Number(layoutBounds?.maxY ?? 0) - padding
    };
  },

  rectIntersectionArea(a, b) {
    const overlapW = Math.max(0, Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x));
    const overlapH = Math.max(0, Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y));
    return overlapW * overlapH;
  },

  scoreCollisionBox(box, placed = [], area = {}) {
    let score = 0;
    placed.forEach((item) => {
      score += this.rectIntersectionArea(box, item.box) * 50;
    });
    if (box.x < area.minX) score += (area.minX - box.x) * 100;
    if (box.y < area.minY) score += (area.minY - box.y) * 100;
    if (box.x + box.width > area.maxX) score += ((box.x + box.width) - area.maxX) * 100;
    if (box.y + box.height > area.maxY) score += ((box.y + box.height) - area.maxY) * 100;
    return score;
  },

  computeCollisionLayout(points = [], options = {}) {
    const normalized = Array.isArray(points) ? points.filter(Boolean) : [];
    if (normalized.length === 0) return [];
    const project = typeof options.project === 'function' ? options.project : ((item) => item.actual || { x: 0, y: 0 });
    const area = this.getCollisionArea(options.bounds || {}, Number(options.padding || 0));
    const maxAttempts = Number(options.maxAttempts || 28);
    const sizeForPoint = typeof options.boxSizeGetter === 'function'
      ? options.boxSizeGetter
      : (() => ({ width: 56, height: 30 }));

    const candidates = normalized.map((item, originalIndex) => {
      const actual = project(item);
      return {
        ...item,
        originalIndex,
        actual
      };
    }).sort((a, b) => (a.actual.y - b.actual.y) || (a.actual.x - b.actual.x));

    const placed = [];
    const layout = [];

    candidates.forEach((item) => {
      const size = sizeForPoint(item);
      let best = null;
      for (let attempt = 0; attempt < maxAttempts; attempt++) {
        const ring = attempt === 0 ? 0 : Math.floor((attempt - 1) / 8) + 1;
        const angleIndex = attempt === 0 ? 0 : (attempt - 1) % 8;
        const angle = (Math.PI / 4) * angleIndex;
        const distanceOffset = ring * Number(options.radiusStep || 18);
        const center = attempt === 0
          ? { x: item.actual.x, y: item.actual.y }
          : {
              x: item.actual.x + Math.cos(angle) * distanceOffset,
              y: item.actual.y + Math.sin(angle) * distanceOffset
            };
        const box = {
          x: center.x - (size.width / 2),
          y: center.y - (size.height / 2),
          width: size.width,
          height: size.height
        };
        const score = this.scoreCollisionBox(box, placed, area) + (distanceOffset * 0.25);
        if (!best || score < best.score) {
          best = { center, box, score };
        }
        if (score === 0) break;
      }
      const chosen = best || {
        center: { x: item.actual.x, y: item.actual.y },
        box: {
          x: item.actual.x - (size.width / 2),
          y: item.actual.y - (size.height / 2),
          width: size.width,
          height: size.height
        },
        score: 0
      };
      const layoutItem = {
        ...item,
        size,
        badge: chosen.center,
        box: chosen.box,
        displaced: Math.hypot(chosen.center.x - item.actual.x, chosen.center.y - item.actual.y) >= 6
      };
      placed.push(layoutItem);
      layout.push(layoutItem);
    });

    return layout.sort((a, b) => a.originalIndex - b.originalIndex);
  },

  buildLiveLayoutPoints(points = []) {
    if (!this.map) return [];
    const container = this.getMapContainer();
    const width = Number(container?.clientWidth || 960);
    const height = Number(container?.clientHeight || 620);
    const configLayout = this.getLiveLayoutConfig(points.length);
    const bounds = { minX: 0, minY: 0, maxX: width, maxY: height };
    const actualPoints = points.map((item) => {
      const point = this.map.latLngToContainerPoint([item.lat, item.lon]);
      return {
        ...item,
        actual: { x: point.x, y: point.y }
      };
    });
    return this.computeCollisionLayout(actualPoints, {
      project: (item) => item.actual,
      bounds,
      padding: 22,
      radiusStep: configLayout.radiusStep,
      maxAttempts: 36,
      boxSizeGetter: (item) => {
        const isSelected = this.getNormalizedLoteKey(item.lote) === this.getNormalizedLoteKey(this.selectedLote);
        const widthBoost = isSelected ? 10 : 0;
        return {
          width: configLayout.badgeWidth + widthBoost,
          height: configLayout.badgeHeight + (isSelected ? 2 : 0)
        };
      }
    }).map((item, index) => ({
      ...item,
      label: this.getMarkerLabel(item.lote, index + 1),
      badgeLatLng: this.map.containerPointToLatLng(L.point(item.badge.x, item.badge.y))
    }));
  },

  getConnectionsForSelection(mapped = [], proximity = null) {
    const selected = this.getSelectedMappedPoint(mapped);
    if (!selected) return [];
    const normalizedKey = this.getNormalizedLoteKey(selected.lote);
    const points = Array.isArray(mapped) ? mapped : [];
    const proximityData = proximity || this.getProximityWeb(points);
    return (proximityData.edges || [])
      .map((edge) => {
        const from = points[edge.from];
        const to = points[edge.to];
        if (!from || !to) return null;
        if (this.getNormalizedLoteKey(from.lote) !== normalizedKey && this.getNormalizedLoteKey(to.lote) !== normalizedKey) return null;
        const other = this.getNormalizedLoteKey(from.lote) === normalizedKey ? to : from;
        return { ...edge, other };
      })
      .filter(Boolean)
      .sort((a, b) => a.distance - b.distance);
  },

  getNavigationUrl(lat, lon) {
    const coords = `${lat},${lon}`;
    const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    return isIOS
      ? `https://maps.apple.com/?daddr=${encodeURIComponent(coords)}&dirflg=w`
      : `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(coords)}&travelmode=walking`;
  },

  enableFallbackBaseLayer(reason = '') {
    if (!this.map || !this.streetLayer) return;
    if (this.activeBaseLayer && this.map.hasLayer(this.activeBaseLayer)) {
      try { this.map.removeLayer(this.activeBaseLayer); } catch (_) {}
    }
    if (!this.map.hasLayer(this.streetLayer)) {
      this.streetLayer.addTo(this.map);
    }
    this.activeBaseLayer = this.streetLayer;
    this.currentBaseLayerKind = 'street';
    if (!this.usingFallbackLayer) {
      this.usingFallbackLayer = true;
      const message = reason || 'Imagem satelital indisponível nesta área/zoom. Exibindo mapa base confiável.';
      debugEngine.log(message, 'warn');
      toastManager.show(message, 'warning', 5200);
    }
  },

  useStreetBaseLayer(silent = false) {
    if (!this.map || !this.streetLayer) return;
    if (this.activeBaseLayer && this.map.hasLayer(this.activeBaseLayer)) {
      try { this.map.removeLayer(this.activeBaseLayer); } catch (_) {}
    }
    if (!this.map.hasLayer(this.streetLayer)) {
      this.streetLayer.addTo(this.map);
    }
    this.activeBaseLayer = this.streetLayer;
    this.currentBaseLayerKind = 'street';
    this.usingFallbackLayer = true;
    this.scheduleOfflineMapSave(true);
    if (!silent) toastManager.show('Mapa base estável ativado.', 'info');
  },

  useSatelliteBaseLayer() {
    if (!this.map || !this.satelliteLayer) return;
    if (this.activeBaseLayer && this.map.hasLayer(this.activeBaseLayer)) {
      try { this.map.removeLayer(this.activeBaseLayer); } catch (_) {}
    }
    if (!this.map.hasLayer(this.satelliteLayer)) {
      this.satelliteLayer.addTo(this.map);
    }
    this.activeBaseLayer = this.satelliteLayer;
    this.currentBaseLayerKind = 'satellite';
    this.usingFallbackLayer = false;
    this.tileFailureCount = 0;
    this.scheduleOfflineMapSave(true);
    toastManager.show('Tentando carregar imagem satelital.', 'info');
  },

  useHybridBaseLayer() {
    this.useSatelliteBaseLayer();
    this.currentBaseLayerKind = 'hybrid';
    toastManager.show('Modo híbrido ativado com base satelital e overlays operacionais da Cysy.', 'info');
  },

  bindTileLayerFallbacks() {
    if (!this.satelliteLayer) return;
    this.satelliteLayer.on('tileerror', () => {
      this.tileFailureCount += 1;
      if (this.tileFailureCount >= 3) {
        this.enableFallbackBaseLayer('Imagem satelital indisponível neste zoom. Exibindo mapa base confiável.');
      }
    });
    this.satelliteLayer.on('load', () => {
      this.tileFailureCount = 0;
    });
  },

  openNavigation(lat, lon) {
    const navUrl = this.getNavigationUrl(lat, lon);
    try {
      const win = window.open(navUrl, '_blank', 'noopener,noreferrer');
      if (!win) window.location.href = navUrl;
    } catch (_) {
      window.location.href = navUrl;
    }
  },

  bindResizeHandlers() {
    if (this.resizeBound) return;
    this.resizeBound = true;
    const bindRefresh = () => {
      this.refreshMapLayout();
      setTimeout(() => this.renderPins(), 90);
    };
    window.addEventListener('resize', bindRefresh);
    window.addEventListener('orientationchange', () => setTimeout(bindRefresh, 260));
    document.addEventListener('fullscreenchange', bindRefresh);
    if (window.visualViewport) {
      window.visualViewport.addEventListener('resize', bindRefresh);
    }
    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && this.isFullscreen) {
        this.toggleFullScreen(false);
      }
    });
  },

  restoreViewport(forceCompany = false) {
    if (!this.map) return;
    try {
      if (!forceCompany && this.currentBounds && Object.keys(this.markers || {}).length > 0) {
        this.map.fitBounds(this.currentBounds, { padding: [42, 42], maxZoom: 19, animate: false });
      } else {
        const baseBounds = this.getBaseFocusBounds();
        if (baseBounds) this.map.fitBounds(baseBounds.pad(0.45), { maxZoom: 18, animate: false });
        else this.map.setView([config.base.lat, config.base.lon], 18, { animate: false });
      }
    } catch (_) {}
  },

  initLeaflet() {
    this.applyRequestedMapReset();
    const mapNode = document.getElementById('patioMap');
    if (!mapNode) {
      debugEngine.log('Container do mapa não encontrado.', 'error');
      return null;
    }

    const host = this.isFullscreen ? this.getFullscreenMount() : this.getMapDock();
    const outerContainer = this.getMapContainer();
    if (outerContainer && host && outerContainer.parentElement !== host) {
      host.appendChild(outerContainer);
    }

    if (!this.map) {
      try {
        this.map = L.map('patioMap', {
          center: [config.base.lat, config.base.lon],
          zoom: 19,
          minZoom: 4,
          maxZoom: 22,
          zoomSnap: 0.25,
          zoomDelta: 0.5,
          zoomControl: true,
          attributionControl: true,
          preferCanvas: true,
          fadeAnimation: false,
          zoomAnimation: true,
          markerZoomAnimation: true
        });

        this.satelliteLayer = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
          maxZoom: 22,
          maxNativeZoom: 18,
          keepBuffer: 10,
          updateWhenIdle: true,
          updateWhenZooming: false,
          attribution: 'Tiles © Esri'
        });

        this.streetLayer = L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', {
          maxZoom: 22,
          maxNativeZoom: 19,
          keepBuffer: 10,
          updateWhenIdle: true,
          updateWhenZooming: false,
          attribution: '&copy; OpenStreetMap'
        });

        this.activeBaseLayer = this.streetLayer;
        this.streetLayer.addTo(this.map);
        this.currentBaseLayerKind = 'street';
        this.usingFallbackLayer = true;
        this.bindTileLayerFallbacks();

        const areaCoords = this.getCompanyAreaCoords();
        if (areaCoords.length >= 3) {
          this.companyAreaLayer = L.polygon(areaCoords, {
            color: '#B45309',
            weight: 2,
            fillColor: '#FACC15',
            fillOpacity: 0.22,
            dashArray: '8 6',
            interactive: false
          }).addTo(this.map);
        }

        this.companyMarker = L.circleMarker([config.base.lat, config.base.lon], {
          color: '#78350F',
          fillColor: '#F59E0B',
          fillOpacity: 0.96,
          radius: 10,
          weight: 3
        }).bindPopup(
          `<div style="font-family:'Montserrat',sans-serif; line-height:1.5; min-width:220px;">
            <b>Base ${escapeHTML(config.base.empresa)}</b><br>
            ${escapeHTML(config.base.enderecoCurto)}<br>
            <span style="font-size:11px;">Atividade: ${escapeHTML(config.base.atividade)}</span><br>
            <span style="font-size:11px;">Coordenadas: ${escapeHTML(config.base.lat)}, ${escapeHTML(config.base.lon)}</span>
          </div>`
        ).addTo(this.map);

        const baseBounds = this.getBaseFocusBounds();
        if (baseBounds) {
          this.map.fitBounds(baseBounds.pad(0.45), { maxZoom: 18, animate: false });
        }
        if (!navigator.onLine) {
          const offlineState = this.getOfflineMapState();
          if (offlineState?.center && Number.isFinite(Number(offlineState.center.lat)) && Number.isFinite(Number(offlineState.center.lng))) {
            this.map.setView([Number(offlineState.center.lat), Number(offlineState.center.lng)], Number(offlineState.zoom || 18), { animate: false });
          }
        }

        this.map.on('moveend', () => {
          this.scheduleOfflineMapSave(false);
        });
        this.map.on('zoomend', () => {
          this.renderPins();
          this.scheduleOfflineMapSave(false);
        });

        this.bindResizeHandlers();
        const container = this.getMapContainer();
        if (container && !this.resizeObs) {
          this.resizeObs = new ResizeObserver(() => this.refreshMapLayout());
          this.resizeObs.observe(container);
        }

        debugEngine.log('Mapa satelital operacional inicializado.', 'success');
      } catch (error) {
        debugEngine.log(`Falha ao inicializar o mapa: ${error.message}`, 'error');
        return null;
      }
    }

    const isMobileViewport = typeof window !== 'undefined' && window.matchMedia && window.matchMedia('(max-width: 768px)').matches;
    if (isMobileViewport) {
      this.sheetExpanded = true;
      safeStorage.setItem('cysyMapSheetExpanded', 'true');
      this.mapStageExpanded = false;
      safeStorage.setItem('cysyMapStageExpanded', 'false');
      this.visualMode = 'clean';
      safeStorage.setItem('cysyMapVisualMode', 'clean');
      this.overlaysExpanded = false;
      safeStorage.setItem('cysyMapOverlaysExpanded', 'false');
    }

    this.syncUiState();
    this.renderPins();
    this.rebuildLocationLogFromLocDb();
    this.refreshMapLayout();
    this.scheduleOfflineMapSave(true);
    this.ensureSharedSyncLoop();
    this.refreshSharedLocations(false);
    return this.map;
  },

  ensureSharedSyncLoop() {
    if (this.sharedSyncTimer) return;
    this.sharedSyncTimer = setInterval(() => {
      const tabMapa = document.getElementById('tab-mapa');
      const mapaAtivo = tabMapa && tabMapa.classList.contains('active');
      if (!mapaAtivo || document.hidden || !navigator.onLine) return;
      this.refreshSharedLocations(false);
    }, 20000);
  },

  async refreshSharedLocations(force = false) {
    if (!navigator.onLine || this.sharedSyncInFlight) return [];
    const now = Date.now();
    if (!force && now - this.lastSharedFetchAt < 12000) return [];
    this.lastSharedFetchAt = now;
    this.sharedSyncInFlight = true;
    try {
      const globalRows = await apiService.fetchGlobalLoc();
      if (Array.isArray(globalRows)) {
        this.lastSharedRows = globalRows;
        this.syncFromGlobal(globalRows);
      }
      return globalRows || [];
    } catch (err) {
      this.lastSharedSyncError = err?.message || 'Falha ao sincronizar localizações compartilhadas.';
      return [];
    } finally {
      this.sharedSyncInFlight = false;
    }
  },

  refreshMapLayout() {
    if (!this.map) return;
    if (this.layoutRefreshRaf) {
      try { cancelAnimationFrame(this.layoutRefreshRaf); } catch (_) {}
      this.layoutRefreshRaf = 0;
    }
    this.layoutRefreshRaf = requestAnimationFrame(() => {
      try {
        this.map.invalidateSize({ pan: false, animate: false });
      } catch (_) {}
      this.layoutRefreshRaf = requestAnimationFrame(() => {
        try {
          this.map.invalidateSize({ pan: false, animate: false });
        } catch (_) {}
      });
    });
  },

  centralizarEmpresa(openPopup = false) {
    this.initLeaflet();
    if (!this.map) return;
    try {
      const baseBounds = this.getBaseFocusBounds();
      if (baseBounds) {
        this.map.fitBounds(baseBounds.pad(0.45), { maxZoom: 18, animate: true });
      } else {
        this.map.setView([config.base.lat, config.base.lon], 18, { animate: true });
      }
      if (openPopup && this.companyMarker) this.companyMarker.openPopup();
      this.refreshMapLayout();
    } catch (_) {}
  },

  ajustarAoPerimetro() {
    this.initLeaflet();
    if (!this.map) return;
    const pts = [[config.base.lat, config.base.lon]];
    this.getCompanyAreaCoords().forEach((coord) => pts.push(coord));
    Object.values(this.markers || {}).forEach((marker) => {
      try {
        const latLng = marker.getLatLng();
        if (latLng) pts.push([latLng.lat, latLng.lng]);
      } catch (_) {}
    });
    if (pts.length <= 1) {
      this.centralizarEmpresa(true);
      return;
    }
    try {
      this.currentBounds = L.latLngBounds(pts);
      this.map.fitBounds(this.currentBounds, { padding: [42, 42], maxZoom: 19, animate: true });
      this.refreshMapLayout();
    } catch (_) {
      this.centralizarEmpresa(true);
    }
  },

  toggleFullScreen(forceState = null) {
    const overlay = this.getFullscreenOverlay();
    const mount = this.getFullscreenMount();
    const dock = this.getMapDock();
    const container = this.getMapContainer();
    if (!overlay || !mount || !dock || !container) return;

    const nextState = typeof forceState === 'boolean' ? forceState : !this.isFullscreen;
    if (nextState === this.isFullscreen) {
      this.refreshMapLayout();
      return;
    }

    if (nextState) {
      this.mapStageExpanded = true;
      safeStorage.setItem('cysyMapStageExpanded', 'true');
      this.updateMapStageUi();
    }

    this.initLeaflet();

    if (nextState) {
      overlay.hidden = false;
      mount.appendChild(container);
      document.body.classList.add('map-fullscreen-open');
      this.isFullscreen = true;
      requestAnimationFrame(() => overlay.classList.add('active'));
    } else {
      overlay.classList.remove('active');
      dock.appendChild(container);
      document.body.classList.remove('map-fullscreen-open');
      this.isFullscreen = false;
      setTimeout(() => {
        if (!this.isFullscreen) overlay.hidden = true;
      }, 180);
    }

    setTimeout(() => {
      this.refreshMapLayout();
      this.restoreViewport(false);
    }, 40);
    setTimeout(() => this.refreshMapLayout(), 240);
  },

  obterLocalizacaoPrecisa(timeoutMs = 18000) {
    return new Promise((resolve, reject) => {
      if (!navigator.geolocation) {
        reject(new Error('GPS_NAO_SUPORTADO'));
        return;
      }
      let best = null;
      let resolved = false;
      let watchId = null;
      const finish = (ok, err = null) => {
        if (resolved) return;
        resolved = true;
        try { if (watchId !== null) navigator.geolocation.clearWatch(watchId); } catch (_) {}
        if (ok && best) resolve(best);
        else reject(err || new Error('GPS_TIMEOUT'));
      };
      const onSuccess = (pos) => {
        if (!best || (pos.coords?.accuracy || Infinity) < (best.coords?.accuracy || Infinity)) {
          best = pos;
        }
        if ((best.coords?.accuracy || Infinity) <= 12) finish(true);
      };
      const onError = (err) => {
        if (best) finish(true);
        else finish(false, err);
      };
      try {
        watchId = navigator.geolocation.watchPosition(
          onSuccess,
          onError,
          { enableHighAccuracy: true, timeout: timeoutMs, maximumAge: 0 }
        );
      } catch (err) {
        finish(false, err);
        return;
      }
      setTimeout(() => {
        if (best) finish(true);
        else finish(false, new Error('GPS_TIMEOUT'));
      }, timeoutMs + 1000);
    });
  },

  async marcarMinhaLocalizacao() {
    if (!navigator.geolocation) {
      this.initLeaflet();
      if (this.map) this.map.setView([config.base.lat, config.base.lon], 19);
      alert("GPS não suportado. Exibindo localização base da Cysy.");
      return;
    }

    this.initLeaflet();

    uiBuilder.toggleLoader(true, "Buscando localização exata por satélite...");
    try {
      const pos = await this.obterLocalizacaoPrecisa(18000);
      const lat = Number(pos.coords.latitude);
      const lon = Number(pos.coords.longitude);
      const acc = Math.round(pos.coords.accuracy || 0);

      if (!this.map) return alert("Mapa ainda não inicializado.");

      this.map.invalidateSize(true);
      this.map.setView([lat, lon], 19);

      if (this.userMarker) this.map.removeLayer(this.userMarker);
      this.userMarker = L.marker([lat, lon]).addTo(this.map)
        .bindPopup(`<div style='text-align:center; font-family:Montserrat;'><b>📍 VOCÊ ESTÁ AQUI</b><br><span style="font-size:11px;">Precisão aprox.: ${acc} m</span></div>`)
        .openPopup();
      toastManager.show(`GPS capturado com precisão aproximada de ${acc} m.`, 'success');
    } catch (err) {
      if (this.map) this.map.setView([config.base.lat, config.base.lon], 19);
      alert(permissionManager.getGpsErrorMessage(err));
    } finally {
      uiBuilder.toggleLoader(false);
    }
  },

  focusLote(lote, openPopup = true) {
    const targetTab = document.getElementById('tab-mapa');
    if (targetTab && !targetTab.classList.contains('active')) {
      uiBuilder.switchTab(null, 'mapa');
      setTimeout(() => this.focusLote(lote, openPopup), 220);
      return;
    }

    this.initLeaflet();
    let parsed = this.parseLocationEntry(this.locDB[lote]);
    if (!parsed) {
      const logEntryRaw = this.locationLog?.[lote] || Object.values(this.locationLog || {}).find((row) =>
        this.getNormalizedLoteKey(row?.lote) === this.getNormalizedLoteKey(lote)
      );
      const logEntry = this.normalizeLocationLogEntry(logEntryRaw || {});
      if (logEntry && Number.isFinite(logEntry.lat) && Number.isFinite(logEntry.lon)) {
        parsed = {
          lat: Number(logEntry.lat),
          lon: Number(logEntry.lon),
          description: logEntry.description || 'Localização registrada no log compartilhado'
        };
      }
    }
    if (!parsed) {
      toastManager.show('Este lote ainda não possui localização salva.', 'warning');
      return;
    }

    try {
      this.selectLote(lote, { center: true, mapped: this.getRenderableMappedPoints(), zoom: 19 });
      if (openPopup && this.markers[lote]) {
        try { this.markers[lote].openTooltip(); } catch (_) {}
      }
      this.refreshMapLayout();
    } catch (_) {}
  },

  renderPanels(mapped = [], pending = [], proximity = null) {
    const visibleCount = document.getElementById('mapVisibleCount');
    const pendingCount = document.getElementById('mapPendingCount');
    const liveStatus = document.getElementById('mapLiveStatus');
    const syncBadge = document.getElementById('mapSyncBadge');
    const sheetHeadline = document.getElementById('mapSheetHeadline');
    const sheetSubhead = document.getElementById('mapSheetSubhead');
    const mappedList = document.getElementById('mapLotesList');
    const selectedCard = document.getElementById('mapSelectedCard');
    const selectedContent = document.getElementById('mapSelectedContent');
    const selectedMode = document.getElementById('mapSelectedModeLabel');
    const proximityData = proximity || this.getProximityWeb(mapped);
    const selected = this.getSelectedMappedPoint(mapped);
    const selectedConnections = this.getConnectionsForSelection(mapped, proximityData);
    const combinedItems = this.getCombinedPanelItems(mapped, pending);
    const filteredItems = this.applySheetFilters(combinedItems);
    const panelWindow = this.getPanelRenderWindow(filteredItems);
    const panelItems = panelWindow.items;

    if (visibleCount) visibleCount.textContent = String(mapped.length);
    if (pendingCount) pendingCount.textContent = String(pending.length);
    if (selectedMode) selectedMode.textContent = this.getVisualModeLabel(this.visualMode);
    if (liveStatus) {
      if (selected) {
        liveStatus.textContent = `Lote ${selected.lote} em foco. O mapa reduziu a densidade visual e destacou as conexões mais relevantes para navegação.`;
      } else if (mapped.length > 0) {
        liveStatus.textContent = `${mapped.length} lote(s) georreferenciado(s). Busque um lote, selecione-o e use o botão principal para registrar ou revisar o GPS sem recarregar o mapa inteiro.`;
      } else if (pending.length > 0) {
        liveStatus.textContent = `${pending.length} lote(s) ativos aguardando georreferenciamento para aparecerem aqui.`;
      } else {
        liveStatus.textContent = 'Nenhum lote ativo com saldo disponível no momento.';
      }
    }
    if (sheetHeadline) {
      sheetHeadline.textContent = selected
        ? `Lote ${selected.lote} em foco`
        : `${filteredItems.length} lote(s) visível(is) no painel`;
    }
    if (sheetSubhead) {
      sheetSubhead.textContent = selected
        ? 'Painel focado no lote selecionado. Use o botão de navegação para seguir até o ponto no terreno.'
        : (panelWindow.truncated
          ? `Mostrando ${panelItems.length} de ${filteredItems.length} lotes para manter o celular fluido. Use a busca ou os filtros para abrir a lista completa.`
          : 'Fluxo recomendado: busque o lote, toque em “Selecionar lote” e só então registre o GPS pelo botão principal.');
    }
    if (syncBadge) {
      if (!navigator.onLine) {
        syncBadge.textContent = '● Offline com dados locais';
        syncBadge.className = 'map-sync-inline map-sync-inline-waiting';
      } else if (this.lastSharedSyncError) {
        syncBadge.textContent = '● Falha ao atualizar mapa compartilhado';
        syncBadge.className = 'map-sync-inline map-sync-inline-blocked';
      } else if (this.lastSharedSyncAt) {
        syncBadge.textContent = `● Compartilhado às ${new Date(this.lastSharedSyncAt).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`;
        syncBadge.className = 'map-sync-inline map-sync-inline-ok';
      } else {
        syncBadge.textContent = '● Mapa online e sincronizado';
        syncBadge.className = 'map-sync-inline map-sync-inline-ok';
      }
    }
    if (selectedCard) selectedCard.hidden = false;

    if (selectedContent) {
      const selectedStatus = selected ? this.getMappedFreshnessStatus(selected) : null;
      const selectedLogEntry = selected ? this.getLocationLogEntryByLote(selected.lote) : null;
      const selectedSyncMeta = selected ? this.getLocationSyncMeta(selectedLogEntry || { lote: selected.lote, source: 'compartilhado' }) : null;
      selectedContent.innerHTML = selected
        ? `<article class="map-focus-card">
            <div class="map-selected-header">
              <div class="map-selected-badge">${escapeHTML(this.getMarkerLabel(selected.lote, mapped.findIndex((item) => item.lote === selected.lote) + 1))}</div>
              <div>
                <div class="map-selected-code">${escapeHTML(selected.lote)}</div>
                <div class="map-selected-product">${escapeHTML(selected.produto || 'Material não informado')}</div>
              </div>
              <span class="${selectedStatus.badgeClass}">${escapeHTML(selectedStatus.label)}</span>
            </div>
            <div class="map-selected-grid">
              <div class="map-selected-chip">
                <span>Quantidade disponível</span>
                <strong>${selected.active ? `${formatTons(selected.saldo)} t` : 'Registro compartilhado'}</strong>
              </div>
              <div class="map-selected-chip">
                <span>Coordenadas</span>
                <strong>${selected.lat.toFixed(6)}, ${selected.lon.toFixed(6)}</strong>
              </div>
            </div>
            <div class="map-selected-sync">
              <span class="${selectedSyncMeta.className}">${escapeHTML(selectedSyncMeta.label)}</span>
              <div class="map-selected-note">${escapeHTML(selectedLogEntry?.syncMessage || selectedSyncMeta.summary)}</div>
            </div>
            <div class="map-selected-note">${escapeHTML(selected.description || 'Sem referência operacional complementar.')}</div>
            ${selected.updatedAt || selected.updatedBy
              ? `<div class="map-selected-meta">Atualizado em ${escapeHTML(selected.updatedAt || 'Horário não informado')}${selected.updatedBy ? ` por ${escapeHTML(selected.updatedBy)}` : ''}</div>`
              : ''
            }
            <div class="map-selected-actions">
              <button type="button" class="btn-primary" onclick="mapController.setRegistrationTarget(decodeURIComponent('${encodeURIComponent(selected.lote)}'), { refreshSearch: true })">Selecionar para GPS</button>
              <button type="button" class="btn-danger" onclick="mapController.deleteSingleLote(decodeURIComponent('${encodeURIComponent(selected.lote)}'))">Apagar do mapa</button>
              <button type="button" class="btn-secondary" onclick="mapController.selectLote(decodeURIComponent('${encodeURIComponent(selected.lote)}'), { center: true, mapped: mapController.lastRenderedMapped, zoom: 19 })">Centralizar lote</button>
              <button type="button" class="btn-success" onclick="mapController.confirmNavigationToLote(decodeURIComponent('${encodeURIComponent(selected.lote)}'))">Ir até este lote</button>
              <button type="button" class="btn-secondary" onclick="mapController.clearSelectedLote(true)">Fechar foco</button>
            </div>
            <div class="map-selected-connections">
              <h4>Conexões relevantes</h4>
              ${selectedConnections.length > 0
                ? selectedConnections.slice(0, 4).map((edge) => `
                    <div class="map-selected-connection">
                      <strong>${escapeHTML(edge.other.lote)}</strong>
                      <span>${escapeHTML(this.formatDistanceLabel(edge.distance))} • ${escapeHTML(edge.direction || 'direção indefinida')}</span>
                    </div>
                  `).join('')
                : `<div class="map-empty-card">Sem conexão operacional relevante para este lote no zoom atual.</div>`
              }
            </div>
          </article>`
        : `<div class="map-focus-empty">Selecione um marcador para abrir o lote em foco e liberar as ações rápidas do painel.</div>`;
    }

    if (mappedList) {
      mappedList.innerHTML = panelItems.length > 0
        ? panelItems.map((item, index) => {
            const loteToken = encodeURIComponent(item.lote);
            const isSelected = Boolean(selected) && this.getNormalizedLoteKey(selected.lote) === this.getNormalizedLoteKey(item.lote);
            const isRegistrationTarget = this.getNormalizedLoteKey(this.selectedRegistrationLote) === this.getNormalizedLoteKey(item.lote);
            const saldoLabel = item.active ? `${formatTons(item.saldo)} t` : (Number(item.saldo || 0) > 0 ? `${formatTons(item.saldo)} t` : 'LOG');
            const status = item.panelStatus || this.getMappedFreshnessStatus(item);
            const markerLabel = item.hasCoordinates ? this.getMarkerLabel(item.lote, index + 1) : 'GPS';
            const logEntry = this.getLocationLogEntryByLote(item.lote);
            const syncMeta = this.getLocationSyncMeta(logEntry || { lote: item.lote, source: item.hasCoordinates ? 'compartilhado' : 'local' });
            const primaryActionButtons = item.hasCoordinates
              ? `<button type="button" class="${isRegistrationTarget ? 'btn-success' : ''}" onclick="mapController.setRegistrationTarget(decodeURIComponent('${loteToken}'), { refreshSearch: true })">${isRegistrationTarget ? '✅ Lote selecionado' : 'Selecionar lote'}</button>
                 <button type="button" onclick="typeof mapOperacionalV2!=='undefined'?mapOperacionalV2.showLoteDetails(decodeURIComponent('${loteToken}')):null" style="background:linear-gradient(135deg,#2563EB,#1D4ED8);color:#fff;font-weight:800;">📋 Ver detalhes</button>`
              : `<button type="button" class="${isRegistrationTarget ? 'btn-success' : ''}" onclick="mapController.setRegistrationTarget(decodeURIComponent('${loteToken}'), { refreshSearch: true })">${isRegistrationTarget ? '✅ Lote selecionado' : 'Selecionar lote'}</button>
                 <button type="button" onclick="mapController.handleRegisterSelectedLote()">⚡ Registrar GPS</button>`;
            const secondaryActionButtons = item.hasCoordinates
              ? `<button type="button" onclick="mapController.selectLote(decodeURIComponent('${loteToken}'), { center: true })">${isSelected ? '🎯 Atualizar foco' : '📍 Focar no mapa'}</button>
                 <button type="button" onclick="mapController.confirmNavigationToLote(decodeURIComponent('${loteToken}'))">Ir até o local</button>
                 <button type="button" class="btn-danger" onclick="mapController.deleteSingleLote(decodeURIComponent('${loteToken}'))">🗑️ Apagar</button>`
              : `<button type="button" onclick="mapController.capturarGPSEditar(decodeURIComponent('${loteToken}'))">Capturar GPS agora</button>`;
            const actionButtons = this.isCompactMobile()
              ? `${primaryActionButtons}
                 <details class="map-card-more">
                   <summary>Mais ações</summary>
                   <div class="map-lote-actions map-lote-actions--nested">${secondaryActionButtons}</div>
                 </details>`
              : `${primaryActionButtons}${secondaryActionButtons}`;
            return `<article class="map-lote-card ${isSelected ? 'is-selected' : ''}" data-status="${escapeHTML(status.code)}">
              <div class="map-lote-top">
                <div>
                  <div class="map-lote-code">${escapeHTML(item.lote)}</div>
                  <div class="map-lote-product">${escapeHTML(item.produto || 'Produto não informado')}</div>
                </div>
                <div>
                  <div class="map-lote-saldo">${escapeHTML(saldoLabel)}</div>
                  <span class="${status.badgeClass}">${escapeHTML(status.label)}</span>
                </div>
              </div>
              <div class="map-lote-meta">
                <span><strong>Marcador</strong> ${escapeHTML(markerLabel)}</span>
                ${item.hasCoordinates ? `<span><strong>Coords</strong> ${item.lat.toFixed(5)}, ${item.lon.toFixed(5)}</span>` : `<span><strong>GPS</strong> pendente</span>`}
              </div>
              <div class="map-lote-location">
                <b>Resumo:</b> ${escapeHTML(status.summary)}<br>
                <b>Material:</b> ${escapeHTML(item.produto || 'Produto não informado')}<br>
                <b>Referência:</b> ${escapeHTML(this.truncateMapText(item.description || 'Sem referência operacional.', 90))}<br>
                <span class="${syncMeta.className}">${escapeHTML(syncMeta.label)}</span>
                ${logEntry?.syncMessage ? `<br><span style="color:#334155;">${escapeHTML(logEntry.syncMessage)}</span>` : ''}
                ${item.updatedAt ? `<br><span style="color:#475569;">Atualizado em ${escapeHTML(item.updatedAt)}${item.updatedBy ? ` por ${escapeHTML(item.updatedBy)}` : ''}</span>` : ''}
              </div>
              <div class="map-lote-actions">
                ${actionButtons}
              </div>
            </article>`;
          }).join('')
        : `<div class="map-empty-card">Nenhum lote corresponde aos filtros atuais. Ajuste a busca ou o status para ampliar a leitura.</div>`;
      if (panelWindow.truncated) {
        mappedList.insertAdjacentHTML('beforeend', `<div class="map-empty-card">Lista resumida no celular: ${panelWindow.hiddenCount} lote(s) ficaram ocultos para preservar a fluidez desta tela.</div>`);
      }
    }

    this.renderTechnicalPanel(mapped, pending, proximityData);
    this.updateSheetUi();
    this.updateStatusFilterControls();
    this.updateRegistrationUi();
  },

  renderPins(force = false) {
    const now = Date.now();
    const minInterval = this.isCompactMobile() ? 240 : 130;
    if (!force && now - this.lastRenderAt < minInterval) {
      if (this.pendingRenderTimer) clearTimeout(this.pendingRenderTimer);
      this.pendingRenderTimer = setTimeout(() => {
        this.pendingRenderTimer = 0;
        this.renderPins(true);
      }, this.isCompactMobile() ? 260 : 140);
      return;
    }
    this.lastRenderAt = now;
    this.purgeIgnoredLocationEntries();
    this.purgeUnmappableLocationEntries();
    this.rebuildLocationLogFromLocDb();
    const visiveis = this.getVisibleLotes();
    const mapped = this.getRenderableMappedPoints();
    const pending = [];

    visiveis.forEach((item) => {
      const parsed = this.parseLocationEntry(this.locDB[item.lote]);
      if (!parsed) pending.push(item);
    });

    const selected = this.getSelectedMappedPoint(mapped);
    if (this.selectedLote && !selected) {
      this.selectedLote = '';
      safeStorage.removeItem('cysySelectedMapLote');
    }
    const proximity = this.isCompactMobile() && mapped.length > 12
      ? { edges: [], averageNearestDistance: 0, threshold: 0 }
      : this.getProximityWeb(mapped);
    this.lastRenderedMapped = mapped;
    this.lastRenderedPending = pending;
    this.lastRenderedProximity = proximity;

    this.renderPanels(mapped, pending, proximity);
    if (this.sheetTab === 'tech') {
      this.renderLocationLog(document.getElementById('inpBuscaMapaLog')?.value || '');
    }

    const emptyState = document.getElementById('mapEmptyState');
    if (!this.map || !this.mapStageExpanded) {
      if (emptyState) emptyState.style.display = 'none';
      return;
    }

    this.clearMapDynamicLayers();

    const bounds = [[config.base.lat, config.base.lon]];
    const layoutPoints = this.buildLiveLayoutPoints(mapped);
    const palette = this.getMarkerPalette();
    const layoutConfig = this.getLiveLayoutConfig(layoutPoints.length);
    const selectedKey = this.getNormalizedLoteKey(this.selectedLote);

    layoutPoints.forEach((item, index) => {
      const isSelected = selectedKey && this.getNormalizedLoteKey(item.lote) === selectedKey;
      const isDimmed = Boolean(selectedKey) && !isSelected;
      const statusMeta = this.getMappedFreshnessStatus(item);
      const color = this.getStatusColor(statusMeta.code, palette[index % palette.length]);
      const leaderLatLng = item.badgeLatLng;

      if (!this.isCompactMobile() && layoutConfig.showLeaders && item.displaced) {
        const leader = L.polyline([
          [item.lat, item.lon],
          [leaderLatLng.lat, leaderLatLng.lng]
        ], {
          color,
          weight: isSelected ? 2.4 : 1.8,
          opacity: isSelected ? 0.72 : (isDimmed ? 0.16 : 0.34),
          dashArray: isSelected ? '' : '4 6',
          interactive: false
        }).addTo(this.map);
        this.decorationLayers.push(leader);
      }

      if (isSelected && !this.isCompactMobile()) {
        const halo = L.circleMarker([item.lat, item.lon], {
          radius: layoutConfig.pointRadius + 11,
          color: color,
          weight: 2,
          opacity: 0.85,
          fillColor: color,
          fillOpacity: 0.12,
          interactive: false
        }).addTo(this.map);
        this.decorationLayers.push(halo);
      }

      const pointDot = L.circleMarker([item.lat, item.lon], {
        radius: isSelected ? layoutConfig.pointRadius + 2 : layoutConfig.pointRadius,
        color: '#FFFFFF',
        weight: 2,
        fillColor: color,
        fillOpacity: isDimmed ? 0.35 : 0.98,
        opacity: isDimmed ? 0.4 : 1,
        interactive: false
      }).addTo(this.map);
      this.decorationLayers.push(pointDot);

      const marker = L.marker([leaderLatLng.lat, leaderLatLng.lng], {
        icon: L.divIcon({
          className: 'map-live-badge-shell',
          html: `<div class="map-live-badge map-live-badge-${escapeHTML(statusMeta.code)} ${item.active ? 'map-live-badge-active' : 'map-live-badge-shared'} ${isSelected ? 'map-live-badge-selected' : ''} ${isDimmed ? 'map-live-badge-muted' : ''}">
            <span class="map-live-badge-code">${escapeHTML(item.label)}</span>
            <button type="button" class="map-live-badge-info" data-map-info="${escapeHTML(item.lote)}" aria-label="Abrir detalhes do lote ${escapeHTML(item.lote)}">i</button>
          </div>`,
          iconSize: [item.size.width, item.size.height],
          iconAnchor: [item.size.width / 2, item.size.height / 2]
        }),
        keyboard: true
      }).addTo(this.map);

      if (!this.isCompactMobile() || isSelected) {
        marker.bindTooltip(
          `<div><strong>Lote ${escapeHTML(item.lote)}</strong><br>${escapeHTML(this.truncateMapText(item.produto || 'Material não informado', 42))}<br>${item.active ? `${formatTons(item.saldo)} t em estoque` : 'Registro do log compartilhado'}</div>`,
          {
            direction: 'top',
            offset: [0, -16],
            className: isSelected ? 'map-live-focus-tooltip' : 'lote-map-tooltip',
            permanent: isSelected && this.visualMode !== 'clean' && !this.isCompactMobile(),
            sticky: !isSelected && !this.isCompactMobile()
          }
        );
      }
      marker.on('click', () => {
        this.selectLote(item.lote, { center: false, mapped });
        if (typeof mapOperacionalV2 !== 'undefined' && typeof mapOperacionalV2.showLoteDetails === 'function') {
          mapOperacionalV2.showLoteDetails(item.lote);
        }
      });
      marker.on('add', () => {
        const markerElement = marker.getElement();
        const infoButton = markerElement?.querySelector('[data-map-info]');
        if (!infoButton) return;
        L.DomEvent.disableClickPropagation(infoButton);
        L.DomEvent.on(infoButton, 'click', (event) => {
          L.DomEvent.stop(event);
          // Abre painel rico com produto, saldo e navegação (mesmo padrão da aba Liberação)
          if (typeof mapOperacionalV2 !== 'undefined' && typeof mapOperacionalV2.showLoteDetails === 'function') {
            mapOperacionalV2.showLoteDetails(item.lote);
          } else {
            this.selectLote(item.lote, { center: false, mapped });
          }
        });
      });
      this.markers[item.lote] = marker;
      bounds.push([item.lat, item.lon]);
    });

    this.currentBounds = bounds.length > 1 ? L.latLngBounds(bounds) : null;
    if (emptyState) emptyState.style.display = 'none';
    this.renderConnectionWeb(mapped);
    this.refreshMapLayout();
    this.scheduleOfflineMapSave(false);
  },

  buildLocationPayload(lote, lat, lon, desc, options = {}) {
    const loteName = String(lote || '').trim();
    const latitude = Number(lat);
    const longitude = Number(lon);
    const currentUser = String(options.usuario || safeStorage.getItem('cysyUser', '') || '').trim();
    const updatedAt = String(options.dataHora || new Date().toLocaleString('pt-BR'));
    const basePayload = {
      tipo: 'LOCALIZACAO',
      acao: 'UPSERT_LOCALIZACAO_LOTE',
      origem: 'MAPA_LOTES',
      usuario: currentUser,
      lote: loteName,
      latitude: Number.isFinite(latitude) ? Number(latitude.toFixed(6)) : '',
      longitude: Number.isFinite(longitude) ? Number(longitude.toFixed(6)) : '',
      descricao: String(desc || '').trim(),
      lat: Number.isFinite(latitude) ? Number(latitude.toFixed(6)) : '',
      lon: Number.isFinite(longitude) ? Number(longitude.toFixed(6)) : '',
      desc: String(desc || '').trim(),
      dataHora: updatedAt
    };
    const meta = dbManager.buildMeta('LOCALIZACAO', basePayload, options.meta || {});
    return {
      payload: {
        ...basePayload,
        syncId: meta.syncId,
        localId: meta.localId,
        fingerprint: meta.fingerprint
      },
      meta,
      updatedAt,
      currentUser
    };
  },

  recordLocationAudit(stage, payload = {}, extra = {}) {
    try {
      backupManager.addEntry('LOCALIZACAO_SYNC', {
        stage,
        lote: String(payload.lote || '').trim(),
        syncId: String(payload.syncId || '').trim(),
        localId: String(payload.localId || '').trim(),
        usuario: String(payload.usuario || '').trim(),
        dataHora: new Date().toLocaleString('pt-BR'),
        ...extra
      });
    } catch (_) {}
  },

  applyLocationDraftFromPayload(payload = {}, options = {}) {
    const lote = String(payload.lote || '').trim();
    const lat = Number(payload.latitude ?? payload.lat ?? '');
    const lon = Number(payload.longitude ?? payload.lon ?? '');
    if (!lote || !Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    const desc = String(payload.descricao || payload.desc || '').trim();
    const updatedAt = String(payload.dataHora || new Date().toLocaleString('pt-BR'));
    const updatedBy = String(payload.usuario || safeStorage.getItem('cysyUser', '') || '').trim();
    this.clearManualDeletedMark(lote);
    this.locDB[lote] = this.getLocationString(lat.toFixed(6), lon.toFixed(6), desc);
    this.setPendingSharedConfirmation(lote, lat, lon, desc, updatedAt, updatedBy);
    this.upsertLocationLogEntry({
      lote,
      lat,
      lon,
      description: desc,
      updatedAt,
      updatedBy,
      source: String(options.source || 'local_draft').trim(),
      product: options.product || '',
      saldo: options.saldo,
      syncStatus: String(options.syncStatus || 'enviando').trim(),
      syncMessage: String(options.syncMessage || '').trim(),
      lastAttemptAt: String(options.lastAttemptAt || new Date().toISOString()).trim(),
      confirmationSource: String(options.confirmationSource || '').trim(),
      syncId: String(payload.syncId || '').trim(),
      localId: String(payload.localId || '').trim(),
      fingerprint: String(payload.fingerprint || '').trim()
    });
    if (options.persist !== false) this.persistLocationState();
    if (options.render === 'full') this.renderPins(true);
    else if (options.render === 'panel') this.renderPanelsFromSnapshot(true);
    return { lote, lat, lon, desc, updatedAt, updatedBy };
  },

  removeStaleSharedLocationEntries(globalData = []) {
    const sharedKeys = new Set();
    const startIndex = String(globalData?.[0]?.[0] || '').trim().toUpperCase() === 'LOTE' ? 1 : 0;
    for (let index = startIndex; index < (Array.isArray(globalData) ? globalData.length : 0); index++) {
      const lote = String(globalData[index]?.[0] || '').trim();
      const key = this.getNormalizedLoteKey(lote);
      if (key) sharedKeys.add(key);
    }
    let changed = false;
    Object.values(this.locationLog || {}).forEach((row) => {
      const entry = this.normalizeLocationLogEntry(row || {});
      if (!entry) return;
      const loteKey = this.getNormalizedLoteKey(entry.lote);
      const sourceKey = String(entry.confirmationSource || entry.source || '').toLowerCase();
      const isSharedOwned = sourceKey.includes('compartilhado') || sourceKey.includes('mapa_compartilhado') || sourceKey.includes('shared') || sourceKey === 'cache_local';
      if (!loteKey || !isSharedOwned || sharedKeys.has(loteKey) || this.getPendingSharedConfirmation(entry.lote)) return;
      delete this.locationLog[entry.lote];
      const parsed = this.parseLocationEntry(this.locDB?.[entry.lote]);
      if (!parsed || !Number.isFinite(entry.lat) || !Number.isFinite(entry.lon) ||
        this.calculateDistanceMeters(parsed.lat, parsed.lon, entry.lat, entry.lon) <= 6) {
        delete this.locDB[entry.lote];
      }
      changed = true;
    });
    if (changed) this.persistLocationState();
    return changed;
  },

  findSharedLocationConfirmation(globalData, lote, lat, lon, toleranceMeters = 6) {
    if (!Array.isArray(globalData) || !lote) return null;
    const loteKey = this.getNormalizedLoteKey(lote);
    const targetLat = Number(lat);
    const targetLon = Number(lon);
    if (!loteKey || !Number.isFinite(targetLat) || !Number.isFinite(targetLon)) return null;
    const startIndex = String(globalData?.[0]?.[0] || '').trim().toUpperCase() === 'LOTE' ? 1 : 0;
    for (let index = startIndex; index < globalData.length; index++) {
      const rowLote = String(globalData[index]?.[0] || '').trim();
      if (this.getNormalizedLoteKey(rowLote) !== loteKey) continue;
      const rowLat = this.normalizeSharedCoordinate(globalData[index]?.[1]);
      const rowLon = this.normalizeSharedCoordinate(globalData[index]?.[2]);
      if (!Number.isFinite(rowLat) || !Number.isFinite(rowLon)) continue;
      const distance = this.calculateDistanceMeters(targetLat, targetLon, rowLat, rowLon);
      if (distance <= Number(toleranceMeters || 6)) {
        return { row: globalData[index], distance };
      }
    }
    return null;
  },

  async waitForSharedLocationSync(lote, lat, lon, options = {}) {
    const timeoutMs = Math.max(4000, Number(options.timeoutMs || this.sharedConfirmationTimeoutMs || 20000));
    const intervalMs = Math.max(800, Number(options.intervalMs || this.sharedConfirmationIntervalMs || 1400));
    const toleranceMeters = Math.max(2, Number(options.toleranceMeters || 6));
    const startedAt = Date.now();
    let attempts = 0;
    let lastRows = Array.isArray(this.lastSharedRows) ? this.lastSharedRows : [];
    const previousInFlightState = this.sharedSyncInFlight;
    this.sharedSyncInFlight = true;

    try {
      while (Date.now() - startedAt <= timeoutMs) {
        attempts += 1;
        try {
          const rows = await apiService.fetchGlobalLoc();
          if (Array.isArray(rows)) {
            lastRows = rows;
            this.lastSharedRows = rows;
            const match = this.findSharedLocationConfirmation(rows, lote, lat, lon, toleranceMeters);
            if (match) {
              this.lastSharedFetchAt = Date.now();
              this.syncFromGlobal(rows, { render: false });
              this.renderPins(true);
              return { confirmed: true, attempts, distanceMeters: match.distance };
            }
          }
        } catch (err) {
          this.lastSharedSyncError = err?.message || 'Falha ao confirmar o mapa compartilhado.';
        }

        if (Date.now() - startedAt + intervalMs > timeoutMs) break;
        await new Promise((resolve) => setTimeout(resolve, intervalMs));
      }

      if (Array.isArray(lastRows)) {
        this.lastSharedRows = lastRows;
      }
      this.renderPins(true);
      return { confirmed: false, attempts };
    } finally {
      this.sharedSyncInFlight = previousInFlightState;
    }
  },

  async confirmLocationPayloadShared(payload = {}, options = {}) {
    const lote = String(payload.lote || '').trim();
    const lat = Number(payload.latitude ?? payload.lat ?? '');
    const lon = Number(payload.longitude ?? payload.lon ?? '');
    if (!lote || !Number.isFinite(lat) || !Number.isFinite(lon)) {
      return {
        ok: false,
        syncStatus: 'pendente_com_falha',
        syncMessage: 'Payload de localização inválido para confirmação compartilhada.',
        confirmed: false,
        attempts: 0
      };
    }
    const waitingMessage = String(options.waitingMessage || 'Aguardando o lote aparecer no mapa compartilhado para todos os usuários.').trim();
    this.setLocationSyncState(lote, {
      syncStatus: 'aguardando_confirmacao_global',
      syncMessage: waitingMessage,
      lastAttemptAt: new Date().toISOString(),
      source: String(options.source || 'local_draft').trim(),
      confirmationSource: '',
      syncId: String(payload.syncId || '').trim(),
      localId: String(payload.localId || '').trim(),
      fingerprint: String(payload.fingerprint || '').trim()
    }, { persist: true, render: 'panel' });
    this.recordLocationAudit('aguardando_confirmacao_global', payload, {
      origemFluxo: String(options.auditOrigin || '').trim()
    });
    const sharedSync = await this.waitForSharedLocationSync(lote, lat, lon, {
      timeoutMs: options.timeoutMs,
      intervalMs: options.intervalMs,
      toleranceMeters: options.toleranceMeters
    });
    if (sharedSync.confirmed) {
      const successMessage = String(options.successMessage || 'GPS confirmado no mapa compartilhado para todos os usuários.').trim();
      this.setLocationSyncState(lote, {
        syncStatus: 'sincronizado',
        syncMessage: successMessage,
        lastAttemptAt: new Date().toISOString(),
        source: 'compartilhado',
        confirmationSource: String(options.confirmationSource || 'mapa_compartilhado').trim(),
        syncId: String(payload.syncId || '').trim(),
        localId: String(payload.localId || '').trim(),
        fingerprint: String(payload.fingerprint || '').trim()
      }, { persist: true, render: 'full' });
      this.recordLocationAudit('confirmado_global', payload, {
        tentativas: Number(sharedSync.attempts || 0),
        distanciaMetros: Number(sharedSync.distanceMeters || 0),
        origemFluxo: String(options.auditOrigin || '').trim()
      });
      return {
        ok: true,
        syncStatus: 'sincronizado',
        syncMessage: successMessage,
        confirmed: true,
        attempts: sharedSync.attempts,
        distanceMeters: sharedSync.distanceMeters
      };
    }
    const failureMessage = String(options.failureMessage || 'Registro enviado, mas o mapa compartilhado ainda não confirmou este lote para todos os usuários.').trim();
    this.setLocationSyncState(lote, {
      syncStatus: 'pendente_com_falha',
      syncMessage: failureMessage,
      lastAttemptAt: new Date().toISOString(),
      source: String(options.source || 'local_draft').trim(),
      confirmationSource: '',
      syncId: String(payload.syncId || '').trim(),
      localId: String(payload.localId || '').trim(),
      fingerprint: String(payload.fingerprint || '').trim()
    }, { persist: true, render: 'full' });
    this.recordLocationAudit('confirmacao_pendente', payload, {
      tentativas: Number(sharedSync.attempts || 0),
      origemFluxo: String(options.auditOrigin || '').trim()
    });
    return {
      ok: true,
      syncStatus: 'pendente_com_falha',
      syncMessage: failureMessage,
      confirmed: false,
      attempts: sharedSync.attempts
    };
  },

  syncFromGlobal(globalData, options = {}) {
    if (!globalData || !Array.isArray(globalData)) return;
    const shouldRender = options.render !== false;
    this.lastSharedRows = globalData;
    this.purgeUnmappableLocationEntries();
    this.removeStaleSharedLocationEntries(globalData);
    let changed = false;
    const startIndex = String(globalData?.[0]?.[0] || '').trim().toUpperCase() === 'LOTE' ? 1 : 0;
    for (let i = startIndex; i < globalData.length; i++) {
      const lote = String(globalData[i]?.[0] || '').trim();
      if (this.shouldRemoveLoteFromMap(lote)) continue;
      const lat = this.normalizeSharedCoordinate(globalData[i]?.[1]);
      const lon = this.normalizeSharedCoordinate(globalData[i]?.[2]);
      if (this.shouldPreservePendingLocation(lote, lat, lon)) continue;
      const desc = globalData[i]?.[3];
      let str = desc || '';
      if (Number.isFinite(lat) && Number.isFinite(lon)) str = `📍 Lat: ${lat.toFixed(6)}, Lon: ${lon.toFixed(6)} \n${str}`.trim();
      if (lote && str && this.locDB[lote] !== str) {
        this.locDB[lote] = str;
        changed = true;
      }
    }
    if (changed) {
      this.persistLocationState();
    }
    for (let i = startIndex; i < globalData.length; i++) {
      const lote = String(globalData[i]?.[0] || '').trim();
      if (this.shouldRemoveLoteFromMap(lote)) continue;
      const lat = this.normalizeSharedCoordinate(globalData[i]?.[1]);
      const lon = this.normalizeSharedCoordinate(globalData[i]?.[2]);
      if (this.shouldPreservePendingLocation(lote, lat, lon)) continue;
      const desc = globalData[i]?.[3];
      const updatedAt = this.normalizeSharedUpdatedAt(globalData[i]?.[4] || '');
      const updatedBy = globalData[i]?.[5] || '';
      if (!lote) continue;
      this.upsertLocationLogEntry({
        lote,
        lat,
        lon,
        description: desc,
        updatedAt,
        updatedBy,
        source: 'compartilhado',
        syncStatus: 'sincronizado',
        syncMessage: 'GPS confirmado no mapa compartilhado para todos os usuários.',
        lastAttemptAt: new Date().toISOString(),
        confirmationSource: 'mapa_compartilhado',
        syncId: String(globalData[i]?.[6] || '').trim()
      });
    }
    this.persistLocationState();
    this.lastSharedSyncAt = new Date().toISOString();
    this.lastSharedSyncError = '';
    if (shouldRender) this.renderPins(true);
  },

  async salvarLocalizacaoServer(lote, lat, lon, desc) {
    const hadPreviousLoc = Object.prototype.hasOwnProperty.call(this.locDB || {}, lote);
    const previousLoc = hadPreviousLoc ? this.locDB[lote] : '';
    const hadPreviousLog = Object.prototype.hasOwnProperty.call(this.locationLog || {}, lote);
    const previousLog = hadPreviousLog ? { ...(this.locationLog[lote] || {}) } : null;
    const previousPendingConfirmation = this.getPendingSharedConfirmation(lote);
    const restorePreviousState = () => {
      if (hadPreviousLoc) this.locDB[lote] = previousLoc;
      else delete this.locDB[lote];
      if (hadPreviousLog && previousLog) this.locationLog[lote] = previousLog;
      else delete this.locationLog[lote];
      if (previousPendingConfirmation) {
        this.setPendingSharedConfirmation(
          previousPendingConfirmation.lote || lote,
          previousPendingConfirmation.lat,
          previousPendingConfirmation.lon,
          previousPendingConfirmation.description,
          previousPendingConfirmation.updatedAt,
          previousPendingConfirmation.updatedBy
        );
      } else {
        this.clearPendingSharedConfirmation(lote);
      }
      this.persistLocationState();
      this.renderPins();
    };
    const { payload, meta } = this.buildLocationPayload(lote, lat, lon, desc);
    const pendingMessage = navigator.onLine
      ? 'O registro foi enviado, mas ainda aguarda confirmação no mapa compartilhado.'
      : 'Sem internet no momento. O lote ficou na fila e será confirmado quando voltar a conexão.';
    this.applyLocationDraftFromPayload(payload, {
      source: navigator.onLine ? 'local_draft' : 'fila_offline',
      syncStatus: 'enviando',
      syncMessage: 'Enviando o GPS do lote para o mapa compartilhado...',
      render: 'full'
    });
    this.recordLocationAudit('captura_concluida', payload, {
      descricao: String(desc || '').trim()
    });
    try {
      const resp = await apiService.sendDataToAppScript(payload);
      if (resp && resp.success) {
        this.recordLocationAudit('post_aceito', payload, {
          resposta: 'success'
        });
        const sharedSync = await this.confirmLocationPayloadShared(payload, {
          source: 'local_draft',
          timeoutMs: this.sharedConfirmationTimeoutMs,
          intervalMs: this.sharedConfirmationIntervalMs,
          toleranceMeters: 6,
          auditOrigin: 'salvamento_direto',
          confirmationSource: 'mapa_compartilhado'
        });
        if (sharedSync.confirmed) {
          dbManager.markSent(meta.syncId, meta.fingerprint, 'LOCALIZACAO', payload);
        }
        backupManager.addEntry('LOCALIZACAO', payload);
        syncManager.refreshSyncView();
        return {
          ok: Boolean(sharedSync.confirmed),
          syncStatus: sharedSync.syncStatus,
          syncMessage: sharedSync.syncMessage,
          confirmed: Boolean(sharedSync.confirmed),
          syncId: payload.syncId,
          attempts: sharedSync.attempts,
          distanceMeters: sharedSync.distanceMeters
        };
      }
    } catch (e) {
      this.recordLocationAudit('post_falhou', payload, {
        erro: String(e?.message || 'Falha ao enviar localização').slice(0, 240)
      });
      const queued = await dbManager.savePending('LOCALIZACAO', payload, {
        syncId: meta.syncId,
        localId: meta.localId,
        fingerprint: meta.fingerprint,
        status: navigator.onLine ? 'ERRO' : 'AGUARDANDO_INTERNET'
      });
      if (!queued?.ok) {
        if (queued.reason === 'OFFLINE_LIMIT' && queued.queueType === 'LOCALIZACAO') {
          restorePreviousState();
          syncManager.refreshSyncView();
          toastManager.show(`Limite offline do mapa atingido: máximo de ${dbManager.maxOfflineLocalizacoes} lotes pendentes sem internet.`, 'warning', 5200);
          try { uiBuilder.switchTab(null, 'sync'); } catch (_) {}
          return {
            ok: false,
            syncStatus: 'pendente_com_falha',
            syncMessage: `Limite offline do mapa atingido: máximo de ${dbManager.maxOfflineLocalizacoes} lotes pendentes sem internet.`,
            confirmed: false,
            blocked: true
          };
        }
        if (queued.reason === 'DUPLICADO_LOCAL') {
          backupManager.addEntry('LOCALIZACAO', payload);
          this.setLocationSyncState(lote, {
            syncStatus: 'pendente_com_falha',
            syncMessage: 'Este lote já estava aguardando sincronização compartilhada na fila local.',
            source: navigator.onLine ? 'local_draft' : 'fila_offline',
            confirmationSource: 'fila_offline',
            lastAttemptAt: new Date().toISOString(),
            syncId: payload.syncId,
            localId: payload.localId,
            fingerprint: payload.fingerprint
          }, { persist: true, render: 'full' });
          syncManager.refreshSyncView();
          return {
            ok: true,
            syncStatus: 'pendente_com_falha',
            syncMessage: 'Este lote já estava aguardando confirmação compartilhada na fila local.',
            confirmed: false,
            queued: true,
            syncId: payload.syncId
          };
        }
        if (queued.reason === 'JA_ENVIADO') {
          this.setLocationSyncState(lote, {
            syncStatus: 'sincronizado',
            syncMessage: 'Esse lote já estava sincronizado anteriormente.',
            source: 'compartilhado',
            confirmationSource: 'sent_index',
            lastAttemptAt: new Date().toISOString(),
            syncId: payload.syncId,
            localId: payload.localId,
            fingerprint: payload.fingerprint
          }, { persist: true, render: 'full' });
          syncManager.refreshSyncView();
          return {
            ok: true,
            syncStatus: 'sincronizado',
            syncMessage: 'Esse lote já estava sincronizado anteriormente.',
            confirmed: true,
            syncId: payload.syncId
          };
        }
      }
      backupManager.addEntry('LOCALIZACAO', payload);
      this.setLocationSyncState(lote, {
        syncStatus: 'pendente_com_falha',
        syncMessage: pendingMessage,
        source: navigator.onLine ? 'local_draft' : 'fila_offline',
        confirmationSource: 'fila_offline',
        lastAttemptAt: new Date().toISOString(),
        syncId: payload.syncId,
        localId: payload.localId,
        fingerprint: payload.fingerprint
      }, { persist: true, render: 'full' });
      this.recordLocationAudit('fila_localizacao', payload, {
        statusFila: navigator.onLine ? 'ERRO' : 'AGUARDANDO_INTERNET'
      });
      syncManager.refreshSyncView();
      return {
        ok: true,
        syncStatus: 'pendente_com_falha',
        syncMessage: pendingMessage,
        confirmed: false,
        queued: true,
        syncId: payload.syncId
      };
    }
    backupManager.addEntry('LOCALIZACAO', payload);
    this.setLocationSyncState(lote, {
      syncStatus: 'pendente_com_falha',
      syncMessage: pendingMessage,
      source: 'local_draft',
      lastAttemptAt: new Date().toISOString(),
      syncId: payload.syncId,
      localId: payload.localId,
      fingerprint: payload.fingerprint
    }, { persist: true, render: 'full' });
    syncManager.refreshSyncView();
    return {
      ok: true,
      syncStatus: 'pendente_com_falha',
      syncMessage: pendingMessage,
      confirmed: false,
      syncId: payload.syncId
    };
  },

  async capturarGPSEditar(loteBusca) {
    const loteTarget = String(loteBusca || this.selectedRegistrationLote || '').trim();
    if (!loteTarget || loteTarget.length < 3) return alert('Digite e selecione um lote na busca antes de registrar o GPS.');
    const lotesAchados = this.encontrarLotes(loteTarget);
    if (lotesAchados.length !== 1) return alert('Refine a busca para encontrar apenas 1 lote.');
    const lote = lotesAchados[0].lote;
    this.setRegistrationTarget(lote, { silent: true, populateSearch: true });

    if (!navigator.geolocation) return alert('GPS não suportado pelo navegador.');
    this.setLocationSyncState(lote, {
      syncStatus: 'capturando',
      syncMessage: 'Buscando a coordenada mais precisa disponível no dispositivo...'
    }, { persist: true, render: 'panel' });
    this.recordLocationAudit('captura_gps_iniciada', {
      lote
    });
    uiBuilder.toggleLoader(true, 'Buscando satélites e capturando coordenadas...');
    try {
      const pos = await this.obterLocalizacaoPrecisa(18000);
      const lat = pos.coords.latitude.toFixed(6);
      const lon = pos.coords.longitude.toFixed(6);
      const acc = Math.round(pos.coords.accuracy || 0);
      const descDefault = `Localizado pelo GPS (precisão ~${acc}m)`;
      const desc = prompt('GPS capturado. Adicione uma descrição complementar do local:', descDefault) || descDefault;
      uiBuilder.toggleLoader(true, 'Enviando localização e aguardando confirmação no mapa compartilhado...');
      const saveResult = await this.salvarLocalizacaoServer(lote, lat, lon, desc);
      if (saveResult?.blocked) return;
      if (saveResult?.syncStatus === 'sincronizado') {
        toastManager.show(saveResult.syncMessage || 'GPS salvo e confirmado no mapa compartilhado.', 'success', 4200);
      } else {
        toastManager.show(saveResult?.syncMessage || 'O lote ficou pendente até a confirmação no mapa compartilhado.', 'warning', 5600);
      }
      this.buscarNaLista();
      this.focusLote(lote, true);
    } catch (err) {
      this.setLocationSyncState(lote, {
        syncStatus: 'pendente_com_falha',
        syncMessage: 'Não foi possível capturar o GPS deste lote no dispositivo.',
        lastAttemptAt: new Date().toISOString()
      }, { persist: true, render: 'panel' });
      alert(permissionManager.getGpsErrorMessage(err));
    } finally {
      uiBuilder.toggleLoader(false);
    }
  },

  irParaCysyNoMapa() {
    this.centralizarEmpresa(true);
  },

  irParaMinhaLocalizacaoAtual() {
    this.marcarMinhaLocalizacao();
  },

  encontrarLotes(termoRaw, sourceList = null) {
    const termo = String(termoRaw || '').toUpperCase();
    const termoNumStr = String(Number(termo.replace(/\D/g, '')));
    const base = Array.isArray(sourceList) ? sourceList : this.getVisibleLotes();
    let matches = base.filter((item) => {
      const produto = String(item.produto || '').toUpperCase();
      const loteStr = String(item.lote || '').toUpperCase();
      const loteNumStr = String(Number(loteStr.replace(/\D/g, '')));
      if (termoNumStr.length >= 4 && (loteNumStr.endsWith(termoNumStr) || termoNumStr.endsWith(loteNumStr))) return true;
      return loteNumStr === termoNumStr || loteStr.includes(termo) || produto.includes(termo);
    });
    if (matches.length === 0) {
      matches = base.filter((item) =>
        String(item.lote || '').toUpperCase().includes(termo) ||
        String(item.produto || '').toUpperCase().includes(termo)
      );
    }
    return matches;
  },

  buscarNaLista() {
    const inpBusca = document.getElementById('inpBuscaMapa');
    const wrapper = document.getElementById('mapaResultWrapper');
    const container = document.getElementById('mapaResultContainer');
    if (!inpBusca || !wrapper || !container) return;
    const termoRaw = inpBusca.value.trim();
    if (termoRaw.length < 3) {
      wrapper.style.display = 'none';
      container.innerHTML = '';
      return;
    }

    const matches = this.encontrarLotes(termoRaw, this.getVisibleLotes());
    if (matches.length === 0) {
      const allMatches = this.encontrarLotes(termoRaw, this.getAllLotes());
      if (allMatches.length > 0 && allMatches.every((m) => Number(m.saldo || 0) <= 0)) {
        container.innerHTML = '<div class="map-empty-card">Lote sem saldo disponível.</div>';
      } else if (allMatches.length > 0) {
        container.innerHTML = '<div class="map-empty-card">O lote foi encontrado, mas ainda não possui localização válida registrada.</div>';
      } else {
        container.innerHTML = '<div class="map-empty-card">Nenhum lote encontrado no estoque atual.</div>';
      }
      wrapper.style.display = 'block';
      return;
    }

    container.innerHTML = matches.map((item) => {
      const loteToken = encodeURIComponent(item.lote);
      const mappedPoint = this.getRenderablePointByLote(item.lote, this.lastRenderedMapped);
      const logEntry = this.getLocationLogEntryByLote(item.lote);
      const syncMeta = this.getLocationSyncMeta(logEntry || { lote: item.lote, source: mappedPoint ? 'compartilhado' : 'local' });
      const isRegistrationTarget = this.getNormalizedLoteKey(this.selectedRegistrationLote) === this.getNormalizedLoteKey(item.lote);
      const locationBlock = mappedPoint
        ? `Coordenadas: ${mappedPoint.lat.toFixed(6)}, ${mappedPoint.lon.toFixed(6)}<br>Descrição: ${escapeHTML(mappedPoint.description)}`
        : 'Localização pendente no sistema';
      const actions = mappedPoint
        ? `<button type="button" class="${isRegistrationTarget ? 'btn-success' : ''}" onclick="mapController.setRegistrationTarget(decodeURIComponent('${loteToken}'), { refreshSearch: true })">${isRegistrationTarget ? '✅ Lote selecionado' : 'Selecionar lote'}</button>
           <button type="button" onclick="mapController.selectLote(decodeURIComponent('${loteToken}'), { center: true })">Abrir no mapa</button>`
        : `<button type="button" class="${isRegistrationTarget ? 'btn-success' : ''}" onclick="mapController.setRegistrationTarget(decodeURIComponent('${loteToken}'), { refreshSearch: true })">${isRegistrationTarget ? '✅ Lote selecionado' : 'Selecionar lote'}</button>
           <button type="button" onclick="mapController.capturarGPSEditar(decodeURIComponent('${loteToken}'))">Registrar agora</button>`;
      return `<article class="map-lote-card">
        <div class="map-lote-top">
          <div>
            <div class="map-lote-code">${escapeHTML(item.lote)}</div>
            <div class="map-lote-product">${escapeHTML(item.produto || 'Produto não informado')}</div>
          </div>
          <div class="map-lote-saldo">${formatTons(item.saldo)} t</div>
        </div>
        <div class="map-lote-location">${locationBlock}<br><span class="${syncMeta.className}">${escapeHTML(syncMeta.label)}</span></div>
        <div class="map-lote-actions">${actions}</div>
      </article>`;
    }).join('');
    wrapper.style.display = 'block';
    this.updateRegistrationUi();
  },

  listarSemLocalizacao() {
    const wrapper = document.getElementById('mapaResultWrapper');
    const container = document.getElementById('mapaResultContainer');
    if (!wrapper || !container) return;
    const missing = this.getVisibleLotes().filter((item) => !this.parseLocationEntry(this.locDB[item.lote]));
    if (missing.length === 0) {
      container.innerHTML = '<div class="map-empty-card">Estoque 100% mapeado. Nenhum lote ativo está sem localização.</div>';
    } else {
      container.innerHTML = missing.map((item) => {
        const loteToken = encodeURIComponent(item.lote);
        return `<article class="map-lote-card">
          <div class="map-lote-top">
            <div>
              <div class="map-lote-code">${escapeHTML(item.lote)}</div>
              <div class="map-lote-product">${escapeHTML(item.produto || 'Produto não informado')}</div>
            </div>
            <div class="map-lote-saldo">${formatTons(item.saldo)} t</div>
          </div>
          <div class="map-lote-location">Sem coordenadas válidas. Capture o GPS do lote para liberar a visualização dinâmica e a rota a pé.</div>
          <div class="map-lote-actions">
            <button type="button" onclick="mapController.setRegistrationTarget(decodeURIComponent('${loteToken}'), { refreshSearch: true })">Selecionar lote</button>
            <button type="button" onclick="mapController.capturarGPSEditar(decodeURIComponent('${loteToken}'))">Capturar GPS</button>
          </div>
        </article>`;
      }).join('');
    }
    wrapper.style.display = 'block';
  },

  extrairDescricao(locRaw) {
    return String(locRaw || '')
      .replace(/📍\s*Lat:\s*[-0-9.]+,\s*Lon:\s*[-0-9.]+\s*/i, '')
      .replace(/\s+/g, ' ')
      .trim();
  },

  getRenderableMappedPoints() {
    const pointsMap = new Map();
    this.purgeUnmappableLocationEntries();
    const allLotes = this.getAllLotes();
    const allLotesMap = new Map();

    allLotes.forEach((item) => {
      allLotesMap.set(this.getNormalizedLoteKey(item.lote), item);
    });

    this.getVisibleLotes().forEach((item) => {
      if (this.shouldRemoveLoteFromMap(item.lote)) return;
      const parsed = this.parseLocationEntry(this.locDB[item.lote]);
      if (!parsed) return;
      pointsMap.set(this.getNormalizedLoteKey(item.lote), {
        lote: String(item.lote || ''),
        produto: String(item.produto || ''),
        saldo: Number(item.saldo || 0),
        lat: parsed.lat,
        lon: parsed.lon,
        description: parsed.description,
        updatedAt: this.locationLog[item.lote]?.updatedAt || '',
        updatedBy: this.locationLog[item.lote]?.updatedBy || '',
        active: true,
        source: 'estoque_ativo'
      });
    });

    Object.values(this.locationLog || {}).forEach((row) => {
      const entry = this.normalizeLocationLogEntry(row);
      if (!entry || !Number.isFinite(entry.lat) || !Number.isFinite(entry.lon)) return;
      const key = this.getNormalizedLoteKey(entry.lote);
      const stockInfo = allLotesMap.get(key);
      const active = Boolean(stockInfo && Number(stockInfo.saldo || 0) > 0);
      if (!active) return;
      const current = pointsMap.get(key);
      const candidate = {
        lote: entry.lote,
        produto: stockInfo?.produto || entry.product || 'Lote registrado no log compartilhado',
        saldo: Number(stockInfo?.saldo ?? entry.saldo ?? 0),
        lat: Number(entry.lat),
        lon: Number(entry.lon),
        description: entry.description || 'Localização registrada no log compartilhado',
        updatedAt: entry.updatedAt || '',
        updatedBy: entry.updatedBy || '',
        active,
        source: active ? 'estoque_ativo' : 'log_compartilhado'
      };
      if (!current || (!current.active && candidate.active)) {
        pointsMap.set(key, candidate);
      }
    });

    return [...pointsMap.values()].sort((a, b) =>
      String(a.lote || '').localeCompare(String(b.lote || ''), 'pt-BR')
    );
  },

  coletarPontosMapeados() {
    return this.getRenderableMappedPoints().map((item) => ({
      lote: String(item.lote || ''),
      produto: String(item.produto || ''),
      saldo: Number(item.saldo || 0),
      lat: Number(item.lat),
      lon: Number(item.lon),
      descricao: String(item.description || ''),
      updatedAt: String(item.updatedAt || ''),
      updatedBy: String(item.updatedBy || ''),
      ativo: Boolean(item.active)
    }));
  },

  calculateDistanceMeters(lat1, lon1, lat2, lon2) {
    const toRad = (value) => (Number(value) * Math.PI) / 180;
    const earthRadius = 6371000;
    const dLat = toRad(Number(lat2) - Number(lat1));
    const dLon = toRad(Number(lon2) - Number(lon1));
    const a = Math.sin(dLat / 2) ** 2 +
      Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return earthRadius * c;
  },

  formatDistanceLabel(distanceMeters) {
    const meters = Number(distanceMeters || 0);
    if (!Number.isFinite(meters) || meters <= 0) return 'n/d';
    if (meters >= 1000) return `${(meters / 1000).toFixed(2)} km`;
    return `${Math.round(meters)} m`;
  },

  truncateMapText(value, maxLength = 28) {
    const text = String(value || '').trim();
    if (!text) return 'Sem descrição';
    if (text.length <= maxLength) return text;
    return `${text.slice(0, Math.max(0, maxLength - 1)).trimEnd()}…`;
  },

  calculateBearingDegrees(lat1, lon1, lat2, lon2) {
    const toRad = (value) => (Number(value) * Math.PI) / 180;
    const toDeg = (value) => (value * 180) / Math.PI;
    const startLat = toRad(lat1);
    const endLat = toRad(lat2);
    const dLon = toRad(Number(lon2) - Number(lon1));
    const y = Math.sin(dLon) * Math.cos(endLat);
    const x = Math.cos(startLat) * Math.sin(endLat) -
      Math.sin(startLat) * Math.cos(endLat) * Math.cos(dLon);
    return (toDeg(Math.atan2(y, x)) + 360) % 360;
  },

  bearingToDirection(bearingDegrees) {
    const bearing = ((Number(bearingDegrees) % 360) + 360) % 360;
    const directions = [
      'norte',
      'nordeste',
      'leste',
      'sudeste',
      'sul',
      'sudoeste',
      'oeste',
      'noroeste'
    ];
    return directions[Math.round(bearing / 45) % 8];
  },

  getDirectionShortLabel(direction) {
    const map = {
      norte: 'N',
      nordeste: 'NE',
      leste: 'L',
      sudeste: 'SE',
      sul: 'S',
      sudoeste: 'SO',
      oeste: 'O',
      noroeste: 'NO'
    };
    return map[String(direction || '').toLowerCase()] || 'N/D';
  },

  getConnectionLabel(edge) {
    if (!edge) return 'n/d';
    return `${this.formatDistanceLabel(edge.distance)} • ${this.getDirectionShortLabel(edge.direction)}`;
  },

  getConnectionVerboseLabel(edge, fromPoint, toPoint) {
    if (!edge || !fromPoint || !toPoint) return 'Conexão operacional';
    return `${fromPoint.lote} → ${toPoint.lote}: ${this.formatDistanceLabel(edge.distance)} para ${edge.direction || 'direção indefinida'}`;
  },

  clearConnectionLayers() {
    (this.connectionLayers || []).forEach((layer) => {
      try { this.map?.removeLayer(layer); } catch (_) {}
    });
    this.connectionLayers = [];
    this.lastRenderedConnectionsKey = '';
  },

  clearConnectionLayersOnly() {
    (this.connectionLayers || []).forEach((layer) => {
      try { this.map?.removeLayer(layer); } catch (_) {}
    });
    this.connectionLayers = [];
  },

  shouldShowConnectionLabels(pointsCount = 0) {
    const zoom = Number(this.map?.getZoom?.() || 0);
    // Rótulos de distância visíveis a partir de zoom 18 para qualquer quantidade,
    // zoom 16 para até 12 lotes, zoom 14 para até 5 lotes
    if (zoom >= 18) return true;
    if (zoom >= 16) return pointsCount <= 12;
    if (zoom >= 14) return pointsCount <= 5;
    return false;
  },

  shouldShowDetailedTooltips(pointsCount = 0) {
    const zoom = Number(this.map?.getZoom?.() || 0);
    if (pointsCount <= 3) return true;
    if (pointsCount <= 5) return zoom >= 18;
    return false;
  },

  getConnectionRenderKey(points = [], edges = []) {
    const zoom = Number(this.map?.getZoom?.() || 0);
    const pointsKey = points.map((item) => `${item.lote}:${Number(item.lat).toFixed(5)}:${Number(item.lon).toFixed(5)}:${Number(item.saldo || 0).toFixed(3)}`).join('|');
    const edgesKey = edges.map((edge) => `${edge.from}:${edge.to}:${Math.round(edge.distance)}:${this.getDirectionShortLabel(edge.direction)}`).join('|');
    return `${zoom}::${this.visualMode}::${this.getNormalizedLoteKey(this.selectedLote)}::${pointsKey}::${edgesKey}`;
  },

  getMidpointLatLng(fromPoint, toPoint, pxOffset = 0) {
    if (!this.map) return null;
    const fromLayer = this.map.latLngToLayerPoint([fromPoint.lat, fromPoint.lon]);
    const toLayer = this.map.latLngToLayerPoint([toPoint.lat, toPoint.lon]);
    const midLayer = L.point((fromLayer.x + toLayer.x) / 2, (fromLayer.y + toLayer.y) / 2);
    if (pxOffset) {
      const dx = toLayer.x - fromLayer.x;
      const dy = toLayer.y - fromLayer.y;
      const length = Math.max(1, Math.hypot(dx, dy));
      midLayer.x += (-dy / length) * pxOffset;
      midLayer.y += (dx / length) * pxOffset;
    }
    return this.map.layerPointToLatLng(midLayer);
  },

  renderConnectionWeb(points = []) {
    if (!this.map) return;
    const normalizedPoints = Array.isArray(points) ? points.filter((item) => Number.isFinite(item.lat) && Number.isFinite(item.lon)) : [];
    if (normalizedPoints.length < 2) {
      this.clearConnectionLayers();
      return;
    }
    const compactMobile = this.isCompactMobile();

    const proximity = this.getProximityWeb(normalizedPoints);
    const selectedKey = this.getNormalizedLoteKey(this.selectedLote);
    const zoom = Number(this.map.getZoom() || 0);
    if (compactMobile && !selectedKey) {
      this.clearConnectionLayers();
      return;
    }
    // Mostrar rótulos de distância: quando há seleção, ou via zoom/modo
    const showLabels = compactMobile
      ? Boolean(selectedKey) && normalizedPoints.length <= 8
      : (Boolean(selectedKey) || this.shouldShowConnectionLabels(normalizedPoints.length));
    // Setas direcionais: seleção com zoom alto ou modo detailed/informative com zoom alto
    const showArrows = compactMobile
      ? false
      : (Boolean(selectedKey)
      ? zoom >= 17
      : ((this.visualMode === 'detailed' || this.visualMode === 'informative') && normalizedPoints.length <= 6 && zoom >= 18));

    const edges = proximity.edges.filter((edge) => {
      if (!selectedKey) {
        if (compactMobile) return false;
        // Teia visível a partir de zoom 13 em qualquer modo
        if (this.visualMode === 'clean') return zoom >= 13;
        // Modo informativo e detalhado: sempre mostrar
        return true;
      }
      const fromPoint = normalizedPoints[edge.from];
      const toPoint = normalizedPoints[edge.to];
      if (!fromPoint || !toPoint) return false;
      return this.getNormalizedLoteKey(fromPoint.lote) === selectedKey || this.getNormalizedLoteKey(toPoint.lote) === selectedKey;
    });

    const renderKey = this.getConnectionRenderKey(normalizedPoints, edges);
    if (renderKey === this.lastRenderedConnectionsKey && (this.connectionLayers || []).length > 0) return;
    this.clearConnectionLayersOnly();
    this.lastRenderedConnectionsKey = renderKey;

    edges.forEach((edge, edgeIndex) => {
      const fromPoint = normalizedPoints[edge.from];
      const toPoint = normalizedPoints[edge.to];
      if (!fromPoint || !toPoint) return;
      const fromSelected = this.getNormalizedLoteKey(fromPoint.lote) === selectedKey;
      const toSelected = this.getNormalizedLoteKey(toPoint.lote) === selectedKey;
      const highlighted = Boolean(selectedKey) && (fromSelected || toSelected);

      // Opacidade e espessura adaptadas por modo visual para não prejudicar a visualização
      const baseOpacity = this.visualMode === 'detailed' ? 0.32
        : this.visualMode === 'informative' ? 0.22
        : 0.12; // clean: linhas bem sutis mas presentes
      const baseWeight = this.visualMode === 'detailed' ? 2.0
        : this.visualMode === 'informative' ? 1.7
        : 1.2; // clean: traço fino

      const line = L.polyline([
        [fromPoint.lat, fromPoint.lon],
        [toPoint.lat, toPoint.lon]
      ], {
        color: highlighted ? '#0EA5E9' : '#38BDF8',
        weight: highlighted ? 3.2 : (showLabels ? baseWeight + 0.4 : baseWeight),
        opacity: highlighted ? 0.68 : (showLabels ? Math.min(baseOpacity + 0.14, 0.55) : baseOpacity),
        dashArray: highlighted ? '10 6' : '5 10',
        interactive: true
      }).addTo(this.map);

      // Tooltip sempre mostra distância em metros ao passar o mouse
      line.bindTooltip(
        '<div style="font-size:12px;font-weight:800;padding:2px 4px;">' + this.getConnectionVerboseLabel(edge, fromPoint, toPoint) + '</div>',
        {
          sticky: true,
          direction: 'center',
          className: 'map-connection-tooltip'
        }
      );
      this.connectionLayers.push(line);

      if (showArrows) {
        const arrowLatLng = this.getMidpointLatLng(fromPoint, toPoint, highlighted ? 14 : 10);
        if (arrowLatLng) {
          const arrow = L.marker(arrowLatLng, {
            interactive: false,
            keyboard: false,
            icon: L.divIcon({
              className: 'map-connection-arrow-shell',
              html: `<div class="map-connection-arrow ${highlighted ? 'map-connection-arrow-active' : ''}" style="transform:rotate(${Number(edge.bearing || 0).toFixed(2)}deg);">➜</div>`,
              iconSize: [24, 24],
              iconAnchor: [12, 12]
            })
          }).addTo(this.map);
          this.connectionLayers.push(arrow);
        }
      }

      if (showLabels) {
        const labelLatLng = this.getMidpointLatLng(fromPoint, toPoint, 24 + (edgeIndex % 2) * 8);
        if (labelLatLng) {
          const label = L.marker(labelLatLng, {
            interactive: false,
            keyboard: false,
            icon: L.divIcon({
              className: 'map-connection-label-shell',
              html: `<div class="map-connection-label ${highlighted ? 'map-connection-label-active' : ''}">${escapeHTML(this.getConnectionLabel(edge))}</div>`,
              iconSize: [112, 28],
              iconAnchor: [56, 14]
            })
          }).addTo(this.map);
          this.connectionLayers.push(label);
        }
      }
    });
  },

  getTileTemplate(layerKind = 'street') {
    if (String(layerKind || '').toLowerCase() === 'satellite') {
      return 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}';
    }
    return 'https://tile.openstreetmap.org/{z}/{x}/{y}.png';
  },

  latLonToTile(lat, lon, zoom) {
    const latRad = (Number(lat) * Math.PI) / 180;
    const n = 2 ** Number(zoom);
    const x = Math.floor(((Number(lon) + 180) / 360) * n);
    const y = Math.floor((1 - Math.log(Math.tan(latRad) + (1 / Math.cos(latRad))) / Math.PI) / 2 * n);
    return { x, y };
  },

  buildOfflineTilePlan(bounds, zoomLevels = [], layerKinds = []) {
    const activeBounds = bounds || this.currentBounds || this.getBaseFocusBounds();
    if (!activeBounds) return [];
    const normalizedZooms = [...new Set((zoomLevels || []).filter((value) => Number.isFinite(value)).map((value) => Math.max(4, Math.min(19, Math.round(value)))))];
    const normalizedLayers = [...new Set((layerKinds || []).filter(Boolean))];
    const urls = [];

    normalizedLayers.forEach((layerKind) => {
      const template = this.getTileTemplate(layerKind);
      normalizedZooms.forEach((zoom) => {
        const northWest = activeBounds.getNorthWest();
        const southEast = activeBounds.getSouthEast();
        const minTile = this.latLonToTile(northWest.lat, northWest.lng, zoom);
        const maxTile = this.latLonToTile(southEast.lat, southEast.lng, zoom);
        const xStart = Math.min(minTile.x, maxTile.x);
        const xEnd = Math.max(minTile.x, maxTile.x);
        const yStart = Math.min(minTile.y, maxTile.y);
        const yEnd = Math.max(minTile.y, maxTile.y);
        for (let x = xStart; x <= xEnd; x++) {
          for (let y = yStart; y <= yEnd; y++) {
            urls.push(template.replace('{z}', String(zoom)).replace('{x}', String(x)).replace('{y}', String(y)));
          }
        }
      });
    });

    return [...new Set(urls)].slice(0, 220);
  },

  persistOfflineMapState(state = {}) {
    const payload = {
      center: state.center || (this.map ? this.map.getCenter() : { lat: config.base.lat, lng: config.base.lon }),
      zoom: Number(state.zoom ?? this.map?.getZoom?.() ?? 18),
      savedAt: state.savedAt || new Date().toISOString(),
      baseLayer: state.baseLayer || this.currentBaseLayerKind || 'street',
      totalLotes: Number(state.totalLotes ?? this.getRenderableMappedPoints().length),
      bounds: state.bounds || null,
      tileCount: Number(state.tileCount || 0)
    };
    safeStorage.setItem(this.offlineMapStateKey, JSON.stringify(payload));
  },

  getOfflineMapState() {
    return safeJsonParse(safeStorage.getItem(this.offlineMapStateKey, '{}'), {});
  },

  scheduleOfflineMapSave(force = false) {
    if (!navigator.onLine) {
      this.persistOfflineMapState();
      return;
    }
    if (this.offlineSaveTimer) clearTimeout(this.offlineSaveTimer);
    const delay = force ? 120 : 900;
    this.offlineSaveTimer = setTimeout(() => {
      this.saveCurrentMapOffline(force).catch(() => {});
    }, delay);
  },

  async saveCurrentMapOffline(force = false) {
    if (!navigator.onLine || !this.map || this.offlineSaveInFlight) return;
    this.offlineSaveInFlight = true;
    try {
      const center = this.map.getCenter();
      const zoom = Number(this.map.getZoom() || 18);
      const layerKinds = ['street'];
      if (this.currentBaseLayerKind === 'satellite') layerKinds.push('satellite');
      const zoomLevels = [zoom - 1, zoom, zoom + 1];
      const urls = this.buildOfflineTilePlan(this.map.getBounds(), zoomLevels, layerKinds);
      if (urls.length > 0 && 'caches' in window) {
        const tileCache = await caches.open(this.offlineTileCacheName);
        await Promise.allSettled(urls.map(async (url) => {
          const req = new Request(url, { mode: 'no-cors', credentials: 'omit' });
          const existing = await tileCache.match(req, { ignoreSearch: false });
          if (existing && !force) return;
          const response = await fetch(req).catch(() => null);
          if (response) await tileCache.put(req, response.clone()).catch(() => {});
        }));
      }
      this.persistOfflineMapState({
        center,
        zoom,
        baseLayer: this.currentBaseLayerKind,
        tileCount: urls.length,
        bounds: this.map.getBounds().toBBoxString()
      });
    } finally {
      this.offlineSaveInFlight = false;
    }
  },

  drawRoundedRect(ctx, x, y, width, height, radius = 10) {
    const safeRadius = Math.min(radius, width / 2, height / 2);
    ctx.beginPath();
    ctx.moveTo(x + safeRadius, y);
    ctx.lineTo(x + width - safeRadius, y);
    ctx.quadraticCurveTo(x + width, y, x + width, y + safeRadius);
    ctx.lineTo(x + width, y + height - safeRadius);
    ctx.quadraticCurveTo(x + width, y + height, x + width - safeRadius, y + height);
    ctx.lineTo(x + safeRadius, y + height);
    ctx.quadraticCurveTo(x, y + height, x, y + height - safeRadius);
    ctx.lineTo(x, y + safeRadius);
    ctx.quadraticCurveTo(x, y, x + safeRadius, y);
    ctx.closePath();
  },

  getProximityWeb(points) {
    const normalized = Array.isArray(points) ? points : [];
    if (normalized.length < 2) {
      return { edges: [], averageNearestDistance: 0, threshold: 0 };
    }

    const sortedLinksByPoint = normalized.map((point, pointIndex) => normalized
      .map((otherPoint, otherIndex) => {
        if (pointIndex === otherIndex) return null;
        const bearing = this.calculateBearingDegrees(point.lat, point.lon, otherPoint.lat, otherPoint.lon);
        return {
          from: pointIndex,
          to: otherIndex,
          distance: this.calculateDistanceMeters(point.lat, point.lon, otherPoint.lat, otherPoint.lon),
          bearing,
          direction: this.bearingToDirection(bearing)
        };
      })
      .filter(Boolean)
      .sort((a, b) => a.distance - b.distance));

    const nearestDistances = sortedLinksByPoint
      .map((links) => links[0]?.distance)
      .filter((value) => Number.isFinite(value) && value > 0);

    const averageNearestDistance = nearestDistances.length
      ? nearestDistances.reduce((sum, value) => sum + value, 0) / nearestDistances.length
      : 0;

    const threshold = averageNearestDistance > 0 ? averageNearestDistance * (normalized.length > 8 ? 1.22 : 1.42) : 120;
    const relaxedThreshold = normalized.length <= 4 ? Number.POSITIVE_INFINITY : Math.max(threshold, normalized.length > 10 ? 18 : 32);
    const maxLinksPerPoint = normalized.length <= 5 ? 2 : 1;
    const edgeMap = new Map();

    sortedLinksByPoint.forEach((links) => {
      links.slice(0, maxLinksPerPoint + 1).forEach((link) => {
        if (!Number.isFinite(link.distance) || link.distance > relaxedThreshold * 1.15) return;
        const key = [link.from, link.to].sort((a, b) => a - b).join(':');
        if (!edgeMap.has(key)) edgeMap.set(key, link);
      });
    });

    return {
      edges: [...edgeMap.values()],
      averageNearestDistance,
      threshold: relaxedThreshold
    };
  },

  dataUrlToBytes(dataUrl) {
    const base64 = String(dataUrl || '').split(',')[1] || '';
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index++) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  },

  getPdfPageSpec() {
    return {
      widthPt: 841.89,
      heightPt: 595.28,
      widthPx: 1684,
      heightPx: 1191,
      marginPx: 46,
      marginPt: 18
    };
  },

  createPdfPageCanvas() {
    const spec = this.getPdfPageSpec();
    const canvas = document.createElement('canvas');
    canvas.width = spec.widthPx;
    canvas.height = spec.heightPx;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, spec.widthPx, spec.heightPx);
    return { canvas, ctx, spec };
  },

  drawPdfPageHeader(ctx, spec, options = {}) {
    const title = String(options.title || 'Mapa Operacional de Lotes');
    const subtitle = String(options.subtitle || '');
    const pageLabel = String(options.pageLabel || '');
    const exportedAt = options.exportedAt instanceof Date ? options.exportedAt : new Date();

    const gradient = ctx.createLinearGradient(0, 0, spec.widthPx, 150);
    gradient.addColorStop(0, '#0F172A');
    gradient.addColorStop(1, '#1D4ED8');
    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, spec.widthPx, 146);

    ctx.fillStyle = '#FFFFFF';
    ctx.font = '800 34px Montserrat, Arial, sans-serif';
    ctx.fillText(title, spec.marginPx, 56);
    ctx.font = '600 17px Inter, Arial, sans-serif';
    ctx.fillStyle = 'rgba(255,255,255,0.92)';
    if (subtitle) ctx.fillText(subtitle, spec.marginPx, 92);
    ctx.fillText(`Gerado em ${exportedAt.toLocaleString('pt-BR')}`, spec.marginPx, 118);

    if (pageLabel) {
      const textW = ctx.measureText(pageLabel).width + 28;
      const x = spec.widthPx - spec.marginPx - textW;
      ctx.fillStyle = 'rgba(255,255,255,0.16)';
      this.drawRoundedRect(ctx, x, 30, textW, 36, 18);
      ctx.fill();
      ctx.fillStyle = '#FFFFFF';
      ctx.font = '800 16px Inter, Arial, sans-serif';
      ctx.fillText(pageLabel, x + 14, 53);
    }
  },

  buildPdfOverviewCanvas(snapshot) {
    const { canvas: pageCanvas, ctx, spec } = this.createPdfPageCanvas();
    const { foco, pontos, termo, now, proximity, canvas: snapshotCanvas } = snapshot;
    this.drawPdfPageHeader(ctx, spec, {
      title: 'Mapa Operacional de Lotes',
      subtitle: `${config.base.empresa} • ${config.base.enderecoCurto}`,
      pageLabel: 'Visão geral',
      exportedAt: now
    });

    const cardY = 172;
    const cardW = 250;
    const cardH = 88;
    const cardGap = 18;
    const cards = [
      { title: 'Pontos no foco', value: String(foco.length), tone: '#2563EB' },
      { title: 'Total mapeado', value: String(pontos.length), tone: '#0891B2' },
      { title: 'Média ao lote mais próximo', value: this.formatDistanceLabel(proximity?.averageNearestDistance || 0), tone: '#B45309' },
      { title: 'Filtro', value: termo || 'SEM FILTRO', tone: '#7C3AED' }
    ];

    cards.forEach((card, index) => {
      const x = spec.marginPx + index * (cardW + cardGap);
      ctx.fillStyle = '#FFFFFF';
      this.drawRoundedRect(ctx, x, cardY, cardW, cardH, 18);
      ctx.fill();
      ctx.strokeStyle = `${card.tone}44`;
      ctx.lineWidth = 2;
      ctx.stroke();
      ctx.fillStyle = card.tone;
      ctx.font = '700 15px Inter, Arial, sans-serif';
      ctx.fillText(card.title, x + 18, cardY + 30);
      ctx.fillStyle = '#0F172A';
      ctx.font = '900 26px Montserrat, Arial, sans-serif';
      ctx.fillText(this.truncateMapText(card.value, 24), x + 18, cardY + 64);
    });

    const frameX = spec.marginPx;
    const frameY = 290;
    const frameW = spec.widthPx - (spec.marginPx * 2);
    const frameH = spec.heightPx - frameY - spec.marginPx;
    ctx.fillStyle = '#FFFFFF';
    this.drawRoundedRect(ctx, frameX, frameY, frameW, frameH, 22);
    ctx.fill();
    ctx.strokeStyle = 'rgba(37, 99, 235, 0.24)';
    ctx.lineWidth = 2;
    ctx.stroke();

    const innerPad = 20;
    const usableW = frameW - (innerPad * 2);
    const usableH = frameH - (innerPad * 2);
    const scale = Math.min(usableW / snapshotCanvas.width, usableH / snapshotCanvas.height);
    const drawW = snapshotCanvas.width * scale;
    const drawH = snapshotCanvas.height * scale;
    const drawX = frameX + innerPad + ((usableW - drawW) / 2);
    const drawY = frameY + innerPad + ((usableH - drawH) / 2);
    ctx.drawImage(snapshotCanvas, drawX, drawY, drawW, drawH);

    ctx.fillStyle = '#475569';
    ctx.font = '600 12px Inter, Arial, sans-serif';
    ctx.fillText('Página 1 do PDF: visão consolidada do mapa, teia comparativa, base Cysy e legenda numerada.', frameX + 16, frameY + frameH - 10);
    return pageCanvas;
  },

  splitPdfRowsIntoPages(rows, rowsPerPage = 12) {
    const pages = [];
    const source = Array.isArray(rows) ? rows.slice() : [];
    for (let start = 0; start < source.length; start += rowsPerPage) {
      pages.push(source.slice(start, start + rowsPerPage));
    }
    if (pages.length > 1) {
      const lastIndex = pages.length - 1;
      const lastPage = pages[lastIndex];
      const previousPage = pages[lastIndex - 1];
      if (lastPage.length > 0 && lastPage.length < 4 && previousPage.length > 6) {
        const moveCount = Math.min(4 - lastPage.length, previousPage.length - 6);
        if (moveCount > 0) {
          const moved = previousPage.splice(previousPage.length - moveCount, moveCount);
          pages[lastIndex] = moved.concat(lastPage);
        }
      }
    }
    return pages.filter((page) => page.length > 0);
  },

  buildPdfBlobFromCanvases(canvases) {
    const pageCanvases = (Array.isArray(canvases) ? canvases : []).filter(Boolean);
    if (pageCanvases.length === 0) {
      throw new Error('Nenhuma página foi gerada para o PDF.');
    }

    const spec = this.getPdfPageSpec();
    const encoder = new TextEncoder();
    const chunks = [];
    let totalLength = 0;
    const objectOffsets = {};

    const pushString = (value) => {
      const bytes = encoder.encode(String(value));
      chunks.push(bytes);
      totalLength += bytes.length;
    };
    const pushBytes = (bytes) => {
      chunks.push(bytes);
      totalLength += bytes.length;
    };
    const startObject = (objectNumber) => {
      objectOffsets[objectNumber] = totalLength;
      pushString(`${objectNumber} 0 obj\n`);
    };
    const endObject = () => {
      pushString('\nendobj\n');
    };

    const pageImages = pageCanvases.map((canvas) => ({
      width: canvas.width,
      height: canvas.height,
      bytes: this.dataUrlToBytes(canvas.toDataURL('image/jpeg', 0.92))
    }));
    const totalObjects = 2 + (pageImages.length * 3);

    pushString('%PDF-1.4\n');
    pushBytes(Uint8Array.from([0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A]));

    startObject(1);
    pushString('<< /Type /Catalog /Pages 2 0 R >>');
    endObject();

    startObject(2);
    pushString(`<< /Type /Pages /Count ${pageImages.length} /Kids [${pageImages.map((_, index) => `${3 + (index * 3)} 0 R`).join(' ')}] >>`);
    endObject();

    pageImages.forEach((page, index) => {
      const pageObject = 3 + (index * 3);
      const contentObject = pageObject + 1;
      const imageObject = pageObject + 2;
      const availableWidth = spec.widthPt - (spec.marginPt * 2);
      const availableHeight = spec.heightPt - (spec.marginPt * 2);
      const scale = Math.min(availableWidth / page.width, availableHeight / page.height);
      const drawWidth = page.width * scale;
      const drawHeight = page.height * scale;
      const drawX = (spec.widthPt - drawWidth) / 2;
      const drawY = spec.heightPt - spec.marginPt - drawHeight;
      const contentStream = `q\n${drawWidth.toFixed(2)} 0 0 ${drawHeight.toFixed(2)} ${drawX.toFixed(2)} ${drawY.toFixed(2)} cm\n/Im${index} Do\nQ`;
      const contentBytes = encoder.encode(contentStream);

      startObject(pageObject);
      pushString(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${spec.widthPt} ${spec.heightPt}] /Rotate 0 /Resources << /XObject << /Im${index} ${imageObject} 0 R >> >> /Contents ${contentObject} 0 R >>`);
      endObject();

      startObject(contentObject);
      pushString(`<< /Length ${contentBytes.length} >>\nstream\n`);
      pushBytes(contentBytes);
      pushString('\nendstream');
      endObject();

      startObject(imageObject);
      pushString(`<< /Type /XObject /Subtype /Image /Width ${page.width} /Height ${page.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${page.bytes.length} >>\nstream\n`);
      pushBytes(page.bytes);
      pushString('\nendstream');
      endObject();
    });

    const xrefOffset = totalLength;
    pushString(`xref\n0 ${totalObjects + 1}\n`);
    pushString('0000000000 65535 f \n');
    for (let index = 1; index <= totalObjects; index++) {
      const offset = objectOffsets[index] || 0;
      pushString(`${String(offset).padStart(10, '0')} 00000 n \n`);
    }
    pushString(`trailer\n<< /Size ${totalObjects + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`);

    return new Blob(chunks, { type: 'application/pdf' });
  },

  downloadBlob(blob, filename) {
    const blobUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = blobUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
  },

  buildPdfDetailCanvases(snapshot) {
    const { foco, termo, now, proximity } = snapshot;
    const rows = foco.map((item, index) => ({
      marcador: index + 1,
      lote: String(item.lote || ''),
      produto: String(item.produto || 'Sem produto'),
      estoque: item.ativo ? `${formatTons(item.saldo)} t` : 'Log compartilhado',
      coordenadas: `${Number(item.lat).toFixed(6)}, ${Number(item.lon).toFixed(6)}`,
      descricao: String(item.descricao || ''),
      atualizacao: String(item.updatedAt || '')
    }));
    const pages = this.splitPdfRowsIntoPages(rows, 18);
    if (pages.length === 0) return [];
    const canvases = [];

    pages.forEach((pageRows, pageIndex) => {
      const { canvas, ctx, spec } = this.createPdfPageCanvas();
      this.drawPdfPageHeader(ctx, spec, {
        title: 'Detalhamento operacional dos lotes',
        subtitle: `${config.base.empresa} • Filtro: ${termo || 'SEM FILTRO'}`,
        pageLabel: `Detalhamento ${pageIndex + 1}/${pages.length}`,
        exportedAt: now
      });

      const summaryY = 160;
      ctx.fillStyle = '#FFFFFF';
      this.drawRoundedRect(ctx, spec.marginPx, summaryY, spec.widthPx - (spec.marginPx * 2), 74, 18);
      ctx.fill();
      ctx.strokeStyle = 'rgba(29, 78, 216, 0.26)';
      ctx.lineWidth = 2;
      ctx.stroke();
      ctx.fillStyle = '#0F172A';
      ctx.font = '700 16px Inter, Arial, sans-serif';
      ctx.fillText(`Base: ${config.base.empresa} • Coordenadas: ${config.base.lat}, ${config.base.lon}`, spec.marginPx + 20, summaryY + 32);
      ctx.font = '600 14px Inter, Arial, sans-serif';
      ctx.fillStyle = '#334155';
      ctx.fillText(`Teia comparativa: média ${this.formatDistanceLabel(proximity?.averageNearestDistance || 0)} ao lote mais próximo • conexões ativas: ${proximity?.edges?.length || 0}`, spec.marginPx + 20, summaryY + 54);
      ctx.fillText('Os números da tabela repetem exatamente os marcadores numerados da página principal do mapa.', spec.marginPx + 20, summaryY + 72);

      const columns = [
        { key: 'marcador', label: '#', width: 72 },
        { key: 'lote', label: 'Lote', width: 126 },
        { key: 'produto', label: 'Material / produto', width: 332 },
        { key: 'estoque', label: 'Quantidade', width: 140 },
        { key: 'coordenadas', label: 'Coordenadas', width: 290 },
        { key: 'descricao', label: 'Referência operacional', width: 430 },
        { key: 'atualizacao', label: 'Atualização', width: 202 }
      ];

      const tableX = spec.marginPx;
      const tableY = 248;
      const tableW = spec.widthPx - (spec.marginPx * 2);
      const rowHeight = 42;
      const headerHeight = 40;
      let cursorX = tableX;
      const mappedColumns = columns.map((column) => {
        const next = { ...column, x: cursorX };
        cursorX += column.width;
        return next;
      });

      ctx.fillStyle = '#0F172A';
      this.drawRoundedRect(ctx, tableX, tableY, tableW, headerHeight, 14);
      ctx.fill();
      ctx.font = '700 14px Inter, Arial, sans-serif';
      ctx.fillStyle = '#FFFFFF';
      mappedColumns.forEach((column) => {
        ctx.fillText(column.label, column.x + 10, tableY + 26);
      });

      pageRows.forEach((row, rowIndex) => {
        const y = tableY + headerHeight + (rowIndex * rowHeight);
        ctx.fillStyle = rowIndex % 2 === 0 ? '#FFFFFF' : '#F8FAFC';
        this.drawRoundedRect(ctx, tableX, y, tableW, rowHeight - 4, 12);
        ctx.fill();
        ctx.strokeStyle = 'rgba(148, 163, 184, 0.22)';
        ctx.lineWidth = 1;
        ctx.stroke();

        ctx.font = '700 14px Inter, Arial, sans-serif';
        ctx.fillStyle = '#0F172A';
        mappedColumns.forEach((column) => {
          const value = String(row[column.key] || '');
          const maxChars = Math.max(8, Math.floor((column.width - 18) / 8.8));
          ctx.save();
          ctx.beginPath();
          ctx.rect(column.x + 8, y + 6, column.width - 16, rowHeight - 12);
          ctx.clip();
          ctx.fillText(this.truncateMapText(value, maxChars), column.x + 10, y + 25);
          ctx.restore();
        });
      });

      ctx.strokeStyle = 'rgba(148,163,184,0.32)';
      ctx.lineWidth = 1;
      for (let idx = 0; idx <= mappedColumns.length; idx++) {
        const x = idx === 0 ? tableX : mappedColumns[idx - 1].x + mappedColumns[idx - 1].width;
        ctx.beginPath();
        ctx.moveTo(x, tableY);
        ctx.lineTo(x, tableY + headerHeight + (pageRows.length * rowHeight));
        ctx.stroke();
      }
      for (let rowIndex = 0; rowIndex <= pageRows.length; rowIndex++) {
        const y = tableY + headerHeight + (rowIndex * rowHeight);
        ctx.beginPath();
        ctx.moveTo(tableX, y);
        ctx.lineTo(tableX + tableW, y);
        ctx.stroke();
      }

      ctx.fillStyle = '#475569';
      ctx.font = '600 13px Inter, Arial, sans-serif';
      ctx.fillText('PDF operacional paginado em layout A4 estável para manter mapa, resumo e tabela legíveis sem rotação nem páginas vazias.', spec.marginPx, spec.heightPx - 24);

      canvases.push(canvas);
    });

    return canvases;
  },

  buildMapSnapshot() {
    const termo = document.getElementById('inpBuscaMapa')?.value?.trim()?.toUpperCase() || '';
    const pontos = this.coletarPontosMapeados();
    const selectedKey = this.getNormalizedLoteKey(this.selectedLote);

    let foco = pontos;
    if (termo.length >= 3 && pontos.length > 0) {
      foco = pontos.filter((ponto) =>
        ponto.lote.toUpperCase().includes(termo) ||
        ponto.produto.toUpperCase().includes(termo) ||
        ponto.descricao.toUpperCase().includes(termo)
      );
      if (foco.length === 0) foco = pontos;
    }
    if (selectedKey) {
      foco = [...foco].sort((a, b) => {
        const aSelected = this.getNormalizedLoteKey(a.lote) === selectedKey ? 1 : 0;
        const bSelected = this.getNormalizedLoteKey(b.lote) === selectedKey ? 1 : 0;
        return bSelected - aSelected || String(a.lote).localeCompare(String(b.lote), 'pt-BR');
      });
    }

    const width = 1800;
    const height = 1100;
    const topH = 250;
    const margin = 56;
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');

    const bg = ctx.createLinearGradient(0, 0, width, height);
    bg.addColorStop(0, '#FFFBEB');
    bg.addColorStop(0.5, '#EFF6FF');
    bg.addColorStop(1, '#FFFFFF');
    ctx.fillStyle = bg;
    ctx.fillRect(0, 0, width, height);

    ctx.fillStyle = '#0F172A';
    ctx.font = '800 42px Montserrat, Arial, sans-serif';
    ctx.fillText('Mapa Operacional de Lotes', margin, 70);
    ctx.font = '600 22px Inter, Arial, sans-serif';
    ctx.fillStyle = '#1D4ED8';
    ctx.fillText(`Terreno: ${config.base.empresa} • Pontos no foco: ${foco.length}`, margin, 110);
    ctx.fillText(`Base operacional: ${config.base.enderecoCurto}`, margin, 144);
    ctx.fillText(`Exportado em: ${new Date().toLocaleString('pt-BR')}`, margin, 178);
    if (termo.length >= 3) ctx.fillText(`Filtro aplicado: ${termo}`, margin, 212);

    const infoX = width - 580;
    const infoW = 520;
    ctx.fillStyle = 'rgba(255,255,255,0.96)';
    this.drawRoundedRect(ctx, infoX, 22, infoW, topH - 40, 22);
    ctx.fill();
    ctx.strokeStyle = 'rgba(59, 130, 246, 0.22)';
    ctx.lineWidth = 2;
    ctx.stroke();
    ctx.fillStyle = '#0F172A';
    ctx.font = '800 22px Montserrat, Arial, sans-serif';
    ctx.fillText('Leitura operacional', infoX + 18, 56);
    ctx.font = '600 15px Inter, Arial, sans-serif';
    ctx.fillStyle = '#334155';
    ctx.fillText(`Modo: PDF limpo • Teia: ${selectedKey ? 'foco selecionado' : 'conexões essenciais'}`, infoX + 18, 88);
    ctx.fillText(`Total mapeado: ${pontos.length} • Média do lote mais próximo: ${this.formatDistanceLabel(this.getProximityWeb(foco).averageNearestDistance || 0)}`, infoX + 18, 114);
    ctx.fillText('A primeira página prioriza a fotografia operacional do terreno.', infoX + 18, 146);
    ctx.fillText('O detalhamento completo segue na tabela paginada.', infoX + 18, 168);

    const mapX = margin;
    const mapY = topH;
    const mapW = width - (margin * 2);
    const mapH = height - topH - margin;
    ctx.fillStyle = '#FFFFFF';
    this.drawRoundedRect(ctx, mapX, mapY, mapW, mapH, 24);
    ctx.fill();
    ctx.strokeStyle = 'rgba(37, 99, 235, 0.34)';
    ctx.lineWidth = 2.5;
    ctx.stroke();

    const gridColor = 'rgba(148, 163, 184, 0.22)';
    ctx.strokeStyle = gridColor;
    ctx.lineWidth = 1;
    for (let i = 1; i < 10; i++) {
      const gx = mapX + (mapW / 10) * i;
      const gy = mapY + (mapH / 10) * i;
      ctx.beginPath(); ctx.moveTo(gx, mapY + 6); ctx.lineTo(gx, mapY + mapH - 6); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(mapX + 6, gy); ctx.lineTo(mapX + mapW - 6, gy); ctx.stroke();
    }

    const areaCoords = this.getCompanyAreaCoords();
    const boundsSource = [...foco, {
      lote: 'BASE',
      produto: config.base.empresa,
      saldo: 0,
      lat: config.base.lat,
      lon: config.base.lon,
      descricao: config.base.enderecoCurto
    }, ...areaCoords.map(([lat, lon], index) => ({
      lote: `AREA_${index}`,
      produto: 'Área da base',
      saldo: 0,
      lat,
      lon,
      descricao: 'Perímetro da Cysy'
    }))];
    let minLat = Math.min(...boundsSource.map((p) => p.lat));
    let maxLat = Math.max(...boundsSource.map((p) => p.lat));
    let minLon = Math.min(...boundsSource.map((p) => p.lon));
    let maxLon = Math.max(...boundsSource.map((p) => p.lon));
    if (minLat === maxLat) { minLat -= 0.0005; maxLat += 0.0005; }
    if (minLon === maxLon) { minLon -= 0.0005; maxLon += 0.0005; }
    const latPad = (maxLat - minLat) * 0.12;
    const lonPad = (maxLon - minLon) * 0.12;
    minLat -= latPad; maxLat += latPad;
    minLon -= lonPad; maxLon += lonPad;

    const project = (pointLat, pointLon) => ({
      x: mapX + ((pointLon - minLon) / (maxLon - minLon)) * mapW,
      y: mapY + (1 - ((pointLat - minLat) / (maxLat - minLat))) * mapH
    });

    const palette = this.getMarkerPalette();
    const rawLayoutPoints = foco.map((ponto, index) => ({
      ...ponto,
      actual: project(ponto.lat, ponto.lon),
      index,
      color: palette[index % palette.length]
    }));
    const layoutPoints = this.computeCollisionLayout(rawLayoutPoints, {
      project: (item) => item.actual,
      bounds: { minX: mapX, minY: mapY, maxX: mapX + mapW, maxY: mapY + mapH },
      padding: 30,
      radiusStep: 24,
      maxAttempts: 42,
      boxSizeGetter: () => ({ width: 42, height: 42 })
    }).map((item) => ({
      ...item,
      color: palette[item.index % palette.length]
    }));

    if (areaCoords.length >= 3) {
      ctx.fillStyle = 'rgba(250, 204, 21, 0.22)';
      ctx.strokeStyle = 'rgba(180, 83, 9, 0.95)';
      ctx.lineWidth = 4;
      ctx.beginPath();
      areaCoords.forEach(([lat, lon], index) => {
        const pt = project(lat, lon);
        if (index === 0) ctx.moveTo(pt.x, pt.y);
        else ctx.lineTo(pt.x, pt.y);
      });
      ctx.closePath();
      ctx.fill();
      ctx.stroke();
    }

    const proximity = this.getProximityWeb(layoutPoints);
    const edgesToRender = proximity.edges.filter((edge) => {
      if (selectedKey) {
        const from = layoutPoints[edge.from];
        const to = layoutPoints[edge.to];
        return this.getNormalizedLoteKey(from?.lote) === selectedKey || this.getNormalizedLoteKey(to?.lote) === selectedKey;
      }
      return layoutPoints.length <= 8;
    });

    ctx.save();
    ctx.setLineDash([8, 8]);
    edgesToRender.forEach((edge, edgeIndex) => {
      const from = layoutPoints[edge.from];
      const to = layoutPoints[edge.to];
      if (!from || !to) return;
      const isHighlighted = selectedKey && (this.getNormalizedLoteKey(from.lote) === selectedKey || this.getNormalizedLoteKey(to.lote) === selectedKey);
      ctx.strokeStyle = isHighlighted ? 'rgba(37, 99, 235, 0.62)' : 'rgba(14, 165, 233, 0.24)';
      ctx.lineWidth = isHighlighted ? 2.8 : 1.6;
      ctx.beginPath();
      ctx.moveTo(from.actual.x, from.actual.y);
      ctx.lineTo(to.actual.x, to.actual.y);
      ctx.stroke();

      const angle = Math.atan2(to.actual.y - from.actual.y, to.actual.x - from.actual.x);
      const arrowLength = isHighlighted ? 12 : 10;
      const arrowWidth = Math.PI / 7;
      ctx.fillStyle = isHighlighted ? 'rgba(37, 99, 235, 0.72)' : 'rgba(14, 165, 233, 0.34)';
      ctx.beginPath();
      ctx.moveTo(to.actual.x, to.actual.y);
      ctx.lineTo(
        to.actual.x - Math.cos(angle - arrowWidth) * arrowLength,
        to.actual.y - Math.sin(angle - arrowWidth) * arrowLength
      );
      ctx.lineTo(
        to.actual.x - Math.cos(angle + arrowWidth) * arrowLength,
        to.actual.y - Math.sin(angle + arrowWidth) * arrowLength
      );
      ctx.closePath();
      ctx.fill();

      if (layoutPoints.length <= 5 || isHighlighted) {
        const midX = (from.actual.x + to.actual.x) / 2;
        const midY = (from.actual.y + to.actual.y) / 2 - 8 - (edgeIndex % 2) * 10;
        const text = this.getConnectionLabel(edge);
        ctx.font = '700 11px Inter, Arial, sans-serif';
        const boxW = Math.max(76, ctx.measureText(text).width + 20);
        ctx.fillStyle = 'rgba(255,255,255,0.96)';
        this.drawRoundedRect(ctx, midX - (boxW / 2), midY - 14, boxW, 24, 10);
        ctx.fill();
        ctx.strokeStyle = isHighlighted ? 'rgba(37, 99, 235, 0.34)' : 'rgba(14, 165, 233, 0.22)';
        ctx.lineWidth = 1;
        ctx.stroke();
        ctx.fillStyle = '#0F172A';
        ctx.fillText(text, midX - (boxW / 2) + 10, midY + 4);
      }
    });
    ctx.restore();

    const basePoint = project(config.base.lat, config.base.lon);
    ctx.fillStyle = '#F59E0B';
    this.drawRoundedRect(ctx, basePoint.x - 14, basePoint.y - 14, 28, 28, 8);
    ctx.fill();
    ctx.strokeStyle = '#78350F';
    ctx.lineWidth = 2;
    ctx.stroke();
    ctx.fillStyle = '#0F172A';
    ctx.font = '800 15px Inter, Arial, sans-serif';
    ctx.fillText('BASE CYSY', basePoint.x + 20, basePoint.y + 6);

    layoutPoints.forEach((item) => {
      const isSelected = selectedKey && this.getNormalizedLoteKey(item.lote) === selectedKey;
      if (item.displaced) {
        ctx.strokeStyle = isSelected ? `${item.color}CC` : `${item.color}88`;
        ctx.lineWidth = isSelected ? 2.6 : 1.8;
        ctx.beginPath();
        ctx.moveTo(item.actual.x, item.actual.y);
        ctx.lineTo(item.badge.x, item.badge.y);
        ctx.stroke();
      }

      ctx.fillStyle = item.color;
      ctx.beginPath();
      ctx.arc(item.actual.x, item.actual.y, isSelected ? 9 : 7, 0, Math.PI * 2);
      ctx.fill();
      ctx.strokeStyle = '#FFFFFF';
      ctx.lineWidth = 2;
      ctx.stroke();

      if (isSelected) {
        ctx.strokeStyle = item.color;
        ctx.lineWidth = 3;
        ctx.beginPath();
        ctx.arc(item.actual.x, item.actual.y, 16, 0, Math.PI * 2);
        ctx.stroke();
      }

      ctx.fillStyle = item.color;
      ctx.beginPath();
      ctx.arc(item.badge.x, item.badge.y, isSelected ? 20 : 17, 0, Math.PI * 2);
      ctx.fill();
      ctx.strokeStyle = '#FFFFFF';
      ctx.lineWidth = 3;
      ctx.stroke();
      ctx.fillStyle = '#0F172A';
      ctx.font = isSelected ? '900 17px Inter, Arial, sans-serif' : '900 15px Inter, Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(`${item.index + 1}`, item.badge.x, item.badge.y + 0.5);
      ctx.textAlign = 'left';
      ctx.textBaseline = 'alphabetic';
    });

    const legendWidth = 462;
    const legendVisible = Math.min(layoutPoints.length, 8);
    const legendRowHeight = 46;
    const legendHeight = Math.min(430, 62 + (legendVisible * legendRowHeight) + (layoutPoints.length > legendVisible ? 22 : 0));
    const legendX = mapX + mapW - legendWidth - 18;
    const legendY = mapY + 18;
    ctx.fillStyle = 'rgba(255,255,255,0.96)';
    this.drawRoundedRect(ctx, legendX, legendY, legendWidth, legendHeight, 18);
    ctx.fill();
    ctx.strokeStyle = 'rgba(59, 130, 246, 0.24)';
    ctx.lineWidth = 2;
    ctx.stroke();
    ctx.fillStyle = '#0F172A';
    ctx.font = '800 18px Montserrat, Arial, sans-serif';
    ctx.fillText('Legenda operacional', legendX + 16, legendY + 30);
    ctx.font = '600 13px Inter, Arial, sans-serif';
    ctx.fillStyle = '#475569';
    ctx.fillText('Use o número no mapa para localizar rapidamente o lote no terreno.', legendX + 16, legendY + 52);
    layoutPoints.slice(0, legendVisible).forEach((item, idx) => {
      const rowY = legendY + 88 + (idx * legendRowHeight);
      ctx.fillStyle = item.color;
      ctx.beginPath();
      ctx.arc(legendX + 22, rowY - 6, 11, 0, Math.PI * 2);
      ctx.fill();
      ctx.strokeStyle = '#FFFFFF';
      ctx.lineWidth = 2;
      ctx.stroke();
      ctx.fillStyle = '#0F172A';
      ctx.font = '900 11px Inter, Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(`${item.index + 1}`, legendX + 22, rowY - 5.5);
      ctx.textAlign = 'left';
      ctx.textBaseline = 'alphabetic';
      ctx.font = '700 12px Inter, Arial, sans-serif';
      ctx.fillStyle = '#0F172A';
      ctx.fillText(`${item.lote} • ${item.ativo ? `${formatTons(item.saldo)} t` : 'LOG'}`, legendX + 44, rowY - 8);
      ctx.font = '600 11px Inter, Arial, sans-serif';
      ctx.fillStyle = '#475569';
      ctx.fillText(this.truncateMapText(item.produto || 'Sem material', 45), legendX + 44, rowY + 10);
    });
    if (layoutPoints.length > legendVisible) {
      ctx.fillStyle = '#475569';
      ctx.font = '600 12px Inter, Arial, sans-serif';
      ctx.fillText(`+ ${layoutPoints.length - legendVisible} lote(s) detalhado(s) na página seguinte.`, legendX + 16, legendY + legendHeight - 14);
    }

    ctx.fillStyle = '#475569';
    ctx.font = '500 14px Inter, Arial, sans-serif';
    ctx.fillText('A página 1 prioriza localização espacial: perímetro da base, lotes numerados, conexões mínimas e legenda lateral.', margin, height - 22);

    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
    const dataUrl = canvas.toDataURL('image/png');
    return { canvas, dataUrl, foco, pontos, termo, now, stamp, proximity };
  },

  async exportarMapaImagem() {
    let snapshot = null;
    try {
      snapshot = this.buildMapSnapshot();
    } catch (err) {
      alert(err.message || 'Nenhum ponto mapeado para exportar.');
      return;
    }
    const { dataUrl, foco, termo, now, stamp } = snapshot;
    const link = document.createElement('a');
    link.href = dataUrl;
    link.download = `mapa-lotes-${stamp}.png`;
    document.body.appendChild(link);
    link.click();
    link.remove();

    await backupManager.addEntry('MAPA_IMAGEM_EXPORTADA', {
      dataHora: now.toLocaleString('pt-BR'),
      filtro: termo || 'SEM_FILTRO',
      totalPontos: foco.length,
      pontos: foco,
      imagemMapa: dataUrl,
      referenciaAutomatica: true,
      baseEmpresa: {
        nome: config.base.empresa,
        lat: config.base.lat,
        lon: config.base.lon
      }
    });
    toastManager.show('Mapa exportado e salvo no log de backup.', 'success');
  },

  async imprimirMapaImagem() {
    let snapshot = null;
    try {
      snapshot = this.buildMapSnapshot();
    } catch (err) {
      alert(err.message || 'Nenhum ponto mapeado para gerar o PDF.');
      return;
    }

    const { foco, termo, now, dataUrl, stamp, proximity } = snapshot;
    const overviewCanvas = this.buildPdfOverviewCanvas(snapshot);
    const detailCanvases = this.buildPdfDetailCanvases(snapshot);
    const pdfBlob = this.buildPdfBlobFromCanvases([overviewCanvas, ...detailCanvases]);
    this.downloadBlob(pdfBlob, `mapa-operacional-lotes-${stamp}.pdf`);

    await backupManager.addEntry('MAPA_PDF_EXPORTADO', {
      dataHora: now.toLocaleString('pt-BR'),
      filtro: termo || 'SEM_FILTRO',
      totalPontos: foco.length,
      pontos: foco,
      imagemMapa: dataUrl,
      paginasPdf: detailCanvases.length + 1,
      mediaDistanciaTeia: this.formatDistanceLabel(proximity?.averageNearestDistance || 0),
      referenciaAutomatica: true,
      baseEmpresa: {
        nome: config.base.empresa,
        lat: config.base.lat,
        lon: config.base.lon
      }
    });
    toastManager.show('PDF do mapa salvo com imagem, teia comparativa e dados completos dos lotes.', 'success');
  }
};

const dashboardEngine = {
  limparNomesGenericos(str) {
    const genericos = ['LTDA', 'S/A', 'SA', 'ME', 'EPP', 'TRANSPORTES', 'TRANSPORTADORA', 'LOGISTICA', 'LOGÍSTICA', 'COMERCIO', 'COMERCIAL', 'AGROPECUARIA', 'AGRO', 'BRASIL', 'CIA', 'EIRELI'];
    return (str || '').split(/[\s,.\-]+/).filter(w => w && !genericos.includes(w)).join(' ').trim();
  },

  renderTopBanners(groups) {
    const bannerContainer = document.getElementById('bannerContainer');
    if (!bannerContainer) return;

    const bloqueados = groups.filter(g => g.bloqueado);
    const atrasados = groups.filter(g => g.statusInteligente?.code === 'ATRASADO');
    const comRnc = groups.filter(g => g.rncMatches?.length > 0);

    let html = '';

    if (bloqueados.length > 0) {
      html += `<div class="smart-banner banner-danger"><div class="banner-icon">⛔</div><div class="banner-content"><h4>Bloqueios críticos encontrados</h4><p>${bloqueados.length} veículo(s) com impedimento operacional ou financeiro.</p></div></div>`;
    }

    if (atrasados.length > 0 || comRnc.length > 0) {
      html += `<div class="smart-banner ${bloqueados.length > 0 ? 'banner-danger' : 'banner-success'}"><div class="banner-icon">📌</div><div class="banner-content"><h4>Resumo executivo do turno</h4><p>${atrasados.length} em atraso • ${comRnc.length} com alerta de qualidade</p></div></div>`;
    }

    bannerContainer.innerHTML = html;
  },

  renderExecTable(parsedData, pendingPlacas = []) {
    const container = document.getElementById('tabelaDiariaContainer');
    if (!container) return;

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const hojeStr = hoje.toLocaleDateString('pt-BR');
    
    debugEngine.log(`[DASHBOARD] Filtrando carregamentos para hoje (${hojeStr})...`, "info");
    debugEngine.log(`[DASHBOARD] Total de registros: ${(parsedData || []).length}`, "info");
    
    const dadosHoje = [];
    const dadosFora = [];
    
    (parsedData || []).forEach(item => {
      const parsedDate = parseSheetDate(item.dataRaw);
      const match = isSameDay(parsedDate, hoje);
      if (debugEngine.verboseMode) {
        debugEngine.log(`[DASHBOARD] Cliente: ${item.cliente} | dataRaw: "${item.dataRaw}" | parsed: ${parsedDate ? parsedDate.toLocaleDateString('pt-BR') : 'NULL'} | match: ${match}`, match ? "success" : "warn");
      }
      if (match) {
        dadosHoje.push(item);
      } else {
        dadosFora.push({
          cliente: item.cliente,
          dataRaw: item.dataRaw,
          parsed: parsedDate ? parsedDate.toLocaleDateString('pt-BR') : 'NULL'
        });
      }
    });

    debugEngine.log(`[DASHBOARD] Encontrados ${dadosHoje.length} carregamentos para hoje`, dadosHoje.length > 0 ? "success" : "warn");
    
    if (dadosFora.length > 0 && dadosHoje.length === 0) {
      debugEngine.log(`[DASHBOARD] MOSTRANDO ${dadosFora.length} registros como fallback (sem match de data)`, "warn");
      dadosHoje.push(...dadosFora);
    }

    if (!dadosHoje || dadosHoje.length === 0) {
      document.getElementById('bannerContainer').innerHTML = '';
      container.innerHTML = `
        <div style="text-align:center; padding:48px 24px; background:#FFFFFF; border:2px solid var(--warning); border-radius:var(--radius-lg); box-shadow:var(--shadow-soft);">
          <div style="font-size:54px; margin-bottom:12px;">📭</div>
          <div style="font-family:'Montserrat'; font-weight:900; font-size:18px; color:var(--primary); margin-bottom:8px;">Nenhum carregamento para hoje (${hojeStr})</div>
          <div style="font-size:14px; color:var(--text-muted); margin-bottom:16px;">Verifique o terminal de debug (🐞) para ver todos os registros.</div>
          <div style="display:flex; flex-direction:column; gap:12px; align-items:center;">
            <button onclick="dashboardEngine.mostrarTodosRegistros()" style="background:var(--secondary); color:white; padding:12px 24px; border-radius:8px; font-weight:800; font-size:13px;">📋 MOSTRAR TODOS OS REGISTROS</button>
          </div>
        </div>`;
      return;
    }

    const grupos = {};
    dadosHoje.forEach(item => {
      const key = item.placa || 'SEM PLACA';
      if (!grupos[key]) grupos[key] = [];
      grupos[key].push(item);
    });
    const pendingSet = new Set((pendingPlacas || []).map(p => normalizeName(p)));

    const arrayGrupos = Object.entries(grupos).map(([placa, itens]) => {
      const obsUnificado = [...new Set(itens.map(i => i.obsStr).filter(Boolean))].join(' | ');
      const timeStr = extrairHorarioObs(obsUnificado);
      let timeMins = 9999;
      if (timeStr) {
        const [h, m] = timeStr.split(':').map(Number);
        timeMins = (h * 60) + m;
      }

      const totalQtd = itens.reduce((sum, i) => sum + Number(i.qtd || 0), 0);
      const bloqueado = itens.some(i => i.bloqueio && i.bloqueio.trim() !== '' && i.bloqueio.toUpperCase() !== 'LIBERADO');

      let statusInteligente = truckStatusManager.definirStatusCaminhao(
        placa,
        timeStr,
        totalQtd,
        itens.some(i => String(i.pallet).toUpperCase() === 'SIM'),
        pendingSet.has(normalizeName(placa))
      );

      if (bloqueado) {
        statusInteligente = { code: 'BLOQUEADO', label: '⛔ BLOQUEADO', msg: 'Há bloqueio informado na programação.' };
      }

      let rncMatches = [];
      if (uiBuilder.rncCargas && Array.isArray(uiBuilder.rncCargas)) {
        const clientesStr = [...new Set(itens.map(i => i.cliente))].join(' • ');
        const bc = this.limparNomesGenericos(clientesStr.toUpperCase());
        const bo = this.limparNomesGenericos(obsUnificado.toUpperCase());

        uiBuilder.rncCargas.forEach(rnc => {
          const rC = this.limparNomesGenericos(rnc.cliente);
          const rT = this.limparNomesGenericos(rnc.transportador);
          const mC = rC.length >= 3 && bc.includes(rC);
          const mT = rT.length >= 3 && bo.includes(rT);
          const pL = placa.toUpperCase().replace(/[^A-Z0-9]/g, '');
          const mP = pL.length >= 7 && (
            rnc.transportador.replace(/[^A-Z0-9]/g, '').includes(pL) ||
            rnc.descricao.toUpperCase().replace(/[^A-Z0-9]/g, '').includes(pL)
          );

          let tr = [];
          if (mP) tr.push('Placa');
          if (mC && !mP) tr.push('Cliente');
          if (mT && !mP) tr.push('Transp.');

          if (tr.length > 0) {
            rncMatches.push({
              data: rnc.data,
              desc: rnc.descricao,
              prod: rnc.produto,
              acao: rnc.acaoImediata,
              trigger: tr.join(' e ')
            });
          }
        });
      }

      return { placa, itens, timeMins, timeStr, totalQtd, statusInteligente, bloqueado, rncMatches };
    }).sort((a, b) => {
      const prioridade = { BLOQUEADO: 0, ATRASADO: 1, CARREGANDO: 2, AGENDADO: 3, AGUARDANDO: 4, FATURAMENTO: 5 };
      return (prioridade[a.statusInteligente?.code] ?? 99) - (prioridade[b.statusInteligente?.code] ?? 99) || a.timeMins - b.timeMins;
    });

    this.renderTopBanners(arrayGrupos);

    let html = '<div class="carga-cards-grid">';
    arrayGrupos.forEach(({ placa, itens, timeStr, statusInteligente, bloqueado, rncMatches }) => {
      const allClientes = [...new Set(itens.map(i => i.cliente))].join(' • ');

      const badgesCarga = [];
      if (itens.some(i => String(i.estado || '').toUpperCase() === 'PY')) badgesCarga.push('<span class="badge-especial badge-py">🇵🇾 PY</span>');
      if (itens.some(i => String(i.pallet || '').toUpperCase() === 'SIM')) badgesCarga.push('<span class="badge-especial badge-palete">🟫 PALETE</span>');
      if (itens.some(i => String(i.filme || '').toUpperCase() === 'SIM')) badgesCarga.push('<span class="badge-especial badge-filme">🎞️ FILME</span>');

      let timeBadgeClass = 'time-badge';
      if (statusInteligente.code === 'AGENDADO') timeBadgeClass += ' alert';
      if (statusInteligente.code === 'ATRASADO' || statusInteligente.code === 'BLOQUEADO') timeBadgeClass += ' late';

      const rncHTML = rncMatches.length > 0 ? `
        <div style="background: rgba(220,38,38,0.12); border-left: 6px solid var(--danger); padding: 16px; border-radius: 12px; color: #111827;">
          <details>
            <summary style="cursor:pointer; font-weight:900; color:#FEE2E2; font-size:14px;">⚠️ ${rncMatches.length} alerta(s) de qualidade</summary>
            <div style="margin-top:16px; display:flex; flex-direction:column; gap:12px;">
              ${rncMatches.map(r => `
                <div style="background:rgba(2,6,23,0.72); border:1px solid #FECACA; border-radius:10px; padding:12px; font-size:12px; color:#F8FAFC; line-height:1.5;">
                  <strong>📅 ${escapeHTML(r.data)}</strong><br>
                  <strong>Alvo:</strong> ${escapeHTML(r.trigger)}<br>
                  <strong>Produto:</strong> ${escapeHTML(r.prod)}<br>
                  <strong>Ocorrência:</strong> ${escapeHTML(r.desc)}<br>
                  <strong>Ação imediata:</strong> ${escapeHTML(r.acao)}
                </div>`).join('')}
            </div>
          </details>
        </div>` : '';

      const produtosHTML = itens.map(item => {
        const badgesLinha = [];
        if (String(item.estado || '').toUpperCase() === 'PY') badgesLinha.push('<span class="badge-especial badge-py">🇵🇾 PY</span>');
        if (String(item.pallet || '').toUpperCase() === 'SIM') badgesLinha.push('<span class="badge-especial badge-palete">🟫 PALETE</span>');
        if (String(item.filme || '').toUpperCase() === 'SIM') badgesLinha.push('<span class="badge-especial badge-filme">🎞️ FILME</span>');

        return `
          <div class="produto-row">
            <div>
              <span class="produto-label">Produto</span>
              <div class="produto-nome">${escapeHTML(item.produtoDesc)}</div>
              <div style="margin-top:4px;">${badgesLinha.join(' ')}</div>
            </div>
            <div>
              <span class="produto-label">Quantidade</span>
              <span class="produto-qtd">${formatTons(item.qtd)} t</span>
            </div>
            <div>
              ${bloqueado && item.bloqueio && item.bloqueio.toUpperCase() !== 'LIBERADO'
                ? `<span class="badge-vermelho">⛔ ${escapeHTML(item.bloqueio)}</span>`
                : `<span class="badge-verde">✔ OK</span>`}
            </div>
          </div>`;
      }).join('');
      const evidenciaHTML = truckStatusManager.obterResumoEvidenciaHTML(placa);

      html += `
        <div class="carga-card ${bloqueado ? 'bloqueado' : 'liberado'}">
          <div class="carga-card-header">
            <div class="carga-placa">🚛 ${escapeHTML(placa)}</div>
            <div class="carga-cliente">
              <div>${escapeHTML(allClientes)}</div>
              <div class="carga-badges">${badgesCarga.join(' ')}</div>
            </div>
            <div class="carga-card-status">
              <span class="${timeBadgeClass}">${timeStr || 'Sem horário'}</span>
            </div>
          </div>

          <div class="carga-card-body">
            <div style="display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap;">
              <div style="font-weight:900; font-size:14px; color:${bloqueado ? '#FCA5A5' : statusInteligente.code === 'ATRASADO' ? '#FCA5A5' : statusInteligente.code === 'CARREGANDO' ? '#93C5FD' : statusInteligente.code === 'FATURAMENTO' ? '#6EE7B7' : '#FCD34D'};">
                ${statusInteligente.label}
              </div>
              ${statusInteligente.code !== 'FATURAMENTO' && statusInteligente.code !== 'BLOQUEADO' ? `<button class="btn-status-manual" onclick="truckStatusManager.iniciarFaturamento('${escapeHTML(placa)}')">📤 ENVIAR P/ FATURAMENTO</button>` : ''}
            </div>

            <div style="font-size:13px; font-weight:700; color:var(--text-muted);">${escapeHTML(statusInteligente.msg)}</div>
            ${evidenciaHTML}

            ${rncHTML}

            ${produtosHTML}
          </div>

          <div class="carga-card-footer">
            <span>📍 Itens: ${itens.length}</span>
          </div>
        </div>`;
    });

    html += '</div>';
    container.innerHTML = html;
  },
  
  mostrarTodosRegistros() {
    const container = document.getElementById('tabelaDiariaContainer');
    if (!container || !appController.lastParsedData) return;
    
    debugEngine.log("[DASHBOARD] Mostrando TODOS os registros (sem filtro de data)", "warn");
    
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const hojeStr = hoje.toLocaleDateString('pt-BR');
    
    let html = `<div style="margin-bottom:16px; padding:12px; background:#FEF3C7; border-radius:8px; border:1px solid #F59E0B;">
      <div style="font-weight:800; color:#92400E; margin-bottom:8px;">⚠️ Modo Debug: Mostrando todos os registros (sem filtro de data)</div>
      <div style="font-size:12px; color:#78350F;">Hoje: ${hojeStr} | Total: ${appController.lastParsedData.length} registros</div>
      <button onclick="appController.handleRefresh(true)" style="margin-top:12px; background:var(--secondary); color:white; padding:8px 16px; border-radius:6px; font-weight:800; font-size:12px;">🔄 ATUALIZAR</button>
    </div>`;
    
    const grupos = {};
    appController.lastParsedData.forEach(item => {
      const key = item.placa || 'SEM PLACA';
      if (!grupos[key]) grupos[key] = [];
      grupos[key].push(item);
    });
    
    html += '<div class="carga-cards-grid">';
    
    Object.entries(grupos).forEach(([placa, itens]) => {
      const allClientes = [...new Set(itens.map(i => i.cliente))].join(' • ');
      const parsedDate = parseSheetDate(itens[0].dataRaw);
      const dateStr = parsedDate ? parsedDate.toLocaleDateString('pt-BR') : 'DATA INVÁLIDA';
      
      html += `
        <div class="carga-card liberado">
          <div class="carga-card-header">
            <div class="carga-placa">🚛 ${escapeHTML(placa)}</div>
            <div class="carga-cliente">
              <div>${escapeHTML(allClientes)}</div>
            </div>
            <div class="carga-card-status">
              <span class="time-badge">📅 ${dateStr}</span>
            </div>
          </div>
          <div class="carga-card-body">
            <div style="font-size:13px; color:var(--text-muted);">
              Total: ${itens.reduce((s,i) => s + Number(i.qtd||0), 0).toFixed(3)}t | ${itens.length} pedido(s)
            </div>
          </div>
          <div class="carga-card-footer">
        <span>📍 dataRaw: "${escapeHTML(itens[0].dataRaw)}"</span>
          </div>
        </div>`;
    });
    
    html += '</div>';
    container.innerHTML = html;
  }
};

