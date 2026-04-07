const appController = {
  isSubmitting: false,
  lastParsedData: [],
  currentPendingPlacas: [],
  cacheKey: 'cysyAppCacheV1',

  hydrateCachedState() {
    const cache = safeJsonParse(safeStorage.getItem(this.cacheKey, '{}'), {});
    if (!cache || typeof cache !== 'object') return false;
    const parsedData = Array.isArray(cache.parsedData) ? cache.parsedData : [];
    const lotesMap = cache.lotesMap && typeof cache.lotesMap === 'object' ? cache.lotesMap : {};
    const lotesFlat = Array.isArray(cache.lotesFlat) ? cache.lotesFlat : [];
    const allLotesFlat = Array.isArray(cache.allLotesFlat) ? cache.allLotesFlat : lotesFlat;
    this.lastParsedData = parsedData;
    uiBuilder.localDataStore = parsedData;
    uiBuilder.historicoCargas = Array.isArray(cache.historicoCargas) ? cache.historicoCargas : [];
    uiBuilder.rncCargas = Array.isArray(cache.rncCargas) ? cache.rncCargas : [];
    uiBuilder.lotesMap = lotesMap;
    uiBuilder.lotesFlat = lotesFlat;
    uiBuilder.allLotesFlat = allLotesFlat;
    if (parsedData.length > 0) uiBuilder.populateSelects(parsedData);
    if (Object.keys(lotesMap).length > 0) {
      perdasController.populateProdutos(lotesMap);
      perdasController.populateGlobais();
    }
    if (parsedData.length > 0) {
      dashboardEngine.renderExecTable(this.lastParsedData, this.currentPendingPlacas);
      uiBuilder.verificarPlacaJaCarregou();
    }
    const updatedAt = cache.updatedAt ? new Date(cache.updatedAt) : new Date();
    uiBuilder.updateGlobalUpdateTimestamp(updatedAt);
    return parsedData.length > 0 || lotesFlat.length > 0;
  },

  persistCachedState(updatedAt = new Date()) {
    safeStorage.setItem(this.cacheKey, JSON.stringify({
      updatedAt: updatedAt.toISOString(),
      parsedData: this.lastParsedData || [],
      historicoCargas: uiBuilder.historicoCargas || [],
      rncCargas: uiBuilder.rncCargas || [],
      lotesMap: uiBuilder.lotesMap || {},
      lotesFlat: uiBuilder.lotesFlat || [],
      allLotesFlat: uiBuilder.allLotesFlat || []
    }));
  },

  async initApp() {
    const compactMobile = typeof window !== 'undefined' &&
      window.matchMedia &&
      (window.matchMedia('(max-width: 768px)').matches || window.matchMedia('(pointer: coarse)').matches);
    if (compactMobile) document.body.classList.add('mobile-optimized');
    soundManager.init();
    loginManager.init();
    envManager.init();
    wakeLockManager.init();
    installManager.init();
    uiBuilder.init();
    perdasController.init();
    truckStatusManager.init();
    priorityAlertManager.init();
    await swManager.init();
    await backupManager.init().catch(() => {});
    await versionManager.ensureVersionRuntime().catch(() => {});
    cacheJanitor.schedule();
    await permissionManager.requestAll(false).catch(() => {});
    await syncManager.init();
    const pending = await dbManager.getAllPending().catch(() => []);
    this.currentPendingPlacas = (pending || []).filter(p => p.type === 'OPERACAO').map(p => p.payload.placa);
    this.hydrateCachedState();
    alertManager.init();
    await this.handleRefresh();
    setTimeout(() => { priorityAlertManager.remindActiveAlerts({ force: true }); }, 2200);
    setInterval(() => { this.silentRefresh(); }, 60000);
  },

  async silentRefresh() {
    if (!navigator.onLine) return;
    try {
      const [rawValues, historicoRaw, rncRaw] = await Promise.all([
        apiService.fetchSheetData().catch(() => null),
        apiService.fetchHistoricoData().catch(() => null),
        apiService.fetchRncData().catch(() => null)
      ]);

      if (historicoRaw) uiBuilder.historicoCargas = historicoRaw;
      if (rncRaw) uiBuilder.rncCargas = dataParser.parseRnc(rncRaw);
      if (rawValues && rawValues.length > 0) {
        this.lastParsedData = dataParser.parse(rawValues);
        uiBuilder.localDataStore = this.lastParsedData;
      }

      const rawEstoque = await apiService.fetchEstoqueData().catch(() => null);
      if (rawEstoque) {
        const parsedE = dataParser.parseEstoque(rawEstoque);
        if (Object.keys(parsedE.map).length > 0) {
          uiBuilder.lotesMap = parsedE.map;
          uiBuilder.lotesFlat = parsedE.flat;
          uiBuilder.allLotesFlat = parsedE.allFlat || parsedE.flat;
          perdasController.populateProdutos(parsedE.map);
          perdasController.populateGlobais();
        }
      }

      const pending = await dbManager.getAllPending();
      this.currentPendingPlacas = (pending || []).filter(p => p.type === 'OPERACAO').map(p => p.payload.placa);

      const dash = document.getElementById('tab-dash');
      if (dash && dash.classList.contains('active')) dashboardEngine.renderExecTable(this.lastParsedData, this.currentPendingPlacas);
      uiBuilder.verificarPlacaJaCarregou();
      uiBuilder.updateGlobalUpdateTimestamp(new Date());
      this.persistCachedState(new Date());
    } catch (error) {}
  },

  async handleRefresh(isManual = false) {
    if (!navigator.onLine) {
      if (isManual) alert("Sem conexão. O app permanece operando em modo offline.");
      return;
    }

    uiBuilder.toggleLoader(true, "Conectando ao Centro de Comando...");
    try {
      const fe = document.getElementById('fatalError');
      if (fe) fe.style.display = 'none';

      const [rawValues, historicoRaw, rncRaw] = await Promise.all([
        apiService.fetchSheetData(),
        apiService.fetchHistoricoData(),
        apiService.fetchRncData()
      ]);

      this.lastParsedData = dataParser.parse(rawValues || []);
      uiBuilder.historicoCargas = historicoRaw || [];
      uiBuilder.rncCargas = dataParser.parseRnc(rncRaw);
      uiBuilder.localDataStore = this.lastParsedData;
      let rawEstoque = null;
      try { rawEstoque = await apiService.fetchEstoqueData(); } catch(e) {}

      if (rawEstoque) {
        const parsedE = dataParser.parseEstoque(rawEstoque);
        if (Object.keys(parsedE.map).length > 0) {
          uiBuilder.lotesMap = parsedE.map;
          uiBuilder.lotesFlat = parsedE.flat;
          uiBuilder.allLotesFlat = parsedE.allFlat || parsedE.flat;
          perdasController.populateProdutos(parsedE.map);
          perdasController.populateGlobais();
        }
      }

      const pending = await dbManager.getAllPending();
      this.currentPendingPlacas = (pending || []).filter(p => p.type === 'OPERACAO').map(p => p.payload.placa);

      uiBuilder.populateSelects(this.lastParsedData);
      dashboardEngine.renderExecTable(this.lastParsedData, this.currentPendingPlacas);
      uiBuilder.updateTimestamp();
      uiBuilder.updateGlobalUpdateTimestamp(new Date());
      syncManager.refreshSyncView();
      this.persistCachedState(new Date());

      debugEngine.log(isManual ? "Atualização manual concluída." : "Atualização automática concluída.", "success");
    } catch (error) {
      const fe = document.getElementById('fatalError');
      if (fe) fe.style.display = 'block';
      debugEngine.log(`Falha no Refresh: ${error.message}`, "error");
    } finally {
      uiBuilder.toggleLoader(false);
    }
  },

  async limpezaCompleta() {
    const confirma = confirm("Esta ação fará uma limpeza forçada: cache, dados locais, fila offline, logs, históricos internos e versões antigas serão removidos deste navegador. Deseja continuar?");
    if (!confirma) return;
    uiBuilder.toggleLoader(true, "Executando limpeza forçada...");
    try {
      try {
        if ('serviceWorker' in navigator) {
          const regs = typeof navigator.serviceWorker.getRegistrations === 'function'
            ? await navigator.serviceWorker.getRegistrations()
            : [];
          await Promise.all(regs.map((registration) => registration.unregister()));
        }
      } catch (_) {}

      try {
        if (typeof cacheJanitor?.clearManagedCaches === 'function') await cacheJanitor.clearManagedCaches(true);
      } catch (_) {}
      try {
        if ('caches' in window) {
          const cacheKeys = await caches.keys();
          await Promise.allSettled(cacheKeys.map((key) => caches.delete(key)));
        }
      } catch (_) {}

      try { await dbManager.clearAll(); } catch (_) {}
      try { await backupManager.clearAll(); } catch (_) {}
      try { if (dbManager.db) dbManager.db.close(); } catch (_) {}
      try { if (backupManager.db) backupManager.db.close(); } catch (_) {}

      const deleteDb = (name) => new Promise((resolve) => {
        try {
          if (!name) return resolve();
          const req = indexedDB.deleteDatabase(name);
          req.onsuccess = () => resolve();
          req.onerror = () => resolve();
          req.onblocked = () => resolve();
        } catch (_) { resolve(); }
      });

      const candidateDbNames = new Set([
        dbManager.dbName,
        backupManager.dbName,
        'CysyDBv5',
        'CysyDBv6',
        'CysyDBv7',
        'CysyDBv8',
        'CysyDBv9',
        'CysyBackupDBv0',
        'CysyBackupDBv1',
        'CysyBackupDBv2'
      ].filter(Boolean));
      try {
        if (typeof indexedDB?.databases === 'function') {
          const knownDatabases = await indexedDB.databases();
          (knownDatabases || []).forEach((dbInfo) => {
            if (dbInfo?.name && /cysy/i.test(dbInfo.name)) candidateDbNames.add(dbInfo.name);
          });
        }
      } catch (_) {}

      await Promise.all([...candidateDbNames].map((name) => deleteDb(name)));
      try {
        if (typeof cacheJanitor?.clearManagedLocalStorage === 'function') cacheJanitor.clearManagedLocalStorage(true);
      } catch (_) {}
      try { localStorage.clear(); } catch (_) {}
      try { sessionStorage.clear(); } catch (_) {}
      try {
        debugEngine.clearLogs();
      } catch (_) {}
      try {
        document.cookie.split(';').forEach((cookie) => {
          const eqPos = cookie.indexOf('=');
          const name = (eqPos > -1 ? cookie.slice(0, eqPos) : cookie).trim();
          if (!name) return;
          document.cookie = `${name}=;expires=Thu, 01 Jan 1970 00:00:00 GMT;path=/`;
        });
      } catch (_) {}

      alert("Limpeza forçada concluída. O sistema será recarregado limpo.");
      location.replace(`${location.pathname}?clean=${Date.now()}`);
    } finally {
      uiBuilder.toggleLoader(false);
    }
  },

  async submitFinal(statusLiberacao) {
    if (this.isSubmitting) return;
    const sa = document.getElementById('selAcomp');
    const acomp = sa ? sa.value : '';
    const itemsCarga = uiBuilder.getSelectedCargaItems();
    if (!acomp || !Array.isArray(itemsCarga) || itemsCarga.length === 0) {
      return alert("⚠️ Preencha Responsável e selecione a carga por placa ou cliente.");
    }
    const cli = [...new Set(itemsCarga.map(i => i.cliente))].join(' | ');
    const ped = [...new Set(itemsCarga.map(i => i.pedido))].join(' | ');

    let missingCheck = false;
    config.checklist.forEach(c => {
      const selectedValue = typeof uiBuilder.getChecklistValue === 'function'
        ? uiBuilder.getChecklistValue(c.id)
        : (document.getElementById(`card_${c.id}`)?.dataset.value || '');
      if (c.type !== 'text' && !selectedValue) missingCheck = true;
    });
    if (missingCheck) return alert("⚠️ Responda todos os itens do Checklist.");

    const base64Frente = document.getElementById('base64Frente')?.value || '';
    const base64Traseira = document.getElementById('base64Traseira')?.value || '';
    const base64Assoalho = document.getElementById('base64Assoalho')?.value || '';

    const p = document.getElementById('inpPlaca');
    const m = document.getElementById('inpMotorista');
    const o = document.getElementById('inpObs');
    const rawPlate = String(p ? p.value : '').trim();
    const hasPlateText = rawPlate.length > 0;
    const isEmptyPlateSentinel = typeof plateRules !== 'undefined' && plateRules.isEmpty(rawPlate);
    const isValidPlate = typeof plateRules === 'undefined' ? Boolean(rawPlate) : plateRules.isValid(rawPlate);
    if (hasPlateText && !isEmptyPlateSentinel && !isValidPlate) {
      return alert("⚠️ Informe uma placa válida. Formatos aceitos: ABC1234, ABC1D23, ABCD123 e CAL042.");
    }
    const normalizedPlate = !hasPlateText
      ? ''
      : (isEmptyPlateSentinel ? 'SEM PLACA' : (typeof plateRules !== 'undefined' ? plateRules.normalize(rawPlate) : rawPlate.toUpperCase()));

    const payload = {
      tipo: "OPERACAO",
      status: statusLiberacao,
      dataHora: new Date().toLocaleString('pt-BR'),
      acompanhamento: acomp,
      cliente: cli,
      pedido: ped,
      produto: '',
      quantidadeTotal: 0,
      placa: normalizedPlate,
      motorista: m ? m.value : '',
      obs: o ? o.value : '',
      linhasOperacao: [],
      checklist: {},
      fotos: { frente: base64Frente, traseira: base64Traseira, assoalho: base64Assoalho }
    };

    let hasMaterial = false;
    for (let i = 1; i <= 10; i++) {
      const p_val = document.getElementById(`selProd_${i}`)?.value || '';
      const q = parseBRNumber(document.getElementById(`inpQtd_${i}`)?.value || '');
      const l = document.getElementById(`selLote_${i}`)?.value || 'S/ Lote';
      const qPerda = parseBRNumber(document.getElementById(`inpPerda_${i}`)?.value || '');
      const mPerda = document.getElementById(`selMotivo_${i}`)?.value || '';

      if (q < 0 || qPerda < 0) return alert("⛔ Valores negativos.");
      if (p_val && (q > 0 || qPerda > 0)) {
        hasMaterial = true;
        if (!payload.produto) payload.produto = p_val;
        if (qPerda > 0 && !mPerda) return alert("⛔ Informe motivo da perda.");
        payload.quantidadeTotal += q;
        payload.linhasOperacao.push({ produto: p_val, quantidade: q, lote: l, perda: qPerda, motivoPerda: mPerda || "-" });
      }
    }

    if (!hasMaterial) return alert("⚠️ Informe a quantidade de pelo menos um produto.");

    config.checklist.forEach(c => {
      const textNode = document.getElementById(`chk_text_${c.id}`);
      const val = typeof uiBuilder.getChecklistValue === 'function'
        ? uiBuilder.getChecklistValue(c.id)
        : (document.getElementById(`card_${c.id}`)?.dataset.value || '');
      payload.checklist[c.lbl] = c.type === 'text'
        ? (textNode ? textNode.value || 'Sem observações adicionais' : 'Sem observações adicionais')
        : (val || '');
    });

    const meta = dbManager.buildMeta('OPERACAO', payload);
    payload.syncId = meta.syncId;
    payload.localId = meta.localId;
    payload.fingerprint = meta.fingerprint;
    payload.imageCount = meta.imageCount;

    if (dbManager.isSent(meta.syncId, meta.fingerprint) || dbManager.isFingerprintSent(meta.fingerprint)) {
      alert("Este checklist já foi enviado anteriormente e não pode ser reenviado.");
      return;
    }
    const dupOpen = await dbManager.findOpenDuplicate(meta.syncId, meta.fingerprint);
    if (dupOpen) {
      toastManager.show('Este checklist já está na fila de sincronização.', 'warning');
      uiBuilder.switchTab(null, 'sync');
      return;
    }

    this.isSubmitting = true;
    uiBuilder.toggleLoader(true, "Transmitindo Operação...");
    try {
      const resp = await apiService.sendDataToAppScript(payload);
      if (resp && resp.success) {
        dbManager.markSent(meta.syncId, meta.fingerprint, 'OPERACAO', payload);
        const open = await dbManager.getAllRecords({ includeSent: false });
        await Promise.all((open || [])
          .filter((row) => row.fingerprint === meta.fingerprint)
          .map((row) => dbManager.markAsSent(row.id, { deduplicado: true, origem: 'online-direto' })));
        backupManager.addEntry('OPERACAO', { statusEnvio: 'ONLINE', payload, syncId: meta.syncId, fingerprint: meta.fingerprint });
        alert("✅ SUCESSO!");
        uiBuilder.clearDependentFields();
        if (sp) sp.value = '';
        uiBuilder.updateGlobalUpdateTimestamp();
        this.silentRefresh();
      } else {
        alert("❌ ERRO DO SCRIPT:\n\n" + (resp.message || "Erro desconhecido."));
      }
    } catch(err) {
      const queued = await dbManager.savePending('OPERACAO', payload, {
        syncId: meta.syncId,
        localId: meta.localId,
        fingerprint: meta.fingerprint,
        imageCount: meta.imageCount,
        status: navigator.onLine ? 'ERRO' : 'AGUARDANDO_INTERNET'
      });
      if (!queued.ok) {
        if (queued.reason === 'OFFLINE_LIMIT') {
          toastManager.show(`Limite offline atingido: máximo de ${queued.limit} carregamentos pendentes com imagem.`, 'error', 5500);
          alert(`Limite offline atingido (${queued.limit}). Sincronize os pendentes para continuar.`);
          uiBuilder.switchTab(null, 'sync');
        } else if (queued.reason === 'JA_ENVIADO') {
          alert("Este checklist já foi enviado anteriormente e não pode ser reenviado.");
        } else if (queued.reason === 'DUPLICADO_LOCAL') {
          toastManager.show('Checklist já existe na fila de sincronização.', 'warning');
          uiBuilder.switchTab(null, 'sync');
        } else {
          alert("Falha ao salvar na fila offline.");
        }
      } else {
        backupManager.addEntry('OPERACAO', { statusEnvio: 'PENDENTE_SYNC', payload, syncId: meta.syncId, fingerprint: meta.fingerprint });
        alert("📡 MODO OFFLINE!\n\nChecklist salvo no dispositivo.");
        uiBuilder.clearDependentFields();
        if (sp) sp.value = '';
        const pending = await dbManager.getAllPending();
        this.currentPendingPlacas = (pending || []).filter(pd => pd.type === 'OPERACAO').map(pd => pd.payload.placa);
        dashboardEngine.renderExecTable(this.lastParsedData, this.currentPendingPlacas);
        await syncManager.updateStatus(navigator.onLine);
      }
    } finally {
      this.isSubmitting = false;
      uiBuilder.toggleLoader(false);
      syncManager.refreshSyncView();
    }
  }
};

const perdasController = {
  maxRows: 10,
  visibleRows: 1,
  isSubmitting: false,
  init() {
    this.buildPerdasGrid();
    this.updateGridVisibility();
  },

  buildPerdasGrid() {
    const c = document.getElementById('perdasContainer');
    if (!c) return;
    if (!document.getElementById('listaProdutosPerda')) c.insertAdjacentHTML('beforebegin', '<datalist id="listaProdutosPerda"></datalist>');
    if (!document.getElementById('listaLotesGlobais')) c.insertAdjacentHTML('beforebegin', '<datalist id="listaLotesGlobais"></datalist>');
    c.querySelectorAll('.perda-row').forEach(row => row.remove());
    for (let i = 1; i <= this.maxRows; i++) {
      c.insertAdjacentHTML('beforeend', `<div class="grid-row perda-row" id="perdaRow_${i}" style="grid-template-columns: minmax(200px,3fr) 2fr 1.4fr 1.2fr 2fr;"><div class="grid-cell" data-label="Produto"><input type="text" list="listaProdutosPerda" id="perdaProd_${i}" placeholder="Comece a digitar..." onchange="perdasController.onPerdaProdChange(${i})" autocomplete="off"></div><div class="grid-cell" data-label="Lote (Saldo)"><input type="text" list="listaLotesGlobais" id="perdaLote_${i}" placeholder="Lote (Ex: 3574)" autocomplete="off" onchange="perdasController.onPerdaLoteChange(${i})" oninput="perdasController.limparSaldo(${i})"><div id="perdaSaldo_${i}" style="font-size: 11px; font-weight: 800; color: var(--text-muted); margin-top: 4px;"></div></div><div class="grid-cell" data-label="Qtd (t)"><input type="number" id="perdaQtd_${i}" step="0.001" min="0" placeholder="Ex: 2.500" oninput="perdasController.updateGridVisibility()"></div><div class="grid-cell" data-label="Tipo de Perda"><select id="perdaTipo_${i}"><option value="">Motivo...</option><option value="Processo">Processo</option><option value="Carregamento">Carregamento</option><option value="Avaria">Avaria</option></select></div><div class="grid-cell" data-label="Observação"><input type="text" id="perdaObs_${i}" placeholder="Detalhes..."></div></div>`);
    }
  },

  addPerdaRow() {
    if (this.visibleRows >= this.maxRows) {
      toastManager.show(`Limite atingido: máximo de ${this.maxRows} blocos de perda.`, 'warning');
      return;
    }
    this.visibleRows++;
    this.updateGridVisibility();
    const next = document.getElementById(`perdaProd_${this.visibleRows}`);
    if (next) next.focus();
  },

  updateGridVisibility() {
    const btn = document.getElementById('btnAddPerdaRow');
    for (let i = 1; i <= this.maxRows; i++) {
      const row = document.getElementById(`perdaRow_${i}`);
      if (!row) continue;
      const pp = document.getElementById(`perdaProd_${i}`);
      const pq = document.getElementById(`perdaQtd_${i}`);
      const pl = document.getElementById(`perdaLote_${i}`);
      const hasContent = (pp && pp.value !== '') || (pq && parseFloat(pq.value) > 0) || (pl && pl.value.trim() !== '');
      row.style.display = (i <= this.visibleRows || hasContent) ? '' : 'none';
    }
    if (btn) {
      const atingiuMax = this.visibleRows >= this.maxRows;
      btn.disabled = atingiuMax;
      btn.style.opacity = atingiuMax ? '0.65' : '1';
      btn.textContent = atingiuMax ? `➕ LIMITE DE ${this.maxRows} BLOCOS` : `➕ ADICIONAR BLOCO DE PERDA (${this.visibleRows}/${this.maxRows})`;
    }
  },

  populateProdutos(lotesMap) {
    const keys = Object.keys(lotesMap).filter(k => /[a-zA-Z]/.test(k)).sort();
    const datalist = document.getElementById('listaProdutosPerda');
    if (datalist) datalist.innerHTML = keys.map(p => `<option value="${escapeHTML(p)}">`).join('');
  },

  populateGlobais() {
    const dl = document.getElementById('listaLotesGlobais');
    if (dl && uiBuilder.lotesFlat) dl.innerHTML = uiBuilder.lotesFlat.map(item => `<option value="${escapeHTML(item.lote)}">${escapeHTML(item.produto)}</option>`).join('');
  },

  onPerdaProdChange(index) {
    const pp = document.getElementById(`perdaProd_${index}`);
    const pl = document.getElementById(`perdaLote_${index}`);
    const ps = document.getElementById(`perdaSaldo_${index}`);
    if (!pp || !pl || !ps) return;
    const pVal = normalizeName(pp.value);
    if (pVal && (!uiBuilder.lotesMap || !uiBuilder.lotesMap[pVal])) pp.style.borderColor = 'var(--danger)';
    else {
      pp.style.borderColor = 'var(--success)';
      if (!pl.value) { ps.innerText = 'Digite o lote para ver saldo...'; ps.style.color = 'var(--text-muted)'; }
    }
    this.updateGridVisibility();
  },

  limparSaldo(index) {
    const ps = document.getElementById(`perdaSaldo_${index}`);
    const pl = document.getElementById(`perdaLote_${index}`);
    if (ps) { ps.innerText = 'Pressione Enter para buscar...'; ps.style.color = 'var(--text-muted)'; }
    if (pl) pl.style.borderColor = 'var(--border)';
  },

  onPerdaLoteChange(index) {
    const loteInput = document.getElementById(`perdaLote_${index}`);
    const prodInput = document.getElementById(`perdaProd_${index}`);
    const lblSaldo = document.getElementById(`perdaSaldo_${index}`);
    if (!loteInput || !prodInput || !lblSaldo) return;
    const val = loteInput.value.trim().toUpperCase();
    const valNumStr = String(Number(val.replace(/\D/g, '')));
    if (!val) {
      lblSaldo.innerText = '';
      loteInput.style.borderColor = 'var(--border)';
      return;
    }

    let matches = (uiBuilder.lotesFlat || []).filter(item => {
      const itemNumStr = String(Number(String(item.lote).replace(/\D/g, '')));
      if (valNumStr.length >= 4 && !isNaN(valNumStr)) return itemNumStr.endsWith(valNumStr) || valNumStr.endsWith(itemNumStr);
      return itemNumStr === valNumStr || item.lote.toUpperCase() === val;
    });
    if (matches.length === 0) matches = (uiBuilder.lotesFlat || []).filter(item => item.lote.toUpperCase().includes(val));

    if (matches.length === 1) {
      loteInput.value = matches[0].lote;
      prodInput.value = matches[0].produto;
      lblSaldo.innerText = `Saldo: ${formatTons(matches[0].saldo)}t`;
      lblSaldo.style.color = 'var(--success)';
      loteInput.style.borderColor = 'var(--success)';
      prodInput.style.borderColor = 'var(--success)';
    } else if (matches.length > 1) {
      lblSaldo.innerText = `⚠️ ${matches.length} lotes com o final ${val}.`;
      lblSaldo.style.color = 'var(--warning)';
      loteInput.style.borderColor = 'var(--warning)';
    } else {
      lblSaldo.innerText = `❌ Não encontrado.`;
      lblSaldo.style.color = 'var(--danger)';
      loteInput.style.borderColor = 'var(--danger)';
    }
    this.updateGridVisibility();
  },

  limparPerda() {
    for (let i = 1; i <= this.maxRows; i++) {
      const pp = document.getElementById(`perdaProd_${i}`); if (pp) { pp.value = ''; pp.style.borderColor = 'var(--border)'; }
      const pq = document.getElementById(`perdaQtd_${i}`); if (pq) pq.value = '';
      const pl = document.getElementById(`perdaLote_${i}`); if (pl) { pl.value = ''; pl.style.borderColor = 'var(--border)'; }
      const ps = document.getElementById(`perdaSaldo_${i}`); if (ps) ps.innerText = '';
      const pt = document.getElementById(`perdaTipo_${i}`); if (pt) pt.value = '';
      const po = document.getElementById(`perdaObs_${i}`); if (po) po.value = '';
    }
    uiBuilder.updateTimestamp();
    this.visibleRows = 1;
    this.updateGridVisibility();
  },

  async submitPerda() {
    if (this.isSubmitting) return;
    const pa = document.getElementById('perdaAcomp');
    const acomp = pa ? pa.value : '';
    if (!acomp) return alert('⚠️ Responsável não identificado.');
    const perdas = [];

    for (let i = 1; i <= this.maxRows; i++) {
      const pRaw = document.getElementById(`perdaProd_${i}`)?.value || '';
      const p = normalizeName(pRaw);
      const lRaw = document.getElementById(`perdaLote_${i}`)?.value.trim() || '';
      const q = parseBRNumber(document.getElementById(`perdaQtd_${i}`)?.value || '');
      const pt = document.getElementById(`perdaTipo_${i}`);
      const po = document.getElementById(`perdaObs_${i}`);
      if (pRaw || lRaw || q > 0) {
        if (!pRaw || !lRaw || q <= 0) return alert(`⚠️ A Linha ${i} está incompleta.`);
        const loteValido = (uiBuilder.lotesFlat || []).find(item => item.lote.toUpperCase() === lRaw.toUpperCase() && normalizeName(item.produto) === p);
        if (!loteValido) return alert(`⚠️ O Lote "${lRaw}" não existe (Linha ${i}).`);
        perdas.push({ produto: pRaw, qtd: q, lote: lRaw, tipo: pt ? pt.value || 'Não informado' : 'Não informado', obs: po ? po.value : '' });
      }
    }

    if (perdas.length === 0) return alert('⚠️ Preencha ao menos uma perda.');
    uiBuilder.toggleLoader(true, "Transmitindo Baixa...");
    const payload = { tipo: 'BAIXA_PERDA', dataHora: new Date().toLocaleString('pt-BR'), acomp, perdas };
    const meta = dbManager.buildMeta('BAIXA_PERDA', payload);
    payload.syncId = meta.syncId;
    payload.localId = meta.localId;
    payload.fingerprint = meta.fingerprint;
    if (dbManager.isSent(meta.syncId, meta.fingerprint) || dbManager.isFingerprintSent(meta.fingerprint)) {
      uiBuilder.toggleLoader(false);
      alert("Esse registro de perda ja foi enviado.");
      return;
    }
    const dupOpen = await dbManager.findOpenDuplicate(meta.syncId, meta.fingerprint);
    if (dupOpen) {
      uiBuilder.toggleLoader(false);
      toastManager.show('Esse registro de perda ja esta na fila.', 'warning');
      syncManager.refreshSyncView();
      return;
    }
    this.isSubmitting = true;

    try {
      const resp = await apiService.sendDataToAppScript(payload);
      if (resp.success) {
        dbManager.markSent(meta.syncId, meta.fingerprint, 'BAIXA_PERDA', payload);
        backupManager.addEntry('BAIXA_PERDA', { statusEnvio: 'ONLINE', payload, syncId: meta.syncId });
        alert('✅ PERDA REGISTRADA COM SUCESSO!');
        uiBuilder.updateGlobalUpdateTimestamp();
        this.limparPerda();
      } else {
        alert("❌ ERRO DO SCRIPT:\n\n" + resp.message);
      }
    } catch(err) {
      const queued = await dbManager.savePending('BAIXA_PERDA', payload, {
        syncId: meta.syncId,
        localId: meta.localId,
        fingerprint: meta.fingerprint,
        status: navigator.onLine ? 'ERRO' : 'AGUARDANDO_INTERNET'
      });
      if (queued.ok) {
        backupManager.addEntry('BAIXA_PERDA', { statusEnvio: 'PENDENTE_SYNC', payload, syncId: meta.syncId });
        alert("📡 MODO OFFLINE!\n\nSalvo no dispositivo.");
        this.limparPerda();
      } else if (queued.reason === 'JA_ENVIADO') {
        alert("Esse registro de perda já foi enviado.");
      } else if (queued.reason === 'DUPLICADO_LOCAL') {
        toastManager.show('Esse registro de perda já está na fila.', 'warning');
      }
      syncManager.updateStatus(navigator.onLine);
    } finally {
      this.isSubmitting = false;
      uiBuilder.toggleLoader(false);
      syncManager.refreshSyncView();
    }
  }
};

window.addEventListener('DOMContentLoaded', async () => {
  try {
    await appController.initApp();
  } catch (error) {
    reportClientError(`Falha na inicialização: ${error.message}`);
    const fe = document.getElementById('fatalError');
    if (fe) fe.style.display = 'block';
  }
});
