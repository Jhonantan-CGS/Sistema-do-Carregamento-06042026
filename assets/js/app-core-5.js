const uiBuilder = {
  localDataStore: [],
  lotesMap: {},
  lotesFlat: [],
  allLotesFlat: [],
  historicoCargas: [],
  rncCargas: [],
  visibleEmptyRows: 1,
  selectedCargaItems: [],
  checklistState: {},

  init() {
    this.buildChecklistUI();
    this.initChecklistInteractions();
    this.buildUnifiedGrid();
    this.updateGridVisibility();
    this.updateTimestamp();
    const startupTab = new URLSearchParams(window.location.search).get('tab')
      || String(window.location.hash || '').replace(/^#tab-?/, '');
    if (startupTab && document.getElementById(`tab-${startupTab}`)) {
      this.switchTab(null, startupTab);
    }
  },

  switchTab(event, tabId) {
    const previousTabId = document.querySelector('.tab-content.active')?.id?.replace(/^tab-/, '') || '';
    const shouldPlayTabSound = Boolean(event?.currentTarget) && previousTabId !== tabId;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

    if (event && event.currentTarget) event.currentTarget.classList.add('active');
    if (!event) {
      const targetBtn = document.querySelector(`.tab-btn[data-tab="${tabId}"]`);
      if (targetBtn) targetBtn.classList.add('active');
    }
    const tabEl = document.getElementById(`tab-${tabId}`);
    if (!tabEl) return;
    tabEl.classList.add('active');
    if (shouldPlayTabSound && typeof soundManager !== 'undefined') soundManager.play('tab');
    if (typeof priorityAlertManager !== 'undefined') {
      if (tabId === 'op' || tabId === 'perdas') priorityAlertManager.hideOverlay();
      else setTimeout(() => priorityAlertManager.remindActiveAlerts(), 120);
    }

    if (tabId === 'sync') {
      syncManager.refreshSyncView();
    }
  },

  updateTimestamp() {
    const d1 = document.getElementById('inpDataHora');
    if (d1) d1.value = new Date().toLocaleString('pt-BR');
  },

  updateGlobalUpdateTimestamp(date = new Date()) {
    const el = document.getElementById('lastUpdateText');
    if (el) el.innerText = `Última atualização: ${date.toLocaleDateString('pt-BR')} ${date.toLocaleTimeString('pt-BR')}`;
  },

  buildChecklistUI() {
    const cg = document.getElementById('checklistGrid');
    if (!cg) return;
    this.checklistState = {};
    cg.innerHTML = config.checklist.map(c => {
      if (c.type === 'text') {
        return `<div class="check-card check-text-card" id="card_${c.id}">
          <div class="check-title">${escapeHTML(c.lbl)}</div>
          <textarea id="chk_text_${c.id}" placeholder="Observações adicionais..." style="min-height: 100px; margin-top: 12px; border-width:2px;"></textarea>
        </div>`;
      }
      const isInv = c.type === 'yn_r';
      const simClass = isInv ? 'check-opt-danger' : 'check-opt-success';
      const naoClass = isInv ? 'check-opt-success' : 'check-opt-danger';
      const groupLabel = escapeHTML(c.lbl).replace(/"/g, '&quot;');
      return `<div class="check-card" id="card_${c.id}" data-inverted="${isInv}" data-value="">
        <div class="check-title">${escapeHTML(c.lbl)}</div>
        <div class="check-options" role="radiogroup" aria-label="${groupLabel}">
          <button type="button" class="check-opt ${simClass}"
            data-check-id="${c.id}" data-check-value="sim"
            role="radio" aria-checked="false">SIM</button>
          <button type="button" class="check-opt ${naoClass}"
            data-check-id="${c.id}" data-check-value="nao"
            role="radio" aria-checked="false">NÃO</button>
        </div>
      </div>`;
    }).join('');
  },

  initChecklistInteractions() {
    const container = document.getElementById('checklistGrid');
    if (!container || container.dataset.bound === 'true') return;
    container.dataset.bound = 'true';
    container.addEventListener('click', (event) => {
      const button = event.target.closest('.check-opt[data-check-id][data-check-value]');
      if (!button) return;
      this.selectChecklistOption(button.dataset.checkId, button.dataset.checkValue);
    });
    container.addEventListener('keydown', (event) => {
      const button = event.target.closest('.check-opt[data-check-id][data-check-value]');
      if (!button) return;
      if (event.key === ' ' || event.key === 'Enter') {
        event.preventDefault();
        this.selectChecklistOption(button.dataset.checkId, button.dataset.checkValue);
        return;
      }
      if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return;
      const options = Array.from(button.parentElement?.querySelectorAll('.check-opt[data-check-id]') || []);
      if (options.length < 2) return;
      event.preventDefault();
      const currentIndex = options.indexOf(button);
      const direction = (event.key === 'ArrowLeft' || event.key === 'ArrowUp') ? -1 : 1;
      const nextButton = options[(currentIndex + direction + options.length) % options.length];
      nextButton?.focus();
      if (nextButton?.dataset.checkId && nextButton?.dataset.checkValue) {
        this.selectChecklistOption(nextButton.dataset.checkId, nextButton.dataset.checkValue);
      }
    });
  },

  getChecklistValue(id) {
    if (Object.prototype.hasOwnProperty.call(this.checklistState, id)) return this.checklistState[id];
    const card = document.getElementById(`card_${id}`);
    return card ? String(card.dataset.value || '') : '';
  },

  selectChecklistOption(id, val, options = {}) {
    const card = document.getElementById(`card_${id}`);
    if (!card) return;
    const normalizedValue = val === 'nao' ? 'nao' : 'sim';
    this.checklistState[id] = normalizedValue;
    const previousValue = String(card.dataset.value || '');
    if (previousValue === normalizedValue) {
      this.updateChecklistVisualState(id, normalizedValue);
      return;
    }
    card.dataset.value = normalizedValue;
    this.updateChecklistVisualState(id, normalizedValue);
    this.onCheckChange(id, normalizedValue);
    if (!options.silent && typeof soundManager !== 'undefined') soundManager.play('click');
  },

  updateChecklistVisualState(id, selectedValue = '') {
    const card = document.getElementById(`card_${id}`);
    if (!card) return;

    const optionButtons = card.querySelectorAll(`.check-opt`);
    optionButtons.forEach((btn) => {
      const active = btn.dataset.checkValue === selectedValue;
      btn.classList.toggle('is-selected', active);
      btn.setAttribute('aria-checked', active ? 'true' : 'false');
      btn.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
    
    // Update card color
    const isInv = card.dataset.inverted === 'true';
    card.classList.remove('sim', 'nao');
    if (selectedValue) {
      const visualType = isInv ? (selectedValue === 'sim' ? 'nao' : 'sim') : selectedValue;
      card.classList.add(visualType);
    }
  },

  onCheckChange(id, val) {
    debugEngine.log(`Checklist [${id}] alterado para: ${val}`, 'info');
  },

  resetChecklistState() {
    this.checklistState = {};
    config.checklist
      .filter((item) => item.type !== 'text')
      .forEach((item) => {
        const card = document.getElementById(`card_${item.id}`);
        if (card) {
          card.dataset.value = '';
          card.classList.remove('sim', 'nao');
        }
        this.updateChecklistVisualState(item.id, '');
      });
  },

  buildUnifiedGrid() {
    const m = document.getElementById('materialsContainer');
    if (!m) return;
    for (let i = 1; i <= 10; i++) {
      m.insertAdjacentHTML('beforeend', `<div class="grid-row" id="gridRow_${i}"><div class="grid-cell" data-label="Produto Liberado"><div id="lblDestino_${i}" style="font-size: 10px; font-weight: 900; color: var(--secondary); margin-bottom: 8px; display: none;"></div><select id="selProd_${i}" onchange="uiBuilder.onProdChange(${i})" title="Selecione o produto liberado nesta linha"><option value="">Selecione o produto liberado...</option></select></div><div class="grid-cell" data-label="Carregado (t)"><input type="number" id="inpQtd_${i}" step="0.001" min="0" placeholder="Ex: 12.500" title="Informe a quantidade carregada em toneladas" oninput="uiBuilder.recalcularLinhas(${i})"></div><div class="grid-cell" data-label="Lote Utilizado"><input type="search" id="inpBuscaLote_${i}" class="lot-search-input" placeholder="Selecione primeiro o produto..." title="Digite parte do lote para localizar mais rapido" autocomplete="off" spellcheck="false" disabled oninput="uiBuilder.onLoteSearchInput(${i})"><select id="selLote_${i}" onchange="uiBuilder.onLoteChange(${i})" title="Escolha o lote fisico usado no carregamento" data-all-options="[]"><option value="">Selecione o lote fisico...</option></select><div id="lblSaldo_${i}" style="font-size: 11px; font-weight: 800; color: var(--success); margin-top: 4px;"></div></div><div class="grid-cell" data-label="Avaria (t)"><input type="number" id="inpPerda_${i}" step="0.001" min="0" placeholder="Ex: 0.150" title="Preencha apenas se houve perda no item"></div><div class="grid-cell" data-label="Motivo Avaria"><select id="selMotivo_${i}" title="Obrigatorio quando houver avaria"><option value="">Selecione o motivo da avaria...</option><option value="Rasgo na embalagem">Rasgo</option><option value="Qualidade">Qualidade</option><option value="Transporte">Transporte</option></select></div></div>`);
    }
    m.parentElement.insertAdjacentHTML('beforeend', `<div style="padding: 16px; text-align: center; background: linear-gradient(180deg, rgba(15, 23, 42, 0.68), rgba(30, 41, 59, 0.52)); border-top: 1px solid rgba(96, 165, 250, 0.12);"><button type="button" id="btnToggleGridRows" onclick="uiBuilder.toggleGridRows()" style="background: transparent; border: 2px dashed rgba(96, 165, 250, 0.42); color: #BFDBFE; box-shadow: none;">⬇ Expandir linha</button></div>`);
  },

  getSortedLoteEntries(lotesObj = {}) {
    return Object.entries(lotesObj || {})
      .filter(([lote]) => Boolean(String(lote || '').trim()))
      .sort((a, b) => String(a[0]).localeCompare(String(b[0]), 'pt-BR', { numeric: true, sensitivity: 'base' }));
  },

  getStoredLoteEntries(index) {
    const loteSelect = document.getElementById(`selLote_${index}`);
    if (!loteSelect || !loteSelect.dataset.allOptions) return [];
    try {
      return JSON.parse(loteSelect.dataset.allOptions).map((item) => [item.value, item.saldo]);
    } catch (error) {
      debugEngine.log(`Falha ao recuperar lotes armazenados da linha ${index}: ${error.message}`, 'warn');
      return [];
    }
  },

  renderLoteOptions(index, loteEntries = [], filterTerm = '', preferredValue = '') {
    const loteSelect = document.getElementById(`selLote_${index}`);
    const loteSearch = document.getElementById(`inpBuscaLote_${index}`);
    if (!loteSelect) return [];

    const normalizedFilter = normalizeName(filterTerm || '');
    const filteredEntries = normalizedFilter
      ? loteEntries.filter(([lote]) => normalizeName(lote).includes(normalizedFilter))
      : loteEntries.slice();

    loteSelect.dataset.allOptions = JSON.stringify(loteEntries.map(([value, saldo]) => ({ value, saldo })));
    loteSelect.innerHTML = `<option value="">${normalizedFilter ? 'Selecione o lote encontrado...' : 'Selecione o lote fisico...'}</option>`
      + filteredEntries.map(([lote, saldo]) => `<option value="${escapeHTML(lote)}" data-saldo="${saldo}">${escapeHTML(lote)}</option>`).join('');

    if (loteSearch) {
      loteSearch.disabled = loteEntries.length === 0;
      loteSearch.placeholder = loteEntries.length === 0
        ? 'Sem lote disponivel para este produto'
        : 'Digite para localizar o lote...';
    }

    if (preferredValue && filteredEntries.some(([lote]) => lote === preferredValue)) {
      loteSelect.value = preferredValue;
    }

    return filteredEntries;
  },

  onLoteSearchInput(index) {
    const loteSearch = document.getElementById(`inpBuscaLote_${index}`);
    const loteSelect = document.getElementById(`selLote_${index}`);
    const lblSaldo = document.getElementById(`lblSaldo_${index}`);
    if (!loteSearch || !loteSelect) return;

    const allEntries = this.getStoredLoteEntries(index);
    const filteredEntries = this.renderLoteOptions(index, allEntries, loteSearch.value, loteSelect.value);
    const normalizedSearch = normalizeName(loteSearch.value || '');
    const exactMatch = normalizedSearch
      ? filteredEntries.find(([lote]) => normalizeName(lote) === normalizedSearch)
      : null;

    if (exactMatch) {
      loteSelect.value = exactMatch[0];
      this.onLoteChange(index);
      return;
    }

    if (filteredEntries.length === 1) {
      loteSelect.value = filteredEntries[0][0];
      this.onLoteChange(index);
      return;
    }

    if (loteSelect.value && !filteredEntries.some(([lote]) => lote === loteSelect.value)) {
      loteSelect.value = '';
    }

    if (lblSaldo) {
      lblSaldo.innerText = filteredEntries.length === 0 && normalizedSearch
        ? 'Nenhum lote encontrado para esta busca.'
        : '';
    }
    this.recalcularLinhas();
  },

  handleImageUpload(event, previewId, hiddenInputId) {
    const file = event.target.files[0];
    const preview = document.getElementById(previewId);
    const hiddenInput = document.getElementById(hiddenInputId);
    if (!file) {
      if (preview) { preview.classList.remove('has-image'); preview.style.backgroundImage = 'none'; preview.innerHTML = '<span>📷 Câmera</span>'; }
      if (hiddenInput) hiddenInput.value = '';
      return;
    }
    if (preview) preview.innerHTML = '<span>Carregando 3D...</span>';
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        let width = img.width;
        let height = img.height;
        if (width > height) { if (width > 1000) { height *= 1000 / width; width = 1000; } }
        else { if (height > 1000) { width *= 1000 / height; height = 1000; } }
        canvas.width = width;
        canvas.height = height;
        canvas.getContext('2d').drawImage(img, 0, 0, width, height);
        const dataUrl = canvas.toDataURL('image/jpeg', 0.6);
        if (preview) { preview.style.backgroundImage = `url(${dataUrl})`; preview.classList.add('has-image'); preview.innerHTML = '<span>Foto Pronta ✔️</span>'; }
        if (hiddenInput) hiddenInput.value = dataUrl;
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  },

  toggleGridRows() {
    const currentShown = Array.from({ length: 10 }, (_, i) => document.getElementById(`gridRow_${i + 1}`)).filter(r => r && r.style.display !== 'none').length;
    const showingAll = currentShown >= 10 || this.visibleEmptyRows >= 10;
    this.visibleEmptyRows = showingAll ? 1 : Math.min(10, this.visibleEmptyRows + 1);
    this.updateGridVisibility();
  },

  updateGridVisibility() {
    const btn = document.getElementById('btnToggleGridRows');
    let hasContent = 0, emptyShown = 0, totalShown = 0;
    const minEmpty = Math.max(1, this.visibleEmptyRows);
    for (let i = 1; i <= 10; i++) {
      const row = document.getElementById(`gridRow_${i}`);
      if (!row) continue;
      const sel = document.getElementById(`selProd_${i}`);
      const qtd = document.getElementById(`inpQtd_${i}`);
      if ((sel && sel.value !== '') || (qtd && parseFloat(qtd.value) > 0)) { row.style.display = ''; hasContent++; totalShown++; }
      else {
        if (emptyShown < minEmpty) { row.style.display = ''; emptyShown++; totalShown++; }
        else row.style.display = 'none';
      }
    }
    if (btn) btn.innerHTML = totalShown >= 10 ? '⬆ Ocultar vazias' : '⬇ Expandir linha';
  },

  populateSelects(parsedData) {
    this.localDataStore = parsedData;
    this.selectedCargaItems = [];
    this.clearDependentFields();
    const placas = [...new Set(parsedData.map(d => d.placa).filter(Boolean))].sort();
    const clientes = [...new Set(parsedData.map(d => String(d.cliente || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const placaSelect = document.getElementById('selPlacaBusca');
    const clienteSelect = document.getElementById('selClienteBusca');
    if (!placaSelect) return;
    const currentPlaca = placaSelect.value;
    const currentCliente = clienteSelect ? clienteSelect.value : '';
    placaSelect.innerHTML = '<option value="">Selecione a Placa...</option>' + placas.map(p => `<option value="${escapeHTML(p)}">${escapeHTML(p)}</option>`).join('');
    if (currentPlaca && placas.includes(currentPlaca)) placaSelect.value = currentPlaca;
    if (clienteSelect) {
      clienteSelect.innerHTML = '<option value="">Selecione o Cliente...</option>' + clientes.map(c => `<option value="${escapeHTML(c)}">${escapeHTML(c)}</option>`).join('');
      if (currentCliente && clientes.includes(currentCliente)) clienteSelect.value = currentCliente;
    }
  },

  applyCargaSelection(items = [], defaultPlaca = '') {
    this.selectedCargaItems = Array.isArray(items) ? items : [];
    if (!Array.isArray(items) || items.length === 0) return;
    const pedContainer = document.getElementById('pedidosContainer');
    if (pedContainer) pedContainer.style.display = 'block';

    const mapPed = new Set();
    const pedidosUnicos = [];
    items.forEach(i => {
      const key = `${i.pedido}-${i.cliente}`;
      if (!mapPed.has(key)) {
        mapPed.add(key);
        pedidosUnicos.push({ cliente: i.cliente, pedido: i.pedido, cidade: i.cidade, estado: i.estado });
      }
    });

    let htmlTabela = `<div class="grid-header" style="grid-template-columns: 2fr 1.2fr 1.5fr;"><div>Cliente</div><div>Pedido</div><div>Destino</div></div>`;
    pedidosUnicos.forEach(p => {
      htmlTabela += `<div class="grid-row" style="grid-template-columns: 2fr 1.2fr 1.5fr; border-left: 4px solid var(--primary);"><div class="grid-cell" style="font-weight: 800;">${escapeHTML(p.cliente)}</div><div class="grid-cell" style="font-family: 'JetBrains Mono', monospace;">${escapeHTML(p.pedido)}</div><div class="grid-cell">${escapeHTML(p.cidade)} / ${escapeHTML(p.estado)}</div></div>`;
    });
    const pedidosGrid = document.getElementById('listaPedidosGrid');
    if (pedidosGrid) pedidosGrid.innerHTML = htmlTabela;

    const allObs = items.map(i => i.obsStr).filter(Boolean);
    let motEncontrado = '';
    for (const obs of allObs) {
      const match = /MOT(?:ORISTA)?[:\s]+([A-ZÀ-Úa-z\s]+?)(?:\s*-|\s*CPF|\s*RG|$)/i.exec(obs);
      if (match) motEncontrado = match[1].trim();
    }

    const primeiraPlacaValida = items.map(i => String(i.placa || '').trim()).find((p) => p && p !== 'SEM PLACA') || '';
    const inpP = document.getElementById('inpPlaca');
    if (inpP) inpP.value = defaultPlaca && defaultPlaca !== 'SEM PLACA' ? defaultPlaca : primeiraPlacaValida;
    const inpM = document.getElementById('inpMotorista');
    if (inpM) inpM.value = motEncontrado;
    const inpO = document.getElementById('inpObs');
    if (inpO) inpO.value = [...new Set(allObs)].join('\n');

    let rowIdx = 1;
    const agrupamentoCliProd = {};
    items.forEach(i => {
      if (i.produtoDesc) {
        const key = `${i.cliente}|${i.cidade}/${i.estado}|${i.produtoDesc}`;
        if (!agrupamentoCliProd[key]) agrupamentoCliProd[key] = { cliente: i.cliente, destino: `${i.cidade}/${i.estado}`, produto: i.produtoDesc, qtd: 0 };
        agrupamentoCliProd[key].qtd += Number(i.qtd);
      }
    });

    for (let i = 1; i <= 10; i++) {
      const s = document.getElementById(`selProd_${i}`);
      if (s) s.innerHTML = '<option value="">Selecione o produto liberado...</option>' + [...new Set(items.map(i => i.produtoDesc).filter(Boolean))].map(p => `<option value="${escapeHTML(p)}">${escapeHTML(p)}</option>`).join('');
    }

    Object.values(agrupamentoCliProd).forEach(item => {
      if (rowIdx > 10) return;
      const sel = document.getElementById(`selProd_${rowIdx}`);
      const lbl = document.getElementById(`lblDestino_${rowIdx}`);
      if (lbl) { lbl.innerText = `📦 ${item.cliente} 📍 ${item.destino}`; lbl.style.display = 'inline-block'; }
      if (sel) { sel.value = item.produto; sel.dataset.sugQtd = item.qtd; }
      this.onProdChange(rowIdx, item.qtd);
      rowIdx++;
    });

    this.verificarPlacaJaCarregou();
    this.visibleEmptyRows = 0;
    this.updateGridVisibility();
  },

  clearDependentFields() {
    this.selectedCargaItems = [];
    ['inpPlaca','inpMotorista','inpObs'].forEach(id => { const e = document.getElementById(id); if (e) e.value = ''; });
    const pedContainer = document.getElementById('pedidosContainer');
    if (pedContainer) pedContainer.style.display = 'none';
    for (let i = 1; i <= 10; i++) {
      const sel = document.getElementById(`selProd_${i}`); if (sel) { sel.innerHTML = '<option value="">Selecione o produto liberado...</option>'; sel.dataset.sugQtd = ''; }
      const q = document.getElementById(`inpQtd_${i}`); if (q) q.value = '';
      const loteBusca = document.getElementById(`inpBuscaLote_${i}`);
      if (loteBusca) {
        loteBusca.value = '';
        loteBusca.disabled = true;
        loteBusca.placeholder = 'Selecione primeiro o produto...';
      }
      const sl = document.getElementById(`selLote_${i}`);
      if (sl) {
        sl.dataset.allOptions = '[]';
        sl.innerHTML = '<option value="">Selecione o lote fisico...</option>';
      }
      const ls = document.getElementById(`lblSaldo_${i}`); if (ls) ls.innerText = '';
      const ip = document.getElementById(`inpPerda_${i}`); if (ip) ip.value = '';
      const sm = document.getElementById(`selMotivo_${i}`); if (sm) sm.value = '';
      const lbl = document.getElementById(`lblDestino_${i}`); if (lbl) { lbl.innerText = ''; lbl.style.display = 'none'; }
    }
    this.resetChecklistState();
    const ta = document.getElementById('chk_text_obs_geral');
    if (ta) ta.value = '';
    ['Frente', 'Traseira', 'Assoalho'].forEach(id => {
      const preview = document.getElementById(`preview${id}`);
      if (preview) { preview.style.backgroundImage = 'none'; preview.classList.remove('has-image'); preview.innerHTML = '<span>📷 Câmera</span>'; }
      const b = document.getElementById(`base64${id}`);
      if (b) b.value = '';
    });
    this.verificarPlacaJaCarregou();
    this.visibleEmptyRows = 1;
    this.updateGridVisibility();
  },

  onPlacaChange() {
    const placaSelect = document.getElementById('selPlacaBusca');
    if (!placaSelect) return;
    const placa = placaSelect.value;
    const clienteSelect = document.getElementById('selClienteBusca');
    if (clienteSelect && placa) clienteSelect.value = '';
    this.clearDependentFields();
    if (!placa) return;

    const items = this.localDataStore.filter(d => d.placa === placa);
    if (items.length > 0) this.applyCargaSelection(items, placa);
  },

  onClienteChange() {
    const clienteSelect = document.getElementById('selClienteBusca');
    if (!clienteSelect) return;
    const cliente = String(clienteSelect.value || '').trim();
    const placaSelect = document.getElementById('selPlacaBusca');
    if (placaSelect && cliente) placaSelect.value = '';
    this.clearDependentFields();
    if (!cliente) return;

    const items = this.localDataStore.filter(d => String(d.cliente || '').trim() === cliente);
    if (items.length === 0) return;

    const confirmar = window.confirm(`Confirmar seleção do cliente ${cliente}?\n\nSerão carregados ${items.length} registro(s) vinculados a este cliente.`);
    if (!confirmar) {
      clienteSelect.value = '';
      return;
    }

    const placas = [...new Set(items.map(i => String(i.placa || '').trim()).filter(Boolean))];
    const placaPrioritaria = placas.find((p) => p === 'SEM PLACA') || placas[0] || '';
    if (placaSelect && placaPrioritaria) placaSelect.value = placaPrioritaria;
    this.applyCargaSelection(items, placaPrioritaria);
  },

  getSelectedCargaItems() {
    if (Array.isArray(this.selectedCargaItems) && this.selectedCargaItems.length > 0) return this.selectedCargaItems;
    const placa = document.getElementById('selPlacaBusca')?.value || '';
    if (placa) return this.localDataStore.filter(d => d.placa === placa);
    const cliente = document.getElementById('selClienteBusca')?.value || '';
    if (cliente) return this.localDataStore.filter(d => String(d.cliente || '').trim() === String(cliente).trim());
    return [];
  },

  verificarPlacaJaCarregou() {
    const p = document.getElementById('inpPlaca');
    const placa = p ? normalizeName(p.value) : '';
    const o = document.getElementById('inpObs');
    const obs = o ? normalizeName(o.value) : '';
    const badge = document.getElementById('badgeJaCarregou');
    if (!badge) return;
    if (!placa || !obs || !this.historicoCargas || this.historicoCargas.length === 0) { badge.style.display = 'none'; return; }

    const ho = new Date();
    const regex = new RegExp(`\\b${String(ho.getDate()).padStart(2, '0')}/${String(ho.getMonth() + 1).padStart(2, '0')}/${ho.getFullYear()}\\b`);
    let jaCarregou = false;
    for (let row of this.historicoCargas) {
      if (!row) continue;
      const str = normalizeName(row.join(' '));
      if (regex.test(str) && str.includes(placa) && str.includes(obs)) { jaCarregou = true; break; }
    }
    badge.style.display = jaCarregou ? 'inline-block' : 'none';
  },

  onProdChange(index, sugQtd = null) {
    const sel = document.getElementById(`selProd_${index}`);
    if (!sel) return;
    const prodKey = normalizeName(sel.value);
    const qtdField = document.getElementById(`inpQtd_${index}`);
    const loteSelect = document.getElementById(`selLote_${index}`);
    const loteSearch = document.getElementById(`inpBuscaLote_${index}`);
    if (loteSearch) {
      loteSearch.value = '';
      loteSearch.disabled = !prodKey;
      loteSearch.placeholder = prodKey ? 'Digite para localizar o lote...' : 'Selecione primeiro o produto...';
    }
    if (loteSelect) {
      loteSelect.dataset.allOptions = '[]';
      loteSelect.innerHTML = '<option value="">Selecione o lote fisico...</option>';
    }
    if (qtdField) qtdField.value = '';
    if (prodKey) {
      const lotesObj = this.lotesMap[prodKey];
      if (lotesObj && Object.keys(lotesObj).length > 0 && loteSelect) {
        const loteEntries = this.getSortedLoteEntries(lotesObj);
        this.renderLoteOptions(index, loteEntries);
        if (loteEntries.length === 1) {
          loteSelect.value = loteEntries[0][0];
          this.onLoteChange(index);
        }
      } else if (loteSelect) {
        loteSelect.dataset.allOptions = '[]';
        loteSelect.innerHTML = '<option value="S/ Lote">Sem Lote no Estoque</option>';
        loteSelect.value = 'S/ Lote';
        if (loteSearch) {
          loteSearch.value = '';
          loteSearch.disabled = true;
          loteSearch.placeholder = 'Sem lote disponivel para este produto';
        }
        this.onLoteChange(index);
      }
      const valQtd = typeof sugQtd === 'number' ? sugQtd : Number(sel.dataset.sugQtd || 0);
      if (valQtd > 0 && qtdField) qtdField.value = formatTons(valQtd);
      this.recalcularLinhas(index);
    } else {
      const l = document.getElementById(`lblDestino_${index}`);
      if (l) l.style.display = 'none';
      this.recalcularLinhas(index);
    }
  },

  onLoteChange(index) {
    const loteSelect = document.getElementById(`selLote_${index}`);
    const loteSearch = document.getElementById(`inpBuscaLote_${index}`);
    const lbl = document.getElementById(`lblSaldo_${index}`);
    if (!loteSelect || !lbl) return;
    const opt = loteSelect.options[loteSelect.selectedIndex];
    if (loteSearch) {
      loteSearch.value = loteSelect.value && loteSelect.value !== 'S/ Lote' ? loteSelect.value : '';
    }
    lbl.innerText = opt && opt.dataset.saldo ? `Saldo: ${formatTons(opt.dataset.saldo)}t` : '';
    this.recalcularLinhas(index);
  },

  recalcularLinhas() {
    const r = (v) => Math.round((Number(v) || 0) * 1000) / 1000;
    const somaUsoPorLote = {};
    for (let i = 1; i <= 10; i++) {
      const p = (document.getElementById(`selProd_${i}`) || {}).value || '';
      const q = r((document.getElementById(`inpQtd_${i}`) || {}).value);
      const l = document.getElementById(`selLote_${i}`)?.value || '';
      if (l && l !== 'S/ Lote' && q > 0) somaUsoPorLote[`${p}||${l}`] = r((somaUsoPorLote[`${p}||${l}`] || 0) + q);
    }
    for (let i = 1; i <= 10; i++) {
      const p = (document.getElementById(`selProd_${i}`) || {}).value || '';
      const qtdField = document.getElementById(`inpQtd_${i}`);
      const lSel = document.getElementById(`selLote_${i}`);
      const lblSaldo = document.getElementById(`lblSaldo_${i}`);
      if (!qtdField) continue;
      if (lSel && lblSaldo && lSel.value && lSel.value !== 'S/ Lote') {
        const opt = lSel.options[lSel.selectedIndex];
        const sOrig = opt ? r(Number(opt.dataset.saldo) || 0) : 0;
        const sReal = r(sOrig - r((somaUsoPorLote[`${p}||${lSel.value}`] || 0) - r(qtdField.value)));
        lblSaldo.innerText = `Disponível: ${formatTons(Math.max(0, sReal))} t`;
        qtdField.max = formatTons(Math.max(0, sReal));
      }
    }
    this.updateGridVisibility();
  },

  formatSyncTime(isoRaw) {
    if (!isoRaw) return '--';
    const dt = new Date(isoRaw);
    if (Number.isNaN(dt.getTime())) return '--';
    return dt.toLocaleString('pt-BR');
  },

  renderSyncCards(items = [], containerId, allowRetry = false) {
    const root = document.getElementById(containerId);
    if (!root) return;
    if (!Array.isArray(items) || items.length === 0) {
      root.innerHTML = `<div style="padding:14px; border:1px dashed rgba(148,163,184,0.45); border-radius:10px; color:#94A3B8; font-size:12px; font-weight:700;">Sem registros neste grupo.</div>`;
      return;
    }
    root.innerHTML = items.map((row) => {
      const payload = row.payload || {};
      const principal = normalizeName(payload.placa || payload.lote || payload.pedido || payload.tipo || row.type || 'SEM ID');
      const typeLabel = String(row.type || payload.tipo || 'REGISTRO');
      const qtdImgs = Number(row.imageCount || 0);
      const tentativas = Number(row.attemptCount || 0);
      const erro = String(row.lastError || '').trim();
      const btnRetry = allowRetry
        ? `<button class="btn-primary" style="min-height:40px; font-size:11px; padding:0 12px;" onclick="syncManager.retryItem('${escapeHTML(String(row.id))}', this)">⟳ Tentar novamente</button>`
        : '';
      const statusColor = row.status === 'ENVIADO' ? '#10B981' : row.status === 'ERRO' ? '#F97316' : row.status === 'ENVIANDO' ? '#0EA5E9' : '#A78BFA';
      return `<div style="border:1px solid rgba(148,163,184,0.35); border-left:4px solid ${statusColor}; border-radius:12px; padding:12px; background:rgba(15,23,42,0.55); display:flex; flex-direction:column; gap:8px;">
        <div style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap;">
          <div style="font-size:13px; font-weight:900; color:#E2E8F0;">${escapeHTML(principal)}</div>
          <div style="font-size:10px; font-weight:800; color:#93C5FD;">${escapeHTML(typeLabel)}</div>
        </div>
        <div style="display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:8px; font-size:11px; color:#CBD5E1;">
          <div><b>Status:</b> ${escapeHTML(String(row.status || '--'))}</div>
          <div><b>Imagens:</b> ${qtdImgs}</div>
          <div><b>Tentativas:</b> ${tentativas}</div>
          <div><b>Criado:</b> ${escapeHTML(this.formatSyncTime(row.createdAt))}</div>
          <div><b>Última tentativa:</b> ${escapeHTML(this.formatSyncTime(row.lastAttemptAt))}</div>
          <div><b>Enviado:</b> ${escapeHTML(this.formatSyncTime(row.sentAt))}</div>
        </div>
        ${erro ? `<div style="font-size:11px; color:#FDBA74; background:rgba(124,45,18,0.25); border:1px solid rgba(251,146,60,0.35); border-radius:8px; padding:8px;"><b>Erro:</b> ${escapeHTML(erro)}</div>` : ''}
        ${btnRetry ? `<div style="display:flex; justify-content:flex-end;">${btnRetry}</div>` : ''}
      </div>`;
    }).join('');
  },

  renderSyncCenter(snapshot = {}) {
    const online = !!snapshot.online;
    const syncing = !!snapshot.syncing;
    const waiting = snapshot.waiting || [];
    const sending = snapshot.sending || [];
    const errors = snapshot.errors || [];
    const sent = snapshot.sent || [];
    const totalPend = Number(snapshot.totalPendentes || 0);
    const pendOps = Number(snapshot.pendentesOperacao || 0);

    const elOnline = document.getElementById('syncOnlineState');
    if (elOnline) {
      elOnline.className = `status-badge ${syncing ? 'syncing' : (online ? 'online' : 'offline')}`;
      elOnline.innerText = syncing ? '● Sincronizando...' : (online ? '● Online' : '● Offline');
    }
    const elLast = document.getElementById('syncLastSyncAt');
    if (elLast) elLast.innerText = snapshot.lastSyncAt ? this.formatSyncTime(snapshot.lastSyncAt) : '--';
    const elTot = document.getElementById('syncCountPending');
    if (elTot) elTot.innerText = String(totalPend);
    const elErr = document.getElementById('syncCountError');
    if (elErr) elErr.innerText = String(errors.length);
    const elSent = document.getElementById('syncCountSent');
    if (elSent) elSent.innerText = String(sent.length);
    const elOps = document.getElementById('syncCountOps');
    if (elOps) elOps.innerText = `${pendOps}/${dbManager.maxOfflineOperacoes}`;
    const mergedWaiting = [...waiting, ...sending].sort((a, b) => new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime());
    this.renderSyncCards(mergedWaiting, 'syncWaitingList', false);
    this.renderSyncCards(errors, 'syncErrorList', true);
    this.renderSyncCards(sent, 'syncSentList', false);
  },

  toggleLoader(show, text = "Sincronizando...") {
    const overlay = document.getElementById('loaderOverlay');
    if (!overlay) return;
    const t = document.getElementById('loaderText');
    if (t) t.innerText = text;
    overlay.style.display = show ? 'flex' : 'none';
  },

  toggleHeaderMenu(event) {
    if (event) event.stopPropagation();
    const menu = document.getElementById('headerMenuPanel');
    if (!menu) return;
    const isShowing = menu.classList.contains('show');
    menu.classList.toggle('show', !isShowing);
    if (!isShowing && !uiBuilder.closeHeaderMenuHandler) {
      uiBuilder.closeHeaderMenuHandler = (e) => {
        if (!e.target.closest('.header-actions')) {
          uiBuilder.closeHeaderMenu();
        }
      };
      setTimeout(() => document.addEventListener('click', uiBuilder.closeHeaderMenuHandler), 10);
    }
  },

  closeHeaderMenu() {
    const menu = document.getElementById('headerMenuPanel');
    if (menu) menu.classList.remove('show');
    if (this.closeHeaderMenuHandler) {
      document.removeEventListener('click', this.closeHeaderMenuHandler);
      this.closeHeaderMenuHandler = null;
    }
  },

  toggleSyncDrawer() {
    const drawer = document.getElementById('syncDrawer');
    if (drawer) {
      drawer.classList.toggle('show');
    }
  }
};

