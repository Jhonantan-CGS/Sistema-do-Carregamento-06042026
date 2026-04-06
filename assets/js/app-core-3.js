const dataParser = {
  parse(rawValues) {
    const parsed = [];
    if (!Array.isArray(rawValues) || rawValues.length === 0) return parsed;

    let headerDate = null;
    if (rawValues[0] && Array.isArray(rawValues[0])) {
      const headerStr = rawValues[0].join(' ');
      const dateMatch = headerStr.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
      if (dateMatch) {
        headerDate = dateMatch[1];
        debugEngine.log(`[PARSER] Data do cabeçalho detectada: ${headerDate}`, "info");
      }
    }

    debugEngine.log(`[PARSER] Iniciando parsing a partir da linha 2...`, "info");
    let countTotal = 0;
    let countSaltado = 0;

    for (let i = 2; i < rawValues.length; i++) {
      const row = rawValues[i];
      if (!row) { countSaltado++; continue; }
      
      const colA = row[0] ? String(row[0]).trim().toUpperCase() : '';
      
      if (colA.includes('TOTAL') || colA.includes('TOTAIS') || colA.includes('SISTEMA') || colA.includes('AGUARDANDO') || colA.includes('PAINEL') || colA.includes('RELATORIO')) {
        debugEngine.log(`[PARSER] Parando na linha ${i+1}: encontrado "${colA.substring(0,50)}"`, "info");
        break;
      }
      
      const cliente = row[8] ? String(row[8]).trim() : '';
      if (!cliente || cliente.toUpperCase() === 'CLIENTE' || cliente.length < 2) {
        countSaltado++;
        continue;
      }
      
      parsed.push({
        dataRaw: headerDate,
        pedido: row[3] ? String(row[3]).trim() : '',
        cliente,
        cidade: row[9] ? String(row[9]).trim() : '-',
        estado: row[10] ? String(row[10]).trim() : '-',
        qtd: parseBRNumber(row[2]),
        bloqueio: row[14] ? String(row[14]).trim() : 'LIBERADO',
        obsStr: row[15] ? String(row[15]).trim() : '',
        produtoDesc: normalizeName(row[1]),
        placa: extractPlaca(row[15] ? String(row[15]).trim() : ''),
        pallet: row[11] ? String(row[11]).trim() : 'NÃO',
        filme: row[12] ? String(row[12]).trim() : 'NÃO'
      });
      countTotal++;
    }
    
    debugEngine.log(`[PARSER] Resultado: ${countTotal} carregamentos parseados, ${countSaltado} linhas ignoradas`, countTotal > 0 ? "success" : "warn");
    return parsed;
  },

  parseEstoque(values) {
    const map = {};
    const flat = [];
    const allFlat = [];
    if (!Array.isArray(values)) return { map, flat, allFlat };
    let ch = null;

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row || !row[0]) continue;
      const colA = String(row[0]).trim();
      if (colA.toUpperCase().startsWith('TOTAL') || colA.toUpperCase() === 'LOTE' || colA.includes('%') || !/[a-zA-Z]/.test(colA)) { ch = null; continue; }
      if (row.length >= 2 && row[1] && String(row[1]).trim().toUpperCase() !== 'LOTE') {
        const saldo = typeof row[5] === 'number' ? row[5] : parseBRNumber(row[5]);
        const kA = normalizeName(colA);
        const kL = String(row[1]).trim();
        if (kL) {
          allFlat.push({
            lote: kL,
            produto: ch ? normalizeName(ch) : kA,
            saldo: Number(saldo || 0)
          });
        }
        if (saldo > 0) {
          if (kA !== normalizeName(kL)) { if (!map[kA]) map[kA] = {}; map[kA][kL] = saldo; }
          if (ch) { const kH = normalizeName(ch); if (!map[kH]) map[kH] = {}; map[kH][kL] = saldo; }
          flat.push({ lote: kL, produto: ch ? normalizeName(ch) : kA, saldo: saldo });
        }
      } else {
        ch = colA;
      }
    }
    return { map, flat, allFlat };
  },

  parseRnc(values) {
    const rncs = [];
    if (!Array.isArray(values) || values.length <= 1) return rncs;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row) continue;
      rncs.push({
        data: String(row[1] || '').trim(),
        cliente: String(row[2] || '').trim().toUpperCase(),
        transportador: String(row[3] || '').trim().toUpperCase(),
        produto: String(row[7] || '').trim(),
        descricao: String(row[8] || '').trim(),
        acaoImediata: String(row[11] || '').trim()
      });
    }
    return rncs;
  }
};

const truckStatusManager = {
  storageConfirmKey: 'cysyFaturamentoConfirmacoes',
  storageEvidenceKey: 'cysyFaturamentoEvidencias',
  manualConfirmations: new Set(),
  evidencias: {},
  placaEmFluxo: '',
  isConfirmando: false,

  chaveDia(date = new Date()) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  },

  chaveRegistro(placa, date = new Date()) {
    return `${this.chaveDia(date)}|${normalizeName(placa)}`;
  },

  init() {
    const savedConfirm = safeJsonParse(safeStorage.getItem(this.storageConfirmKey, '[]'), []);
    if (Array.isArray(savedConfirm)) {
      savedConfirm.forEach(item => this.manualConfirmations.add(String(item)));
    }
    this.evidencias = safeJsonParse(safeStorage.getItem(this.storageEvidenceKey, '{}'), {});
    this.limparHistoricoAntigo();
  },

  limparHistoricoAntigo() {
    const hoje = this.chaveDia();
    this.manualConfirmations = new Set([...this.manualConfirmations].filter(item => String(item).startsWith(`${hoje}|`)));
    const evidenciasHoje = {};
    Object.entries(this.evidencias || {}).forEach(([k, v]) => {
      if (String(k).startsWith(`${hoje}|`)) evidenciasHoje[k] = v;
    });
    this.evidencias = evidenciasHoje;
    this.persistir();
  },

  persistir() {
    safeStorage.setItem(this.storageConfirmKey, JSON.stringify([...this.manualConfirmations]));
    safeStorage.setItem(this.storageEvidenceKey, JSON.stringify(this.evidencias));
  },

  temConfirmacaoManual(placa) {
    return this.manualConfirmations.has(this.chaveRegistro(placa));
  },

  obterEvidencia(placa) {
    return this.evidencias[this.chaveRegistro(placa)] || null;
  },

  obterResumoEvidenciaHTML(placa) {
    const evid = this.obterEvidencia(placa);
    if (!evid) return '';
    return `<div style="font-size:11px; font-weight:800; color:var(--success); background:var(--success-bg); border:1px solid var(--success); border-radius:8px; padding:6px 10px; display:inline-flex; gap:6px; align-items:center;">📸 Imagem da carga pronta anexada em ${escapeHTML(evid.dataHora || '--')}</div>`;
  },

  iniciarFaturamento(placa) {
    const p = normalizeName(placa);
    if (!p || p === 'SEM PLACA') {
      alert("Não é possível enviar para faturamento sem placa válida.");
      return;
    }
    this.placaEmFluxo = p;
    const overlay = document.getElementById('faturamentoOverlay');
    const placaEl = document.getElementById('faturamentoPlaca');
    const hidden = document.getElementById('faturamentoImgBase64');
    const err = document.getElementById('faturamentoError');
    const preview = document.getElementById('faturamentoImgPreview');
    const input = document.getElementById('faturamentoImgInput');
    if (placaEl) placaEl.innerText = p;
    if (hidden) hidden.value = '';
    if (err) err.style.display = 'none';
    if (input) input.value = '';
    if (preview) {
      preview.style.backgroundImage = 'none';
      preview.classList.remove('has-image');
      preview.innerHTML = '<span>📷 Anexar imagem da carga pronta</span>';
    }
    if (overlay) overlay.style.display = 'flex';
  },

  fecharModalFaturamento() {
    const overlay = document.getElementById('faturamentoOverlay');
    if (overlay) overlay.style.display = 'none';
    this.placaEmFluxo = '';
  },

  onImagemFaturamentoChange(event) {
    const file = event?.target?.files?.[0];
    const hidden = document.getElementById('faturamentoImgBase64');
    const preview = document.getElementById('faturamentoImgPreview');
    const err = document.getElementById('faturamentoError');
    if (err) err.style.display = 'none';
    if (!file) {
      if (hidden) hidden.value = '';
      if (preview) {
        preview.style.backgroundImage = 'none';
        preview.classList.remove('has-image');
        preview.innerHTML = '<span>📷 Anexar imagem da carga pronta</span>';
      }
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        let width = img.width;
        let height = img.height;
        const max = 1280;
        if (width > height) { if (width > max) { height *= max / width; width = max; } }
        else { if (height > max) { width *= max / height; height = max; } }
        canvas.width = width;
        canvas.height = height;
        canvas.getContext('2d').drawImage(img, 0, 0, width, height);
        const dataUrl = canvas.toDataURL('image/jpeg', 0.72);
        if (hidden) hidden.value = dataUrl;
        if (preview) {
          preview.style.backgroundImage = `url(${dataUrl})`;
          preview.classList.add('has-image');
          preview.innerHTML = '<span>Imagem pronta ✔️</span>';
        }
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  },

  async confirmarFaturamentoComImagem() {
    if (this.isConfirmando) return;
    const placa = this.placaEmFluxo;
    if (!placa) return;
    const hidden = document.getElementById('faturamentoImgBase64');
    const err = document.getElementById('faturamentoError');
    const imagemCargaPronta = hidden ? hidden.value : '';
    if (!imagemCargaPronta) {
      if (err) {
        err.style.display = 'block';
        err.innerText = "Anexe a imagem da carga pronta para concluir o envio ao faturamento.";
      }
      return;
    }

    const acomp = document.getElementById('selAcomp')?.value || safeStorage.getItem('cysyUser', '');
    const dataHora = new Date().toLocaleString('pt-BR');
    const payload = {
      tipo: 'FATURAMENTO_EVIDENCIA',
      placa,
      dataHora,
      acompanhamento: acomp,
      imagemCargaPronta
    };
    const meta = dbManager.buildMeta('FATURAMENTO_EVIDENCIA', payload);
    payload.syncId = meta.syncId;
    payload.localId = meta.localId;
    payload.fingerprint = meta.fingerprint;
    if (dbManager.isSent(meta.syncId, meta.fingerprint) || dbManager.isFingerprintSent(meta.fingerprint)) {
      toastManager.show('Esse envio para faturamento ja foi registrado anteriormente.', 'info');
      return;
    }
    const dupOpen = await dbManager.findOpenDuplicate(meta.syncId, meta.fingerprint);
    if (dupOpen) {
      toastManager.show('Essa evidencia ja esta na fila de sincronizacao.', 'warning');
      syncManager.refreshSyncView();
      return;
    }

    this.isConfirmando = true;
    uiBuilder.toggleLoader(true, "Registrando envio para faturamento...");
    let enviadoOnline = false;
    try {
      const resp = await apiService.sendDataToAppScript(payload);
      enviadoOnline = !!(resp && resp.success);
    } catch (_) {
      const queued = await dbManager.savePending('FATURAMENTO_EVIDENCIA', payload, {
        syncId: meta.syncId,
        localId: meta.localId,
        fingerprint: meta.fingerprint,
        status: navigator.onLine ? 'ERRO' : 'AGUARDANDO_INTERNET'
      });
      if (!queued.ok && queued.reason === 'JA_ENVIADO') {
        enviadoOnline = true;
      }
    } finally {
      this.isConfirmando = false;
      uiBuilder.toggleLoader(false);
    }

    this.manualConfirmations.add(this.chaveRegistro(placa));
    this.evidencias[this.chaveRegistro(placa)] = {
      dataHora,
      acompanhamento: acomp,
      status: enviadoOnline ? 'ONLINE' : 'PENDENTE_SYNC'
    };
    this.persistir();
    backupManager.addEntry('FATURAMENTO_EVIDENCIA', {
      placa,
      dataHora,
      acompanhamento: acomp,
      status: enviadoOnline ? 'ONLINE' : 'PENDENTE_SYNC',
      imagemCargaPronta
    });
    this.fecharModalFaturamento();
    dashboardEngine.renderExecTable(appController.lastParsedData, appController.currentPendingPlacas);
    toastManager.show(enviadoOnline ? 'Lote enviado para faturamento com imagem anexada.' : 'Imagem vinculada e lote marcado. Envio pendente para sincronização.', enviadoOnline ? 'success' : 'warning');
    syncManager.refreshSyncView();
  },

  calcularHorarioFinal(hStr, mins) {
    if (!hStr) return null;
    const [h, m] = hStr.split(':').map(Number);
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), h, m, 0);
    const end = new Date(start.getTime() + mins * 60000);
    return { start, end, endStr: `${String(end.getHours()).padStart(2, '0')}:${String(end.getMinutes()).padStart(2, '0')}` };
  },

  verificarSeJaCarregouPlanilha(placa) {
    const p = normalizeName(placa);
    const ho = new Date();
    const r = new RegExp(`\\b${String(ho.getDate()).padStart(2, '0')}/${String(ho.getMonth() + 1).padStart(2, '0')}/${ho.getFullYear()}\\b`);
    if (!uiBuilder.historicoCargas) return false;
    for (let row of uiBuilder.historicoCargas) {
      if (!row) continue;
      const str = normalizeName(row.join(' '));
      if (r.test(str) && str.includes(p)) return true;
    }
    return false;
  },

  definirStatusCaminhao(placa, hStr, qtd, isPallet, pending) {
    if (this.verificarSeJaCarregouPlanilha(placa) || this.temConfirmacaoManual(placa) || pending) {
      return { code: 'FATURAMENTO', label: '✅ ENVIADO', msg: 'Processo finalizado.' };
    }
    if (!hStr) {
      return { code: 'AGUARDANDO', label: 'Aguardando', msg: 'Horário não programado.' };
    }
    const tempos = this.calcularHorarioFinal(hStr, (qtd * 120) + (isPallet ? qtd * 5 : 0));
    if (!tempos) return null;
    const now = new Date();
    if (now >= tempos.start && now <= tempos.end) {
      return { code: 'CARREGANDO', label: '🚛 CARREGANDO', msg: `Término previsto às ${tempos.endStr}` };
    }
    if (now > tempos.end) {
      return { code: 'ATRASADO', label: '⏳ EM ATRASO', msg: `Deveria ter finalizado às ${tempos.endStr}` };
    }
    return { code: 'AGENDADO', label: '⏰ AGENDADO', msg: `Início previsto às ${hStr}` };
  },

  confirmarCarregamentoManual(placa) {
    this.iniciarFaturamento(placa);
  }
};

