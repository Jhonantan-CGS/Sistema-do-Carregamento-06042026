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
    const rawPlate = String(placa || '').trim();
    if (!rawPlate || rawPlate.toUpperCase() === 'SEM PLACA' || (typeof plateRules !== 'undefined' && !plateRules.isValid(rawPlate))) {
      alert("Não é possível enviar para faturamento sem placa válida.");
      return;
    }
    const p = typeof plateRules !== 'undefined' ? plateRules.normalize(rawPlate) : normalizeName(rawPlate);
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

const dashboardEngine = {
  limparNomesGenericos(str) {
    const genericos = ['LTDA', 'S/A', 'SA', 'ME', 'EPP', 'TRANSPORTES', 'TRANSPORTADORA', 'LOGISTICA', 'LOGISTICA', 'COMERCIO', 'COMERCIAL', 'AGROPECUARIA', 'AGRO', 'BRASIL', 'CIA', 'EIRELI'];
    return (str || '').split(/[\s,.\-]+/).filter((w) => w && !genericos.includes(w)).join(' ').trim();
  },

  renderTopBanners(groups) {
    const bannerContainer = document.getElementById('bannerContainer');
    if (!bannerContainer) return;

    const bloqueados = groups.filter((g) => g.bloqueado);
    const atrasados = groups.filter((g) => g.statusInteligente?.code === 'ATRASADO');
    const comRnc = groups.filter((g) => g.rncMatches?.length > 0);

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
    const bannerContainer = document.getElementById('bannerContainer');
    if (!container) return;

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const hojeStr = hoje.toLocaleDateString('pt-BR');

    const dadosHoje = [];
    const dadosFora = [];
    (parsedData || []).forEach((item) => {
      const parsedDate = parseSheetDate(item.dataRaw);
      if (isSameDay(parsedDate, hoje)) {
        dadosHoje.push(item);
      } else {
        dadosFora.push(item);
      }
    });

    if (dadosFora.length > 0 && dadosHoje.length === 0) {
      dadosHoje.push(...dadosFora);
    }

    if (!dadosHoje.length) {
      if (bannerContainer) bannerContainer.innerHTML = '';
      container.innerHTML = `
        <div style="text-align:center; padding:48px 24px; background:#FFFFFF; border:2px solid var(--warning); border-radius:var(--radius-lg); box-shadow:var(--shadow-soft);">
          <div style="font-size:54px; margin-bottom:12px;">📭</div>
          <div style="font-family:'Montserrat'; font-weight:900; font-size:18px; color:var(--primary); margin-bottom:8px;">Nenhum carregamento para hoje (${hojeStr})</div>
          <div style="font-size:14px; color:#475569; margin-bottom:16px;">Verifique o terminal de debug (🐞) para ver todos os registros.</div>
          <div style="display:flex; flex-direction:column; gap:12px; align-items:center;">
            <button onclick="dashboardEngine.mostrarTodosRegistros()" style="background:var(--secondary); color:white; padding:12px 24px; border-radius:8px; font-weight:800; font-size:13px;">📋 MOSTRAR TODOS OS REGISTROS</button>
          </div>
        </div>`;
      return;
    }

    const grupos = {};
    dadosHoje.forEach((item) => {
      const key = item.placa || 'SEM PLACA';
      if (!grupos[key]) grupos[key] = [];
      grupos[key].push(item);
    });
    const pendingSet = new Set((pendingPlacas || []).map((p) => normalizeName(p)));

    const arrayGrupos = Object.entries(grupos).map(([placa, itens]) => {
      const obsUnificado = [...new Set(itens.map((i) => i.obsStr).filter(Boolean))].join(' | ');
      const timeStr = extrairHorarioObs(obsUnificado);
      let timeMins = 9999;
      if (timeStr) {
        const [h, m] = timeStr.split(':').map(Number);
        timeMins = (h * 60) + m;
      }

      const totalQtd = itens.reduce((sum, i) => sum + Number(i.qtd || 0), 0);
      const bloqueado = itens.some((i) => i.bloqueio && i.bloqueio.trim() !== '' && i.bloqueio.toUpperCase() !== 'LIBERADO');
      let statusInteligente = truckStatusManager.definirStatusCaminhao(
        placa,
        timeStr,
        totalQtd,
        itens.some((i) => String(i.pallet).toUpperCase() === 'SIM'),
        pendingSet.has(normalizeName(placa))
      );

      if (bloqueado) {
        statusInteligente = { code: 'BLOQUEADO', label: '⛔ BLOQUEADO', msg: 'Há bloqueio informado na programação.' };
      }

      let rncMatches = [];
      if (Array.isArray(uiBuilder.rncCargas)) {
        const clientesStr = [...new Set(itens.map((i) => i.cliente))].join(' • ');
        const bc = this.limparNomesGenericos(clientesStr.toUpperCase());
        const bo = this.limparNomesGenericos(obsUnificado.toUpperCase());
        uiBuilder.rncCargas.forEach((rnc) => {
          const rC = this.limparNomesGenericos(rnc.cliente);
          const rT = this.limparNomesGenericos(rnc.transportador);
          const mC = rC.length >= 3 && bc.includes(rC);
          const mT = rT.length >= 3 && bo.includes(rT);
          const pL = placa.toUpperCase().replace(/[^A-Z0-9]/g, '');
          const mP = pL.length >= 7 && (
            rnc.transportador.replace(/[^A-Z0-9]/g, '').includes(pL) ||
            rnc.descricao.toUpperCase().replace(/[^A-Z0-9]/g, '').includes(pL)
          );
          const tr = [];
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

      return { placa, itens, timeMins, timeStr, statusInteligente, bloqueado, rncMatches };
    }).sort((a, b) => {
      const prioridade = { BLOQUEADO: 0, ATRASADO: 1, CARREGANDO: 2, AGENDADO: 3, AGUARDANDO: 4, FATURAMENTO: 5 };
      return (prioridade[a.statusInteligente?.code] ?? 99) - (prioridade[b.statusInteligente?.code] ?? 99) || a.timeMins - b.timeMins;
    });

    this.renderTopBanners(arrayGrupos);

    let html = '<div class="carga-cards-grid">';
    arrayGrupos.forEach(({ placa, itens, timeStr, statusInteligente, bloqueado, rncMatches }) => {
      const allClientes = [...new Set(itens.map((i) => i.cliente))].join(' • ');
      const badgesCarga = [];
      const statusColor = bloqueado
        ? '#B91C1C'
        : statusInteligente.code === 'ATRASADO'
          ? '#B91C1C'
          : statusInteligente.code === 'CARREGANDO'
            ? '#1D4ED8'
            : statusInteligente.code === 'FATURAMENTO'
              ? '#047857'
              : '#B45309';
      if (itens.some((i) => String(i.estado || '').toUpperCase() === 'PY')) badgesCarga.push('<span class="badge-especial badge-py">🇵🇾 PY</span>');
      if (itens.some((i) => String(i.pallet || '').toUpperCase() === 'SIM')) badgesCarga.push('<span class="badge-especial badge-palete">🟫 PALETE</span>');
      if (itens.some((i) => String(i.filme || '').toUpperCase() === 'SIM')) badgesCarga.push('<span class="badge-especial badge-filme">🎞️ FILME</span>');

      let timeBadgeClass = 'time-badge';
      if (statusInteligente.code === 'AGENDADO') timeBadgeClass += ' alert';
      if (statusInteligente.code === 'ATRASADO' || statusInteligente.code === 'BLOQUEADO') timeBadgeClass += ' late';
      const statusProgress = bloqueado
        ? 16
        : ({
            ATRASADO: 32,
            AGENDADO: 46,
            AGUARDANDO: 58,
            CARREGANDO: 78,
            FATURAMENTO: 100
          }[statusInteligente.code] ?? 52);
      const progressTone = bloqueado
        ? 'danger'
        : ({
            ATRASADO: 'danger',
            AGENDADO: 'warning',
            AGUARDANDO: 'neutral',
            CARREGANDO: 'info',
            FATURAMENTO: 'success'
          }[statusInteligente.code] ?? 'neutral');

      const rncHTML = rncMatches.length ? `
        <div style="background: rgba(239, 68, 68, 0.10); border: 1px solid rgba(239, 68, 68, 0.24); border-left: 6px solid #EF4444; padding: 16px; border-radius: 12px; color: #FEE2E2;">
          <details>
            <summary style="cursor:pointer; font-weight:900; color:#991B1B; font-size:14px;">⚠️ ${rncMatches.length} alerta(s) de qualidade</summary>
            <div style="margin-top:16px; display:flex; flex-direction:column; gap:12px;">
              ${rncMatches.map((r) => `
                <div style="background: rgba(12, 14, 20, 0.82); border:1px solid rgba(239, 68, 68, 0.18); border-radius:10px; padding:12px; font-size:12px; color:#F8FAFC; line-height:1.5;">
                  <strong>📅 ${escapeHTML(r.data)}</strong><br>
                  <strong>Alvo:</strong> ${escapeHTML(r.trigger)}<br>
                  <strong>Produto:</strong> ${escapeHTML(r.prod)}<br>
                  <strong>Ocorrência:</strong> ${escapeHTML(r.desc)}<br>
                  <strong>Ação imediata:</strong> ${escapeHTML(r.acao)}
                </div>`).join('')}
            </div>
          </details>
        </div>` : '';

      const produtosHTML = itens.map((item) => {
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
                : '<span class="badge-verde">✔ OK</span>'}
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
              <div style="font-weight:900; font-size:14px; color:${statusColor};">
                ${statusInteligente.label}
              </div>
              ${statusInteligente.code !== 'FATURAMENTO' && statusInteligente.code !== 'BLOQUEADO' ? `<button class="btn-status-manual" onclick="truckStatusManager.iniciarFaturamento('${escapeHTML(placa)}')">📤 ENVIAR P/ FATURAMENTO</button>` : ''}
            </div>
            <div style="font-size:13px; font-weight:700; color:#B8C1D1;">${escapeHTML(statusInteligente.msg)}</div>
            <div class="status-progress status-progress-${progressTone}">
              <div class="status-progress-copy">
                <span>Andamento operacional</span>
                <strong>${statusProgress}%</strong>
              </div>
              <div class="status-progress-track">
                <span class="status-progress-bar" style="width:${statusProgress}%"></span>
              </div>
            </div>
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

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const hojeStr = hoje.toLocaleDateString('pt-BR');

    let html = `<div style="margin-bottom:16px; padding:12px; background:#FEF3C7; border-radius:8px; border:1px solid #F59E0B;">
      <div style="font-weight:800; color:#92400E; margin-bottom:8px;">⚠️ Modo Debug: Mostrando todos os registros (sem filtro de data)</div>
      <div style="font-size:12px; color:#78350F;">Hoje: ${hojeStr} | Total: ${appController.lastParsedData.length} registros</div>
      <button onclick="appController.handleRefresh(true)" style="margin-top:12px; background:var(--secondary); color:white; padding:8px 16px; border-radius:6px; font-weight:800; font-size:12px;">🔄 ATUALIZAR</button>
    </div>`;

    const grupos = {};
    appController.lastParsedData.forEach((item) => {
      const key = item.placa || 'SEM PLACA';
      if (!grupos[key]) grupos[key] = [];
      grupos[key].push(item);
    });

    html += '<div class="carga-cards-grid">';
    Object.entries(grupos).forEach(([placa, itens]) => {
      const allClientes = [...new Set(itens.map((i) => i.cliente))].join(' • ');
      const parsedDate = parseSheetDate(itens[0].dataRaw);
      const dateStr = parsedDate ? parsedDate.toLocaleDateString('pt-BR') : 'DATA INVÁLIDA';
      html += `
        <div class="carga-card liberado">
          <div class="carga-card-header">
            <div class="carga-placa">🚛 ${escapeHTML(placa)}</div>
            <div class="carga-cliente"><div>${escapeHTML(allClientes)}</div></div>
            <div class="carga-card-status"><span class="time-badge">📅 ${dateStr}</span></div>
          </div>
          <div class="carga-card-body">
            <div style="font-size:13px; color:#B8C1D1;">
              Total: ${itens.reduce((s, i) => s + Number(i.qtd || 0), 0).toFixed(3)}t | ${itens.length} pedido(s)
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
