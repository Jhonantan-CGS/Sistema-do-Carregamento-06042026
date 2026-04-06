const mapOperacionalV2 = {
  version: '20260330r21',
  expandedLote: null,
  infoPanel: null,
  infoPanelBackdrop: null,
  initialized: false,

  init() {
    if (this.initialized) return;
    this.createInfoPanelElements();
    this.setupGlobalClickHandler();
    this.setupKeyboardShortcuts();
    this.initialized = true;
    console.log('[MapV2] Inicializado', this.version);
  },

  findStockInfo(normalizedLote) {
    if (typeof uiBuilder === 'undefined') return null;
    const source = uiBuilder.allLotesFlat || uiBuilder.lotesFlat || [];
    return source.find(l => normalizeName(l.lote) === normalizedLote) || null;
  },

  findLoteLocation(lote) {
    if (!lote) return null;
    const parsed = mapController.parseLocationEntry(mapController.locDB[lote]);
    if (parsed) {
      return { lat: parsed.lat, lon: parsed.lon, description: parsed.description };
    }
    const logEntry = mapController.locationLog?.[lote];
    if (logEntry && logEntry.lat && logEntry.lon) {
      return { lat: logEntry.lat, lon: logEntry.lon, description: logEntry.description };
    }
    return null;
  },

  getLoteData(lote) {
    if (!lote) return null;
    const normalizedLote = normalizeName(lote);
    const stockInfo = this.findStockInfo(normalizedLote);
    const locationData = this.findLoteLocation(lote);
    
    const mapped = mapController.lastRenderedMapped || [];
    const mappedData = mapped.find(m => normalizeName(m.lote) === normalizedLote);
    
    const produto = stockInfo?.produto || mappedData?.produto || 'Material não identificado';
    const saldo = Number(stockInfo?.saldo ?? mappedData?.saldo ?? 0);
    
    return {
      lote: lote,
      produto: produto,
      saldo: saldo,
      active: saldo > 0,
      lat: locationData?.lat,
      lon: locationData?.lon,
      description: locationData?.description || '',
      hasLocation: Boolean(locationData)
    };
  },

  getRelatedConnections(lote) {
    if (!lote || typeof mapController === 'undefined') return [];
    const normalizedLote = normalizeName(lote);
    const mapped = Array.isArray(mapController.lastRenderedMapped) ? mapController.lastRenderedMapped : [];
    const proximity = mapController.lastRenderedProximity?.edges?.length
      ? mapController.lastRenderedProximity
      : (typeof mapController.getProximityWeb === 'function' ? mapController.getProximityWeb(mapped) : { edges: [] });
    const edges = Array.isArray(proximity?.edges) ? proximity.edges : [];

    return edges.map((edge) => {
      const fromPoint = mapped[edge.from];
      const toPoint = mapped[edge.to];
      if (!fromPoint || !toPoint) return null;
      const fromKey = normalizeName(fromPoint.lote);
      const toKey = normalizeName(toPoint.lote);
      if (fromKey !== normalizedLote && toKey !== normalizedLote) return null;
      const otherPoint = fromKey === normalizedLote ? toPoint : fromPoint;
      return {
        lote: otherPoint.lote,
        produto: otherPoint.produto || 'Material não identificado',
        distance: Number(edge.distance || 0),
        direction: edge.direction || '',
        label: typeof mapController.getConnectionLabel === 'function'
          ? mapController.getConnectionLabel(edge)
          : `${Math.round(Number(edge.distance || 0))} m`
      };
    }).filter(Boolean).sort((a, b) => a.distance - b.distance);
  },

  createInfoPanelElements() {
    const compactMobile = typeof mapController !== 'undefined' && typeof mapController.isCompactMobile === 'function'
      ? mapController.isCompactMobile()
      : false;
    const existing = document.getElementById('map-operacional-info-backdrop');
    if (existing) existing.remove();
    const existingPanel = document.getElementById('map-operacional-info-panel');
    if (existingPanel) existingPanel.remove();

    const backdrop = document.createElement('div');
    backdrop.id = 'map-operacional-info-backdrop';
    Object.assign(backdrop.style, {
      position: 'fixed', top: '0', left: '0', right: '0', bottom: '0',
      background: 'rgba(0, 0, 0, 0.55)', zIndex: '10000',
      opacity: '0', visibility: 'hidden', transition: compactMobile ? 'opacity 0.14s ease' : 'opacity 0.25s, visibility 0.25s'
    });
    backdrop.addEventListener('click', () => this.closeInfoPanel());
    document.body.appendChild(backdrop);
    this.infoPanelBackdrop = backdrop;

    const panel = document.createElement('div');
    panel.id = 'map-operacional-info-panel';
    Object.assign(panel.style, {
      position: 'fixed', top: '50%', left: '50%',
      transform: 'translate(-50%, -50%) scale(0.92)',
      background: 'white', borderRadius: '18px', padding: '0',
      minWidth: compactMobile ? '300px' : '340px', maxWidth: compactMobile ? '400px' : '440px', width: '90vw',
      boxShadow: compactMobile ? '0 10px 28px rgba(15, 23, 42, 0.18)' : '0 24px 80px rgba(0, 0, 0, 0.35)',
      zIndex: '10001', opacity: '0', visibility: 'hidden',
      transition: compactMobile ? 'opacity 0.14s ease, transform 0.14s ease' : 'opacity 0.25s, transform 0.25s, visibility 0.25s',
      fontFamily: 'Inter, -apple-system, BlinkMacSystemFont, sans-serif',
      overflow: 'hidden'
    });
    panel.innerHTML = `
      <div id="info-panel-header" style="background:linear-gradient(135deg,#0F172A,#1E40AF);padding:20px 24px;color:#fff;">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;">
          <div>
            <div id="info-panel-lote" style="font-size:22px;font-weight:800;margin-bottom:2px;"></div>
            <div id="info-panel-label" style="font-size:12px;opacity:0.7;"></div>
          </div>
          <button id="info-panel-close" type="button" style="background:rgba(255,255,255,0.15);border:none;color:#fff;font-size:20px;cursor:pointer;border-radius:8px;width:36px;height:36px;display:flex;align-items:center;justify-content:center;line-height:1;">&times;</button>
        </div>
      </div>
      <div id="info-panel-body" style="padding:20px 24px;"></div>
    `;
    document.body.appendChild(panel);
    this.infoPanel = panel;
    panel.querySelector('#info-panel-close').addEventListener('click', () => this.closeInfoPanel());
  },

  openInfoPanel(data) {
    if (!this.infoPanel) this.createInfoPanelElements();

    const loteEl = this.infoPanel.querySelector('#info-panel-lote');
    const labelEl = this.infoPanel.querySelector('#info-panel-label');
    const bodyEl = this.infoPanel.querySelector('#info-panel-body');

    loteEl.textContent = data.lote || 'Lote';
    labelEl.textContent = data.label || '';

    const saldoFormatado = Number(data.saldo || 0).toLocaleString('pt-BR');
    const saldoCor = Number(data.saldo) > 0 ? '#16A34A' : '#DC2626';
    const statusTexto = Number(data.saldo) > 0 ? 'Ativo' : 'Esgotado';
    const statusCor = Number(data.saldo) > 0 ? '#16A34A' : '#DC2626';
    const relatedConnections = this.getRelatedConnections(data.lote);
    const connectionsHtml = relatedConnections.length > 0
      ? relatedConnections.slice(0, 5).map((item) => `
        <div style="display:flex;justify-content:space-between;gap:10px;padding:10px 12px;background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;">
          <div>
            <div style="font-size:13px;font-weight:800;color:#0F172A;">${escapeHTML(item.lote)}</div>
            <div style="font-size:12px;color:#475569;line-height:1.4;">${escapeHTML(item.produto)}</div>
          </div>
          <div style="text-align:right;min-width:90px;">
            <div style="font-size:13px;font-weight:800;color:#0369A1;">${escapeHTML(item.label)}</div>
            <div style="font-size:11px;color:#64748B;">${escapeHTML(item.direction || 'sem direção')}</div>
          </div>
        </div>
      `).join('')
      : `<div style="padding:12px;background:#F8FAFC;border:1px dashed #CBD5E1;border-radius:10px;font-size:12px;color:#64748B;">Sem conexões próximas suficientes para montar a teia deste lote agora.</div>`;

    bodyEl.innerHTML = `
      <div style="margin-bottom:16px;">
        <div style="font-size:11px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:6px;font-weight:600;">Produto</div>
        <div style="font-size:17px;font-weight:700;color:#0F172A;line-height:1.3;">${escapeHTML(data.produto || 'Não identificado')}</div>
      </div>
      <div style="display:flex;gap:14px;margin-bottom:16px;">
        <div style="flex:1;padding:14px;background:#F1F5F9;border-radius:10px;">
          <div style="font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:4px;font-weight:600;">Saldo disponível</div>
          <div style="font-size:26px;font-weight:800;color:${saldoCor};">${saldoFormatado} <span style="font-size:14px;font-weight:600;">t</span></div>
        </div>
        <div style="flex:0.7;padding:14px;background:#F1F5F9;border-radius:10px;">
          <div style="font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:4px;font-weight:600;">Status</div>
          <div style="font-size:15px;font-weight:700;color:${statusCor};">${statusTexto}</div>
        </div>
      </div>
      ${data.description ? `<div style="padding:12px;background:#FEF3C7;border-radius:10px;font-size:13px;color:#92400E;line-height:1.5;margin-bottom:16px;">${escapeHTML(data.description)}</div>` : ''}
      <div style="margin-bottom:16px;">
        <div style="font-size:11px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:8px;font-weight:700;">Teia de proximidade</div>
        <div style="display:grid;gap:8px;">${connectionsHtml}</div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
        <button type="button" id="info-panel-navigate" data-lat="${data.lat}" data-lon="${data.lon}" data-lote="${escapeHTML(data.lote)}" style="padding:14px;background:linear-gradient(135deg,#2563EB,#1D4ED8);color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;transition:transform 0.1s;">
          Ir até o local
        </button>
        <button type="button" id="info-panel-delete" data-lote="${escapeHTML(data.lote)}" style="padding:14px;background:rgba(220,38,38,0.08);color:#B91C1C;border:1px solid rgba(220,38,38,0.26);border-radius:10px;font-size:14px;font-weight:800;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;">
          Apagar do mapa
        </button>
      </div>
    `;

    const navBtn = bodyEl.querySelector('#info-panel-navigate');
    navBtn.addEventListener('click', () => {
      this.navigateToLocation(data.lote);
    });
    const deleteBtn = bodyEl.querySelector('#info-panel-delete');
    deleteBtn?.addEventListener('click', () => {
      const deleted = mapController.deleteSingleLote(data.lote);
      if (deleted) this.closeInfoPanel();
    });

    this.infoPanel.style.opacity = '1';
    this.infoPanel.style.visibility = 'visible';
    this.infoPanel.style.transform = 'translate(-50%, -50%) scale(1)';
    this.infoPanelBackdrop.style.opacity = '1';
    this.infoPanelBackdrop.style.visibility = 'visible';
    this.expandedLote = data.lote;
  },

  closeInfoPanel() {
    if (this.infoPanel) {
      this.infoPanel.style.opacity = '0';
      this.infoPanel.style.visibility = 'hidden';
      this.infoPanel.style.transform = 'translate(-50%, -50%) scale(0.92)';
    }
    if (this.infoPanelBackdrop) {
      this.infoPanelBackdrop.style.opacity = '0';
      this.infoPanelBackdrop.style.visibility = 'hidden';
    }
    this.expandedLote = null;
  },

  showLoteDetails(lote) {
    if (!lote) return;
    const data = this.getLoteData(lote);
    if (!data) {
      toastManager.show('Dados do lote não encontrados', 'error');
      return;
    }
    const mapped = mapController.lastRenderedMapped || [];
    const layoutPoints = mapController.buildLiveLayoutPoints ? mapController.buildLiveLayoutPoints(mapped) : [];
    const layoutItem = layoutPoints.find(p => normalizeName(p.lote) === normalizeName(lote));
    data.label = layoutItem?.label || mapController.getMarkerLabel(lote);
    this.openInfoPanel(data);
  },

  navigateToLocation(lote) {
    if (!lote) return;
    mapController.selectLote(lote, { center: true, zoom: 19 });
    const mapped = mapController.lastRenderedMapped || [];
    const item = mapped.find(m => normalizeName(m.lote) === normalizeName(lote));
    if (item && item.lat && item.lon) {
      toastManager.show('Lote selecionado no mapa. Use Google Maps para navegação.', 'info');
    }
    this.closeInfoPanel();
  },

  setupGlobalClickHandler() {
    document.addEventListener('click', (e) => {
      const infoBtn = e.target.closest('.map-live-badge-info');
      if (infoBtn) {
        e.stopPropagation();
        e.preventDefault();
        const lote = infoBtn.getAttribute('data-map-info');
        if (lote) {
          this.showLoteDetails(lote);
        }
      }
    });
  },

  setupKeyboardShortcuts() {
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && this.expandedLote) {
        this.closeInfoPanel();
      }
    });
  },

  finalize() {
    this.init();
    console.log('[MapV2] Finalizado', this.version);
  }
};

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => mapOperacionalV2.finalize());
} else {
  mapOperacionalV2.finalize();
}
