(function(){
  const state = { stops: [], map: null, layerGroup: null, order: null, trip: null, delivered: [], userLocation: null };
  const STORAGE_KEY = 'rutaOptima_state_v1';
  const MAP_EMPTY_DEFAULT_TEXT = 'El mapa y la ruta optimizada van a aparecer acá una vez que cargues las direcciones.';

  const dropzone = document.getElementById('rp-dropzone');
  const fileInput = document.getElementById('rp-file-input');
  const stopListEl = document.getElementById('rp-stop-list');
  const optimizeBtn = document.getElementById('rp-optimize-btn');
  const resetBtn = document.getElementById('rp-reset-btn');
  const errorBanner = document.getElementById('rp-error-banner');
  const resultsEl = document.getElementById('rp-results');
  const mapEmpty = document.getElementById('rp-map-empty');
  const mapLoading = document.getElementById('rp-map-loading');
  const mapLoadingText = document.getElementById('rp-map-loading-text');
  let mapLoadingShownAt = 0;
  const useLocationCheckbox = document.getElementById('rp-use-location');

  const scanOpenBtn = document.getElementById('rp-scan-open-btn');
  const scanPanel = document.getElementById('rp-scan-panel');
  const scanVideo = document.getElementById('rp-scan-video');
  const scanCanvas = document.getElementById('rp-scan-canvas');
  const scanLiveControls = document.getElementById('rp-scan-live-controls');
  const scanCaptureBtn = document.getElementById('rp-scan-capture-btn');
  const scanCancelBtn = document.getElementById('rp-scan-cancel-btn');
  const scanOcrLoading = document.getElementById('rp-scan-ocr-loading');
  const scanResult = document.getElementById('rp-scan-result');
  const scanDetectedLabel = document.getElementById('rp-scan-detected-label');
  const scanEditInput = document.getElementById('rp-scan-edit');
  const scanRawText = document.getElementById('rp-scan-raw-text');
  const scanConfirmBtn = document.getElementById('rp-scan-confirm-btn');
  const scanRetryBtn = document.getElementById('rp-scan-retry-btn');

  function showMapLoading(text){
    mapLoadingText.textContent = text || 'Calculando la mejor ruta...';
    mapLoading.classList.add('show');
    mapLoadingShownAt = Date.now();
  }
  function hideMapLoading(){
    const MIN_VISIBLE_MS = 800;
    const elapsed = Date.now() - mapLoadingShownAt;
    const wait = Math.max(0, MIN_VISIBLE_MS - elapsed);
    return new Promise(resolve => {
      setTimeout(() => {
        mapLoading.classList.remove('show');
        resolve();
      }, wait);
    });
  }

  function saveState(){
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        stops: state.stops,
        order: state.order,
        trip: state.trip,
        delivered: state.delivered
      }));
    } catch (e) { /* localStorage no disponible en este entorno, sin problema */ }
  }

  function loadState(){
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const data = JSON.parse(raw);
      if (data.stops && data.stops.length) {
        state.stops = data.stops;
        renderStopList();
        checkReady();
      }
      state.delivered = data.delivered || [];
      if (data.order && data.trip && data.order.length >= 2) {
        state.order = data.order;
        state.trip = data.trip;
        renderResults(data.trip, data.order);
        renderMap(data.trip, data.order);
      } else if (data.order && data.order.length === 1) {
        state.order = data.order;
        renderSingleStop(data.order[0]);
      } else if (state.delivered.length && data.order && data.order.length === 0) {
        renderRouteComplete();
      }
      renderDeliveredList();
    } catch (e) { /* nada guardado o corrupto, arranca de cero */ }
  }

  resetBtn.addEventListener('click', () => {
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) {}
    state.stops = [];
    state.order = null;
    state.trip = null;
    state.delivered = [];
    fileInput.value = '';
    stopListEl.innerHTML = '';
    resultsEl.classList.remove('show');
    clearError();
    hideMapLoading();
    if (state.layerGroup) state.layerGroup.clearLayers();
    mapEmpty.style.display = 'flex';
    mapEmpty.textContent = MAP_EMPTY_DEFAULT_TEXT;
    document.getElementById('rp-details-upload').setAttribute('open', '');
    document.getElementById('rp-details-nav').removeAttribute('open');
    renderDeliveredList();
    checkReady();
  });

  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('drag'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag'));
  dropzone.addEventListener('drop', e => {
    e.preventDefault();
    dropzone.classList.remove('drag');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', e => {
    if (e.target.files.length) handleFile(e.target.files[0]);
  });

  function showError(msg){
    errorBanner.textContent = msg;
    errorBanner.classList.add('show');
  }
  function clearError(){
    errorBanner.classList.remove('show');
    errorBanner.textContent = '';
  }

  function handleFile(file){
    clearError();
    resultsEl.classList.remove('show');
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        let addresses = rows
          .map(r => (r && r[0] !== undefined ? String(r[0]).trim() : ''))
          .filter(a => a.length > 0);

        if (addresses.length && /direcci|address|domicilio/i.test(addresses[0])) {
          addresses = addresses.slice(1);
        }
        if (!addresses.length) {
          showError('No encontré direcciones en la primera columna del archivo.');
          return;
        }
        const spaceLeft = 80 - state.stops.length;
        if (spaceLeft <= 0) {
          showError('Ya tenés 80 direcciones cargadas, el máximo por tanda. Optimizá o empezá de nuevo antes de sumar más.');
          return;
        }
        if (addresses.length > spaceLeft) {
          showError('Encontré ' + addresses.length + ' direcciones pero solo entran ' + spaceLeft + ' más (tope de 80 por tanda) — el resto se ignora por ahora.');
          addresses = addresses.slice(0, spaceLeft);
        }

        const newStops = addresses.map((addr, i) => ({
          id: 'stop-xlsx-' + Date.now() + '-' + i,
          raw: addr,
          lat: null, lon: null,
          status: 'pending'
        }));
        state.stops = state.stops.concat(newStops);
        renderStopList();
        geocodeAll(newStops);
      } catch (err) {
        showError('No pude leer el archivo. Probá exportarlo como .xlsx o .csv simple.');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function renderStopList(){
    stopListEl.innerHTML = '';
    state.stops.forEach((stop, idx) => {
      const li = document.createElement('li');
      li.className = 'rp-stop';

      const dot = document.createElement('div');
      dot.className = 'rp-dot ' + (stop.status === 'ok' ? 'ok' : stop.status === 'error' ? 'err' : 'pending');

      const addrWrap = document.createElement('div');
      addrWrap.className = 'rp-stop-addr';
      const rawEl = document.createElement('span');
      rawEl.className = 'raw';
      rawEl.textContent = stop.raw;
      const statusEl = document.createElement('span');
      statusEl.className = 'status' + (stop.status === 'error' ? ' err' : '');
      statusEl.textContent = stop.status === 'ok' ? 'Ubicada' : stop.status === 'error' ? 'No se pudo ubicar' : 'Buscando...';
      addrWrap.appendChild(rawEl);
      addrWrap.appendChild(statusEl);

      if (stop.status === 'error') {
        const retryWrap = document.createElement('div');
        retryWrap.className = 'rp-retry';
        const input = document.createElement('input');
        input.type = 'text';
        input.placeholder = 'Corregí la dirección...';
        input.value = stop.raw;
        const btn = document.createElement('button');
        btn.className = 'rp-btn-mini';
        btn.textContent = 'Reintentar';
        btn.addEventListener('click', async () => {
          stop.raw = input.value.trim();
          stop.status = 'pending';
          renderStopList();
          await geocodeOne(stop);
          renderStopList();
          checkReady();
        });
        retryWrap.appendChild(input);
        retryWrap.appendChild(btn);
        addrWrap.appendChild(retryWrap);
      }

      li.appendChild(dot);
      li.appendChild(addrWrap);
      stopListEl.appendChild(li);
    });
  }

  async function geocodeOne(stop){
    try {
      const q = /argentina|buenos aires|caba|amba/i.test(stop.raw) ? stop.raw : stop.raw + ', Buenos Aires, Argentina';
      const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=ar&q=' + encodeURIComponent(q);
      const res = await fetch(url, { headers: { 'Accept-Language': 'es' } });
      const json = await res.json();
      if (json && json.length) {
        stop.lat = parseFloat(json[0].lat);
        stop.lon = parseFloat(json[0].lon);
        stop.status = 'ok';
      } else {
        stop.status = 'error';
      }
    } catch (e) {
      stop.status = 'error';
    }
  }

  async function geocodeAll(stopsToGeocode){
    const list = stopsToGeocode || state.stops.filter(s => s.status === 'pending');
    optimizeBtn.disabled = true;
    const total = list.length;
    let done = 0;
    for (const stop of list) {
      const remaining = total - done;
      const secsLeft = Math.round(remaining * 1.1);
      optimizeBtn.textContent = 'Ubicando ' + (done + 1) + '/' + total + '... (~' + secsLeft + 's)';
      await geocodeOne(stop);
      done++;
      renderStopList();
      if (done < total) {
        await new Promise(r => setTimeout(r, 1100)); // respetar límite de Nominatim (1 req/seg)
      }
    }
    checkReady();
    saveState();
  }

  function checkReady(){
    const ok = state.stops.filter(s => s.status === 'ok');
    if (ok.length >= 2) {
      optimizeBtn.disabled = false;
      optimizeBtn.textContent = 'Optimizar ruta (' + ok.length + ' paradas)';
    } else {
      optimizeBtn.disabled = true;
      optimizeBtn.textContent = 'Necesito al menos 2 direcciones ubicadas';
    }
  }

  // Cola separada para geocodificar direcciones agregadas de a una (por escaneo),
  // respetando igual el límite de 1 req/seg de Nominatim sin pisar una carga masiva de Excel en curso.
  let scanGeocodeChain = Promise.resolve();
  function addStopFromScan(rawAddress){
    const clean = rawAddress.trim();
    if (!clean) return;
    const stop = {
      id: 'stop-scan-' + Date.now() + '-' + Math.random().toString(36).slice(2, 7),
      raw: clean,
      lat: null, lon: null,
      status: 'pending'
    };
    state.stops.push(stop);
    renderStopList();
    checkReady();

    scanGeocodeChain = scanGeocodeChain.then(async () => {
      await geocodeOne(stop);
      renderStopList();
      checkReady();
      saveState();
      await new Promise(r => setTimeout(r, 1100));
    });
  }

  async function fetchDurationMatrix(points){
    const coordStr = points.map(p => p.lon + ',' + p.lat).join(';');
    const url = 'https://router.project-osrm.org/table/v1/driving/' + coordStr + '?annotations=duration';
    const res = await fetch(url);
    const json = await res.json();
    if (json.code !== 'Ok' || !json.durations) throw new Error('table service failed');
    return json.durations;
  }

  // Arma un primer orden razonable: en cada paso salta al punto no visitado más cercano en tiempo.
  function nearestNeighborOrder(matrix, startIdx, n){
    const visited = new Array(n).fill(false);
    const order = [startIdx];
    visited[startIdx] = true;
    let current = startIdx;
    for (let step = 1; step < n; step++) {
      let best = -1, bestTime = Infinity;
      for (let j = 0; j < n; j++) {
        if (!visited[j] && matrix[current][j] != null && matrix[current][j] < bestTime) {
          bestTime = matrix[current][j];
          best = j;
        }
      }
      if (best === -1) { for (let j = 0; j < n; j++) { if (!visited[j]) { best = j; break; } } }
      order.push(best);
      visited[best] = true;
      current = best;
    }
    return order;
  }

  // Mejora el orden probando invertir tramos si eso reduce el tiempo total (2-opt), con límite de tiempo para no colgar el navegador.
  function twoOptImprove(matrix, order, maxMs){
    const start = Date.now();
    const n = order.length;
    const time = (a, b) => matrix[a][b] == null ? 0 : matrix[a][b];
    let improved = true;
    while (improved && (Date.now() - start) < maxMs) {
      improved = false;
      for (let i = 1; i < n - 2; i++) {
        for (let k = i + 1; k < n - 1; k++) {
          const a = order[i - 1], b = order[i], c = order[k], d = order[k + 1];
          const before = time(a, b) + time(c, d);
          const after = time(a, c) + time(b, d);
          if (after < before - 0.5) {
            let lo = i, hi = k;
            while (lo < hi) { const tmp = order[lo]; order[lo] = order[hi]; order[hi] = tmp; lo++; hi--; }
            improved = true;
          }
        }
        if ((Date.now() - start) >= maxMs) break;
      }
    }
    return order;
  }

  function getUserLocation(){
    return new Promise((resolve) => {
      if (!useLocationCheckbox.checked || !navigator.geolocation) { resolve(null); return; }
      navigator.geolocation.getCurrentPosition(
        pos => resolve({ lat: pos.coords.latitude, lon: pos.coords.longitude }),
        () => resolve(null),
        { timeout: 6000 }
      );
    });
  }

  optimizeBtn.addEventListener('click', async () => {
    clearError();
    optimizeBtn.disabled = true;
    optimizeBtn.textContent = 'Calculando la mejor ruta...';
    showMapLoading('Calculando la mejor ruta...');

    const userLoc = await getUserLocation();
    const okStops = state.stops.filter(s => s.status === 'ok');
    const points = userLoc ? [{ lat: userLoc.lat, lon: userLoc.lon, raw: 'Tu ubicación actual', isStart: true }, ...okStops] : okStops;

    try {
      const matrix = await fetchDurationMatrix(points);
      let orderIdx = nearestNeighborOrder(matrix, 0, points.length);
      orderIdx = twoOptImprove(matrix, orderIdx, 1800);
      const ordered = orderIdx.map(i => points[i]);

      await recomputeRouteForOrder(ordered);
      optimizeBtn.textContent = 'Recalcular ruta';
      document.getElementById('rp-details-upload').removeAttribute('open');
      document.getElementById('rp-details-route').setAttribute('open', '');
    } catch (e) {
      showError('No pude calcular la ruta óptima. Probá de nuevo en unos segundos.');
      optimizeBtn.textContent = 'Reintentar optimización';
      await hideMapLoading();
    } finally {
      optimizeBtn.disabled = false;
    }
  });

  function renderResults(trip, ordered){
    resultsEl.classList.add('show');
    document.getElementById('rp-stat-dist').textContent = (trip.distance / 1000).toFixed(1) + ' km';
    const mins = Math.round(trip.duration / 60);
    document.getElementById('rp-stat-time').textContent = mins < 60 ? mins + ' min' : Math.floor(mins/60) + 'h ' + (mins%60) + 'm';

    const orderList = document.getElementById('rp-order-list');
    orderList.innerHTML = '';
    let dragFromIndex = null;

    ordered.forEach((p, i) => {
      const li = document.createElement('li');
      li.className = 'rp-order-item';
      li.draggable = true;
      li.style.cursor = 'grab';
      li.dataset.index = i;

      const num = document.createElement('div');
      num.className = 'rp-num';
      num.textContent = i + 1;
      const addr = document.createElement('div');
      addr.textContent = p.raw;
      addr.style.flex = '1';
      li.appendChild(num);
      li.appendChild(addr);
      li.appendChild(makeCheckBtn(p));

      li.addEventListener('dragstart', () => {
        dragFromIndex = i;
        li.style.opacity = '0.5';
      });
      li.addEventListener('dragend', () => { li.style.opacity = '1'; });
      li.addEventListener('dragover', (e) => e.preventDefault());
      li.addEventListener('drop', (e) => {
        e.preventDefault();
        if (dragFromIndex === null || dragFromIndex === i) return;
        const newOrder = ordered.slice();
        const [moved] = newOrder.splice(dragFromIndex, 1);
        newOrder.splice(i, 0, moved);
        recomputeRouteForOrder(newOrder);
      });

      orderList.appendChild(li);
    });

    renderNavLinks(ordered);
  }

  async function recomputeRouteForOrder(order){
    clearError();
    showMapLoading('Recalculando ruta...');
    try {
      const coordStr = order.map(p => p.lon + ',' + p.lat).join(';');
      const url = 'https://router.project-osrm.org/route/v1/driving/' + coordStr + '?overview=full&geometries=geojson';
      const res = await fetch(url);
      const json = await res.json();
      if (json.code !== 'Ok' || !json.routes || !json.routes.length) {
        showError('No pude recalcular con ese orden. Probá de nuevo.');
        return;
      }
      const route = json.routes[0];
      state.order = order;
      state.trip = route;
      renderResults(route, order);
      renderMap(route, order);
      saveState();
    } catch (e) {
      showError('Falló la conexión al recalcular. Probá de nuevo.');
    } finally {
      await hideMapLoading();
    }
  }

  function makeCheckBtn(point){
    const btn = document.createElement('button');
    btn.className = 'rp-btn-mini';
    btn.style.marginLeft = 'auto';
    btn.style.flexShrink = '0';
    btn.textContent = '✓ Entregado';
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      markDelivered(point);
    });
    return btn;
  }

  async function markDelivered(point){
    const current = state.order || [];
    const newOrder = current.filter(p => p !== point);

    state.delivered = state.delivered || [];
    state.delivered.unshift(point);
    const stopRef = state.stops.find(s => s.lat === point.lat && s.lon === point.lon && s.raw === point.raw);
    if (stopRef) stopRef.delivered = true;
    renderDeliveredList();

    if (newOrder.length >= 2) {
      await recomputeRouteForOrder(newOrder);
    } else if (newOrder.length === 1) {
      state.order = newOrder;
      state.trip = null;
      renderSingleStop(newOrder[0]);
      saveState();
    } else {
      state.order = [];
      state.trip = null;
      renderRouteComplete();
      saveState();
    }
  }

  async function undoDelivered(point){
    state.delivered = (state.delivered || []).filter(p => p !== point);
    const stopRef = state.stops.find(s => s.lat === point.lat && s.lon === point.lon && s.raw === point.raw);
    if (stopRef) stopRef.delivered = false;
    renderDeliveredList();

    const newOrder = (state.order || []).concat([point]);
    if (newOrder.length >= 2) {
      await recomputeRouteForOrder(newOrder);
    } else {
      state.order = newOrder;
      renderSingleStop(newOrder[0]);
      saveState();
    }
  }

  function renderDeliveredList(){
    const wrap = document.getElementById('rp-delivered-wrap');
    const list = document.getElementById('rp-delivered-list');
    const countEl = document.getElementById('rp-delivered-count');
    const delivered = state.delivered || [];
    if (!delivered.length) { wrap.style.display = 'none'; return; }
    resultsEl.classList.add('show');
    wrap.style.display = 'block';
    countEl.textContent = '(' + delivered.length + ')';
    list.innerHTML = '';
    delivered.forEach((p) => {
      const li = document.createElement('li');
      li.className = 'rp-order-item';
      const check = document.createElement('div');
      check.textContent = '✓';
      check.style.color = '#1f5c4f';
      check.style.fontWeight = '700';
      check.style.width = '22px';
      check.style.textAlign = 'center';
      const addr = document.createElement('div');
      addr.textContent = p.raw;
      addr.style.flex = '1';
      addr.style.textDecoration = 'line-through';
      addr.style.color = 'var(--ink-soft)';
      const undoBtn = document.createElement('button');
      undoBtn.className = 'rp-btn-mini';
      undoBtn.style.marginLeft = 'auto';
      undoBtn.style.flexShrink = '0';
      undoBtn.textContent = 'Deshacer';
      undoBtn.addEventListener('click', () => undoDelivered(p));
      li.appendChild(check);
      li.appendChild(addr);
      li.appendChild(undoBtn);
      list.appendChild(li);
    });
  }

  function renderSingleStop(point){
    resultsEl.classList.add('show');
    document.getElementById('rp-stat-dist').textContent = '–';
    document.getElementById('rp-stat-time').textContent = 'última parada';

    const orderList = document.getElementById('rp-order-list');
    orderList.innerHTML = '';
    const li = document.createElement('li');
    li.className = 'rp-order-item';
    const num = document.createElement('div');
    num.className = 'rp-num';
    num.textContent = '1';
    const addr = document.createElement('div');
    addr.textContent = point.raw;
    addr.style.flex = '1';
    li.appendChild(num);
    li.appendChild(addr);
    li.appendChild(makeCheckBtn(point));
    orderList.appendChild(li);

    ensureMap();
    state.layerGroup.clearLayers();
    const icon = L.divIcon({ className: '', html: '<div class="rp-marker-num">1</div>', iconSize: [26, 26], iconAnchor: [13, 13] });
    L.marker([point.lat, point.lon], { icon }).addTo(state.layerGroup).bindPopup('Última parada: ' + point.raw);
    state.map.setView([point.lat, point.lon], 15);

    const wrap = document.getElementById('rp-nav-links');
    wrap.innerHTML = '';
    const a = document.createElement('a');
    a.href = 'https://www.google.com/maps/dir/?api=1&destination=' + point.lat + ',' + point.lon + '&travelmode=driving';
    a.target = '_blank';
    a.rel = 'noopener';
    a.className = 'rp-nav-link';
    a.textContent = 'Abrir última parada en Google Maps';
    wrap.appendChild(a);
  }

  function renderRouteComplete(){
    document.getElementById('rp-stat-dist').textContent = '–';
    document.getElementById('rp-stat-time').textContent = '–';
    document.getElementById('rp-order-list').innerHTML = '';
    document.getElementById('rp-nav-links').innerHTML = '';
    if (state.layerGroup) state.layerGroup.clearLayers();
    mapEmpty.style.display = 'flex';
    mapEmpty.textContent = '¡Listo! Entregaste todos los paquetes de esta tanda. Subí un nuevo archivo para la próxima.';
  }

  function renderNavLinks(ordered){
    const wrap = document.getElementById('rp-nav-links');
    wrap.innerHTML = '';
    const stopsOnly = ordered.filter(p => !p.isStart);
    const startPoint = ordered.find(p => p.isStart) || ordered[0];
    const restStart = ordered[0].isStart ? 1 : 1;

    const chunkSize = 9; // cada tramo = origen + 8 paradas intermedias + destino = 10 puntos totales por link (el máximo que acepta Google Maps)
    let idx = 0;
    let leg = 1;
    while (idx < ordered.length - 1) {
      const origin = ordered[idx];
      const remaining = ordered.slice(idx + 1);
      const chunk = remaining.slice(0, chunkSize);
      const destination = chunk[chunk.length - 1];
      const waypoints = chunk.slice(0, -1);

      let mapsUrl = 'https://www.google.com/maps/dir/?api=1' +
        '&origin=' + origin.lat + ',' + origin.lon +
        '&destination=' + destination.lat + ',' + destination.lon +
        '&travelmode=driving';
      if (waypoints.length) {
        mapsUrl += '&waypoints=' + waypoints.map(p => p.lat + ',' + p.lon).join('|');
      }

      const a = document.createElement('a');
      a.href = mapsUrl;
      a.target = '_blank';
      a.rel = 'noopener';
      a.className = 'rp-nav-link';
      a.textContent = ordered.length > chunkSize + 1
        ? 'Abrir tramo ' + leg + ' en Google Maps (' + chunk.length + ' paradas)'
        : 'Abrir ruta completa en Google Maps';
      wrap.appendChild(a);

      idx += chunk.length;
      leg++;
    }
  }

  function ensureMap(){
    if (state.map) return;
    mapEmpty.style.display = 'none';
    state.map = L.map('rp-map');
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '&copy; OpenStreetMap contributors',
      maxZoom: 19
    }).addTo(state.map);
    state.layerGroup = L.layerGroup().addTo(state.map);
  }

  function renderMap(trip, ordered){
    ensureMap();
    state.layerGroup.clearLayers();

    ordered.forEach((p, i) => {
      const icon = L.divIcon({
        className: '',
        html: '<div class="rp-marker-num">' + (i + 1) + '</div>',
        iconSize: [26, 26],
        iconAnchor: [13, 13]
      });
      L.marker([p.lat, p.lon], { icon }).addTo(state.layerGroup).bindPopup((i + 1) + '. ' + p.raw);
    });

    if (trip.geometry && trip.geometry.coordinates) {
      const latlngs = trip.geometry.coordinates.map(c => [c[1], c[0]]);
      L.polyline(latlngs, { color: '#c4432b', weight: 4, opacity: 0.85 }).addTo(state.layerGroup);
      state.map.fitBounds(L.latLngBounds(latlngs), { padding: [24, 24] });
    } else {
      const bounds = L.latLngBounds(ordered.map(p => [p.lat, p.lon]));
      state.map.fitBounds(bounds, { padding: [24, 24] });
    }
  }

  // ===================== ESCANEAR PAQUETE (cámara + OCR) =====================

  const STREET_PREFIXES = [
    'avenida', 'av\\.?', 'calle', 'pasaje', 'pje\\.?', 'boulevard', 'bv\\.?',
    'ruta', 'autopista', 'diagonal', 'diag\\.?', 'bulevar'
  ];
  const STREET_REGEX = new RegExp(
    '(' + STREET_PREFIXES.join('|') + ')\\s+([a-záéíóúñü0-9°ª\'\\.\\s]{2,40}?)\\s*n?°?\\s*(\\d{1,6})',
    'i'
  );

  const KNOWN_LOCALITIES = [
    'CABA', 'Ciudad Autónoma de Buenos Aires', 'Capital Federal', 'Buenos Aires',
    'Vicente López', 'San Isidro', 'San Martín', 'San Fernando', 'Tigre', 'Escobar',
    'Pilar', 'Malvinas Argentinas', 'José C. Paz', 'San Miguel', 'Moreno', 'Merlo',
    'Morón', 'Ituzaingó', 'Hurlingham', 'Tres de Febrero', 'La Matanza', 'Ezeiza',
    'Esteban Echeverría', 'Almirante Brown', 'Lanús', 'Lomas de Zamora', 'Avellaneda',
    'Quilmes', 'Berazategui', 'Florencio Varela', 'San Vicente', 'La Plata',
    'Ensenada', 'Berisso'
  ];

  let ocrWorkerPromise = null;
  let cameraStream = null;

  function getOcrWorker(){
    if (!ocrWorkerPromise) {
      ocrWorkerPromise = Tesseract.createWorker('spa');
    }
    return ocrWorkerPromise;
  }

  async function openScanPanel(){
    clearError();
    scanPanel.style.display = 'block';
    scanLiveControls.style.display = 'flex';
    scanResult.style.display = 'none';
    scanOcrLoading.style.display = 'none';
    scanVideo.style.display = 'block';
    try {
      cameraStream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: 'environment' } },
        audio: false
      });
      scanVideo.srcObject = cameraStream;
    } catch (e) {
      showError('No pude acceder a la cámara. Revisá los permisos del navegador para este sitio.');
      closeScanPanel();
    }
  }

  function closeScanPanel(){
    if (cameraStream) {
      cameraStream.getTracks().forEach(t => t.stop());
      cameraStream = null;
    }
    scanPanel.style.display = 'none';
  }

  function normalizeForMatch(text){
    return text.replace(/\r/g, '').replace(/[ \t]+/g, ' ');
  }

  function detectAddressFromText(rawText){
    const text = normalizeForMatch(rawText);
    const flat = text.replace(/\n/g, ' ');

    // Estrategia 1 (más confiable): las etiquetas de Mercado Libre/Flex/Andreani/Correo
    // suelen traer un campo explícito "Dirección:" — no depende de que la calle
    // empiece con "Av./Calle/etc.", que en la mayoría de los casos ni está.
    let streetPart = null;
    const dirFieldMatch = text.match(/direcci[oó]n\s*:?\s*([^\n]+)/i);
    if (dirFieldMatch) {
      streetPart = dirFieldMatch[1].trim().replace(/\s{2,}/g, ' ');
    }

    // Estrategia 2 (respaldo): buscar prefijo de calle conocido + número.
    if (!streetPart) {
      const m = flat.match(STREET_REGEX);
      if (m) {
        const prefix = m[1].replace(/\.$/, '');
        const streetName = m[2].trim().replace(/\s{2,}/g, ' ');
        const number = m[3];
        streetPart = prefix.charAt(0).toUpperCase() + prefix.slice(1) + ' ' + streetName + ' ' + number;
      }
    }

    // Localidad: primero el campo "Barrio:" si existe, si no, buscar por nombre conocido.
    let locality = null;
    const barrioMatch = text.match(/barrio\s*:?\s*([^\n]+)/i);
    if (barrioMatch) {
      locality = barrioMatch[1].trim();
    } else {
      for (const loc of KNOWN_LOCALITIES) {
        const re = new RegExp('\\b' + loc.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i');
        if (re.test(flat)) { locality = loc; break; }
      }
    }

    if (streetPart) {
      return { found: true, address: streetPart + (locality ? ', ' + locality : '') };
    }
    // No matcheó ningún patrón: devolvemos el texto crudo para que lo edite a mano
    return { found: false, address: flat.trim().slice(0, 120) };
  }

  async function captureAndReadLabel(){
    if (!scanVideo.videoWidth) {
      showError('La cámara todavía no cargó la imagen. Esperá un segundo y volvé a intentar.');
      return;
    }

    scanLiveControls.style.display = 'none';
    scanVideo.style.display = 'none';
    scanOcrLoading.style.display = 'flex';

    // Redimensionamos: las fotos de celular son enormes y eso hace más lento (o hasta falla) el OCR.
    const MAX_DIM = 1400;
    let w = scanVideo.videoWidth;
    let h = scanVideo.videoHeight;
    if (Math.max(w, h) > MAX_DIM) {
      const scale = MAX_DIM / Math.max(w, h);
      w = Math.round(w * scale);
      h = Math.round(h * scale);
    }
    scanCanvas.width = w;
    scanCanvas.height = h;
    const ctx = scanCanvas.getContext('2d');
    ctx.drawImage(scanVideo, 0, 0, w, h);

    if (cameraStream) {
      cameraStream.getTracks().forEach(t => t.stop());
      cameraStream = null;
    }

    try {
      const worker = await getOcrWorker();
      const { data } = await worker.recognize(scanCanvas);
      const result = detectAddressFromText(data.text || '');
      scanOcrLoading.style.display = 'none';
      scanResult.style.display = 'flex';
      scanEditInput.value = result.address;
      scanRawText.value = data.text || '(vacío)';
      scanDetectedLabel.textContent = result.found
        ? '📍 Dirección detectada'
        : '🤔 No pude reconocer el patrón — revisá/completá a mano';
      scanEditInput.focus();
    } catch (e) {
      // No cerramos el panel: dejamos que escriba la dirección a mano sobre la misma foto que sacó,
      // así no pierde el paso aunque el OCR haya fallado.
      scanOcrLoading.style.display = 'none';
      scanResult.style.display = 'flex';
      scanEditInput.value = '';
      scanRawText.value = '(no se pudo leer nada — error de OCR)';
      scanDetectedLabel.textContent = '⚠️ Falló la lectura automática — escribí la dirección a mano';
      showError('Detalle técnico del error de OCR: ' + (e && e.message ? e.message : e));
      scanEditInput.focus();
    }
  }

  scanOpenBtn.addEventListener('click', openScanPanel);
  scanCancelBtn.addEventListener('click', closeScanPanel);
  scanCaptureBtn.addEventListener('click', captureAndReadLabel);

  scanRetryBtn.addEventListener('click', () => {
    scanResult.style.display = 'none';
    openScanPanel();
  });

  scanConfirmBtn.addEventListener('click', () => {
    const addr = scanEditInput.value.trim();
    if (!addr) return;
    addStopFromScan(addr);
    scanResult.style.display = 'none';
    openScanPanel(); // vuelve a la cámara para escanear el próximo paquete
  });

  loadState();
})();