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
  const useLocationCheckbox = document.getElementById('rp-use-location');

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
    if (state.layerGroup) state.layerGroup.clearLayers();
    mapEmpty.style.display = 'flex';
    mapEmpty.textContent = MAP_EMPTY_DEFAULT_TEXT;
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
        if (addresses.length > 30) {
          showError('Encontré ' + addresses.length + ' direcciones. Funciona mejor con hasta 30 por tanda — el resto se ignora por ahora.');
          addresses = addresses.slice(0, 30);
        }

        state.stops = addresses.map((addr, i) => ({
          id: 'stop-' + i,
          raw: addr,
          lat: null, lon: null,
          status: 'pending'
        }));
        renderStopList();
        geocodeAll();
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

  async function geocodeAll(){
    optimizeBtn.disabled = true;
    optimizeBtn.textContent = 'Ubicando direcciones...';
    for (const stop of state.stops) {
      await geocodeOne(stop);
      renderStopList();
      await new Promise(r => setTimeout(r, 1100)); // respetar límite de Nominatim (1 req/seg)
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

    const userLoc = await getUserLocation();
    const okStops = state.stops.filter(s => s.status === 'ok');
    const points = userLoc ? [{ lat: userLoc.lat, lon: userLoc.lon, raw: 'Tu ubicación actual', isStart: true }, ...okStops] : okStops;

    const coordStr = points.map(p => p.lon + ',' + p.lat).join(';');
    const url = 'https://router.project-osrm.org/trip/v1/driving/' + coordStr +
      '?source=first&roundtrip=false&overview=full&geometries=geojson';

    try {
      const res = await fetch(url);
      const json = await res.json();
      if (json.code !== 'Ok' || !json.trips || !json.trips.length) {
        showError('No pude calcular la ruta óptima. Probá de nuevo en unos segundos.');
        optimizeBtn.disabled = false;
        optimizeBtn.textContent = 'Reintentar optimización';
        return;
      }
      const trip = json.trips[0];
      const ordered = json.waypoints
        .map((wp, i) => ({ point: points[i], order: wp.waypoint_index }))
        .sort((a, b) => a.order - b.order)
        .map(x => x.point);

      state.order = ordered;
      state.trip = trip;
      renderResults(trip, ordered);
      renderMap(trip, ordered);
      optimizeBtn.disabled = false;
      optimizeBtn.textContent = 'Recalcular ruta';
      saveState();
    } catch (e) {
      showError('Falló la conexión con el servicio de rutas. Probá de nuevo.');
      optimizeBtn.disabled = false;
      optimizeBtn.textContent = 'Reintentar optimización';
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

    const chunkSize = 23; // límite práctico de waypoints en el link de Google Maps
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

  loadState();
})();