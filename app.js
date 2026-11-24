/**
 * Arkan Group — Tour Expenses Interactive Web Sheet
 * Exact Excel replica with monthly automation, dynamic routes,
 * searchable From/To dropdowns, day copy/paste, signature manager,
 * print/PDF and native Excel (.xlsx) export.
 */
(function () {
  'use strict';

  const STORAGE_KEY = 'arkan_tour_expenses_v1';
  const MONTH_NAMES = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  // ======================== Utilities ========================
  const $ = (id) => document.getElementById(id);

  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  function numVal(v) {
    if (v === '' || v === null || v === undefined) return 0;
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }

  function fmtNumCell(v) {
    const n = numVal(v);
    if (!n) return '';
    return Number.isInteger(n) ? String(n) : String(Number(n.toFixed(2)));
  }

  function fmtTotalNum(n) {
    return Number.isInteger(n) ? String(n) : String(Number(n.toFixed(2)));
  }

  function parseIsoDate(iso) {
    if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return null;
    const [y, m, d] = iso.split('-').map(Number);
    return { y, m, d, utc: Date.UTC(y, m - 1, d) };
  }

  function toIsoDate(y, m, d) {
    return `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }

  // Excel Duration format: 1-8-2026
  function formatDashDate(iso) {
    const p = parseIsoDate(iso);
    if (!p) return iso || '';
    return `${p.d}-${p.m}-${p.y}`;
  }

  // Excel Table Date format: 1/8/2026
  function formatSlashDate(iso) {
    const p = parseIsoDate(iso);
    if (!p) return iso || '';
    return `${p.d}/${p.m}/${p.y}`;
  }

  // Signature handwritten date format: 31-08-2026
  function formatSigDate(iso) {
    const p = parseIsoDate(iso);
    if (!p) return iso || '';
    return `${String(p.d).padStart(2, '0')}-${String(p.m).padStart(2, '0')}-${p.y}`;
  }

  function calcDurationDays(startIso, endIso) {
    const s = parseIsoDate(startIso);
    const e = parseIsoDate(endIso);
    if (!s || !e || e.utc < s.utc) return 0;
    return Math.round((e.utc - s.utc) / 86400000) + 1;
  }

  function getDayOfWeekUtc(iso) {
    const p = parseIsoDate(iso);
    if (!p) return -1;
    return new Date(p.utc).getUTCDay(); // 0=Sun, ..., 5=Fri, 6=Sat
  }

  function addDaysIso(iso, delta) {
    const p = parseIsoDate(iso);
    if (!p) return iso;
    const dt = new Date(p.utc + delta * 86400000);
    return toIsoDate(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
  }

  function getNextWorkingDate(iso) {
    let next = addDaysIso(iso, 1);
    if (getDayOfWeekUtc(next) === 5) {
      next = addDaysIso(next, 1); // Skip Friday
    }
    return next;
  }

  function createEmptyRoute(fromPlace = '') {
    return { from: fromPlace, to: '', bus: '', metro: '', auto: '', bike: '', rickshaw: '' };
  }

  function createEmptyDay(dateIso, routeCount = 2) {
    const routes = [];
    for (let i = 0; i < Math.max(1, routeCount); i++) {
      routes.push(createEmptyRoute());
    }
    return {
      date: dateIso,
      routes,
      breakfast: '',
      lunch: '',
      dinner: '',
      teaBreak: '',
      hotel: ''
    };
  }

  function computeDayStats(day) {
    let bus = 0, metro = 0, auto = 0, bike = 0, rickshaw = 0;
    for (const r of day.routes) {
      bus += numVal(r.bus);
      metro += numVal(r.metro);
      auto += numVal(r.auto);
      bike += numVal(r.bike);
      rickshaw += numVal(r.rickshaw);
    }
    const transportTotal = bus + metro + auto + bike + rickshaw;
    const breakfast = numVal(day.breakfast);
    const lunch = numVal(day.lunch);
    const dinner = numVal(day.dinner);
    const teaBreak = numVal(day.teaBreak);
    const mealTotal = breakfast + lunch + dinner + teaBreak;
    const hotel = numVal(day.hotel);
    const dayExpense = transportTotal + mealTotal + hotel;
    return {
      bus, metro, auto, bike, rickshaw, transportTotal,
      breakfast, lunch, dinner, teaBreak, mealTotal,
      hotel, dayExpense
    };
  }

  function isDayEmpty(day) {
    const stats = computeDayStats(day);
    if (stats.dayExpense > 0) return false;
    return day.routes.every(r => !(r.from && r.from.trim()) && !(r.to && r.to.trim()));
  }

  // ======================== App State ========================
  let state = null;
  let clipboardDay = null; // { sourceIndex, sourceDate, data: { routes, breakfast, lunch, dinner, teaBreak, hotel } }
  let saveTimer = null;
  let toastTimer = null;

  function buildInitialState() {
    const def = window.TOUR_DEFAULTS;
    return {
      profile: deepClone(def.profile),
      signature: {
        dataUrl: def.defaultSignature || '',
        width: 118,
        blendMultiply: true
      },
      places: [...def.places],
      activeMonthId: def.sampleMonth.id,
      months: [deepClone(def.sampleMonth)]
    };
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return buildInitialState();
      const parsed = JSON.parse(raw);
      if (!parsed || !Array.isArray(parsed.months) || !parsed.months.length) {
        return buildInitialState();
      }
      if (!parsed.profile) parsed.profile = deepClone(window.TOUR_DEFAULTS.profile);
      if (!parsed.signature) {
        parsed.signature = { dataUrl: window.TOUR_DEFAULTS.defaultSignature || '', width: 118, blendMultiply: true };
      }
      if (!Array.isArray(parsed.places)) parsed.places = [...window.TOUR_DEFAULTS.places];
      if (!parsed.months.some(m => m.id === parsed.activeMonthId)) {
        parsed.activeMonthId = parsed.months[0].id;
      }
      return parsed;
    } catch (e) {
      return buildInitialState();
    }
  }

  function persistState(immediate = false, skipCloud = false) {
    state.updatedAt = Date.now();
    const write = () => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        if (!window.TourCloudSync || !window.TourCloudSync.isConnected()) {
          const st = $('saveStatus');
          if (st) {
            st.textContent = 'Saved locally';
            st.style.color = '#059669';
          }
        }
      } catch (err) {
        const st = $('saveStatus');
        if (st) {
          st.textContent = 'Storage full (use Backup)';
          st.style.color = '#dc2626';
        }
      }
      if (!skipCloud && window.TourCloudSync) {
        window.TourCloudSync.notifyLocalChange(state, state.updatedAt);
      }
    };
    if (saveTimer) clearTimeout(saveTimer);
    if (immediate) {
      write();
    } else {
      const st = $('saveStatus');
      if (st) st.textContent = 'Saving...';
      saveTimer = setTimeout(write, 220);
    }
  }

  function getActiveMonth() {
    let m = state.months.find(x => x.id === state.activeMonthId);
    if (!m) {
      m = state.months[0];
      state.activeMonthId = m.id;
    }
    return m;
  }

  function registerPlaceIfNew(name) {
    const trimmed = (name || '').trim();
    if (!trimmed) return;
    const exists = state.places.some(p => p.toLowerCase() === trimmed.toLowerCase());
    if (!exists) {
      state.places.push(trimmed);
      updatePlacesBadge();
      persistState();
    }
  }

  function updatePlacesBadge() {
    $('placesCountBadge').textContent = String(state.places.length);
  }

  function showToast(message) {
    const el = $('toast');
    el.textContent = message;
    el.hidden = false;
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      el.hidden = true;
    }, 2600);
  }

  // ======================== Rendering ========================
  function renderMonthSelector() {
    const sel = $('monthSelect');
    sel.innerHTML = '';
    for (const m of state.months) {
      const opt = document.createElement('option');
      opt.value = m.id;
      opt.textContent = m.label || m.id;
      if (m.id === state.activeMonthId) opt.selected = true;
      sel.appendChild(opt);
    }
    $('btnDeleteMonth').disabled = state.months.length <= 1;
  }

  function renderHeaderAndMeta() {
    const month = getActiveMonth();
    $('companyName').textContent = state.profile.companyName || 'Arkan Group';
    $('reportTitle').textContent = state.profile.reportTitle || 'Tour Expenses (Out of Khulna)';

    $('metaDestination').value = month.destination ?? state.profile.destination ?? '';
    $('metaPurpose').value = month.purpose ?? state.profile.purpose ?? '';
    $('metaDesignation').value = month.designation ?? state.profile.designation ?? '';

    $('metaStartDate').value = month.startDate || '';
    $('metaEndDate').value = month.endDate || '';
    $('startDateDisplay').textContent = formatDashDate(month.startDate);
    $('endDateDisplay').textContent = formatDashDate(month.endDate);
    $('metaTotalDays').textContent = String(calcDurationDays(month.startDate, month.endDate));

    // Set sensible default for custom date adder
    const lastDay = month.days[month.days.length - 1];
    $('customAddDateInput').value = lastDay && lastDay.date
      ? getNextWorkingDate(lastDay.date)
      : (month.startDate || new Date().toISOString().slice(0, 10));

    const lockChk = $('chkLockDatesOnMove');
    if (lockChk) lockChk.checked = state.profile.lockDatesOnMove !== false;
  }

  function renderSignatures() {
    const month = getActiveMonth();
    const img = $('preparedSigImg');
    if (state.signature.dataUrl) {
      img.src = state.signature.dataUrl;
      img.style.width = `${state.signature.width || 118}px`;
      img.classList.toggle('blend-multiply', state.signature.blendMultiply !== false);
    } else {
      img.removeAttribute('src');
    }

    const sigDateEl = $('preparedSigDate');
    if (month.showSigDate && month.sigDate) {
      sigDateEl.textContent = month.sigDate;
      sigDateEl.hidden = false;
    } else {
      sigDateEl.textContent = '';
      sigDateEl.hidden = true;
    }

    const labels = state.profile.signatureLabels || window.TOUR_DEFAULTS.profile.signatureLabels;
    document.querySelectorAll('.sig-label[data-sig-idx]').forEach((el) => {
      const idx = Number(el.dataset.sigIdx);
      if (labels[idx] !== undefined) el.textContent = labels[idx];
    });
  }

  function createDayTbody(day, dayIdx) {
    const tbody = document.createElement('tbody');
    tbody.className = 'day-group';
    tbody.dataset.dayIdx = String(dayIdx);

    if (clipboardDay && clipboardDay.sourceIndex === dayIdx && clipboardDay.monthId === state.activeMonthId) {
      tbody.classList.add('is-copied');
    }

    const routes = day.routes && day.routes.length ? day.routes : [createEmptyRoute()];
    day.routes = routes;
    const rowSpan = routes.length;
    const stats = computeDayStats(day);
    const hasClipboard = Boolean(clipboardDay);
    const month = getActiveMonth();

    routes.forEach((route, rIdx) => {
      const tr = document.createElement('tr');
      tr.dataset.routeIdx = String(rIdx);

      // Col 1: Date (first row shows date + picker; subsequent rows blank in print + hover route delete button)
      const tdDate = document.createElement('td');
      if (rIdx === 0) {
        tdDate.className = 'td-date td-date-main';
        tdDate.innerHTML = `
          <div class="date-cell-wrap" title="তারিখ পরিবর্তন করতে ক্লিক করুন">
            <span class="date-cell-text">${formatSlashDate(day.date)}</span>
            <input type="date" class="date-cell-input" value="${day.date || ''}" data-action="change-date" />
          </div>
        `;
      } else {
        tdDate.className = 'td-date td-date-sub';
        tdDate.innerHTML = `
          <button type="button" class="route-del-btn no-print" data-action="remove-route" data-ridx="${rIdx}" title="এই রুটটি মুছুন">
            &times; R${rIdx + 1}
          </button>
        `;
      }
      tr.appendChild(tdDate);

      // Col 2: From (with searchable combobox dropdown)
      tr.appendChild(buildPlaceCell('from', route.from || '', dayIdx, rIdx));
      // Col 3: To (with searchable combobox dropdown)
      tr.appendChild(buildPlaceCell('to', route.to || '', dayIdx, rIdx));

      // Cols 4..8: Bus, Metro/Rail, Auto, Bike, Rickshaw
      for (const field of ['bus', 'metro', 'auto', 'bike', 'rickshaw']) {
        const td = document.createElement('td');
        const inp = document.createElement('input');
        inp.type = 'number';
        inp.min = '0';
        inp.step = 'any';
        inp.className = 'cell-input cell-input-num';
        inp.value = route[field] !== '' && route[field] !== null && route[field] !== undefined ? route[field] : '';
        inp.dataset.dayIdx = String(dayIdx);
        inp.dataset.routeIdx = String(rIdx);
        inp.dataset.field = field;
        inp.dataset.scope = 'route';
        td.appendChild(inp);
        tr.appendChild(td);
      }

      // Cols 9..16 + Action Col: Merged across all routes of the day (rowspan = routes.length)
      if (rIdx === 0) {
        // Col 9: Transport Total (auto-calculated)
        const tdTrans = document.createElement('td');
        tdTrans.rowSpan = rowSpan;
        tdTrans.className = 'td-merged-calc';
        tdTrans.dataset.calc = 'transportTotal';
        tdTrans.textContent = fmtNumCell(stats.transportTotal);
        tr.appendChild(tdTrans);

        // Cols 10..13: Breakfast, Lunch, Dinner, Tea Break
        for (const mField of ['breakfast', 'lunch', 'dinner', 'teaBreak']) {
          const tdM = document.createElement('td');
          tdM.rowSpan = rowSpan;
          tdM.className = 'td-merged-input';
          const inp = document.createElement('input');
          inp.type = 'number';
          inp.min = '0';
          inp.step = 'any';
          inp.className = 'cell-input cell-input-num';
          inp.value = day[mField] !== '' && day[mField] !== null && day[mField] !== undefined ? day[mField] : '';
          inp.dataset.dayIdx = String(dayIdx);
          inp.dataset.field = mField;
          inp.dataset.scope = 'day';
          tdM.appendChild(inp);
          tr.appendChild(tdM);
        }

        // Col 14: Meal Total (auto-calculated)
        const tdMeal = document.createElement('td');
        tdMeal.rowSpan = rowSpan;
        tdMeal.className = 'td-merged-calc';
        tdMeal.dataset.calc = 'mealTotal';
        tdMeal.textContent = fmtNumCell(stats.mealTotal);
        tr.appendChild(tdMeal);

        // Col 15: Hotel (Living Cost)
        const tdHotel = document.createElement('td');
        tdHotel.rowSpan = rowSpan;
        tdHotel.className = 'td-merged-input';
        const inpH = document.createElement('input');
        inpH.type = 'number';
        inpH.min = '0';
        inpH.step = 'any';
        inpH.className = 'cell-input cell-input-num';
        inpH.value = day.hotel !== '' && day.hotel !== null && day.hotel !== undefined ? day.hotel : '';
        inpH.dataset.dayIdx = String(dayIdx);
        inpH.dataset.field = 'hotel';
        inpH.dataset.scope = 'day';
        tdHotel.appendChild(inpH);
        tr.appendChild(tdHotel);

        // Col 16: Total Day Expense (auto-calculated)
        const tdGrand = document.createElement('td');
        tdGrand.rowSpan = rowSpan;
        tdGrand.className = 'td-merged-calc';
        tdGrand.dataset.calc = 'dayExpense';
        tdGrand.textContent = fmtNumCell(stats.dayExpense);
        tr.appendChild(tdGrand);

        // Col 17 (Screen only): Day Tools
        // Layout: [Drag ⋮⋮] | [▲ Up][Copy] / [▼ Down][Paste] / [+Route][Del]
        const tdAct = document.createElement('td');
        tdAct.rowSpan = rowSpan;
        tdAct.className = 'td-actions no-print';
        const isCopied = clipboardDay && clipboardDay.sourceIndex === dayIdx && clipboardDay.monthId === state.activeMonthId;
        const isFirst = dayIdx === 0;
        const isLast = dayIdx === month.days.length - 1;
        tdAct.innerHTML = `
          <div class="day-tools-wrap">
            <div class="drag-handle" title="টেনে উপরে বা নিচে সরান (Drag Destination → Grand Total)">&#8942;&#8942;</div>
            <div class="day-tools-grid">
              <button type="button" class="tool-btn" data-action="move-up" ${isFirst ? 'disabled' : ''} title="উপরে সরান (তারিখ স্থির থাকবে)">&#9650; Up</button>
              <button type="button" class="tool-btn ${isCopied ? 'btn-copy-active' : ''}" data-action="copy-day" title="এই দিনের সব রুট ও হিসাব কপি করুন">Copy</button>
              <button type="button" class="tool-btn" data-action="move-down" ${isLast ? 'disabled' : ''} title="নিচে সরান (তারিখ স্থির থাকবে)">&#9660; Down</button>
              <button type="button" class="tool-btn ${hasClipboard ? 'btn-paste-ready' : ''}" data-action="paste-day" ${hasClipboard ? '' : 'disabled'} title="কপি করা দিনের রুট এখানে পেস্ট করুন">Paste</button>
              <button type="button" class="tool-btn tool-btn-add" data-action="add-route" title="এই তারিখে নতুন রুট যোগ করুন">+Route</button>
              <button type="button" class="tool-btn tool-btn-del" data-action="delete-day" ${month.days.length <= 1 ? 'disabled' : ''} title="এই তারিখটি মুছুন">Del</button>
            </div>
          </div>
        `;
        tr.appendChild(tdAct);
      }

      tbody.appendChild(tr);
    });

    return tbody;
  }

  function buildPlaceCell(field, value, dayIdx, routeIdx) {
    const td = document.createElement('td');
    td.className = 'td-place';

    const inp = document.createElement('input');
    inp.type = 'text';
    inp.className = 'cell-input cell-input-place';
    inp.value = value;
    inp.autocomplete = 'off';
    inp.dataset.dayIdx = String(dayIdx);
    inp.dataset.routeIdx = String(routeIdx);
    inp.dataset.field = field;
    inp.dataset.scope = 'place';

    const caret = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    caret.setAttribute('viewBox', '0 0 12 12');
    caret.setAttribute('class', 'combo-caret no-print');
    caret.innerHTML = '<path d="M3 4.5 6 7.5 9 4.5" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>';

    td.appendChild(inp);
    td.appendChild(caret);
    return td;
  }

  function renderTable() {
    const table = $('expenseTable');
    // Remove all existing day tbodies
    table.querySelectorAll('tbody.day-group').forEach(el => el.remove());

    const tfoot = table.querySelector('tfoot');
    const month = getActiveMonth();
    const frag = document.createDocumentFragment();

    month.days.forEach((day, idx) => {
      frag.appendChild(createDayTbody(day, idx));
    });

    table.insertBefore(frag, tfoot);
    recalcAllGrandTotals();
  }

  function rerenderSingleDay(dayIdx, flashPaste = false) {
    const table = $('expenseTable');
    const oldTbody = table.querySelector(`tbody.day-group[data-day-idx="${dayIdx}"]`);
    const month = getActiveMonth();
    if (!oldTbody || !month.days[dayIdx]) {
      renderTable();
      return;
    }
    const newTbody = createDayTbody(month.days[dayIdx], dayIdx);
    if (flashPaste) newTbody.classList.add('just-pasted');
    oldTbody.replaceWith(newTbody);
    recalcAllGrandTotals();
  }

  function updateSingleDayCalcCells(dayIdx) {
    const month = getActiveMonth();
    const day = month.days[dayIdx];
    if (!day) return;
    const stats = computeDayStats(day);
    const tbody = $('expenseTable').querySelector(`tbody.day-group[data-day-idx="${dayIdx}"]`);
    if (tbody) {
      const elTr = tbody.querySelector('[data-calc="transportTotal"]');
      const elMl = tbody.querySelector('[data-calc="mealTotal"]');
      const elGd = tbody.querySelector('[data-calc="dayExpense"]');
      if (elTr) elTr.textContent = fmtNumCell(stats.transportTotal);
      if (elMl) elMl.textContent = fmtNumCell(stats.mealTotal);
      if (elGd) elGd.textContent = fmtNumCell(stats.dayExpense);
    }
    recalcAllGrandTotals();
  }

  function recalcAllGrandTotals() {
    const month = getActiveMonth();
    let bus = 0, metro = 0, auto = 0, bike = 0, rickshaw = 0, transport = 0;
    let breakfast = 0, lunch = 0, dinner = 0, teaBreak = 0, meal = 0, hotel = 0, grand = 0;
    let totalRoutes = 0;

    for (const day of month.days) {
      totalRoutes += day.routes.length;
      const s = computeDayStats(day);
      bus += s.bus;
      metro += s.metro;
      auto += s.auto;
      bike += s.bike;
      rickshaw += s.rickshaw;
      transport += s.transportTotal;
      breakfast += s.breakfast;
      lunch += s.lunch;
      dinner += s.dinner;
      teaBreak += s.teaBreak;
      meal += s.mealTotal;
      hotel += s.hotel;
      grand += s.dayExpense;
    }

    $('sumBus').textContent = fmtTotalNum(bus);
    $('sumMetro').textContent = fmtTotalNum(metro);
    $('sumAuto').textContent = fmtTotalNum(auto);
    $('sumBike').textContent = fmtTotalNum(bike);
    $('sumRickshaw').textContent = fmtTotalNum(rickshaw);
    $('sumTransport').textContent = fmtTotalNum(transport);
    $('sumBreakfast').textContent = fmtTotalNum(breakfast);
    $('sumLunch').textContent = fmtTotalNum(lunch);
    $('sumDinner').textContent = fmtTotalNum(dinner);
    $('sumTeaBreak').textContent = fmtTotalNum(teaBreak);
    $('sumMeal').textContent = fmtTotalNum(meal);
    $('sumHotel').textContent = fmtTotalNum(hotel);
    $('sumGrand').textContent = fmtTotalNum(grand);

    // Top summary ribbon
    $('statDays').textContent = String(month.days.length);
    $('statRoutes').textContent = String(totalRoutes);
    $('statTransport').textContent = fmtTotalNum(transport);
    $('statMeal').textContent = fmtTotalNum(meal);
    $('statHotel').textContent = fmtTotalNum(hotel);
    $('statGrand').textContent = fmtTotalNum(grand);
  }

  function updateClipboardUI() {
    const bar = $('clipboardBar');
    if (!clipboardDay) {
      bar.hidden = true;
    } else {
      const s = computeDayStats(clipboardDay.data);
      $('clipboardText').textContent =
        `Copied ${formatSlashDate(clipboardDay.sourceDate)} (${clipboardDay.data.routes.length} routes, Transport ${fmtTotalNum(s.transportTotal)}, Meal ${fmtTotalNum(s.mealTotal)}, Total ${fmtTotalNum(s.dayExpense)}) — ক্লিক করুন যেকোনো ডেটের Paste বাটনে`;
      bar.hidden = false;
    }

    // Update all Copy/Paste button states without losing focus
    document.querySelectorAll('tbody.day-group').forEach((tbody) => {
      const idx = Number(tbody.dataset.dayIdx);
      const isSrc = clipboardDay && clipboardDay.sourceIndex === idx && clipboardDay.monthId === state.activeMonthId;
      tbody.classList.toggle('is-copied', Boolean(isSrc));
      const copyBtn = tbody.querySelector('[data-action="copy-day"]');
      const pasteBtn = tbody.querySelector('[data-action="paste-day"]');
      if (copyBtn) copyBtn.classList.toggle('btn-copy-active', Boolean(isSrc));
      if (pasteBtn) {
        pasteBtn.disabled = !clipboardDay;
        pasteBtn.classList.toggle('btn-paste-ready', Boolean(clipboardDay));
      }
    });
  }

  function renderAll() {
    renderMonthSelector();
    renderHeaderAndMeta();
    renderTable();
    renderSignatures();
    updatePlacesBadge();
    updateClipboardUI();
  }

  // ======================== Searchable From / To Combobox Dropdown ========================
  const dropdown = $('placeDropdown');
  let activePlaceInput = null;
  let dropdownActiveIdx = -1;
  let currentFilteredOptions = [];

  function positionDropdown(input) {
    const rect = input.getBoundingClientRect();
    const spaceBelow = window.innerHeight - rect.bottom;
    dropdown.style.width = `${Math.max(rect.width + 26, 175)}px`;
    dropdown.style.left = `${Math.min(rect.left, window.innerWidth - 195)}px`;
    if (spaceBelow < 220 && rect.top > 240) {
      dropdown.style.top = 'auto';
      dropdown.style.bottom = `${window.innerHeight - rect.top + 3}px`;
    } else {
      dropdown.style.bottom = 'auto';
      dropdown.style.top = `${rect.bottom + 2}px`;
    }
  }

  function openPlaceDropdown(input, filterByValue = false) {
    activePlaceInput = input;
    const rawTrim = (input.value || '').trim();
    const query = filterByValue ? rawTrim.toLowerCase() : '';

    // Combine saved places + any places currently in active month (deduplicated)
    const pool = [...state.places];
    const seen = new Set(pool.map(p => p.toLowerCase()));
    const month = getActiveMonth();
    for (const d of month.days) {
      for (const r of d.routes) {
        for (const p of [r.from, r.to]) {
          const t = (p || '').trim();
          if (t && !seen.has(t.toLowerCase())) {
            seen.add(t.toLowerCase());
            pool.push(t);
          }
        }
      }
    }

    let matches = pool;
    if (query) {
      const starts = pool.filter(p => p.toLowerCase().startsWith(query));
      const contains = pool.filter(p => !p.toLowerCase().startsWith(query) && p.toLowerCase().includes(query));
      matches = [...starts, ...contains];
    }

    currentFilteredOptions = matches;
    const existingIdx = rawTrim ? matches.findIndex(p => p.toLowerCase() === rawTrim.toLowerCase()) : -1;
    dropdownActiveIdx = existingIdx >= 0 ? existingIdx : (matches.length ? 0 : -1);

    dropdown.innerHTML = '';
    matches.forEach((place, idx) => {
      const div = document.createElement('div');
      div.className = 'place-option' + (idx === dropdownActiveIdx ? ' is-active' : '');
      div.textContent = place;
      div.dataset.value = place;
      div.addEventListener('mousedown', (e) => {
        e.preventDefault();
        selectPlaceOption(place);
      });
      dropdown.appendChild(div);
    });

    const exactMatch = rawTrim && pool.some(p => p.toLowerCase() === rawTrim.toLowerCase());
    if (rawTrim && !exactMatch) {
      const addDiv = document.createElement('div');
      addDiv.className = 'place-option place-option-add';
      addDiv.textContent = `+ Add "${rawTrim}" to dropdown`;
      addDiv.addEventListener('mousedown', (e) => {
        e.preventDefault();
        registerPlaceIfNew(rawTrim);
        selectPlaceOption(rawTrim);
        showToast(`"${rawTrim}" added to dropdown list`);
      });
      dropdown.appendChild(addDiv);
    }

    if (!matches.length && !rawTrim) {
      const empty = document.createElement('div');
      empty.className = 'place-option-empty';
      empty.textContent = 'Type a location name...';
      dropdown.appendChild(empty);
    }

    positionDropdown(input);
    dropdown.hidden = false;
  }

  function closePlaceDropdown() {
    dropdown.hidden = true;
    activePlaceInput = null;
    dropdownActiveIdx = -1;
  }

  function selectPlaceOption(value) {
    if (!activePlaceInput) return;
    const inp = activePlaceInput;
    inp.value = value;
    const dIdx = Number(inp.dataset.dayIdx);
    const rIdx = Number(inp.dataset.routeIdx);
    const field = inp.dataset.field;
    const month = getActiveMonth();
    if (month.days[dIdx] && month.days[dIdx].routes[rIdx]) {
      month.days[dIdx].routes[rIdx][field] = value;
      registerPlaceIfNew(value);
      persistState();
    }
    closePlaceDropdown();
  }

  function highlightDropdownOption(newIdx) {
    const items = dropdown.querySelectorAll('.place-option:not(.place-option-add)');
    if (!items.length) return;
    dropdownActiveIdx = (newIdx + items.length) % items.length;
    items.forEach((el, i) => el.classList.toggle('is-active', i === dropdownActiveIdx));
    items[dropdownActiveIdx].scrollIntoView({ block: 'nearest' });
  }

  // ======================== Table Event Delegation ========================
  function bindTableEvents() {
    const table = $('expenseTable');

    // Input typing inside cells (numbers & From/To places)
    table.addEventListener('input', (e) => {
      const target = e.target;
      if (!target.classList.contains('cell-input')) return;
      const month = getActiveMonth();
      const dIdx = Number(target.dataset.dayIdx);
      const day = month.days[dIdx];
      if (!day) return;

      const scope = target.dataset.scope;
      const field = target.dataset.field;

      if (scope === 'place') {
        const rIdx = Number(target.dataset.routeIdx);
        if (day.routes[rIdx]) {
          day.routes[rIdx][field] = target.value;
          persistState();
          openPlaceDropdown(target, true);
        }
      } else if (scope === 'route') {
        const rIdx = Number(target.dataset.routeIdx);
        if (day.routes[rIdx]) {
          day.routes[rIdx][field] = target.value === '' ? '' : Number(target.value);
          updateSingleDayCalcCells(dIdx);
          persistState();
        }
      } else if (scope === 'day') {
        day[field] = target.value === '' ? '' : Number(target.value);
        updateSingleDayCalcCells(dIdx);
        persistState();
      }
    });

    // Focus opens From/To dropdown automatically
    table.addEventListener('focusin', (e) => {
      const target = e.target;
      if (target.dataset && target.dataset.scope === 'place') {
        openPlaceDropdown(target);
      } else {
        closePlaceDropdown();
      }
    });

    // Blur saves any newly typed place name into dropdown list
    table.addEventListener('focusout', (e) => {
      const target = e.target;
      if (target.dataset && target.dataset.scope === 'place') {
        const val = target.value.trim();
        if (val) registerPlaceIfNew(val);
        setTimeout(() => {
          if (document.activeElement !== target) closePlaceDropdown();
        }, 120);
      }
    });

    // Keyboard navigation in From/To dropdown
    table.addEventListener('keydown', (e) => {
      const target = e.target;
      if (!target.dataset || target.dataset.scope !== 'place') return;
      if (dropdown.hidden) {
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          openPlaceDropdown(target);
        }
        return;
      }
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        highlightDropdownOption(dropdownActiveIdx + 1);
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        highlightDropdownOption(dropdownActiveIdx - 1);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        if (dropdownActiveIdx >= 0 && currentFilteredOptions[dropdownActiveIdx]) {
          selectPlaceOption(currentFilteredOptions[dropdownActiveIdx]);
        } else {
          registerPlaceIfNew(target.value);
          closePlaceDropdown();
        }
      } else if (e.key === 'Escape') {
        closePlaceDropdown();
      }
    });

    // Change Date picker on any Day row
    table.addEventListener('change', (e) => {
      const target = e.target;
      if (target.dataset.action !== 'change-date') return;
      const tbody = target.closest('tbody.day-group');
      if (!tbody) return;
      const dIdx = Number(tbody.dataset.dayIdx);
      const month = getActiveMonth();
      if (!month.days[dIdx]) return;
      month.days[dIdx].date = target.value;
      const txt = tbody.querySelector('.date-cell-text');
      if (txt) txt.textContent = formatSlashDate(target.value);
      persistState();
    });

    // Click actions inside Day Tools & Route Delete
    table.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-action]');
      if (!btn) return;
      const tbody = btn.closest('tbody.day-group');
      if (!tbody) return;
      const dIdx = Number(tbody.dataset.dayIdx);
      const action = btn.dataset.action;
      const month = getActiveMonth();
      const day = month.days[dIdx];
      if (!day) return;

      if (action === 'add-route') {
        const lastRoute = day.routes[day.routes.length - 1];
        const autoFrom = lastRoute && lastRoute.to ? lastRoute.to : '';
        day.routes.push(createEmptyRoute(autoFrom));
        rerenderSingleDay(dIdx);
        persistState();
        // Focus the new route's To (if From was auto-filled) or From
        const updatedBody = table.querySelector(`tbody.day-group[data-day-idx="${dIdx}"]`);
        if (updatedBody) {
          const newIdx = day.routes.length - 1;
          const focusField = autoFrom ? 'to' : 'from';
          const inp = updatedBody.querySelector(`input[data-route-idx="${newIdx}"][data-field="${focusField}"]`);
          if (inp) inp.focus();
        }
      } else if (action === 'remove-route') {
        const rIdx = Number(btn.dataset.ridx);
        if (day.routes.length <= 1) return;
        day.routes.splice(rIdx, 1);
        rerenderSingleDay(dIdx);
        persistState();
      } else if (action === 'copy-day') {
        clipboardDay = {
          monthId: state.activeMonthId,
          sourceIndex: dIdx,
          sourceDate: day.date,
          data: {
            routes: deepClone(day.routes),
            breakfast: day.breakfast,
            lunch: day.lunch,
            dinner: day.dinner,
            teaBreak: day.teaBreak,
            hotel: day.hotel
          }
        };
        updateClipboardUI();
        showToast(`Copied ${formatSlashDate(day.date)} (${day.routes.length} routes)`);
      } else if (action === 'paste-day') {
        if (!clipboardDay) return;
        applyDayPayload(day, clipboardDay.data);
        rerenderSingleDay(dIdx, true);
        persistState();
        showToast(`Pasted into ${formatSlashDate(day.date)}`);
      } else if (action === 'move-up') {
        moveDayData(dIdx, dIdx - 1);
      } else if (action === 'move-down') {
        moveDayData(dIdx, dIdx + 1);
      } else if (action === 'delete-day') {
        if (month.days.length <= 1) return;
        if (!isDayEmpty(day)) {
          if (!confirm(`Delete ${formatSlashDate(day.date)} and its ${day.routes.length} routes?`)) return;
        }
        month.days.splice(dIdx, 1);
        if (clipboardDay && clipboardDay.monthId === state.activeMonthId && clipboardDay.sourceIndex === dIdx) {
          clipboardDay = null;
        }
        renderTable();
        updateClipboardUI();
        persistState();
      }
    });

    // ---- Drag & Drop Reorder (Destination -> Grand Total moves; Dates stay fixed by default) ----
    let dragSourceIdx = null;
    let dropTargetIdx = null;
    let dropPosition = null; // 'before' | 'after'
    let dragScrollTimer = null;

    function clearDropIndicators() {
      table.querySelectorAll('tbody.day-group.drop-before, tbody.day-group.drop-after').forEach((el) => {
        el.classList.remove('drop-before', 'drop-after');
      });
    }

    function stopAutoScroll() {
      if (dragScrollTimer) {
        clearInterval(dragScrollTimer);
        dragScrollTimer = null;
      }
    }

    function updateAutoScroll(clientY) {
      const topZone = 145;
      const bottomZone = window.innerHeight - 95;
      let speed = 0;
      if (clientY < topZone) speed = -14;
      else if (clientY > bottomZone) speed = 14;

      if (speed === 0) {
        stopAutoScroll();
      } else if (!dragScrollTimer) {
        dragScrollTimer = setInterval(() => {
          window.scrollBy(0, speed);
        }, 20);
      }
    }

    // Only arm draggable="true" when pressing the drag handle so cell inputs remain 100% selectable
    table.addEventListener('mousedown', (e) => {
      const handle = e.target.closest('.drag-handle');
      if (!handle) return;
      const tbody = handle.closest('tbody.day-group');
      if (tbody) tbody.setAttribute('draggable', 'true');
    });

    window.addEventListener('mouseup', () => {
      table.querySelectorAll('tbody.day-group[draggable="true"]').forEach((el) => {
        el.removeAttribute('draggable');
      });
    });

    table.addEventListener('dragstart', (e) => {
      const tbody = e.target.closest('tbody.day-group');
      if (!tbody || tbody.getAttribute('draggable') !== 'true') {
        e.preventDefault();
        return;
      }
      closePlaceDropdown();
      dragSourceIdx = Number(tbody.dataset.dayIdx);
      tbody.classList.add('is-dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', String(dragSourceIdx));
    });

    table.addEventListener('dragover', (e) => {
      if (dragSourceIdx === null) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      updateAutoScroll(e.clientY);

      const overBody = e.target.closest('tbody.day-group');
      if (!overBody) return;
      const overIdx = Number(overBody.dataset.dayIdx);
      const rect = overBody.getBoundingClientRect();
      const pos = (e.clientY < rect.top + rect.height / 2) ? 'before' : 'after';

      if (dropTargetIdx !== overIdx || dropPosition !== pos) {
        clearDropIndicators();
        dropTargetIdx = overIdx;
        dropPosition = pos;
        overBody.classList.add(pos === 'before' ? 'drop-before' : 'drop-after');
      }
    });

    table.addEventListener('drop', (e) => {
      if (dragSourceIdx === null) return;
      e.preventDefault();
      stopAutoScroll();
      clearDropIndicators();

      const src = dragSourceIdx;
      const overIdx = dropTargetIdx;
      const pos = dropPosition;
      dragSourceIdx = null;
      dropTargetIdx = null;
      dropPosition = null;

      if (overIdx === null || pos === null) return;
      let insertAt = pos === 'after' ? overIdx + 1 : overIdx;
      if (src < insertAt) insertAt -= 1;
      moveDayData(src, insertAt);
    });

    table.addEventListener('dragend', () => {
      stopAutoScroll();
      clearDropIndicators();
      table.querySelectorAll('tbody.day-group').forEach((el) => {
        el.classList.remove('is-dragging');
        el.removeAttribute('draggable');
      });
      dragSourceIdx = null;
      dropTargetIdx = null;
      dropPosition = null;
    });

    window.addEventListener('scroll', () => {
      if (activePlaceInput && !dropdown.hidden) positionDropdown(activePlaceInput);
    }, true);
    window.addEventListener('resize', () => {
      if (activePlaceInput && !dropdown.hidden) positionDropdown(activePlaceInput);
    });
  }

  function moveDayData(fromIdx, toIdx) {
    const month = getActiveMonth();
    if (!month || !month.days || month.days.length <= 1) return;
    if (fromIdx < 0 || fromIdx >= month.days.length) return;
    if (toIdx < 0 || toIdx >= month.days.length) return;
    if (fromIdx === toIdx) return;

    const lockDates = state.profile.lockDatesOnMove !== false;
    const fixedDates = month.days.map(d => d.date);

    const [moved] = month.days.splice(fromIdx, 1);
    month.days.splice(toIdx, 0, moved);

    if (lockDates) {
      // Keep every row's calendar date unchanged; only Destination -> Grand Total shifts
      // Row counts (2, 3, 4, 5, 6 routes) automatically adjust for every date!
      month.days.forEach((d, i) => {
        d.date = fixedDates[i];
      });
    }

    if (clipboardDay && clipboardDay.monthId === state.activeMonthId) {
      clipboardDay.sourceIndex = -1;
    }

    renderTable();
    updateClipboardUI();

    const movedBody = $('expenseTable').querySelector(`tbody.day-group[data-day-idx="${toIdx}"]`);
    if (movedBody) {
      movedBody.classList.add('just-moved');
      const rect = movedBody.getBoundingClientRect();
      if (rect.top < 110 || rect.bottom > window.innerHeight - 40) {
        movedBody.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }

    persistState();
    const targetDateStr = formatSlashDate(month.days[toIdx].date);
    showToast(
      lockDates
        ? `Moved ${moved.routes.length} route(s) to ${targetDateStr} (dates kept fixed)`
        : `Moved ${targetDateStr} block to row #${toIdx + 1}`
    );
  }

  function applyDayPayload(targetDay, source) {
    targetDay.routes = deepClone(source.routes);
    targetDay.breakfast = source.breakfast;
    targetDay.lunch = source.lunch;
    targetDay.dinner = source.dinner;
    targetDay.teaBreak = source.teaBreak;
    targetDay.hotel = source.hotel;
  }

  // ======================== Header, Meta & Toolbar Events ========================
  function bindHeaderAndToolbarEvents() {
    // Editable Company Name & Subtitle (saved globally)
    $('companyName').addEventListener('input', (e) => {
      state.profile.companyName = e.target.textContent.trim();
      persistState();
    });

    const lockChk = $('chkLockDatesOnMove');
    if (lockChk) {
      lockChk.addEventListener('change', (e) => {
        state.profile.lockDatesOnMove = e.target.checked;
        persistState();
      });
    }
    $('reportTitle').addEventListener('input', (e) => {
      state.profile.reportTitle = e.target.textContent.trim();
      persistState();
    });

    // Destination, Purpose, Travellers Designation (saved on month + persisted as default!)
    $('metaDestination').addEventListener('input', (e) => {
      const m = getActiveMonth();
      m.destination = e.target.value;
      state.profile.destination = e.target.value;
      persistState();
    });
    $('metaPurpose').addEventListener('input', (e) => {
      const m = getActiveMonth();
      m.purpose = e.target.value;
      state.profile.purpose = e.target.value;
      persistState();
    });
    $('metaDesignation').addEventListener('input', (e) => {
      const m = getActiveMonth();
      m.designation = e.target.value;
      state.profile.designation = e.target.value;
      persistState();
    });

    // Duration Start & End Date Pickers -> Auto-calculate Total Day
    $('metaStartDate').addEventListener('change', (e) => {
      const m = getActiveMonth();
      m.startDate = e.target.value;
      $('startDateDisplay').textContent = formatDashDate(m.startDate);
      $('metaTotalDays').textContent = String(calcDurationDays(m.startDate, m.endDate));
      persistState();
    });
    $('metaEndDate').addEventListener('change', (e) => {
      const m = getActiveMonth();
      m.endDate = e.target.value;
      $('endDateDisplay').textContent = formatDashDate(m.endDate);
      $('metaTotalDays').textContent = String(calcDurationDays(m.startDate, m.endDate));
      persistState();
    });

    // Editable Signature Labels
    document.querySelectorAll('.sig-label[data-sig-idx]').forEach((el) => {
      el.addEventListener('input', () => {
        const idx = Number(el.dataset.sigIdx);
        state.profile.signatureLabels[idx] = el.textContent.trim();
        persistState();
      });
    });

    // Month Switcher
    $('monthSelect').addEventListener('change', (e) => {
      state.activeMonthId = e.target.value;
      persistState(true);
      renderAll();
    });

    // Delete current month
    $('btnDeleteMonth').addEventListener('click', () => {
      if (state.months.length <= 1) return;
      const m = getActiveMonth();
      if (!confirm(`Delete sheet "${m.label}" permanently?`)) return;
      state.months = state.months.filter(x => x.id !== m.id);
      state.activeMonthId = state.months[0].id;
      if (clipboardDay && clipboardDay.monthId === m.id) clipboardDay = null;
      persistState(true);
      renderAll();
      showToast(`Deleted "${m.label}"`);
    });

    // Sort Dates
    $('btnSortDates').addEventListener('click', () => {
      const m = getActiveMonth();
      m.days.sort((a, b) => (a.date || '').localeCompare(b.date || ''));
      renderTable();
      persistState();
      showToast('Dates sorted chronologically');
    });

    // Quick Add Next Working Date
    const addNextWorkingDay = () => {
      const m = getActiveMonth();
      const last = m.days[m.days.length - 1];
      const nextDate = last && last.date
        ? getNextWorkingDate(last.date)
        : (m.startDate || new Date().toISOString().slice(0, 10));
      m.days.push(createEmptyDay(nextDate, 2));
      renderTable();
      $('customAddDateInput').value = getNextWorkingDate(nextDate);
      const lastTbody = $('expenseTable').querySelector(`tbody.day-group[data-day-idx="${m.days.length - 1}"]`);
      if (lastTbody) lastTbody.scrollIntoView({ behavior: 'smooth', block: 'center' });
      persistState();
    };
    $('btnAddNextDate').addEventListener('click', addNextWorkingDay);
    $('btnQuickAddDate').addEventListener('click', addNextWorkingDay);

    // Add Custom Date with chosen route count
    $('btnAddCustomDate').addEventListener('click', () => {
      const m = getActiveMonth();
      const dateVal = $('customAddDateInput').value || m.startDate || new Date().toISOString().slice(0, 10);
      const count = Math.max(1, Number($('customAddRoutesCount').value) || 2);
      m.days.push(createEmptyDay(dateVal, count));
      renderTable();
      $('customAddDateInput').value = getNextWorkingDate(dateVal);
      const lastTbody = $('expenseTable').querySelector(`tbody.day-group[data-day-idx="${m.days.length - 1}"]`);
      if (lastTbody) lastTbody.scrollIntoView({ behavior: 'smooth', block: 'center' });
      persistState();
      showToast(`Added ${formatSlashDate(dateVal)} (${count} routes)`);
    });

    // Clipboard Clear
    $('btnClearClipboard').addEventListener('click', () => {
      clipboardDay = null;
      updateClipboardUI();
    });

    // Print / PDF
    $('btnPrint').addEventListener('click', () => {
      closePlaceDropdown();
      window.print();
    });

    // Export Excel (.xlsx)
    $('btnExportXlsx').addEventListener('click', async () => {
      const m = getActiveMonth();
      const safeName = (m.label || m.id || 'Tour_Expenses').replace(/[^a-zA-Z0-9_-]+/g, '_');
      const filename = `Arkan_Group_Tour_Expenses_${safeName}.xlsx`;
      try {
        await window.TourXlsxExporter.download({
          companyName: state.profile.companyName,
          reportTitle: state.profile.reportTitle,
          destination: m.destination ?? state.profile.destination,
          purpose: m.purpose ?? state.profile.purpose,
          designation: m.designation ?? state.profile.designation,
          startDate: m.startDate,
          endDate: m.endDate,
          days: m.days,
          signatureLabels: state.profile.signatureLabels,
          signatureDataUrl: state.signature.dataUrl,
          showSigDate: m.showSigDate,
          sigDate: m.sigDate,
          sigBlendMultiply: state.signature.blendMultiply !== false
        }, filename);
        showToast(`Downloaded ${filename}`);
      } catch (err) {
        console.error(err);
        alert('Excel export failed: ' + err.message);
      }
    });

    // Backup / Restore JSON
    $('btnBackup').addEventListener('click', () => {
      const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Arkan_Tour_Expenses_Backup_${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 200);
      showToast('Backup JSON downloaded');
    });

    $('btnRestore').addEventListener('click', () => $('restoreFileInput').click());
    $('restoreFileInput').addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const imported = JSON.parse(reader.result);
          if (!imported || !Array.isArray(imported.months) || !imported.months.length) {
            throw new Error('Invalid backup file');
          }
          state = imported;
          clipboardDay = null;
          persistState(true);
          renderAll();
          showToast(`Restored ${state.months.length} month sheet(s)`);
        } catch (err) {
          alert('Could not restore backup: ' + err.message);
        }
        e.target.value = '';
      };
      reader.readAsText(file);
    });
  }

  // ======================== Modals Common ========================
  function openModal(id) {
    closePlaceDropdown();
    $(id).hidden = false;
  }

  function closeModal(id) {
    $(id).hidden = true;
  }

  function bindModalCloseHandlers() {
    document.querySelectorAll('[data-close-modal]').forEach((btn) => {
      btn.addEventListener('click', () => closeModal(btn.dataset.closeModal));
    });
    document.querySelectorAll('.modal-backdrop').forEach((backdrop) => {
      backdrop.addEventListener('click', (e) => {
        if (e.target === backdrop) backdrop.hidden = true;
      });
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        document.querySelectorAll('.modal-backdrop:not([hidden])').forEach(m => { m.hidden = true; });
      }
    });
  }

  // ======================== MODAL 1: New Month Generator ========================
  function initNewMonthModal() {
    const picker = $('nmMonthPicker');
    const labelInp = $('nmLabel');
    const startInp = $('nmStartDate');
    const endInp = $('nmEndDate');

    function updateNewMonthPreview() {
      const total = calcDurationDays(startInp.value, endInp.value);
      let work = 0;
      if (total > 0) {
        let cur = startInp.value;
        for (let i = 0; i < total; i++) {
          if (getDayOfWeekUtc(cur) !== 5) work++;
          cur = addDaysIso(cur, 1);
        }
      }
      $('nmTotalDaysPreview').textContent = String(total);
      $('nmWorkDaysPreview').textContent = String(work);
    }

    function syncFromMonthValue(ym) {
      if (!ym || !/^\d{4}-\d{2}$/.test(ym)) return;
      const [y, m] = ym.split('-').map(Number);
      const lastDay = new Date(Date.UTC(y, m, 0)).getUTCDate();
      startInp.value = toIsoDate(y, m, 1);
      endInp.value = toIsoDate(y, m, lastDay);
      labelInp.value = `${MONTH_NAMES[m - 1]} ${y}`;
      updateNewMonthPreview();
    }

    $('btnNewMonth').addEventListener('click', () => {
      // Suggest the month immediately following the latest saved month
      const cur = getActiveMonth();
      const p = parseIsoDate(cur.startDate) || { y: new Date().getFullYear(), m: new Date().getMonth() + 1 };
      let nextM = p.m + 1;
      let nextY = p.y;
      if (nextM > 12) { nextM = 1; nextY += 1; }
      const ym = `${nextY}-${String(nextM).padStart(2, '0')}`;
      picker.value = ym;
      syncFromMonthValue(ym);
      openModal('newMonthModal');
    });

    picker.addEventListener('change', () => syncFromMonthValue(picker.value));
    startInp.addEventListener('change', updateNewMonthPreview);
    endInp.addEventListener('change', updateNewMonthPreview);

    document.querySelectorAll('input[name="nmMode"]').forEach((radio) => {
      radio.addEventListener('change', () => {
        const mode = document.querySelector('input[name="nmMode"]:checked').value;
        $('nmDefaultRoutesGroup').hidden = (mode === 'clone');
      });
    });

    $('btnConfirmNewMonth').addEventListener('click', () => {
      const startIso = startInp.value;
      const endIso = endInp.value;
      const total = calcDurationDays(startIso, endIso);
      if (!startIso || !endIso || total <= 0) {
        alert('Please pick a valid start date and end date.');
        return;
      }

      const mode = document.querySelector('input[name="nmMode"]:checked').value;
      const defaultRoutes = Math.max(1, Number($('nmDefaultRoutes').value) || 2);
      const label = (labelInp.value || '').trim() || `${formatDashDate(startIso)} to ${formatDashDate(endIso)}`;

      // Generate unique month ID
      const baseId = startIso.slice(0, 7);
      let id = baseId;
      let suffix = 2;
      while (state.months.some(x => x.id === id)) {
        id = `${baseId}-${suffix++}`;
      }

      const dates = [];
      let curIso = startIso;
      for (let i = 0; i < total; i++) {
        if (mode === 'alldays') {
          dates.push(curIso);
        } else if (mode === 'workdays' || mode === 'clone') {
          if (getDayOfWeekUtc(curIso) !== 5) dates.push(curIso);
        }
        curIso = addDaysIso(curIso, 1);
      }

      const curMonth = getActiveMonth();
      let days = [];

      if (mode === 'blank') {
        days = [createEmptyDay(startIso, defaultRoutes)];
      } else if (mode === 'clone') {
        const srcDays = curMonth.days || [];
        days = dates.map((dIso, i) => {
          if (srcDays[i]) {
            return {
              date: dIso,
              routes: deepClone(srcDays[i].routes),
              breakfast: srcDays[i].breakfast,
              lunch: srcDays[i].lunch,
              dinner: srcDays[i].dinner,
              teaBreak: srcDays[i].teaBreak,
              hotel: srcDays[i].hotel
            };
          }
          return createEmptyDay(dIso, defaultRoutes);
        });
      } else {
        days = dates.map(dIso => createEmptyDay(dIso, defaultRoutes));
      }

      const newMonth = {
        id,
        label,
        startDate: startIso,
        endDate: endIso,
        destination: state.profile.destination || curMonth.destination || 'Working',
        purpose: state.profile.purpose || curMonth.purpose || 'Delar Visit & Surcing',
        designation: state.profile.designation || curMonth.designation || 'ASM',
        showSigDate: curMonth.showSigDate !== false,
        sigDate: formatSigDate(endIso),
        days
      };

      state.months.push(newMonth);
      state.activeMonthId = id;
      persistState(true);
      renderAll();
      closeModal('newMonthModal');
      showToast(`Created "${label}" (${days.length} days)`);
    });
  }

  // ======================== MODAL 2: Manage Places ========================
  function initPlacesModal() {
    function renderPlacesChips() {
      const container = $('placesListContainer');
      container.innerHTML = '';
      state.places.forEach((place, idx) => {
        const chip = document.createElement('span');
        chip.className = 'place-chip';
        chip.innerHTML = `<span></span><button type="button" title="Remove">&times;</button>`;
        chip.querySelector('span').textContent = place;
        chip.querySelector('button').addEventListener('click', () => {
          state.places.splice(idx, 1);
          renderPlacesChips();
          updatePlacesBadge();
          persistState();
        });
        container.appendChild(chip);
      });
    }

    $('btnManagePlaces').addEventListener('click', () => {
      renderPlacesChips();
      openModal('placesModal');
      setTimeout(() => $('newPlaceInput').focus(), 60);
    });

    const addFromInput = () => {
      const inp = $('newPlaceInput');
      const val = inp.value.trim();
      if (!val) return;
      registerPlaceIfNew(val);
      inp.value = '';
      renderPlacesChips();
      inp.focus();
    };

    $('btnAddPlace').addEventListener('click', addFromInput);
    $('newPlaceInput').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        addFromInput();
      }
    });

    $('btnResetPlaces').addEventListener('click', () => {
      if (!confirm('Reset dropdown list to the default places from the Excel format?')) return;
      state.places = [...window.TOUR_DEFAULTS.places];
      renderPlacesChips();
      updatePlacesBadge();
      persistState();
      showToast('Default places restored');
    });
  }

  // ======================== MODAL 3: Signature Manager (Upload + Draw) ========================
  function initSignatureModal() {
    const previewImg = $('sigModalPreview');
    const emptyMsg = $('sigModalEmpty');
    const widthRange = $('sigWidthRange');
    const widthVal = $('sigWidthVal');
    const blendCheck = $('sigBlendMultiply');
    const showDateCheck = $('sigShowDateCheck');
    const dateInput = $('sigDateInput');

    function syncSigModalUI() {
      const m = getActiveMonth();
      if (state.signature.dataUrl) {
        previewImg.src = state.signature.dataUrl;
        previewImg.hidden = false;
        emptyMsg.hidden = true;
      } else {
        previewImg.removeAttribute('src');
        previewImg.hidden = true;
        emptyMsg.hidden = false;
      }
      widthRange.value = String(state.signature.width || 118);
      widthVal.textContent = `${state.signature.width || 118}px`;
      blendCheck.checked = state.signature.blendMultiply !== false;
      showDateCheck.checked = m.showSigDate !== false;
      dateInput.value = m.sigDate || formatSigDate(m.endDate);
    }

    const openSig = () => {
      syncSigModalUI();
      openModal('signatureModal');
    };

    $('btnSignature').addEventListener('click', openSig);
    $('preparedSigBox').addEventListener('click', (e) => {
      if (e.target.classList.contains('sig-label')) return;
      openSig();
    });

    // Tabs
    document.querySelectorAll('.sig-tab').forEach((tab) => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.sig-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        const target = tab.dataset.sigTab;
        $('sigTabUpload').hidden = (target !== 'upload');
        $('sigTabDraw').hidden = (target !== 'draw');
      });
    });

    // Upload File
    $('sigFileInput').addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        state.signature.dataUrl = reader.result;
        syncSigModalUI();
        renderSignatures();
        persistState(true);
        showToast('Signature image uploaded');
        e.target.value = '';
      };
      reader.readAsDataURL(file);
    });

    const restoreDefBtn = $('btnSigRestoreDefault');
    if (restoreDefBtn && !window.TOUR_DEFAULTS.defaultSignature) {
      restoreDefBtn.hidden = true;
    }

    $('btnSigRestoreDefault').addEventListener('click', () => {
      state.signature.dataUrl = window.TOUR_DEFAULTS.defaultSignature || '';
      syncSigModalUI();
      renderSignatures();
      persistState(true);
      showToast('Default Excel signature restored');
    });

    $('btnSigRemove').addEventListener('click', () => {
      state.signature.dataUrl = '';
      syncSigModalUI();
      renderSignatures();
      persistState(true);
      showToast('Signature removed');
    });

    widthRange.addEventListener('input', () => {
      state.signature.width = Number(widthRange.value);
      widthVal.textContent = `${state.signature.width}px`;
      renderSignatures();
      persistState();
    });

    blendCheck.addEventListener('change', () => {
      state.signature.blendMultiply = blendCheck.checked;
      renderSignatures();
      persistState();
    });

    showDateCheck.addEventListener('change', () => {
      const m = getActiveMonth();
      m.showSigDate = showDateCheck.checked;
      renderSignatures();
      persistState();
    });

    dateInput.addEventListener('input', () => {
      const m = getActiveMonth();
      m.sigDate = dateInput.value;
      renderSignatures();
      persistState();
    });

    $('btnSyncSigDate').addEventListener('click', () => {
      const m = getActiveMonth();
      m.sigDate = formatSigDate(m.endDate);
      dateInput.value = m.sigDate;
      m.showSigDate = true;
      showDateCheck.checked = true;
      renderSignatures();
      persistState();
    });

    // Freehand Signature Drawing Canvas
    const canvas = $('sigCanvas');
    const ctx = canvas.getContext('2d');
    let drawing = false;
    let inkColor = '#111111';
    let hasStroke = false;

    function getCanvasPos(ev) {
      const rect = canvas.getBoundingClientRect();
      const clientX = ev.touches ? ev.touches[0].clientX : ev.clientX;
      const clientY = ev.touches ? ev.touches[0].clientY : ev.clientY;
      return {
        x: (clientX - rect.left) * (canvas.width / rect.width),
        y: (clientY - rect.top) * (canvas.height / rect.height)
      };
    }

    function startDraw(ev) {
      ev.preventDefault();
      drawing = true;
      hasStroke = true;
      const pos = getCanvasPos(ev);
      ctx.beginPath();
      ctx.moveTo(pos.x, pos.y);
    }

    function moveDraw(ev) {
      if (!drawing) return;
      ev.preventDefault();
      const pos = getCanvasPos(ev);
      ctx.strokeStyle = inkColor;
      ctx.lineWidth = 2.6;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.lineTo(pos.x, pos.y);
      ctx.stroke();
    }

    function endDraw() {
      drawing = false;
    }

    canvas.addEventListener('mousedown', startDraw);
    canvas.addEventListener('mousemove', moveDraw);
    window.addEventListener('mouseup', endDraw);
    canvas.addEventListener('touchstart', startDraw, { passive: false });
    canvas.addEventListener('touchmove', moveDraw, { passive: false });
    canvas.addEventListener('touchend', endDraw);

    document.querySelectorAll('.ink-dot').forEach((dot) => {
      dot.addEventListener('click', () => {
        document.querySelectorAll('.ink-dot').forEach(d => d.classList.remove('active'));
        dot.classList.add('active');
        inkColor = dot.dataset.ink;
      });
    });

    $('btnClearCanvas').addEventListener('click', () => {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      hasStroke = false;
    });

    $('btnUseDrawnSig').addEventListener('click', () => {
      if (!hasStroke) {
        alert('Please draw your signature first.');
        return;
      }
      state.signature.dataUrl = canvas.toDataURL('image/png');
      syncSigModalUI();
      renderSignatures();
      persistState(true);
      // Switch back to upload preview tab so user sees the applied signature
      document.querySelector('.sig-tab[data-sig-tab="upload"]').click();
      showToast('Drawn signature applied');
    });
  }

  // ======================== MODAL 4: Paste to Multiple Dates ========================
  function initMultiPasteModal() {
    $('btnMultiPaste').addEventListener('click', () => {
      if (!clipboardDay) return;
      const m = getActiveMonth();
      const s = computeDayStats(clipboardDay.data);
      $('multiPasteSummary').textContent =
        `Pasting routes from ${formatSlashDate(clipboardDay.sourceDate)} (${clipboardDay.data.routes.length} routes, Total ${fmtTotalNum(s.dayExpense)}) into selected dates:`;

      const list = $('multiPasteList');
      list.innerHTML = '';
      m.days.forEach((day, idx) => {
        const empty = isDayEmpty(day);
        const dayStats = computeDayStats(day);
        const item = document.createElement('div');
        item.className = 'mp-item';
        item.innerHTML = `
          <label>
            <input type="checkbox" value="${idx}" data-empty="${empty ? '1' : '0'}" />
            <span><strong>${formatSlashDate(day.date)}</strong></span>
          </label>
          <span class="mp-badge ${empty ? 'empty' : ''}">${empty ? 'Empty' : `${day.routes.length} routes • ${fmtTotalNum(dayStats.dayExpense)}`}</span>
        `;
        list.appendChild(item);
      });

      openModal('multiPasteModal');
    });

    $('btnSelectEmptyDates').addEventListener('click', () => {
      $('multiPasteList').querySelectorAll('input[type="checkbox"]').forEach((cb) => {
        cb.checked = (cb.dataset.empty === '1');
      });
    });

    $('btnSelectNoneDates').addEventListener('click', () => {
      $('multiPasteList').querySelectorAll('input[type="checkbox"]').forEach((cb) => {
        cb.checked = false;
      });
    });

    $('btnConfirmMultiPaste').addEventListener('click', () => {
      if (!clipboardDay) return;
      const checked = [...$('multiPasteList').querySelectorAll('input[type="checkbox"]:checked')]
        .map(cb => Number(cb.value));
      if (!checked.length) {
        alert('Select at least one date.');
        return;
      }
      const m = getActiveMonth();
      for (const idx of checked) {
        if (m.days[idx]) {
          applyDayPayload(m.days[idx], clipboardDay.data);
        }
      }
      renderTable();
      checked.forEach((idx) => {
        const tb = $('expenseTable').querySelector(`tbody.day-group[data-day-idx="${idx}"]`);
        if (tb) tb.classList.add('just-pasted');
      });
      persistState(true);
      closeModal('multiPasteModal');
      showToast(`Pasted into ${checked.length} date(s)`);
    });
  }

  // ======================== Boot ========================
  function init() {
    state = loadState();
    bindTableEvents();
    bindHeaderAndToolbarEvents();
    bindModalCloseHandlers();
    initNewMonthModal();
    initPlacesModal();
    initSignatureModal();
    initMultiPasteModal();
    renderAll();

    if (window.TourCloudSync) {
      window.TourCloudSync.init({
        getState: () => state,
        getLocalUpdatedAt: () => Number(state.updatedAt) || 0,
        applyCloudState: (cloudState, updatedAt) => {
          if (!cloudState || !Array.isArray(cloudState.months) || !cloudState.months.length) return;
          state = cloudState;
          if (!state.profile) state.profile = deepClone(window.TOUR_DEFAULTS.profile);
          if (!state.signature) {
            state.signature = { dataUrl: window.TOUR_DEFAULTS.defaultSignature || '', width: 118, blendMultiply: true };
          }
          if (!Array.isArray(state.places)) state.places = [...window.TOUR_DEFAULTS.places];
          if (!state.months.some(m => m.id === state.activeMonthId)) {
            state.activeMonthId = state.months[0].id;
          }
          state.updatedAt = Number(updatedAt) || Date.now();
          clipboardDay = null;
          persistState(true, true);
          renderAll();
        },
        showToast
      });
    }
  }

  document.addEventListener('DOMContentLoaded', init);
})();
