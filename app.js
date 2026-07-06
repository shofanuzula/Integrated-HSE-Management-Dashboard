window.HSE = window.HSE || {};
HSE.app = {
  async init() {
    this.loadSchedules();
    HSE.state.search = localStorage.getItem('hseLastSearch') || HSE.state.search || '';
    HSE.data.set(this.initialModel(), HSE.workbookSeed ? 'HSE_Dashboard_Database.xlsx (Synced)' : 'No workbook loaded');
    this.renderNav();
    this.bind();
    this.populateFilters();
    this.tickClock();
    setInterval(() => this.tickClock(), 1000);
    this.render();
    await this.loadPublishedWorkbook();
  },
  initialModel() {
    return HSE.workbookSeed ? HSE.excel.parseRows(HSE.workbookSeed) : HSE.data.empty();
  },
  async loadPublishedWorkbook() {
    if (!/^https?:$/.test(window.location.protocol)) return;
    try {
      const response = await fetch(`./HSE_Dashboard_Database.xlsx?v=${Date.now()}`, { cache: 'no-store' });
      if (!response.ok) return;
      const file = new File([await response.arrayBuffer()], 'HSE_Dashboard_Database.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const model = await HSE.excel.load(file);
      const rowCount = Object.values(model).reduce((total, rows) => total + (Array.isArray(rows) ? rows.length : 0), 0);
      if (!rowCount) return;
      HSE.data.set(model, 'HSE_Dashboard_Database.xlsx (Published)', file);
      this.populateFilters();
      this.render();
    } catch (_) {
      // Snapshot data remains active when the published workbook is unavailable.
    }
  },
  renderNav() {
    document.getElementById('sideNav').innerHTML = HSE.config.pages.map(([id, label, icon]) => `<button data-page="${id}" class="${id === HSE.state.page ? 'active' : ''}"><i data-lucide="${icon}"></i><span class="nav-label">${label}</span></button>`).join('');
    if (window.lucide) lucide.createIcons();
  },
  bind() {
    document.getElementById('collapseSidebar').addEventListener('click', () => document.getElementById('sidebar').classList.toggle('collapsed'));
    document.getElementById('sideNav').addEventListener('click', (e) => { const btn = e.target.closest('button[data-page]'); if (btn) this.navigate(btn.dataset.page); });
    const searchInput = document.getElementById('globalSearch');
    const clearSearch = document.getElementById('clearSearch');
    const applySearch = HSE.helpers.debounce((value) => {
      HSE.state.search = value;
      localStorage.setItem('hseLastSearch', value);
      HSE.data.applyFilters();
      this.render();
    }, 300);
    searchInput.value = localStorage.getItem('hseLastSearch') || HSE.state.search || '';
    searchInput.closest('.search-box')?.classList.toggle('has-value', Boolean(searchInput.value));
    HSE.state.search = searchInput.value;
    searchInput.addEventListener('input', (e) => {
      e.currentTarget.closest('.search-box')?.classList.toggle('has-value', Boolean(e.currentTarget.value));
      applySearch(e.currentTarget.value);
    });
    clearSearch?.addEventListener('click', (e) => {
      e.preventDefault();
      searchInput.value = '';
      searchInput.closest('.search-box')?.classList.remove('has-value');
      HSE.state.search = '';
      localStorage.removeItem('hseLastSearch');
      document.querySelectorAll('#quickFilters button').forEach((b) => b.classList.remove('active'));
      HSE.data.applyFilters();
      this.render();
    });
    document.getElementById('quickFilters')?.addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-search]');
      if (!btn) return;
      document.querySelectorAll('#quickFilters button').forEach((b) => b.classList.toggle('active', b === btn));
      searchInput.value = btn.dataset.search;
      searchInput.closest('.search-box')?.classList.add('has-value');
      HSE.state.search = btn.dataset.search;
      localStorage.setItem('hseLastSearch', btn.dataset.search);
      HSE.data.applyFilters();
      this.render();
    });
    document.querySelectorAll('.filter-bar select, .filter-bar input').forEach((el) => el.addEventListener('change', () => { HSE.data.applyFilters(); this.render(); }));
    document.getElementById('excelInput').addEventListener('change', async (e) => this.importExcel(e.target.files?.[0]));
    document.getElementById('importExcelButton')?.addEventListener('click', (e) => {
      if (e.target?.id !== 'excelInput') {
        e.preventDefault();
        document.getElementById('excelInput').click();
      }
    });
    document.getElementById('importExcelButton')?.addEventListener('keydown', (e) => { if (e.key === 'Enter') document.getElementById('excelInput').click(); });
    document.getElementById('exportPdf').addEventListener('click', () => HSE.exporter.pdf());
    document.getElementById('exportExcel').addEventListener('click', () => this.exportChoice());
    document.getElementById('refreshData').addEventListener('click', () => this.refreshWorkbook());
    document.getElementById('darkMode').addEventListener('click', () => document.body.classList.toggle('dark'));
    document.getElementById('notificationButton').addEventListener('click', () => this.openNotifications());
    document.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        this.openCommandPalette();
      }
    });
  },
  async importExcel(file) {
    if (!file) return;
    try {
      this.setLoading(true, 'Importing workbook...');
      const model = await HSE.excel.load(file);
      const rowCount = Object.values(model).reduce((total, rows) => total + (Array.isArray(rows) ? rows.length : 0), 0);
      if (!rowCount) throw new Error('Workbook terbaca, tetapi tidak ada data master yang dapat diproses.');
      HSE.data.set(model, file.name, file);
      this.populateFilters();
      this.render();
      this.showValidation();
      HSE.ui.toast('Import berhasil', `${rowCount.toLocaleString('id-ID')} baris data dimuat dari ${file.name}.`);
    } catch (error) {
      HSE.ui.toast('Import failed', error.message || 'Excel format cannot be read.');
    } finally {
      this.setLoading(false);
      const input = document.getElementById('excelInput');
      if (input) input.value = '';
    }
  },
  async refreshWorkbook() {
    if (!HSE.state.workbookFile) {
      HSE.data.applyFilters();
      this.render();
      HSE.ui.toast('Refresh berhasil', 'Dashboard diperbarui dari cache saat ini.');
      return;
    }
    try {
      const activeFilters = HSE.data.getFilters();
      this.setLoading(true, 'Refreshing workbook...');
      const model = await HSE.excel.load(HSE.state.workbookFile);
      HSE.data.set(model, HSE.state.workbookFile.name, HSE.state.workbookFile);
      this.populateFilters();
      this.restoreFilterValues(activeFilters);
      this.render();
      this.showValidation();
      HSE.ui.toast('Refresh berhasil', 'Workbook berhasil dibaca ulang tanpa reload browser.');
    } catch (error) {
      HSE.ui.toast('Refresh gagal', error.message || 'Workbook tidak dapat dibaca ulang.');
    } finally {
      this.setLoading(false);
    }
  },
  restoreFilterValues(filters) {
    const map = { year: 'filterYear', month: 'filterMonth', date: 'filterDate', department: 'filterDepartment', area: 'filterArea', pic: 'filterPic', risk: 'filterRisk', status: 'filterStatus', priority: 'filterPriority', auditType: 'filterAuditType', incidentType: 'filterIncidentType', permitType: 'filterPermitType', project: 'filterProject', vendor: 'filterVendor' };
    Object.entries(map).forEach(([key, id]) => {
      const el = document.getElementById(id);
      if (el && filters[key] !== undefined && [...(el.options || [])].some((o) => o.value === filters[key])) el.value = filters[key];
      if (el?.type === 'date') el.value = filters[key] || '';
    });
  },
  showValidation() {
    const v = HSE.state.validation || { missingSheets: [], missingColumns: [], rowErrors: [] };
    const details = [...v.missingSheets.map((x) => `Sheet: ${x}`), ...v.missingColumns, ...(v.rowErrors || [])];
    HSE.state.validation.summary = details;
    return details.length;
  },
  exportChoice() {
    const pick = window.prompt('Export: workbook / schedule / both', 'workbook');
    if (pick === 'schedule') HSE.exporter.schedules();
    else if (pick === 'both') HSE.exporter.combined();
    else HSE.exporter.xlsx('hse-command-center-export.xlsx');
    HSE.ui.toast('Export berhasil', 'Data diekspor sesuai filter saat ini.');
  },
  setLoading(active, text = 'Loading...') {
    HSE.state.loading = active;
    const host = document.getElementById('pageHost');
    if (active && host) host.insertAdjacentHTML('afterbegin', `<div id="loadingOverlay" class="card skeleton mb-3"><strong>${HSE.helpers.esc(text)}</strong><div class="progress-line"><span style="width:72%"></span></div></div>`);
    if (!active) document.getElementById('loadingOverlay')?.remove();
  },
  openCommandPalette() {
    const commands = [
      ['Cari Incident', () => this.quickSearch('Incident')],
      ['Cari Finding', () => this.quickSearch('Finding')],
      ['Cari CAPA', () => this.quickSearch('CAPA')],
      ['Cari Permit', () => this.quickSearch('Permit')],
      ['Cari Project', () => this.quickSearch('Project')],
      ['Cari Department', () => this.quickSearch('Department')],
      ...HSE.config.pages.map(([id, label]) => [`Buka ${label}`, () => this.navigate(id)]),
      ['Import Workbook', () => document.getElementById('excelInput').click()],
      ['Export PDF', () => HSE.exporter.pdf()],
      ['Export Excel', () => HSE.exporter.xlsx()],
      ['Tambah Schedule', () => this.showCalendarModal(new Date(), [])],
    ];
    document.getElementById('detailModalTitle').textContent = 'Command Palette';
    document.getElementById('detailModalBody').innerHTML = `<div class="command-palette"><input id="commandSearch" class="form-control mb-3" autofocus aria-label="Command search"><div id="commandList" class="activity-list">${commands.map(([label], i) => `<button class="activity-item command-item" data-index="${i}"><strong>⌘K</strong><div><strong>${label}</strong><span>Execute command</span></div><span class="badge-soft badge-info">Enter</span></button>`).join('')}</div></div>`;
    const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('detailModal'));
    modal.show();
    const render = () => {
      const q = HSE.helpers.norm(document.getElementById('commandSearch').value);
      document.querySelectorAll('.command-item').forEach((btn) => {
        btn.style.display = HSE.helpers.norm(btn.innerText).includes(q) ? '' : 'none';
      });
    };
    setTimeout(() => document.getElementById('commandSearch')?.focus(), 180);
    document.getElementById('commandSearch').addEventListener('input', render);
    document.querySelectorAll('.command-item').forEach((btn) => btn.addEventListener('click', () => {
      const fn = commands[Number(btn.dataset.index)][1];
      modal.hide();
      fn();
    }));
  },
  quickSearch(type) {
    const input = document.getElementById('globalSearch');
    input.value = type;
    HSE.state.search = type;
    HSE.data.applyFilters();
    this.render();
  },
  populateFilters() {
    const all = Object.values(HSE.state.raw).flat(), h = HSE.helpers;
    const fill = (id, values) => { document.getElementById(id).innerHTML = values.map((v) => `<option>${h.esc(v)}</option>`).join(''); };
    fill('filterYear', ['All', '2026', ...h.unique(all.map((r) => r.date?.getFullYear()).filter(Boolean))]);
    fill('filterMonth', ['All', ...HSE.config.months]);
    fill('filterDepartment', ['All', ...h.unique([...HSE.config.departments, ...all.map((r) => r.department)])]);
    fill('filterArea', ['All', ...h.unique([...HSE.config.areas, ...all.map((r) => r.area)])]);
    fill('filterPic', ['All', ...h.unique(all.map((r) => r.pic))]);
    fill('filterRisk', ['All', ...HSE.config.risks]);
    fill('filterStatus', ['All', ...h.unique([...HSE.config.statuses, ...all.map((r) => r.status)])]);
    fill('filterPriority', ['All', ...HSE.config.priorities]);
    fill('filterAuditType', ['All', ...HSE.config.auditTypes]);
    fill('filterIncidentType', ['All', ...HSE.config.incidentTypes]);
    fill('filterPermitType', ['All', ...HSE.config.permitTypes]);
    fill('filterProject', ['All', ...h.unique(HSE.state.raw.project.map((r) => r.name))]);
    fill('filterVendor', ['All', ...h.unique([...HSE.state.raw.environment.map((r) => r.vendor), ...HSE.state.raw.permit.map((r) => r.vendor), ...HSE.state.raw.project.map((r) => r.vendor)])]);
  },
  render() {
    HSE.data.applyFilters();
    const reminders = this.scheduleReminders();
    HSE.state.notifications = this.buildNotifications(reminders);
    document.getElementById('notificationCount').textContent = HSE.state.notifications.length;
    this.renderNav();
    HSE.pages.render();
    this.drawSparks();
  },
  navigate(page) { HSE.state.page = page; this.render(); },
  openDetail(id) { this.showDetail({ Chart: id, Page: HSE.state.page, Filter: HSE.data.getFilters() }); },
  showDetail(data) {
    document.getElementById('detailModalTitle').textContent = 'Detail';
    document.getElementById('detailModalBody').innerHTML = `<pre class="mb-0">${HSE.helpers.esc(JSON.stringify(data, null, 2))}</pre>`;
    bootstrap.Modal.getOrCreateInstance(document.getElementById('detailModal')).show();
  },
  buildNotifications(reminders = []) {
    const d = HSE.state.filtered;
    const findings = [...d.finding, ...d.audit];
    return [
      ...findings.filter((r) => r.status === 'Overdue').slice(0, 5).map((r) => ({ title: `Finding Overdue: ${r.id}`, page: 'audit', tone: 'Critical' })),
      ...d.capa.filter((r) => r.status === 'Overdue').slice(0, 5).map((r) => ({ title: `CAPA Overdue: ${r.id}`, page: 'capa', tone: 'Critical' })),
      ...d.apar.filter((r) => /expired/i.test(`${r.status} ${r.condition}`)).slice(0, 5).map((r) => ({ title: `APAR Expired: ${r.id || r.aparNumber}`, page: 'analytics', tone: 'Critical' })),
      ...d.permit.filter((r) => r.status === 'Expired').slice(0, 5).map((r) => ({ title: `Permit Expired: ${r.id}`, page: 'permit', tone: 'Critical' })),
      ...d.project.filter((r) => r.status === 'Delayed').slice(0, 5).map((r) => ({ title: `Project Delayed: ${r.id}`, page: 'project', tone: 'High' })),
      ...d.safety.filter((r) => ['Critical', 'Fatality', 'Lost Time Injury'].includes(r.severity) || ['Fatality', 'Lost Time Injury'].includes(r.incidentType)).slice(0, 5).map((r) => ({ title: `Critical Incident: ${r.id}`, page: 'safety', tone: 'Critical' })),
      ...reminders.map((s) => {
        const today = HSE.helpers.fmtDate(new Date());
        const tomorrow = new Date(); tomorrow.setDate(tomorrow.getDate() + 1);
        const label = s.date === today ? 'Schedule Hari Ini' : HSE.helpers.fmtDate(tomorrow) === s.date ? 'Schedule Besok' : 'Schedule';
        return { title: `${label}: ${s.startTime || '-'} - ${s.title}`, page: 'executive', tone: s.priority || 'Medium' };
      }),
    ];
  },
  openNotifications() {
    const items = HSE.state.notifications;
    document.getElementById('detailModalTitle').textContent = 'Notification Center';
    document.getElementById('detailModalBody').innerHTML = items.length ? `<div class="activity-list">${items.map((n, i) => `<button class="activity-item notification-item" data-page="${n.page}"><strong>${i + 1}</strong><div><strong>${HSE.helpers.esc(n.title)}</strong><span>${HSE.helpers.esc(n.page)}</span></div><span class="badge-soft ${HSE.ui.badge(n.tone)}">${HSE.helpers.esc(n.tone)}</span></button>`).join('')}</div>` : '<div class="empty-state">Tidak ada notifikasi.</div>';
    document.querySelectorAll('.notification-item').forEach((el) => el.addEventListener('click', () => { bootstrap.Modal.getOrCreateInstance(document.getElementById('detailModal')).hide(); this.navigate(el.dataset.page); }));
    bootstrap.Modal.getOrCreateInstance(document.getElementById('detailModal')).show();
  },
  showCalendarModal(date, activities = []) {
    const selected = date || new Date();
    document.getElementById('detailModalTitle').textContent = 'HSE Calendar';
    document.getElementById('detailModalBody').innerHTML = `
      <ul class="nav nav-tabs" role="tablist">
        <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#calActivities" type="button">Activities</button></li>
        <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#calAdd" type="button">Add Schedule</button></li>
      </ul>
      <div class="tab-content pt-3">
        <div id="calActivities" class="tab-pane fade show active">${this.scheduleList(activities, selected)}</div>
        <div id="calAdd" class="tab-pane fade">${this.scheduleForm({ date: HSE.helpers.fmtDate(selected) })}</div>
      </div>`;
    this.bindScheduleModal();
    bootstrap.Modal.getOrCreateInstance(document.getElementById('detailModal')).show();
  },
  scheduleList(items, date) {
    if (!items.length) return '<div class="empty-state">Tidak ada aktivitas hari ini.</div>';
    return `<div class="activity-list">${items.sort((a,b) => String(a.startTime || a.time).localeCompare(String(b.startTime || b.time))).map((a) => `<div class="activity-item"><strong>${a.startTime || a.time || '-'}</strong><div><strong>${HSE.helpers.esc(a.title || a.name)}</strong><span>${HSE.helpers.esc(a.department || '-')} • ${HSE.helpers.esc(a.area || '-')} • ${HSE.helpers.esc(a.pic || '-')}</span></div><span class="badge-soft ${HSE.ui.badge(a.priority || a.status)}">${HSE.helpers.esc(a.status || '-')}</span>${a.manual ? `<div class="schedule-actions full"><button class="action-button schedule-edit" data-id="${a.id}">Edit</button><button class="action-button schedule-duplicate" data-id="${a.id}">Duplicate</button><button class="action-button schedule-copy" data-id="${a.id}">Copy</button><button class="action-button schedule-complete" data-id="${a.id}">Completed</button><button class="action-button schedule-cancel" data-id="${a.id}">Cancelled</button><button class="action-button schedule-delete" data-id="${a.id}">Delete</button></div>` : ''}</div>`).join('')}</div>`;
  },
  scheduleForm(schedule = {}) {
    const c = ['Audit', 'Inspection', '5R Audit', 'APD Inspection', 'APAR Inspection', 'Safety Patrol', 'Safety Talk', 'Meeting', 'Project', 'Maintenance', 'Vendor Visit', 'Calibration', 'Emergency Drill', 'Training', 'Campaign', 'Other'];
    return `<form id="scheduleForm" class="schedule-form">
      <input type="hidden" name="id" value="${schedule.id || ''}">
      <label>Judul<input required name="title" value="${HSE.helpers.esc(schedule.title || '')}"></label>
      <label>Kategori<select name="category">${c.map((x) => `<option ${schedule.category === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
      <label>Tanggal<input required type="date" name="date" value="${schedule.date || HSE.helpers.fmtDate(new Date())}"></label>
      <label>Jam Mulai<input required type="time" name="startTime" value="${schedule.startTime || '08:00'}"></label>
      <label>Jam Selesai<input required type="time" name="endTime" value="${schedule.endTime || '09:00'}"></label>
      <label>Department<select name="department">${HSE.config.departments.map((x) => `<option ${schedule.department === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
      <label>Area<select name="area">${HSE.config.areas.map((x) => `<option ${schedule.area === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
      <label>PIC<input name="pic" value="${HSE.helpers.esc(schedule.pic || '')}"></label>
      <label>Prioritas<select name="priority">${['Low', 'Medium', 'High', 'Critical'].map((x) => `<option ${schedule.priority === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
      <label>Status<select name="status">${['Scheduled', 'In Progress', 'Completed', 'Cancelled'].map((x) => `<option ${schedule.status === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
      <label>Reminder<input name="reminder" value="${HSE.helpers.esc(schedule.reminder || '')}"></label>
      <label>Lampiran (Opsional)<input name="attachment" value="${HSE.helpers.esc(schedule.attachment || '')}"></label>
      <label class="full">Deskripsi<textarea name="description">${HSE.helpers.esc(schedule.description || '')}</textarea></label>
      <div class="schedule-actions full"><button class="action-button" type="submit">Save Schedule</button></div>
    </form>`;
  },
  bindScheduleModal() {
    document.getElementById('scheduleForm')?.addEventListener('submit', (e) => {
      e.preventDefault();
      const data = Object.fromEntries(new FormData(e.target).entries());
      this.saveSchedule(data);
      bootstrap.Modal.getOrCreateInstance(document.getElementById('detailModal')).hide();
    });
    document.querySelectorAll('.schedule-delete').forEach((b) => b.addEventListener('click', () => this.deleteSchedule(b.dataset.id)));
    document.querySelectorAll('.schedule-complete').forEach((b) => b.addEventListener('click', () => this.updateScheduleStatus(b.dataset.id, 'Completed')));
    document.querySelectorAll('.schedule-cancel').forEach((b) => b.addEventListener('click', () => this.updateScheduleStatus(b.dataset.id, 'Cancelled')));
    document.querySelectorAll('.schedule-duplicate').forEach((b) => b.addEventListener('click', () => this.duplicateSchedule(b.dataset.id)));
    document.querySelectorAll('.schedule-copy').forEach((b) => b.addEventListener('click', () => this.copySchedule(b.dataset.id)));
    document.querySelectorAll('.schedule-edit').forEach((b) => b.addEventListener('click', () => {
      const item = HSE.state.schedules.find((s) => s.id === b.dataset.id);
      document.querySelector('[data-bs-target="#calAdd"]').click();
      document.getElementById('calAdd').innerHTML = this.scheduleForm(item);
      this.bindScheduleModal();
    }));
  },
  saveSchedule(data) {
    const item = { ...data, id: data.id || `SCH-${Date.now()}`, status: data.status || 'Scheduled', manual: true, type: 'Schedule', dateObj: new Date(`${data.date}T${data.startTime || '08:00'}`) };
    const idx = HSE.state.schedules.findIndex((s) => s.id === item.id);
    if (idx >= 0) HSE.state.schedules[idx] = item; else HSE.state.schedules.push(item);
    this.persistSchedules();
    this.render();
    HSE.ui.toast('Schedule saved', `${item.title} tersimpan.`);
  },
  deleteSchedule(id) {
    if (!window.confirm('Hapus schedule ini?')) return;
    HSE.state.schedules = HSE.state.schedules.filter((s) => s.id !== id);
    this.persistSchedules();
    this.render();
    HSE.ui.toast('Schedule deleted', 'Schedule berhasil dihapus.');
  },
  updateScheduleStatus(id, status) { const item = HSE.state.schedules.find((s) => s.id === id); if (item) item.status = status; this.persistSchedules(); this.render(); HSE.ui.toast('Schedule updated', `Status menjadi ${status}.`); },
  duplicateSchedule(id) {
    const item = HSE.state.schedules.find((s) => s.id === id);
    if (!item) return;
    HSE.state.schedules.push({ ...item, id: `SCH-${Date.now()}`, title: `${item.title} Copy` });
    this.persistSchedules();
    this.render();
    HSE.ui.toast('Schedule duplicated', 'Schedule berhasil diduplikasi.');
  },
  copySchedule(id) {
    const item = HSE.state.schedules.find((s) => s.id === id);
    if (!item) return;
    navigator.clipboard?.writeText(JSON.stringify(item, null, 2));
    HSE.ui.toast('Schedule copied', 'Detail schedule disalin.');
  },
  loadSchedules() { try { HSE.state.schedules = JSON.parse(localStorage.getItem('hseSchedules') || '[]'); } catch (_) { HSE.state.schedules = []; } },
  persistSchedules() { localStorage.setItem('hseSchedules', JSON.stringify(HSE.state.schedules)); },
  scheduleReminders() {
    const today = new Date(); today.setHours(0,0,0,0);
    const tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);
    return HSE.state.schedules.filter((s) => { const d = new Date(s.date); d.setHours(0,0,0,0); return d.getTime() === today.getTime() || d.getTime() === tomorrow.getTime(); });
  },
  reset() {
    HSE.state.search = '';
    document.getElementById('globalSearch').value = '';
    HSE.data.set(this.initialModel(), HSE.workbookSeed ? 'HSE_Dashboard_Database.xlsx (Synced)' : 'No workbook loaded');
    this.populateFilters();
    this.render();
    HSE.ui.toast('Dashboard reset', 'Dashboard kembali ke dummy data.');
  },
  resetFilters() {
    HSE.state.search = '';
    document.getElementById('globalSearch').value = '';
    document.querySelectorAll('.filter-bar select').forEach((el) => el.value = 'All');
    document.getElementById('filterDate').value = '';
    HSE.data.applyFilters();
    this.render();
    HSE.ui.toast('Filter reset', 'Seluruh filter dikembalikan ke All.');
  },
  backupSchedules() {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(HSE.state.schedules, null, 2)], { type: 'application/json' }));
    a.download = 'hse-schedule-backup.json';
    a.click();
  },
  async restoreSchedules(file) {
    if (!file) return;
    try {
      HSE.state.schedules = JSON.parse(await file.text());
      this.persistSchedules();
      this.render();
      HSE.ui.toast('Schedule restored', 'Backup schedule berhasil dipulihkan.');
    } catch (error) {
      HSE.ui.toast('Restore gagal', 'File backup schedule tidak valid.');
    }
  },
  clearCache() {
    localStorage.removeItem('hseSchedules');
    HSE.state.schedules = [];
    this.render();
    HSE.ui.toast('Cache cleared', 'Cache schedule lokal dibersihkan.');
  },
  tickClock() {
    const now = new Date();
    document.getElementById('digitalClock').textContent = now.toLocaleTimeString('id-ID');
    document.getElementById('todayDate').textContent = now.toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  },
  drawSparks() {
    const base = HSE.helpers.byMonth(HSE.state.filtered.audit);
    document.querySelectorAll('.sparkline').forEach((canvas, i) => {
      new Chart(canvas, {
        type: 'line',
        data: { labels: HSE.config.months, datasets: [{ data: base.map((v, idx) => Math.max(0, v + ((i + idx) % 5) - 2)), borderColor: '#2563eb', borderWidth: 1.5, pointRadius: 0, tension: .35 }] },
        options: { responsive: false, plugins: { legend: { display: false }, tooltip: { enabled: false } }, scales: { x: { display: false }, y: { display: false } } },
      });
    });
  },
};
document.addEventListener('DOMContentLoaded', () => HSE.app.init());
