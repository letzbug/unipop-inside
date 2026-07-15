(() => {
  'use strict';

  const cfg = window.UNIPOP_CONFIG || {};
  const state = {
    supabase: null,
    online: false,
    activeImport: null,
    courses: [],
    modifications: [],
    imports: [],
    generatedBlob: null,
    generatedFilename: '',
    generatedFilePath: '',
    originalWorkbook: null,
    selectedCourse: null,
    passwordAction: null,
  };

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

  const routes = ['home', 'import', 'export-links', 'courses', 'modifications', 'history', 'settings'];
  const fieldAliases = {
    courseId: ['cours id', 'id cours'],
    schoolYear: ['annee scolaire'],
    title: ['intitule'],
    level: ['niveau'],
    startDate: ['date de debut'],
    endDate: ['date de fin'],
    totalDuration: ['duree totale heures', 'duree totale'],
    schedule: ['horaires', 'horaire'],
    places: ['nb places', 'nombre de places'],
    description: ['description du cours'],
    additionalInfo: ['renseignements complementaires'],
    locationName: ['lieu de formation nom'],
    locationRoom: ['lieu de formation salle'],
    link: ['link', 'lien'],
    qr: ['qr code', 'qrcode'],
    trainer: ['formateur s paiement', 'formateurs paiement', 'formateur'],
    category: ['categorie'],
    subject: ['matiere']
  };

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    bindNavigation();
    bindUpload();
    bindSearch();
    bindPasswords();
    bindMisc();
    setupSupabase();
    await loadData();
    renderAll();
    routeFromHash();
  }

  function setupSupabase() {
    if (cfg.supabaseUrl && cfg.supabaseAnonKey && window.supabase) {
      state.supabase = window.supabase.createClient(cfg.supabaseUrl, cfg.supabaseAnonKey);
      state.online = true;
    }
  }

  async function loadData() {
    if (state.online) {
      try {
        const { data: imports, error: impErr } = await state.supabase.from('imports').select('*').order('created_at', { ascending: false });
        if (impErr) throw impErr;
        state.imports = imports || [];
        state.activeImport = state.imports.find(i => i.is_active) || null;
      state.generatedFilename = state.activeImport?.generated_filename || '';
      state.generatedFilePath = state.activeImport?.generated_file_path || '';
        if (state.activeImport) {
          const [{ data: courses, error: cErr }, { data: mods, error: mErr }] = await Promise.all([
            state.supabase.from('courses').select('*').eq('import_id', state.activeImport.id).order('title'),
            state.supabase.from('modifications').select('*').eq('import_id', state.activeImport.id).order('created_at', { ascending: false })
          ]);
          if (cErr) throw cErr;
          if (mErr) throw mErr;
          state.courses = courses || [];
          state.modifications = mods || [];
        }
      } catch (err) {
        console.error(err);
        state.online = false;
        toast('Connexion Supabase indisponible. Passage en mode local.', true);
        loadLocal();
      }
    } else {
      loadLocal();
    }
  }

  function loadLocal() {
    state.activeImport = safeJson(localStorage.getItem('unipop_active_import'));
    state.courses = safeJson(localStorage.getItem('unipop_courses')) || [];
    state.modifications = safeJson(localStorage.getItem('unipop_modifications')) || [];
    state.imports = safeJson(localStorage.getItem('unipop_imports')) || [];
    const cached = localStorage.getItem('unipop_generated_file');
    if (cached) {
      try {
        const meta = JSON.parse(cached);
        state.generatedFilename = meta.filename || '';
      } catch (_) {}
    }
  }

  function saveLocal() {
    localStorage.setItem('unipop_active_import', JSON.stringify(state.activeImport));
    localStorage.setItem('unipop_courses', JSON.stringify(state.courses));
    localStorage.setItem('unipop_modifications', JSON.stringify(state.modifications));
    localStorage.setItem('unipop_imports', JSON.stringify(state.imports));
    localStorage.setItem('unipop_generated_file', JSON.stringify({ filename: state.generatedFilename }));
  }

  function bindNavigation() {
    window.addEventListener('hashchange', routeFromHash);
    $$('[data-go]').forEach(btn => btn.addEventListener('click', () => location.hash = btn.dataset.go));
    $('#mobileMenu').addEventListener('click', () => $('#sidebar').classList.toggle('open'));
  }

  function routeFromHash() {
    const route = (location.hash || '#home').slice(1);
    showRoute(routes.includes(route) ? route : 'home');
  }

  function showRoute(route) {
    $$('.view').forEach(v => v.classList.toggle('active', v.id === `view-${route}`));
    $$('.nav-link').forEach(n => n.classList.toggle('active', n.dataset.route === route));
    $('#sidebar').classList.remove('open');
    if (route === 'modifications') renderModifications();
    if (route === 'history') renderHistory();
    if (route === 'courses') $('#courseSearch').focus();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function bindUpload() {
    const input = $('#excelInput');
    const zone = $('#dropZone');
    $('#chooseFile').addEventListener('click', () => input.click());
    input.addEventListener('change', () => input.files[0] && handleUpload(input.files[0]));
    ['dragenter', 'dragover'].forEach(evt => zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach(evt => zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.remove('dragover'); }));
    zone.addEventListener('drop', e => e.dataTransfer.files[0] && handleUpload(e.dataTransfer.files[0]));
  }

  async function handleUpload(file) {
    if (!file.name.toLowerCase().endsWith('.xlsx')) return toast('Veuillez sélectionner un fichier .xlsx.', true);
    if (file.size > (cfg.maxFileSizeMb || 25) * 1024 * 1024) return toast(`Le fichier dépasse ${cfg.maxFileSizeMb || 25} Mo.`, true);

    $('#importResult').classList.add('hidden');
    $('#importProgress').classList.remove('hidden');
    updateProgress(5, 'Lecture du fichier Excel…', file.name);

    try {
      const buffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      state.originalWorkbook = workbook;
      updateProgress(15, 'Analyse des onglets…', `${file.name} · ${workbook.worksheets.length} onglet(s) : ${workbook.worksheets.map(ws => ws.name).join(' | ')}`);

      const courses = [];
      const sheetsInfo = [];
      let totalRows = workbook.worksheets.reduce((n, ws) => n + Math.max(0, ws.rowCount - 1), 0);
      let processed = 0;
      const skippedSheets = [];

      for (const worksheet of workbook.worksheets) {
        const headerMap = detectHeaders(worksheet);
        const missingHeaders = getMissingRequiredHeaders(headerMap);
        if (missingHeaders.length) {
          skippedSheets.push({ name: worksheet.name, missing: missingHeaders });
          sheetsInfo.push({ name: worksheet.name, count: 0, skipped: true, missing: missingHeaders });
          continue;
        }
        ensureOutputColumns(worksheet, headerMap);
        const sheetStartCount = courses.length;
        const firstDataRow = (headerMap.__headerRow || 1) + 1;

        for (let r = firstDataRow; r <= worksheet.rowCount; r++) {
          const row = worksheet.getRow(r);
          const courseId = cellText(row.getCell(headerMap.courseId));
          const title = fixMojibake(cellText(row.getCell(headerMap.title)));
          if (!courseId && !title) continue;

          const course = extractCourse(row, headerMap, worksheet.name, r);
          course.link = buildFormationUrl(course);
          row.getCell(headerMap.link).value = course.link;
          row.getCell(headerMap.link).alignment = { wrapText: true, vertical: 'middle' };
          row.getCell(headerMap.link).font = { color: { argb: 'FF1455D9' }, underline: true };

          const qrDataUrl = await createQrDataUrl(course.link);
          const imageId = workbook.addImage({ base64: qrDataUrl, extension: 'png' });
          worksheet.addImage(imageId, {
            tl: { col: headerMap.qr - 1 + 0.08, row: r - 1 + 0.08 },
            ext: { width: 86, height: 86 }
          });
          row.height = Math.max(row.height || 15, 70);
          worksheet.getColumn(headerMap.link).width = Math.max(worksheet.getColumn(headerMap.link).width || 10, 48);
          worksheet.getColumn(headerMap.qr).width = Math.max(worksheet.getColumn(headerMap.qr).width || 10, 14);
          course.qr_data = qrDataUrl;
          courses.push(course);
          processed++;
          if (processed % 10 === 0 || processed === totalRows) {
            const pct = Math.min(82, 15 + Math.round((processed / Math.max(1, totalRows)) * 67));
            updateProgress(pct, 'Création des liens et QR codes…', `${processed} ligne(s) analysée(s)`);
            await pause(0);
          }
        }
        sheetsInfo.push({ name: worksheet.name, count: courses.length - sheetStartCount });
      }

      if (!courses.length) {
        const details = skippedSheets.map(s => `« ${s.name} » (${s.missing.join(', ')})`).join(' ; ');
        throw new Error(`Le fichier « ${file.name} » ne contient aucun onglet de cours reconnu. Onglets réellement lus : ${workbook.worksheets.map(ws => `« ${ws.name} »`).join(', ')}.${details ? ` Détail : ${details}` : ''}`);
      }

      updateProgress(86, 'Préparation du nouveau fichier…', `${courses.length} cours trouvés${skippedSheets.length ? ` · ${skippedSheets.length} onglet(s) ignoré(s)` : ''}`);
      const outputBuffer = await workbook.xlsx.writeBuffer();
      state.generatedBlob = new Blob([outputBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      state.generatedFilename = `${file.name.replace(/\.xlsx$/i, '')}_Links_QR.xlsx`;

      const importRecord = {
        id: crypto.randomUUID(),
        original_filename: file.name,
        generated_filename: state.generatedFilename,
        generated_file_path: '',
        school_year: mostCommon(courses.map(c => c.school_year).filter(Boolean)) || '',
        course_count: courses.length,
        sheet_count: workbook.worksheets.length,
        is_active: true,
        created_at: new Date().toISOString()
      };

      updateProgress(91, 'Activation de la nouvelle base…', state.online ? 'Enregistrement sécurisé…' : 'Enregistrement local…');
      await activateImport(importRecord, courses);
      updateProgress(100, 'Importation terminée', `${courses.length} cours · ${workbook.worksheets.length} onglet(s)`);
      renderImportResult(importRecord, sheetsInfo);
      renderAll();
      toast('Le nouveau fichier est maintenant la base active.');
    } catch (err) {
      console.error(err);
      updateProgress(0, 'Importation impossible', err.message || 'Erreur inconnue');
      toast(err.message || 'Le fichier n’a pas pu être traité.', true);
    }
  }

  async function activateImport(importRecord, courses) {
    if (state.online) {
      const { error: archiveErr } = await state.supabase.from('imports').update({ is_active: false }).eq('is_active', true);
      if (archiveErr) throw archiveErr;
      const { data: inserted, error: importErr } = await state.supabase.from('imports').insert(importRecord).select().single();
      if (importErr) throw importErr;
      importRecord = inserted;

      if (state.generatedBlob) {
        const safeName = slugify(state.generatedFilename.replace(/\.xlsx$/i, '')) || 'fichier-links-qr';
        const storagePath = `${importRecord.id}/${safeName}.xlsx`;
        const { error: uploadErr } = await state.supabase.storage
          .from(cfg.storageBucket || 'unipop-files')
          .upload(storagePath, state.generatedBlob, {
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            upsert: true
          });
        if (uploadErr) throw new Error(`Stockage du fichier impossible : ${uploadErr.message}`);
        const { data: updated, error: updateErr } = await state.supabase
          .from('imports')
          .update({ generated_file_path: storagePath })
          .eq('id', importRecord.id)
          .select()
          .single();
        if (updateErr) throw updateErr;
        importRecord = updated;
        state.generatedFilePath = storagePath;
      }

      const payload = courses.map(c => ({ ...c, import_id: importRecord.id, id: crypto.randomUUID() }));
      for (let i = 0; i < payload.length; i += 100) {
        const { error } = await state.supabase.from('courses').insert(payload.slice(i, i + 100));
        if (error) throw error;
      }
      state.imports.forEach(i => i.is_active = false);
      state.imports.unshift(importRecord);
      state.activeImport = importRecord;
      state.courses = payload;
      state.modifications = [];
    } else {
      state.imports.forEach(i => i.is_active = false);
      state.imports.unshift(importRecord);
      state.activeImport = importRecord;
      state.courses = courses.map(c => ({ ...c, id: crypto.randomUUID(), import_id: importRecord.id }));
      state.modifications = [];
      saveLocal();
    }
  }

  function detectHeaders(worksheet) {
    const required = ['courseId', 'schoolYear', 'title', 'level'];
    const maxHeaderRow = Math.min(50, Math.max(1, worksheet.rowCount || 1));
    let bestMap = {};
    let bestScore = -1;

    for (let rowNumber = 1; rowNumber <= maxHeaderRow; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      const candidate = { __headerRow: rowNumber };
      const normalizedByColumn = {};
      const maxColumn = Math.max(worksheet.columnCount || 0, row.cellCount || 0, 40);

      for (let col = 1; col <= maxColumn; col++) {
        const normalized = normalizeHeader(cellText(row.getCell(col)));
        normalizedByColumn[col] = normalized;
        if (!normalized) continue;

        for (const [key, aliases] of Object.entries(fieldAliases)) {
          if (candidate[key]) continue;
          if (aliases.some(alias => headerMatches(normalized, alias))) {
            candidate[key] = col;
          }
        }
      }

      // Structure exacte des exports UniPop vérifiée sur le fichier réel :
      // A=Cours ID, B=Année scolaire, E=Intitulé, F=Niveau.
      // Le fallback n'est activé que si la ligne ressemble réellement à une
      // ligne d'en-tête UniPop, afin de ne jamais confondre une ligne de données.
      const uniPopHeaderSignals = [
        headerMatches(normalizedByColumn[1] || '', 'cours id'),
        headerMatches(normalizedByColumn[2] || '', 'annee scolaire'),
        headerMatches(normalizedByColumn[5] || '', 'intitule'),
        headerMatches(normalizedByColumn[6] || '', 'niveau')
      ].filter(Boolean).length;

      if (uniPopHeaderSignals >= 3) {
        const knownColumns = {
          courseId: 1,
          schoolYear: 2,
          category: 3,
          subject: 4,
          title: 5,
          level: 6,
          description: 8,
          schedule: 9,
          startDate: 11,
          endDate: 12,
          totalDuration: 13,
          places: 15,
          additionalInfo: 17,
          locationName: 19,
          locationRoom: 20,
          trainer: 35,
          link: 36,
          qr: 37
        };
        for (const [key, col] of Object.entries(knownColumns)) {
          if (!candidate[key] && col <= Math.max(worksheet.columnCount || 0, 37)) {
            candidate[key] = col;
          }
        }
      }

      const requiredHits = required.filter(key => Number.isInteger(candidate[key])).length;
      const optionalHits = Object.keys(candidate)
        .filter(key => key !== '__headerRow' && !required.includes(key) && Number.isInteger(candidate[key]))
        .length;
      const nonEmptyHeaders = Object.values(normalizedByColumn).filter(Boolean).length;
      const score = requiredHits * 1000 + optionalHits * 10 + Math.min(nonEmptyHeaders, 50);

      if (score > bestScore) {
        bestScore = score;
        bestMap = candidate;
      }

      if (requiredHits === required.length && optionalHits >= 6) break;
    }

    return bestMap;
  }

  function headerMatches(value, alias) {
    const normalizedValue = normalizeHeader(value);
    const normalizedAlias = normalizeHeader(alias);
    if (!normalizedValue || !normalizedAlias) return false;
    return normalizedValue === normalizedAlias ||
      normalizedValue.startsWith(`${normalizedAlias} `) ||
      normalizedValue.endsWith(` ${normalizedAlias}`);
  }

  function ensureOutputColumns(worksheet, map) {
    if (!map.link) {
      map.link = worksheet.columnCount + 1;
      const headerRow = map.__headerRow || 1;
      const c = worksheet.getCell(headerRow, map.link);
      c.value = 'Link';
      cloneHeaderStyle(worksheet.getCell(headerRow, Math.max(1, map.link - 1)), c);
    }
    if (!map.qr) {
      map.qr = Math.max(worksheet.columnCount + 1, map.link + 1);
      const headerRow = map.__headerRow || 1;
      const c = worksheet.getCell(headerRow, map.qr);
      c.value = 'QR-Code';
      cloneHeaderStyle(worksheet.getCell(headerRow, Math.max(1, map.qr - 1)), c);
    }
  }

  function cloneHeaderStyle(source, target) {
    try {
      target.style = JSON.parse(JSON.stringify(source.style || {}));
      target.alignment = { ...(target.alignment || {}), horizontal: 'center', vertical: 'middle' };
    } catch (_) {
      target.font = { bold: true };
    }
  }

  function getMissingRequiredHeaders(map) {
    const labels = {
      courseId: 'Cours ID',
      schoolYear: 'Année Scolaire',
      title: 'Intitulé',
      level: 'Niveau'
    };
    return Object.keys(labels).filter(key => !map[key]).map(key => labels[key]);
  }

  function extractCourse(row, map, sheetName, rowNumber) {
    const get = key => map[key] ? fixMojibake(cellText(row.getCell(map[key]))) : '';
    return {
      source_sheet: sheetName,
      source_row: rowNumber,
      course_id: get('courseId'),
      school_year: get('schoolYear'),
      title: get('title'),
      level: get('level'),
      start_date: excelDateToIso(row.getCell(map.startDate)?.value),
      end_date: excelDateToIso(row.getCell(map.endDate)?.value),
      total_duration: get('totalDuration'),
      schedule: get('schedule'),
      places: get('places'),
      description: get('description'),
      additional_info: get('additionalInfo'),
      location_name: get('locationName'),
      location_room: get('locationRoom'),
      trainer: get('trainer'),
      category: get('category'),
      subject: get('subject'),
      link: '',
      qr_data: ''
    };
  }

  function buildFormationUrl(course) {
    const raw = [course.title, course.level, course.course_id, (course.school_year || '').replace('/', '-')].filter(Boolean).join('-');
    return `${cfg.baseFormationUrl || 'https://www.unipop.lu/formations/'}${slugify(raw)}/`;
  }

  function slugify(value) {
    return String(value)
      .replace(/[’']/g, '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-|-$/g, '');
  }

  async function createQrDataUrl(text) {
    await ensureQrCodeLibrary();
    if (!window.QRCode || typeof window.QRCode.toDataURL !== 'function') {
      throw new Error('Le générateur de QR codes n’a pas pu être chargé. Vérifiez la connexion Internet puis rechargez la page.');
    }
    return window.QRCode.toDataURL(text, {
      width: 240,
      margin: 1,
      errorCorrectionLevel: 'M'
    });
  }

  async function ensureQrCodeLibrary() {
    if (window.QRCode && typeof window.QRCode.toDataURL === 'function') return;

    const sources = [
      'https://cdnjs.cloudflare.com/ajax/libs/qrcode/1.5.3/qrcode.min.js',
      'https://cdn.jsdelivr.net/npm/qrcode@1.5.3/build/qrcode.min.js',
      'https://unpkg.com/qrcode@1.5.3/build/qrcode.min.js'
    ];

    for (const src of sources) {
      try {
        await loadExternalScript(src);
        if (window.QRCode && typeof window.QRCode.toDataURL === 'function') return;
      } catch (error) {
        console.warn(`Chargement QR impossible depuis ${src}`, error);
      }
    }

    throw new Error('Le module QR code est indisponible. Rechargez la page ou vérifiez que le réseau autorise cdnjs, jsDelivr ou unpkg.');
  }

  function loadExternalScript(src) {
    return new Promise((resolve, reject) => {
      const existing = [...document.scripts].find(script => script.src === src);
      if (existing) {
        if (window.QRCode && typeof window.QRCode.toDataURL === 'function') return resolve();
        existing.addEventListener('load', resolve, { once: true });
        existing.addEventListener('error', reject, { once: true });
        return;
      }
      const script = document.createElement('script');
      script.src = src;
      script.async = true;
      script.onload = resolve;
      script.onerror = () => reject(new Error(`Impossible de charger ${src}`));
      document.head.appendChild(script);
    });
  }

  function bindSearch() {
    const input = $('#courseSearch');
    let timer;
    input.addEventListener('input', () => {
      clearTimeout(timer);
      timer = setTimeout(() => renderCourseResults(input.value), 160);
    });
    $('#clearSearch').addEventListener('click', () => { input.value = ''; renderCourseResults(''); input.focus(); });
    $('#closeEditor').addEventListener('click', closeEditor);
    $('#courseEditor').addEventListener('click', e => { if (e.target.id === 'courseEditor') closeEditor(); });
    $('#modSearch').addEventListener('input', renderModifications);
  }

  function renderCourseResults(query) {
    const root = $('#courseResults');
    const q = normalizeSearch(query);
    root.innerHTML = '';
    if (q.length < 2) return;
    const terms = q.split(/\s+/).filter(Boolean);
    const results = state.courses.filter(c => {
      const haystack = normalizeSearch([c.title, c.course_id, c.level, c.location_name, c.location_room, c.category, c.subject].join(' '));
      return terms.every(t => haystack.includes(t));
    }).slice(0, 60);
    if (!results.length) {
      root.innerHTML = '<div class="panel empty-state" style="grid-column:1/-1">Aucun cours ne correspond à votre recherche.</div>';
      return;
    }
    root.innerHTML = results.map(c => `
      <article class="course-card" data-course="${escapeAttr(c.id)}">
        <div class="course-code">${escapeHtml(c.course_id || 'Cours')}</div>
        <h3>${escapeHtml(c.title || 'Sans intitulé')}</h3>
        <div class="course-tags"><span class="tag">${escapeHtml(c.level || 'Niveau non indiqué')}</span>${c.school_year ? `<span class="tag">${escapeHtml(c.school_year)}</span>` : ''}</div>
        <div class="course-meta"><span>◷ ${escapeHtml(formatDate(c.start_date) || 'Date non indiquée')}</span><span>⌖ ${escapeHtml([c.location_name, c.location_room].filter(Boolean).join(' · ') || 'Lieu non indiqué')}</span><span>→ Vérifier ce cours</span></div>
      </article>`).join('');
    $$('.course-card', root).forEach(card => card.addEventListener('click', () => openEditor(card.dataset.course)));
  }

  function openEditor(id) {
    const course = state.courses.find(c => c.id === id);
    if (!course) return;
    state.selectedCourse = course;
    $('#editorContent').innerHTML = `
      <div class="editor-head"><span class="editor-code">${escapeHtml(course.course_id)}</span><h2>${escapeHtml(course.title)}</h2><p class="muted">Vérifiez les informations avant leur publication.</p></div>
      <form id="courseForm">
        <div class="readonly-grid">
          <div class="field"><label>Intitulé</label><input class="readonly" readonly value="${escapeAttr(course.title)}"></div>
          <div class="field"><label>Cours ID</label><input class="readonly" readonly value="${escapeAttr(course.course_id)}"></div>
        </div>
        <h3>Informations modifiables</h3>
        <div class="edit-grid">
          <div class="field"><label for="editStart">Date de début</label><input id="editStart" type="date" value="${escapeAttr(course.start_date || '')}"></div>
          <div class="field"><label for="editEnd">Date de fin</label><input id="editEnd" type="date" value="${escapeAttr(course.end_date || '')}"></div>
          <div class="field full"><label for="editDescription">Description du cours</label><textarea id="editDescription">${escapeHtml(course.description || '')}</textarea></div>
          <div class="field full"><label for="editAdditional">Renseignements complémentaires</label><textarea id="editAdditional">${escapeHtml(course.additional_info || '')}</textarea></div>
          <div class="field"><label for="editLocation">Lieu de formation – Nom</label><input id="editLocation" value="${escapeAttr(course.location_name || '')}"></div>
          <div class="field"><label for="editRoom">Lieu de formation – Salle</label><input id="editRoom" value="${escapeAttr(course.location_room || '')}"></div>
        </div>
        <h3>Promotion du cours</h3>
        <div class="promo-box"><img class="qr-preview" src="${course.qr_data || ''}" alt="QR code"><div><div class="promo-link">${escapeHtml(course.link)}</div><button type="button" class="btn ghost small" id="copyLink">Copier le lien</button> <button type="button" class="btn ghost small" id="downloadQr">Télécharger le QR code</button></div></div>
        <h3>Identification du formateur</h3>
        <div class="edit-grid"><div class="field"><label for="trainerName">Nom et prénom *</label><input id="trainerName" required placeholder="Votre nom et prénom"></div><div class="field"><label for="trainerEmail">Adresse e-mail</label><input id="trainerEmail" type="email" placeholder="nom@exemple.lu"></div></div>
        <div id="formError" class="field-error"></div>
        <div class="editor-actions"><button type="button" class="btn ghost" id="confirmNoChange">Confirmer sans modification</button><button type="submit" class="btn primary">Envoyer mes modifications</button></div>
      </form>`;
    $('#courseEditor').classList.remove('hidden');
    document.body.style.overflow = 'hidden';
    $('#courseForm').addEventListener('submit', submitCourseChanges);
    $('#confirmNoChange').addEventListener('click', confirmNoChange);
    $('#copyLink').addEventListener('click', async () => { await navigator.clipboard.writeText(course.link); toast('Lien copié.'); });
    $('#downloadQr').addEventListener('click', () => downloadDataUrl(course.qr_data, `${slugify(course.course_id)}-qr.png`));
  }

  async function submitCourseChanges(e) {
    e.preventDefault();
    const name = $('#trainerName').value.trim();
    const email = $('#trainerEmail').value.trim();
    if (!name) return $('#formError').textContent = 'Veuillez indiquer votre nom et prénom.';
    const c = state.selectedCourse;
    const proposed = {
      start_date: $('#editStart').value,
      end_date: $('#editEnd').value,
      description: $('#editDescription').value.trim(),
      additional_info: $('#editAdditional').value.trim(),
      location_name: $('#editLocation').value.trim(),
      location_room: $('#editRoom').value.trim()
    };
    const labels = {
      start_date: 'Date de début', end_date: 'Date de fin', description: 'Description du cours',
      additional_info: 'Renseignements complémentaires', location_name: 'Lieu de formation – Nom', location_room: 'Lieu de formation – Salle'
    };
    const changes = Object.keys(proposed).filter(k => normalizeComparable(proposed[k]) !== normalizeComparable(c[k])).map(k => ({
      id: crypto.randomUUID(), import_id: state.activeImport.id, course_id_ref: c.id, course_code: c.course_id,
      course_title: c.title, trainer_name: name, trainer_email: email, field_key: k, field_label: labels[k],
      original_value: c[k] || '', proposed_value: proposed[k] || '', created_at: new Date().toISOString()
    }));
    if (!changes.length) return $('#formError').textContent = 'Aucune modification n’a été détectée. Utilisez « Confirmer sans modification ».';
    try {
      await persistModifications(changes);
      closeEditor();
      renderAll();
      toast(`${changes.length} modification(s) enregistrée(s). Merci.`);
    } catch (err) {
      console.error(err);
      $('#formError').textContent = 'Les modifications n’ont pas pu être enregistrées.';
    }
  }

  async function confirmNoChange() {
    const name = $('#trainerName').value.trim();
    if (!name) return $('#formError').textContent = 'Veuillez indiquer votre nom et prénom.';
    const c = state.selectedCourse;
    const confirmation = {
      id: crypto.randomUUID(), import_id: state.activeImport.id, course_id_ref: c.id, course_code: c.course_id,
      course_title: c.title, trainer_name: name, trainer_email: $('#trainerEmail').value.trim(),
      field_key: '_verified', field_label: 'Cours vérifié sans modification', original_value: '', proposed_value: 'Confirmé', created_at: new Date().toISOString()
    };
    await persistModifications([confirmation]);
    closeEditor();
    renderAll();
    toast('Le cours a été confirmé sans modification. Merci.');
  }

  async function persistModifications(changes) {
    if (state.online) {
      const { error } = await state.supabase.from('modifications').insert(changes);
      if (error) throw error;
    }
    state.modifications.unshift(...changes);
    if (!state.online) saveLocal();
  }

  function closeEditor() {
    $('#courseEditor').classList.add('hidden');
    document.body.style.overflow = '';
    state.selectedCourse = null;
  }

  function bindPasswords() {
    $('#downloadLinksBtn').addEventListener('click', () => requestPassword('links'));
    $('#downloadModsBtn').addEventListener('click', () => requestPassword('mods'));
    $('#closePassword').addEventListener('click', closePassword);
    $('#passwordSubmit').addEventListener('click', validatePassword);
    $('#passwordInput').addEventListener('keydown', e => e.key === 'Enter' && validatePassword());
    $('#passwordModal').addEventListener('click', e => e.target.id === 'passwordModal' && closePassword());
  }

  function requestPassword(action) {
    if (action === 'links' && !state.generatedBlob && !state.generatedFilePath && !state.activeImport?.generated_file_path) {
      return toast('Aucun fichier avec liens et QR codes n’est disponible.', true);
    }
    if (action === 'mods' && !state.modifications.some(m => m.field_key !== '_verified')) {
      return toast('Aucune modification à télécharger.', true);
    }
    state.passwordAction = action;
    $('#passwordText').textContent = action === 'links' ? 'Saisissez le mot de passe pour télécharger le fichier avec liens et QR codes.' : 'Saisissez le mot de passe pour télécharger les modifications des formateurs.';
    $('#passwordInput').value = '';
    $('#passwordError').textContent = '';
    $('#passwordModal').classList.remove('hidden');
    setTimeout(() => $('#passwordInput').focus(), 50);
  }

  async function validatePassword() {
    const expected = state.passwordAction === 'links' ? cfg.downloadPasswords?.linksQr : cfg.downloadPasswords?.modifications;
    if ($('#passwordInput').value !== expected) {
      $('#passwordError').textContent = 'Mot de passe incorrect.';
      return;
    }
    const action = state.passwordAction;
    closePassword();
    if (action === 'links') await downloadGeneratedWorkbook();
    if (action === 'mods') generateModificationsWorkbook();
  }


  async function downloadGeneratedWorkbook() {
    if (state.generatedBlob) {
      downloadBlob(state.generatedBlob, state.generatedFilename || state.activeImport?.generated_filename || 'UniPop_Links_QR.xlsx');
      return;
    }
    const path = state.generatedFilePath || state.activeImport?.generated_file_path;
    if (!state.online || !path) {
      toast('Le fichier n’est plus disponible dans cette session.', true);
      return;
    }
    try {
      const { data, error } = await state.supabase.storage
        .from(cfg.storageBucket || 'unipop-files')
        .download(path);
      if (error) throw error;
      downloadBlob(data, state.activeImport?.generated_filename || state.generatedFilename || 'UniPop_Links_QR.xlsx');
    } catch (err) {
      console.error(err);
      toast('Le fichier n’a pas pu être téléchargé depuis Supabase.', true);
    }
  }

  function closePassword() {
    $('#passwordModal').classList.add('hidden');
    state.passwordAction = null;
  }

  async function generateModificationsWorkbook() {
    const mods = state.modifications.filter(m => m.field_key !== '_verified');
    const wb = new ExcelJS.Workbook();
    wb.creator = 'UniPop Inside';
    const ws = wb.addWorksheet('Modifications formateurs', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = [
      { header: 'Cours ID', key: 'course_code', width: 18 },
      { header: 'Intitulé', key: 'course_title', width: 42 },
      { header: 'Formateur', key: 'trainer_name', width: 25 },
      { header: 'E-mail', key: 'trainer_email', width: 30 },
      { header: 'Champ modifié', key: 'field_label', width: 30 },
      { header: 'Valeur originale', key: 'original_value', width: 55 },
      { header: 'Nouvelle valeur', key: 'proposed_value', width: 55 },
      { header: 'Date de soumission', key: 'created_at', width: 22 }
    ];
    mods.forEach(m => ws.addRow({ ...m, created_at: formatDateTime(m.created_at) }));
    const header = ws.getRow(1);
    header.height = 27;
    header.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF062650' } };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      cell.alignment = { vertical: 'middle' };
    });
    ws.eachRow((row, n) => {
      if (n > 1) {
        row.alignment = { vertical: 'top', wrapText: true };
        row.height = Math.max(32, Math.min(100, 18 + Math.ceil(Math.max(String(row.getCell(6).value || '').length, String(row.getCell(7).value || '').length) / 80) * 15));
      }
      row.eachCell(cell => cell.border = { bottom: { style: 'thin', color: { argb: 'FFDCE5F1' } } });
    });
    ws.autoFilter = { from: 'A1', to: 'H1' };
    const buffer = await wb.xlsx.writeBuffer();
    downloadBlob(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Modifications_Formateurs_${state.activeImport?.school_year?.replace('/', '-') || 'UniPop'}.xlsx`);
  }

  function bindMisc() {
    $('#downloadModsBtn').disabled = false;
  }

  function renderAll() {
    renderHome();
    renderExport();
    renderModifications();
    renderHistory();
    renderSettings();
  }

  function renderHome() {
    $('#schoolYear').textContent = state.activeImport?.school_year || '—';
    const active = state.activeImport;
    $('#activeStatus').textContent = active ? 'Actif' : 'Aucun fichier';
    $('#activeStatus').classList.toggle('active', !!active);
    $('#activeFileSummary').innerHTML = active ? `<div class="file-summary"><div class="file-icon">XLS</div><div><h4>${escapeHtml(active.original_filename)}</h4><p>Importé le ${escapeHtml(formatDateTime(active.created_at))} · ${active.sheet_count || 0} onglet(s)</p></div></div>` : 'Importez une première fiche Excel pour démarrer.';
    const realMods = state.modifications.filter(m => m.field_key !== '_verified');
    const verifiedIds = new Set(state.modifications.filter(m => m.field_key === '_verified').map(m => m.course_id_ref));
    const modifiedIds = new Set(realMods.map(m => m.course_id_ref));
    $('#statCourses').textContent = state.courses.length;
    $('#statVerified').textContent = verifiedIds.size;
    $('#statModified').textContent = modifiedIds.size;
    $('#statPending').textContent = realMods.length;
    const recent = realMods.slice(0, 5);
    $('#recentModifications').innerHTML = recent.length ? recent.map(m => `<div class="history-item" style="border:0;border-bottom:1px solid var(--line);border-radius:0;padding:10px 0"><div style="width:4px;height:40px;background:var(--blue);border-radius:4px"></div><div><h3>${escapeHtml(m.course_code)} · ${escapeHtml(m.field_label)}</h3><p>${escapeHtml(m.course_title)} · ${escapeHtml(m.trainer_name)}</p></div><small>${escapeHtml(formatDateTime(m.created_at))}</small></div>`).join('') : 'Aucune modification enregistrée.';
  }

  function renderExport() {
    const active = state.activeImport;
    $('#exportFileName').textContent = active ? (state.generatedFilename || active.generated_filename || 'Fichier généré') : 'Aucun fichier prêt';
    $('#exportFileMeta').textContent = active ? `${active.course_count || state.courses.length} cours · ${active.sheet_count || 0} onglet(s) · Mot de passe requis` : 'Importez d’abord un fichier principal.';
    $('#downloadLinksBtn').disabled = !(state.generatedBlob || state.generatedFilePath || active?.generated_file_path);
  }

  function renderModifications() {
    const query = normalizeSearch($('#modSearch')?.value || '');
    const mods = state.modifications.filter(m => m.field_key !== '_verified').filter(m => !query || normalizeSearch([m.course_code, m.course_title, m.trainer_name, m.field_label].join(' ')).includes(query));
    $('#modCount').textContent = `${mods.length} modification${mods.length === 1 ? '' : 's'}`;
    $('#modTableBody').innerHTML = mods.map(m => `<tr><td><strong>${escapeHtml(m.course_code)}</strong><br>${escapeHtml(m.course_title)}</td><td>${escapeHtml(m.trainer_name)}<br><small>${escapeHtml(m.trainer_email || '')}</small></td><td>${escapeHtml(m.field_label)}</td><td class="value-cell">${escapeHtml(m.original_value || '—')}</td><td class="value-cell new-value">${escapeHtml(m.proposed_value || '—')}</td><td>${escapeHtml(formatDateTime(m.created_at))}</td></tr>`).join('');
    $('#modsEmpty').classList.toggle('hidden', mods.length > 0);
  }

  function renderHistory() {
    const root = $('#historyList');
    if (!state.imports.length) return root.innerHTML = '<div class="panel empty-state">Aucun import enregistré.</div>';
    root.innerHTML = state.imports.map(i => `<article class="history-item"><div class="file-icon">XLS</div><div><h3>${escapeHtml(i.original_filename)}</h3><p>${escapeHtml(formatDateTime(i.created_at))} · ${i.course_count || 0} cours · ${i.sheet_count || 0} onglet(s)</p></div><span class="status-pill ${i.is_active ? 'active' : ''}">${i.is_active ? 'Base active' : 'Archivé'}</span></article>`).join('');
  }

  function renderSettings() {
    const box = $('#supabaseState');
    box.className = `connection-state ${state.online ? 'online' : 'offline'}`;
    box.textContent = state.online ? '● Connecté à Supabase – les données sont partagées entre tous les utilisateurs.' : '● Mode local – les données restent uniquement dans ce navigateur.';
  }

  function renderImportResult(record, sheetsInfo) {
    const el = $('#importResult');
    el.classList.remove('hidden');
    el.innerHTML = `<div class="result-card"><h3>✓ Le nouveau fichier est actif</h3><p>${escapeHtml(record.original_filename)} a été analysé et remplace désormais la base précédente.</p><div class="result-stats"><span><strong>${record.course_count}</strong> cours</span><span><strong>${record.sheet_count}</strong> onglet(s)</span>${sheetsInfo.map(s => `<span>${escapeHtml(s.name)} : <strong>${s.count}</strong></span>`).join('')}</div><div style="margin-top:18px"><button class="btn primary" data-result-download>Télécharger le fichier Links & QR</button></div></div>`;
    $('[data-result-download]', el).addEventListener('click', () => requestPassword('links'));
  }

  function updateProgress(pct, title, log) {
    $('#progressPct').textContent = `${pct}%`;
    $('#progressBar').style.width = `${pct}%`;
    $('#progressTitle').textContent = title;
    $('#progressLog').textContent = log || '';
  }

  function excelDateToIso(value) {
    if (!value) return '';
    if (value instanceof Date && !isNaN(value)) return localIsoDate(value);
    if (typeof value === 'number') {
      const date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
      return date.toISOString().slice(0, 10);
    }
    if (typeof value === 'object' && value.result) return excelDateToIso(value.result);
    const text = fixMojibake(String(value).trim());
    const dmy = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
    if (dmy) return `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
    const parsed = new Date(text);
    return isNaN(parsed) ? '' : localIsoDate(parsed);
  }

  function localIsoDate(d) {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }

  function cellText(cell) {
    if (!cell) return '';
    const v = cell.value;
    if (v == null) return '';
    if (v instanceof Date) return localIsoDate(v);
    if (typeof v === 'object') {
      if (v.text != null) return String(v.text);
      if (v.result != null) return String(v.result);
      if (Array.isArray(v.richText)) return v.richText.map(x => x.text).join('');
      if (v.hyperlink) return v.text || v.hyperlink;
    }
    const s = String(v);
    return s.toLowerCase() === 'null' ? '' : s;
  }

  function fixMojibake(str) {
    if (!/[ÃÂâ]/.test(str || '')) return str || '';
    try {
      const bytes = Uint8Array.from([...str].map(ch => ch.charCodeAt(0) & 0xff));
      const repaired = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
      return repaired.includes('�') ? str : repaired;
    } catch (_) { return str; }
  }

  function normalizeHeader(value) {
    return fixMojibake(String(value || '')).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  }
  function normalizeSearch(value) { return normalizeHeader(value); }
  function normalizeComparable(value) { return String(value || '').trim().replace(/\r\n/g, '\n'); }
  function formatDate(value) { if (!value) return ''; const d = new Date(`${value}T00:00:00`); return isNaN(d) ? value : new Intl.DateTimeFormat('fr-LU').format(d); }
  function formatDateTime(value) { if (!value) return ''; const d = new Date(value); return isNaN(d) ? value : new Intl.DateTimeFormat('fr-LU', { dateStyle: 'short', timeStyle: 'short' }).format(d); }
  function mostCommon(arr) { return arr.sort((a,b) => arr.filter(v => v === a).length - arr.filter(v => v === b).length).pop(); }
  function safeJson(v) { try { return JSON.parse(v); } catch (_) { return null; } }
  function pause(ms) { return new Promise(r => setTimeout(r, ms)); }
  function escapeHtml(v) { return String(v ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
  function escapeAttr(v) { return escapeHtml(v); }
  function toast(message, error = false) { const el = $('#toast'); el.textContent = message; el.className = `toast show${error ? ' error' : ''}`; clearTimeout(toast.t); toast.t = setTimeout(() => el.className = 'toast', 3600); }
  function downloadBlob(blob, filename) { const a = document.createElement('a'); const url = URL.createObjectURL(blob); a.href = url; a.download = filename; a.click(); setTimeout(() => URL.revokeObjectURL(url), 2000); }
  function downloadDataUrl(url, filename) { const a = document.createElement('a'); a.href = url; a.download = filename; a.click(); }
})();
