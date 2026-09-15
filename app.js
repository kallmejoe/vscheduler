import { Parser } from './parser.js';

/**
 * app.js
 *
 * DOM wiring for the Schedule Conflict Checker. All Excel-reading logic
 * lives in parser.js; this file owns application state and rendering.
 *
 * Sessions that span more than one period (a merged cell in the source
 * sheet) are represented as a single item with a `periods` array (e.g.
 * [3, 4]) rather than two separate items - see parser.js's
 * mergeConsecutivePeriods. Everything here treats `periods` as the unit of
 * selection/conflict-checking, not a single period number.
 */

(function () {
  'use strict';

  // By default, load this file from the same folder as index.html. Update
  // this if next term's workbook is named differently, or just use the
  // "Load a different schedule file" picker on the page.
  var DEFAULT_FILE = 'CESS CAIE Schedule Fall 2026.xlsm';

  // ---- state -------------------------------------------------------------
  var sheets = null; // { sheetName: parsedSheet }
  var currentLevel = null;
  var mySchedule = []; // flat list of chosen session items (see makeItem)
  var allDays = [];
  var allPeriods = [];
  var gridPeriodTimes = {};

  // ---- small helpers -------------------------------------------------------

  function slug(s) {
    return String(s).replace(/[^a-zA-Z0-9_-]/g, '-');
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function periodRangeLabel(periods) {
    return periods.length > 1 ? ('P' + periods[0] + '\u2013P' + periods[periods.length - 1]) : ('P' + periods[0]);
  }

  function timeRangeLabel(periods) {
    var start = gridPeriodTimes[periods[0]];
    var end = gridPeriodTimes[periods[periods.length - 1]];
    if (!start) return '';
    if (periods.length === 1 || !end) return start;
    var startPart = String(start).split('-')[0].trim();
    var endPart = String(end).split('-').slice(-1)[0].trim();
    return startPart + ' - ' + endPart;
  }

  function makeItem(session, groupKey, sourceLabel) {
    return {
      groupKey: groupKey,
      itemKey: groupKey + '::' + session.day + '::' + session.periods.join(',') + '::' + session.raw,
      day: session.day,
      periods: session.periods.slice(),
      code: session.code,
      title: session.title,
      type: session.type,
      room: session.room,
      note: session.note,
      raw: session.raw,
      sourceLabel: sourceLabel
    };
  }

  function roomText(item) {
    var parts = [];
    if (item.room) parts.push(item.room);
    if (item.note) parts.push(item.note);
    return parts.length ? ' (' + parts.join(' \u00b7 ') + ')' : '';
  }

  function wouldConflict(day, periods, excludeGroupKey) {
    return mySchedule.some(function (x) {
      if (x.groupKey === excludeGroupKey || x.day !== day) return false;
      return x.periods.some(function (p) { return periods.indexOf(p) !== -1; });
    });
  }

  function hasFlatSheet() {
    return Object.keys(sheets).some(function (n) { return sheets[n].mode === 'flat'; });
  }

  function uniqueGroupKeys(items) {
    var seen = {};
    var out = [];
    items.forEach(function (it) {
      if (!seen[it.groupKey]) { seen[it.groupKey] = true; out.push(it.groupKey); }
    });
    return out;
  }

  function groupShortLabel(gk) {
    if (gk.indexOf('section:') === 0) return 'section';
    if (gk.indexOf('floating:') === 0) return 'time';
    if (gk.indexOf('elective:') === 0) return 'elective';
    return 'choice';
  }

  function groupToElement(gk) {
    if (gk.indexOf('section:') === 0) return document.getElementById('levelCard');
    return document.getElementById('group-' + slug(gk));
  }

  // ---- state mutation -------------------------------------------------------

  function setGroupSelection(groupKey, items) {
    mySchedule = mySchedule.filter(function (x) { return x.groupKey !== groupKey; });
    mySchedule = mySchedule.concat(items);
    render();
  }

  // ---- grid metadata (computed once per loaded workbook) --------------------

  function computeGridMeta() {
    var daysSet = {};
    var periodTimes = {};
    Object.keys(sheets).forEach(function (name) {
      var sheet = sheets[name];
      var all = [];
      if (sheet.mode === 'sections') {
        Object.keys(sheet.sections).forEach(function (k) { all = all.concat(sheet.sections[k]); });
        sheet.floating.forEach(function (g) { all = all.concat(g.items); });
      } else if (sheet.mode === 'flat') {
        all = sheet.sessions;
      }
      all.forEach(function (s) {
        if (s.day) daysSet[s.day] = true;
        s.periods.forEach(function (p, i) {
          if (periodTimes[p] == null && s.times[i] != null) periodTimes[p] = s.times[i];
        });
      });
    });
    allDays = Parser.DAY_ORDER.filter(function (d) { return daysSet[d]; });
    allPeriods = Object.keys(periodTimes).map(Number).sort(function (a, b) { return a - b; });
    gridPeriodTimes = periodTimes;
  }

  // Group a flat sheet's sessions into: [{ label, courses: [{ code, title, types: { typeName: [session...] } }] }]
  function groupElectiveSessions(sessions) {
    var byLabel = {};
    var labelOrder = [];
    sessions.forEach(function (s) {
      var label = s.table || 'Electives';
      if (!byLabel[label]) { byLabel[label] = {}; labelOrder.push(label); }
      var course = byLabel[label][s.code] || (byLabel[label][s.code] = { code: s.code, title: s.title, types: {} });
      if (!course.title && s.title) course.title = s.title;
      (course.types[s.type] = course.types[s.type] || []).push(s);
    });
    return labelOrder.map(function (label) {
      var courses = Object.keys(byLabel[label]).sort().map(function (k) { return byLabel[label][k]; });
      return { label: label, courses: courses };
    });
  }

  // ---- rendering -------------------------------------------------------

  function populateLevelSelect(names) {
    var sel = document.getElementById('levelSelect');
    sel.innerHTML = '';
    names.forEach(function (n) {
      var opt = document.createElement('option');
      opt.value = n;
      opt.textContent = n;
      sel.appendChild(opt);
    });
  }

  function onLevelChange() {
    currentLevel = document.getElementById('levelSelect').value;
    var sheet = sheets[currentLevel];

    var sectionSel = document.getElementById('sectionSelect');
    sectionSel.innerHTML = '';
    var placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = 'Select a section…';
    sectionSel.appendChild(placeholder);
    sheet.sectionNumbers.forEach(function (n) {
      var opt = document.createElement('option');
      opt.value = n;
      opt.textContent = 'Section ' + n;
      sectionSel.appendChild(opt);
    });
    sectionSel.value = '';

    // Only one level can be "active" at a time.
    mySchedule = mySchedule.filter(function (x) {
      return x.groupKey.indexOf('section:') !== 0 && x.groupKey.indexOf('floating:') !== 0;
    });
    render();
  }

  function onSectionChange() {
    var val = document.getElementById('sectionSelect').value;
    var groupKey = 'section:' + currentLevel;
    if (!val) {
      setGroupSelection(groupKey, []);
      return;
    }
    var sheet = sheets[currentLevel];
    var sessions = sheet.sections[val] || [];
    var label = currentLevel + ' \u00b7 Section ' + val;
    var items = sessions.map(function (s) { return makeItem(s, groupKey, label); });
    setGroupSelection(groupKey, items);
  }

  function makeOptionButton(item, groupKey, selected) {
    var conflicting = wouldConflict(item.day, item.periods, groupKey);
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'option-btn' + (selected ? ' selected' : '') + (conflicting ? ' warn' : '');
    btn.textContent = item.day + ' ' + periodRangeLabel(item.periods) + ' (' + timeRangeLabel(item.periods) + ')';
    btn.addEventListener('click', function () {
      setGroupSelection(groupKey, selected ? [] : [item]);
    });
    return btn;
  }

  function renderFloating(levelSheet) {
    var container = document.getElementById('floatingContainer');
    container.innerHTML = '';
    if (!levelSheet || !levelSheet.floating || levelSheet.floating.length === 0) return;

    levelSheet.floating.forEach(function (group) {
      var groupKey = 'floating:' + currentLevel + ':' + group.col;
      var card = document.createElement('div');
      card.className = 'group-card';
      card.id = 'group-' + slug(groupKey);

      var h3 = document.createElement('h3');
      h3.textContent = group.label + ' \u2014 choose a time';
      card.appendChild(h3);

      var list = document.createElement('div');
      list.className = 'option-list';
      group.items.forEach(function (session) {
        var item = makeItem(session, groupKey, group.label);
        var selected = mySchedule.some(function (x) { return x.itemKey === item.itemKey; });
        list.appendChild(makeOptionButton(item, groupKey, selected));
      });
      card.appendChild(list);
      container.appendChild(card);
    });
  }

  function renderElectives() {
    var container = document.getElementById('electiveContainer');
    container.innerHTML = '';
    document.getElementById('electiveCard').hidden = !hasFlatSheet();
    if (!hasFlatSheet()) return;

    var filterVal = (document.getElementById('electiveFilter').value || '').trim().toLowerCase();

    Object.keys(sheets).forEach(function (name) {
      var sheet = sheets[name];
      if (sheet.mode !== 'flat') return;

      groupElectiveSessions(sheet.sessions).forEach(function (group) {
        var matching = group.courses.filter(function (course) {
          if (!filterVal) return true;
          var hay = (course.code + ' ' + (course.title || '')).toLowerCase();
          return hay.indexOf(filterVal) !== -1;
        });
        if (matching.length === 0) return;

        var details = document.createElement('details');
        details.className = 'elective-group';
        details.open = true;

        var summary = document.createElement('summary');
        summary.textContent = group.label + ' (' + matching.length + ' course' + (matching.length > 1 ? 's' : '') + ')';
        details.appendChild(summary);

        matching.forEach(function (course) {
          var card = document.createElement('div');
          card.className = 'course-card';

          var h4 = document.createElement('h4');
          h4.textContent = course.code + (course.title ? ': ' + course.title : '');
          card.appendChild(h4);

          Object.keys(course.types).sort().forEach(function (typeName) {
            var groupKey = 'elective:' + name + ':' + course.code + ':' + typeName;

            var row = document.createElement('div');
            row.className = 'type-row';
            row.id = 'group-' + slug(groupKey);

            var typeLabel = document.createElement('span');
            typeLabel.className = 'type-label';
            typeLabel.textContent = typeName;
            row.appendChild(typeLabel);

            var list = document.createElement('div');
            list.className = 'option-list';
            course.types[typeName].forEach(function (session) {
              var item = makeItem(session, groupKey, course.code + ' ' + typeName);
              var selected = mySchedule.some(function (x) { return x.itemKey === item.itemKey; });
              list.appendChild(makeOptionButton(item, groupKey, selected));
            });
            row.appendChild(list);
            card.appendChild(row);
          });

          details.appendChild(card);
        });

        container.appendChild(details);
      });
    });
  }

  function renderScheduleList() {
    var ul = document.getElementById('scheduleList');
    ul.innerHTML = '';
    if (mySchedule.length === 0) {
      var empty = document.createElement('li');
      empty.textContent = 'Nothing selected yet.';
      empty.style.border = 'none';
      empty.classList.add('muted');
      ul.appendChild(empty);
      return;
    }

    var conflictingKeys = {};
    Parser.findConflicts(mySchedule).forEach(function (c) {
      c.items.forEach(function (it) { conflictingKeys[it.itemKey] = true; });
    });

    var sorted = mySchedule.slice().sort(function (a, b) {
      return Parser.dayIndex(a.day) - Parser.dayIndex(b.day) || a.periods[0] - b.periods[0];
    });

    sorted.forEach(function (item) {
      var li = document.createElement('li');
      if (conflictingKeys[item.itemKey]) li.className = 'conflict';

      var left = document.createElement('span');
      left.innerHTML = '<strong>' + escapeHtml(item.code) + '</strong> ' + escapeHtml(item.type) + escapeHtml(roomText(item)) +
        ' <span class="meta">\u00b7 ' + escapeHtml(item.day) + ' ' + periodRangeLabel(item.periods) + ' (' + escapeHtml(timeRangeLabel(item.periods)) +
        ') \u00b7 ' + escapeHtml(item.sourceLabel) + '</span>';
      li.appendChild(left);

      var removeBtn = document.createElement('button');
      removeBtn.className = 'remove';
      removeBtn.title = 'Remove';
      removeBtn.textContent = '\u2715';
      removeBtn.addEventListener('click', function () {
        mySchedule = mySchedule.filter(function (x) { return x.itemKey !== item.itemKey; });
        render();
      });
      li.appendChild(removeBtn);
      ul.appendChild(li);
    });
  }

  function renderGrid() {
    var wrap = document.getElementById('scheduleGridWrap');
    wrap.innerHTML = '';
    if (!allDays.length || !allPeriods.length) return;

    var table = document.createElement('table');
    table.className = 'grid';

    var thead = document.createElement('thead');
    var headRow = document.createElement('tr');
    headRow.appendChild(document.createElement('th'));
    allDays.forEach(function (d) {
      var th = document.createElement('th');
      th.textContent = d;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);

    // A multi-period item occupies every period it spans - place it in each
    // of those grid rows so it's visible wherever it blocks time, but it's
    // still the same single session (see the "Selected sessions" list).
    var byDayPeriod = {};
    mySchedule.forEach(function (item) {
      item.periods.forEach(function (p) {
        var key = item.day + '|' + p;
        (byDayPeriod[key] = byDayPeriod[key] || []).push(item);
      });
    });

    var tbody = document.createElement('tbody');
    allPeriods.forEach(function (p) {
      var tr = document.createElement('tr');
      var th = document.createElement('th');
      th.textContent = 'P' + p;
      th.title = gridPeriodTimes[p];
      tr.appendChild(th);

      allDays.forEach(function (d) {
        var td = document.createElement('td');
        var items = byDayPeriod[d + '|' + p] || [];
        if (items.length > 1) td.className = 'cell-conflict';
        items.forEach(function (item) {
          var span = document.createElement('span');
          span.className = 'cell-session';
          span.textContent = item.code + ' ' + item.type + roomText(item);
          td.appendChild(span);
        });
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(table);
  }

  function renderConflicts() {
    var conflicts = Parser.findConflicts(mySchedule);
    var summary = document.getElementById('conflictsSummary');
    var list = document.getElementById('conflictsList');
    list.innerHTML = '';

    if (conflicts.length === 0) {
      summary.textContent = mySchedule.length ? 'No conflicts found.' : 'Select your courses below to build your schedule.';
      summary.className = mySchedule.length ? 'ok' : 'muted';
      return;
    }

    summary.textContent = conflicts.length + ' conflict' + (conflicts.length > 1 ? 's' : '') + ' found:';
    summary.className = 'bad';

    conflicts.forEach(function (c) {
      var li = document.createElement('li');
      var names = c.items.map(function (it) { return it.code + ' ' + it.type + ' (' + it.sourceLabel + ')'; });
      li.innerHTML = '<strong>' + escapeHtml(c.day) + ' ' + periodRangeLabel(c.periods) + '</strong> (' + escapeHtml(timeRangeLabel(c.periods)) +
        ') \u2014 <span class="vs">' + names.map(escapeHtml).join(' \u2715 ') + '</span>';

      uniqueGroupKeys(c.items).forEach(function (gk) {
        var el = groupToElement(gk);
        if (!el) return;
        var btn = document.createElement('button');
        btn.className = 'link-btn';
        btn.style.marginLeft = '8px';
        btn.textContent = 'change ' + groupShortLabel(gk);
        btn.addEventListener('click', function () {
          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          el.classList.add('highlight');
          setTimeout(function () { el.classList.remove('highlight'); }, 1500);
        });
        li.appendChild(btn);
      });
      list.appendChild(li);
    });
  }

  function render() {
    renderFloating(sheets[currentLevel]);
    renderElectives();
    renderScheduleList();
    renderGrid();
    renderConflicts();
  }

  // ---- loading -------------------------------------------------------

  function loadWorkbook(workbook, label) {
    sheets = Parser.parseWorkbook(workbook);
    mySchedule = [];
    computeGridMeta();

    var levelNames = Object.keys(sheets).filter(function (n) { return sheets[n].mode === 'sections'; });
    var status = document.getElementById('fileStatus');
    if (levelNames.length === 0) {
      status.textContent = 'Could not find any level/section tables in "' + label + '". Is this the right workbook shape?';
      document.getElementById('appBody').hidden = true;
      return;
    }

    populateLevelSelect(levelNames);
    status.textContent = 'Loaded "' + label + '" \u2014 ' + Object.keys(sheets).length + ' sheet(s) found.';
    document.getElementById('appBody').hidden = false;
    onLevelChange();
  }

  function loadArrayBuffer(buf, label) {
    try {
      var workbook = XLSX.read(new Uint8Array(buf), { type: 'array' });
      loadWorkbook(workbook, label);
    } catch (err) {
      console.error(err);
      document.getElementById('fileStatus').textContent = 'Failed to read "' + label + '": ' + err.message;
    }
  }

  function loadDefaultFile() {
    var status = document.getElementById('fileStatus');
    document.getElementById('defaultFileName').textContent = DEFAULT_FILE;
    status.textContent = 'Loading "' + DEFAULT_FILE + '"\u2026';
    fetch(encodeURIComponent(DEFAULT_FILE))
      .then(function (res) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return res.arrayBuffer();
      })
      .then(function (buf) { loadArrayBuffer(buf, DEFAULT_FILE); })
      .catch(function (err) {
        console.warn('Could not auto-load default file:', err);
        status.textContent = 'Could not auto-load "' + DEFAULT_FILE + '" (this only works when the page is served ' +
          'over a local server, not opened directly as a file). Use "Load a different schedule file" below to pick it manually.';
        document.getElementById('fileLoader').open = true;
      });
  }

  // ---- wiring -------------------------------------------------------

  document.getElementById('fileInput').addEventListener('change', function (e) {
    var file = e.target.files[0];
    if (!file) return;
    var status = document.getElementById('fileStatus');
    status.textContent = 'Reading "' + file.name + '"\u2026';
    var reader = new FileReader();
    reader.onload = function (ev) { loadArrayBuffer(ev.target.result, file.name); };
    reader.readAsArrayBuffer(file);
  });

  document.getElementById('levelSelect').addEventListener('change', onLevelChange);
  document.getElementById('sectionSelect').addEventListener('change', onSectionChange);
  document.getElementById('electiveFilter').addEventListener('input', renderElectives);
  document.getElementById('clearBtn').addEventListener('click', function () {
    mySchedule = [];
    var sectionSel = document.getElementById('sectionSelect');
    if (sectionSel) sectionSel.value = '';
    render();
  });

  loadDefaultFile();
})();
