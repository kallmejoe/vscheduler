/**
 * parser.js
 *
 * Pure, DOM-free logic for turning a SheetJS workbook (as produced from the
 * ICHEP-style "Schedule" .xlsm/.xlsx files) into structured schedule data.
 *
 * The workbook shape this understands (confirmed on the Fall 2026 file, and
 * expected to hold for future terms as long as the layout is unchanged):
 *
 *   - Each sheet has one or more tables that start with a "Day" header cell,
 *     followed by "Period", "Time", and a "Sections" column.
 *   - Below the header, a row lists the section numbers as plain numbers
 *     (e.g. 5, 6, 7, 8 ...). Those numbers label the columns that follow.
 *   - Data rows follow: one row per (day, period), with the day only written
 *     on the first row of each day-block (rest are blank/merged).
 *   - A cell under a section column holds the session text, e.g.
 *     "CSE141: Introduction to computer Programming - Lec (921A)".
 *   - Merged cells are used both for sessions spanning multiple periods
 *     (vertical merges) and sessions shared by multiple sections
 *     (horizontal merges). We expand every merge into a full grid before
 *     reading, so both cases "just work" without special-casing.
 *   - Some sheets (e.g. "Elective courses") lay out several such tables
 *     side by side instead of a single one with numbered sections. In that
 *     case column numbers don't identify a personal "section" - each cell is
 *     an independent, individually pickable elective session ("flat" mode).
 *   - A level sheet may also have "floating" columns after the last numbered
 *     section: a course (usually a Tutorial) that isn't tied to any section,
 *     listed with several alternative day/period options to choose from.
 */

var Parser = (function () {
  var DAY_ORDER = ['Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

  // Matches things like:
  //   "CSE141 - Lab (350)"
  //   "CSE121: Introduction to Logic Design - Lec (911A)"
  //   "CSE433 - Tut. (9xx)"
  //   "CSE346 - Online"          (no room)
  var SESSION_RE = /^([A-Za-z]{2,6}\d{2,4})\s*:?\s*(.*)-\s*(Lec|Lab|Tut(?:orial)?|Online)\.?\s*(?:\(([^)]*)\))?\s*(?:[-\s]+(.*))?$/i;

  function normalizeType(t) {
    var lower = t.toLowerCase();
    if (lower === 'lec') return 'Lecture';
    if (lower === 'lab') return 'Lab';
    if (lower.indexOf('tut') === 0) return 'Tutorial';
    if (lower === 'online') return 'Online';
    return t;
  }

  function parseSessionText(raw) {
    var text = String(raw).trim();
    var m = text.match(SESSION_RE);
    if (m) {
      return {
        code: m[1].toUpperCase(),
        title: m[2].trim().replace(/[:\-]+$/, '').trim(),
        type: normalizeType(m[3]),
        room: m[4] ? m[4].trim() : null,
        note: m[5] ? m[5].trim() : null,
        raw: text
      };
    }
    var codeMatch = text.match(/^([A-Za-z]{2,6}\d{2,4})/);
    return {
      code: codeMatch ? codeMatch[1].toUpperCase() : text,
      title: '',
      type: 'Other',
      room: null,
      note: null,
      raw: text
    };
  }

  // Turn a worksheet into a fully rectangular, merge-expanded 2D array
  // (0-based [row][col]) so every merged cell reads as if it were repeated
  // across its whole range.
  function sheetToMatrix(ws) {
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    var width = 0;
    for (var i = 0; i < aoa.length; i++) width = Math.max(width, aoa[i].length);

    var matrix = aoa.map(function (row) {
      var copy = row.slice(0, width);
      while (copy.length < width) copy.push(null);
      return copy;
    });

    var merges = ws['!merges'] || [];
    for (var mi = 0; mi < merges.length; mi++) {
      var m = merges[mi];
      var topRow = matrix[m.s.r];
      var value = topRow ? topRow[m.s.c] : null;
      for (var r = m.s.r; r <= m.e.r; r++) {
        if (!matrix[r]) matrix[r] = new Array(width).fill(null);
        for (var c = m.s.c; c <= m.e.c; c++) {
          matrix[r][c] = value;
        }
      }
    }
    return matrix;
  }

  function findDayHeaders(matrix) {
    // A table's "Day" header cell is often itself merged vertically (e.g. to
    // visually span the Day/Period/Time/Sections header block), which after
    // merge-expansion repeats the text down several rows in the same
    // column. Only the top-most occurrence per column is a real table
    // anchor, so keep the minimum row for each column.
    var byCol = {};
    for (var r = 0; r < matrix.length; r++) {
      for (var c = 0; c < matrix[r].length; c++) {
        var v = matrix[r][c];
        if (typeof v === 'string' && v.trim().toLowerCase() === 'day') {
          if (byCol[c] === undefined || r < byCol[c]) byCol[c] = r;
        }
      }
    }
    return Object.keys(byCol).map(function (c) {
      return { row: byCol[c], col: Number(c) };
    });
  }

  // Work out the geometry of one Day/Period/Time/Sections table starting at
  // (hr, hc): where the section-number row is, which columns are numbered
  // sections, and which rows hold actual data.
  function analyzeTable(matrix, hr, hc, boundaryCol) {
    var nCols = boundaryCol != null ? boundaryCol : matrix[hr].length;

    var numbersRow = -1;
    for (var r = hr + 1; r < matrix.length; r++) {
      var periodVal = matrix[r][hc + 1];
      if (typeof periodVal === 'number') break;
      for (var c = hc + 3; c < nCols; c++) {
        if (typeof matrix[r][c] === 'number') { numbersRow = r; break; }
      }
      if (numbersRow >= 0) break;
    }

    var sectionCols = [];
    if (numbersRow >= 0) {
      for (var c2 = hc + 3; c2 < nCols; c2++) {
        var v = matrix[numbersRow][c2];
        if (typeof v === 'number') sectionCols.push({ col: c2, number: v });
        else break;
      }
    }

    var dataStart = -1;
    var searchFrom = (numbersRow >= 0 ? numbersRow : hr) + 1;
    for (var r2 = searchFrom; r2 < matrix.length; r2++) {
      if (typeof matrix[r2][hc + 1] === 'number') { dataStart = r2; break; }
    }

    var dataEnd = dataStart - 1;
    if (dataStart >= 0) {
      for (var r3 = dataStart; r3 < matrix.length; r3++) {
        if (typeof matrix[r3][hc + 1] !== 'number') break;
        dataEnd = r3;
      }
    }

    return { hr: hr, hc: hc, nCols: nCols, numbersRow: numbersRow, sectionCols: sectionCols, dataStart: dataStart, dataEnd: dataEnd };
  }

  function fillDayColumn(matrix, table) {
    var days = [];
    var last = null;
    for (var r = table.dataStart; r <= table.dataEnd; r++) {
      var v = matrix[r][table.hc];
      if (typeof v === 'string' && v.trim()) last = v.trim();
      days[r] = last;
    }
    return days;
  }

  function findLabel(matrix, table, col) {
    for (var r = table.hr; r < table.dataStart; r++) {
      var v = matrix[r][col];
      if (typeof v === 'string' && v.trim()) return v.trim();
    }
    return null;
  }

  function dedupeSessions(sessions) {
    var seen = {};
    var out = [];
    for (var i = 0; i < sessions.length; i++) {
      var s = sessions[i];
      var key = s.day + '|' + s.period + '|' + s.raw;
      if (seen[key]) continue;
      seen[key] = true;
      out.push(s);
    }
    return out;
  }

  // A lecture/lab/tutorial that occupies more than one period is stored as a
  // merged cell spanning several rows. After merge-expansion that shows up
  // as several consecutive per-period rows with identical text - group them
  // back into a single logical session with a `periods` array instead of
  // treating it as several unrelated bookings.
  function mergeConsecutivePeriods(sessions) {
    var sorted = sessions.slice().sort(function (a, b) {
      return dayIndex(a.day) - dayIndex(b.day) || a.period - b.period;
    });
    var out = [];
    sorted.forEach(function (s) {
      var last = out[out.length - 1];
      var contiguous = last &&
        last.day === s.day &&
        last.raw === s.raw &&
        (last.table || null) === (s.table || null) &&
        s.period === last.periods[last.periods.length - 1] + 1;
      if (contiguous) {
        last.periods.push(s.period);
        last.times.push(s.time);
      } else {
        var clone = Object.assign({}, s);
        clone.periods = [s.period];
        clone.times = [s.time];
        out.push(clone);
      }
    });
    out.forEach(function (o) {
      o.period = o.periods[0];
      o.time = o.times[0];
    });
    return out;
  }

  // Parse one worksheet into either:
  //   { mode: 'sections', sectionNumbers, sections: {num: [session...]}, floating: [{col,label,items}] }
  //   { mode: 'flat', sessions: [session...] }
  //   { mode: 'empty' }
  function parseSheet(ws, sheetName) {
    var matrix = sheetToMatrix(ws);
    var headers = findDayHeaders(matrix).sort(function (a, b) { return a.col - b.col; });

    if (headers.length === 0) {
      return { name: sheetName, mode: 'empty' };
    }

    if (headers.length > 1) {
      var sessions = [];
      headers.forEach(function (h, i) {
        var boundaryCol = i + 1 < headers.length ? headers[i + 1].col : undefined;
        var table = analyzeTable(matrix, h.row, h.col, boundaryCol);
        if (table.dataStart < 0) return;
        var days = fillDayColumn(matrix, table);
        // Each side-by-side table usually has its own label (e.g. a bylaw
        // name) sitting between the header row and the section-number row.
        var tableLabel = findLabel(matrix, table, h.col + 3);
        for (var r = table.dataStart; r <= table.dataEnd; r++) {
          for (var c = h.col + 3; c < table.nCols; c++) {
            var raw = matrix[r][c];
            if (raw === null || raw === '' || typeof raw === 'number') continue;
            var parsed = parseSessionText(raw);
            sessions.push(Object.assign({
              day: days[r],
              period: matrix[r][h.col + 1],
              time: matrix[r][h.col + 2],
              table: tableLabel
            }, parsed));
          }
        }
      });
      return { name: sheetName, mode: 'flat', sessions: mergeConsecutivePeriods(dedupeSessions(sessions)) };
    }

    var h0 = headers[0];
    var table0 = analyzeTable(matrix, h0.row, h0.col);
    if (table0.dataStart < 0 || table0.sectionCols.length === 0) {
      return { name: sheetName, mode: 'empty' };
    }
    var days0 = fillDayColumn(matrix, table0);

    var sections = {};
    table0.sectionCols.forEach(function (sc) { sections[sc.number] = []; });
    for (var r0 = table0.dataStart; r0 <= table0.dataEnd; r0++) {
      for (var si = 0; si < table0.sectionCols.length; si++) {
        var sc0 = table0.sectionCols[si];
        var raw0 = matrix[r0][sc0.col];
        if (raw0 === null || raw0 === '') continue;
        var parsed0 = parseSessionText(raw0);
        sections[sc0.number].push(Object.assign({
          day: days0[r0],
          period: matrix[r0][h0.col + 1],
          time: matrix[r0][h0.col + 2]
        }, parsed0));
      }
    }
    Object.keys(sections).forEach(function (num) {
      sections[num] = mergeConsecutivePeriods(dedupeSessions(sections[num]));
    });

    var lastSectionCol = Math.max.apply(null, table0.sectionCols.map(function (s) { return s.col; }));
    var floating = [];
    for (var c1 = lastSectionCol + 1; c1 < matrix[h0.row].length; c1++) {
      var items = [];
      for (var r1 = table0.dataStart; r1 <= table0.dataEnd; r1++) {
        var raw1 = matrix[r1][c1];
        if (raw1 === null || raw1 === '') continue;
        var parsed1 = parseSessionText(raw1);
        items.push(Object.assign({
          day: days0[r1],
          period: matrix[r1][h0.col + 1],
          time: matrix[r1][h0.col + 2]
        }, parsed1));
      }
      if (items.length > 0) {
        var deduped = mergeConsecutivePeriods(dedupeSessions(items));
        var label = findLabel(matrix, table0, c1) || (deduped[0].code + ' - ' + deduped[0].type);
        floating.push({ col: c1, label: label, items: deduped });
      }
    }

    return {
      name: sheetName,
      mode: 'sections',
      sectionNumbers: table0.sectionCols.map(function (s) { return s.number; }).sort(function (a, b) { return a - b; }),
      sections: sections,
      floating: floating
    };
  }

  function parseWorkbook(workbook) {
    var sheets = {};
    workbook.SheetNames.forEach(function (name) {
      sheets[name] = parseSheet(workbook.Sheets[name], name);
    });
    return sheets;
  }

  function dayIndex(day) {
    var i = DAY_ORDER.indexOf(day);
    return i === -1 ? DAY_ORDER.length : i;
  }

  // Given a flat list of chosen sessions ({day, periods: [...], ...}), find
  // every (day, period) slot occupied by more than one session. Slots that
  // share the exact same set of conflicting sessions across consecutive
  // periods are merged into a single conflict with a period range, so a
  // clash between two 2-period lectures is reported once, not twice.
  function findConflicts(mySchedule) {
    var bySlot = {};
    mySchedule.forEach(function (item) {
      item.periods.forEach(function (p) {
        var key = item.day + '|' + p;
        (bySlot[key] = bySlot[key] || []).push(item);
      });
    });

    var slots = Object.keys(bySlot)
      .map(function (key) {
        var parts = key.split('|');
        return { day: parts[0], period: Number(parts[1]), items: bySlot[key] };
      })ime: items[0].time, items: items });
      }
    });
    conflicts.sort(function (a, b) { return dayIndex(a.day) - dayIndex(b.day) || a.period - b.period; });
    return conflicts;
  }

  return {
    DAY_ORDER: DAY_ORDER,
    parseSessionText: parseSessionText,
    sheetToMatrix: sheetToMatrix,
    findDayHeaders: findDayHeaders,
    analyzeTable: analyzeTable,
    parseSheet: parseSheet,
    parseWorkbook: parseWorkbook,
    findConflicts: findConflicts,
    dayIndex: dayIndex
  };
})();
