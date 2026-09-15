/* Run with: node parser.test.cjs */
const assert = require('node:assert/strict');

globalThis.XLSX = {
  utils: {
    sheet_to_json: (worksheet) => worksheet.aoa
  }
};

function worksheet(merges) {
  return {
    aoa: [
      ['Day', 'Period', 'Time', 'Sections'],
      [null, null, null, 1],
      ['Monday', 1, '08:00 - 09:30', 'CSE141 - Lec (911A)'],
      [null, 2, '09:45 - 11:15', 'CSE141 - Lec (911A)']
    ],
    '!merges': merges || []
  };
}

(async function () {
  const { Parser } = await import('./parser.js');
  const verticalMerge = [{ s: { r: 2, c: 3 }, e: { r: 3, c: 3 } }];
  const mergedSessions = Parser.parseSheet(worksheet(verticalMerge), 'Test').sections[1];
  assert.equal(mergedSessions.length, 1);
  assert.deepEqual(mergedSessions[0].periods, [1, 2]);

  const separateSessions = Parser.parseSheet(worksheet(), 'Test').sections[1];
  assert.equal(separateSessions.length, 2);
  assert.deepEqual(separateSessions[0].periods, [1]);
  assert.deepEqual(separateSessions[1].periods, [2]);

  console.log('parser merge handling: passed');
})();
