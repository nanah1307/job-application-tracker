// ---------- AUTOMATED REJECTIONS ----------
const REJECTION_QUERY = 'from:jobstreet "unlikely to progress further" after:2026/01/01';
const REJECTION_HEADERS = ['Date Received', 'Job Title', 'Company', 'Email Link', 'Message ID'];

function scrapeRejectionEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('automated rejections') || ss.insertSheet('automated rejections');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(REJECTION_HEADERS);
    sheet.getRange(1, 1, 1, REJECTION_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Message IDs already in the sheet (column 5) to avoid duplicates
  const lastRow = sheet.getLastRow();
  const existingIds = new Set(
    lastRow > 1 ? sheet.getRange(2, 5, lastRow - 1, 1).getValues().flat() : []
  );

  const cutoff = new Date('2026-01-01T00:00:00+08:00');
  const rows = [];
  const pageSize = 100;
  let start = 0;
  let threads;

  do {
    threads = GmailApp.search(REJECTION_QUERY, start, pageSize);
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        const id = msg.getId();
        if (existingIds.has(id)) return;
        if (msg.getDate() < cutoff) return;

        // Try the plain-text body first, fall back to stripped HTML
        let parsed = parseRejectionText(msg.getSubject() + ' ' + msg.getPlainBody());
        if (!parsed) {
          const htmlText = msg.getBody()
            .replace(/<[^>]+>/g, ' ')
            .replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&');
          parsed = parseRejectionText(htmlText);
        }
        if (!parsed) return;

        rows.push([msg.getDate(), parsed.title, parsed.company, thread.getPermalink(), id]);
        existingIds.add(id);
      });
    });
    start += pageSize;
  } while (threads.length === pageSize);

  if (rows.length) {
    rows.sort((a, b) => a[0] - b[0]);
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, REJECTION_HEADERS.length).setValues(rows);
    sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  }

  notify(`Done! Added ${rows.length} new rejection(s).`);
}

// "Thank you for your interest in the X job at Y. Unfortunately, ..."
function parseRejectionText(text) {
  const flat = text.replace(/\s+/g, ' ');
  const m = flat.match(/interest in the\s+(.+?)\s+job at\s+(.+?)\.\s*Unfortunately/i);
  if (!m) return null;

  return {
    title: m[1].trim(),
    company: m[2].trim().replace(/\.+$/, '') // "Inc.." -> "Inc", same style as the JobStreet sheet
  };
}