const SEARCH_QUERY = 'from:jobstreet "successfully submitted" after:2026/01/01';
const HEADERS = ['Date Applied', 'Job Title', 'Company', 'Email Link', 'Message ID'];

function scrapeJobStreetEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Job Street') || ss.insertSheet('Job Street');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  const existingIds = new Set(
    lastRow > 1 ? sheet.getRange(2, 5, lastRow - 1, 1).getValues().flat() : []
  );

  const rows = [];
  const pageSize = 100;
  let start = 0;
  let threads;

  do {
    threads = GmailApp.search(SEARCH_QUERY, start, pageSize);
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        const id = msg.getId();
        if (existingIds.has(id)) return;

        const subject = msg.getSubject();
        const plain = subject + ' ' + msg.getPlainBody();
        const html = subject + ' ' + msg.getBody()
          .replace(/<(style|script)[\s\S]*?<\/\1>/gi, ' ')
          .replace(/<[^>]+>/g, ' ')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&');

        // Every message is an application: parse it, and keep the row even if parsing fails
        const parsed = parseJobStreetText(plain) || parseJobStreetText(html) || { title: '', company: '' };

        rows.push([msg.getDate(), parsed.title, parsed.company, thread.getPermalink(), id]);
        existingIds.add(id);
      });
    });
    start += pageSize;
  } while (threads.length === pageSize);

  if (rows.length) {
    rows.sort((a, b) => a[0] - b[0]);
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
    sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');

    // Highlight rows where the title or company couldn't be read
    rows.forEach((r, i) => {
      if (!r[1] || !r[2]) {
        sheet.getRange(startRow + i, 1, 1, HEADERS.length).setBackground('#fff2cc');
      }
    });
  }

  const unreadable = rows.filter(r => !r[1] || !r[2]).length;
  notify(`Done! Added ${rows.length} new application(s). ${unreadable} highlighted row(s) need a quick check.`);
}

function parseJobStreetText(text) {
  const flat = text.replace(/\s+/g, ' ');
  const m = flat.match(
    /Your application for\s+(.+?)\s+was successfully submitted to\s+(.{1,120}?)\s*(?:Each employer|Keep track|$)/i
  );
  if (!m) return null;
  return {
    title: m[1].trim(),
    company: m[2].trim().replace(/\.+$/, '')
  };
}