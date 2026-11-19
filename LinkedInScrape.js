const LINKEDIN_QUERY = 'from:linkedin "Your application was sent to" after:2026/01/01';
const LINKEDIN_HEADERS = ['Date Applied', 'Job Title', 'Company', 'Location', 'Email Link', 'Message ID'];

function scrapeLinkedInEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('LinkedIn') || ss.insertSheet('LinkedIn');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(LINKEDIN_HEADERS);
    sheet.getRange(1, 1, 1, LINKEDIN_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Message IDs already in the sheet (column 6) to avoid duplicates
  const lastRow = sheet.getLastRow();
  const existingIds = new Set(
    lastRow > 1 ? sheet.getRange(2, 6, lastRow - 1, 1).getValues().flat() : []
  );

  const cutoff = new Date('2026-01-01T00:00:00+08:00');
  const rows = [];
  const pageSize = 100;
  let start = 0;
  let threads;

  do {
    threads = GmailApp.search(LINKEDIN_QUERY, start, pageSize);
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        const id = msg.getId();
        if (existingIds.has(id)) return;
        if (msg.getDate() < cutoff) return;

        // Try the plain-text body first, fall back to stripped HTML
        let parsed = parseLinkedInText(msg.getPlainBody());
        if (!parsed) {
          const htmlText = msg.getBody()
            .replace(/<(br|\/p|\/div|\/tr|\/td|\/h\d)[^>]*>/gi, '\n')
            .replace(/<[^>]+>/g, '')
            .replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&')
            .replace(/&middot;/g, '·');
          parsed = parseLinkedInText(htmlText);
        }
        if (!parsed) return;

        // Use the "Applied on ..." date if present, otherwise the email date
        const dateApplied = parsed.appliedOn || msg.getDate();

        rows.push([dateApplied, parsed.title, parsed.company, parsed.location, thread.getPermalink(), id]);
        existingIds.add(id);
      });
    });
    start += pageSize;
  } while (threads.length === pageSize);

  if (rows.length) {
    rows.sort((a, b) => a[0] - b[0]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, LINKEDIN_HEADERS.length).setValues(rows);
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd');
  }

  notify(`Done! Added ${rows.length} new LinkedIn application(s).`);
}

// Pulls title, company, location and applied date out of the email text
function parseLinkedInText(text) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);

  const headerIdx = lines.findIndex(l => /^your application was sent to\s+/i.test(l));
  if (headerIdx === -1) return null;

  // The "Company · Location" line; the job title is the line right above it
  const infoIdx = lines.findIndex((l, i) => i > headerIdx && l.includes('·'));
  if (infoIdx < 1) return null;

  const title = lines[infoIdx - 1];
  const parts = lines[infoIdx].split('·').map(p => p.trim());
  const company = parts[0];
  const location = parts.slice(1).join(' · ');

  // "Applied on October 14, 2026"
  let appliedOn = null;
  const appliedLine = lines.find(l => /^applied on\s+/i.test(l));
  if (appliedLine) {
    const d = new Date(appliedLine.replace(/^applied on\s+/i, '') + ' 00:00:00 GMT+0800');
    if (!isNaN(d)) appliedOn = d;
  }

  return { title, company, location, appliedOn };
}