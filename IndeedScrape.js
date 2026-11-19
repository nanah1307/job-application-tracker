const INDEED_QUERY = 'from:indeed "Application submitted" after:2026/01/01';
const INDEED_HEADERS = ['Date Applied', 'Job Title', 'Company', 'Location', 'Email Link', 'Message ID'];

function scrapeIndeedEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Indeed') || ss.insertSheet('Indeed');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(INDEED_HEADERS);
    sheet.getRange(1, 1, 1, INDEED_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Message IDs already in the sheet (column 6) to avoid duplicates
  const lastRow = sheet.getLastRow();
  const existingIds = new Set(
    lastRow > 1 ? sheet.getRange(2, 6, lastRow - 1, 1).getValues().flat() : []
  );

  const rows = [];
  const pageSize = 100;
  let start = 0;
  let threads;

  do {
    threads = GmailApp.search(INDEED_QUERY, start, pageSize);
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        const id = msg.getId();
        if (existingIds.has(id)) return;

        // Try the plain-text body first, fall back to stripped HTML
        let parsed = parseIndeedText(msg.getPlainBody());
        if (!parsed) {
          const htmlText = msg.getBody()
            .replace(/<(br|\/p|\/div|\/tr|\/td|\/h\d)[^>]*>/gi, '\n')
            .replace(/<[^>]+>/g, '')
            .replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&');
          parsed = parseIndeedText(htmlText);
        }
        if (!parsed) return;

        rows.push([msg.getDate(), parsed.title, parsed.company, parsed.location, thread.getPermalink(), id]);
        existingIds.add(id);
      });
    });
    start += pageSize;
  } while (threads.length === pageSize);

  if (rows.length) {
    rows.sort((a, b) => a[0] - b[0]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, INDEED_HEADERS.length).setValues(rows);
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd');
  }

  notify(`Done! Added ${rows.length} new Indeed application(s).`);
}

// Pulls title, company and location out of the email text
function parseIndeedText(text) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const i = lines.findIndex(l => /^application submitted$/i.test(l));
  if (i === -1 || !lines[i + 1]) return null;

  const title = lines[i + 1];
  const companyLine = lines.slice(i + 2).find(l => !/^company logo$/i.test(l));
  if (!companyLine) return null;

  // "Anthro Group - Pasig" -> company + location (split on the LAST " - ")
  const sep = companyLine.lastIndexOf(' - ');
  const company = sep > -1 ? companyLine.slice(0, sep).trim() : companyLine;
  const location = sep > -1 ? companyLine.slice(sep + 3).trim() : '';

  return { title, company, location };
}