// ---------- MANUAL REJECTIONS ----------
const MANUAL_REJECTION_QUERY =
  '("regret to inform" OR "pushing through" OR "not be moving forward" OR "not moving forward" OR ' +
  '"will not be proceeding" OR "not be proceeding" OR "unable to move forward" OR "not to proceed" OR ' +
  '"move forward with other" OR "other candidates" OR "not selected" OR "position has been filled" OR ' +
  '"decided not to" OR "not be able to offer") ' +
  '-"unlikely to progress further" -category:promotions -category:social -from:me after:2026/01/01';

const MANUAL_REJECTION_HEADERS = ['Email Date', 'Company', 'Job Title', 'Source', 'Match', 'Sender', 'Subject', 'Email Link', 'Message ID'];

// Code-side check (second layer after the Gmail search)
const MANUAL_REJECTION_REGEX = new RegExp([
  'regret to (inform|advise|let you know|tell you)',
  '(will not|won[\'’]t|cannot|can[\'’]t|unable to)\\s+(be\\s+)?(pushing through|push through|moving forward|move forward|proceeding|proceed|continuing|continue)',
  '(not|no longer)\\s+(be\\s+)?(pushing through|moving forward|proceeding|continuing)',
  'not (been )?selected',
  'move forward with (other|another)',
  'decided (not )?to (move forward|proceed|pursue)[^.\\n]{0,40}(other|another)',
  'decided not to',
  'other candidates',
  'position (has been|is now) filled',
  'not be able to offer'
].join('|'), 'i');

const AUTOMATED_REJECTION_REGEX = /unlikely to progress further/i;
const MANUAL_REJECTION_NOISE_REGEX = /job alert|jobs? (similar|for you)|recommended|new jobs|top picks|webinar|newsletter/i;

function scrapeManualRejectionEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('rejection') || ss.insertSheet('rejection');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(MANUAL_REJECTION_HEADERS);
    sheet.getRange(1, 1, 1, MANUAL_REJECTION_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Message IDs already in the sheet (column 9) to avoid duplicates
  const lastRow = sheet.getLastRow();
  const existingIds = new Set(
    lastRow > 1 ? sheet.getRange(2, 9, lastRow - 1, 1).getValues().flat() : []
  );

  const apps = loadApplications(); // reuses the helper from the interview scraper
  const me = Session.getEffectiveUser().getEmail().toLowerCase();
  const cutoff = new Date('2026-01-01T00:00:00+08:00');
  const rows = [];
  const pageSize = 100;
  let start = 0;
  let threads;

  do {
    threads = GmailApp.search(MANUAL_REJECTION_QUERY, start, pageSize);
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        const id = msg.getId();
        if (existingIds.has(id)) return;
        if (msg.getDate() < cutoff) return;

        const from = msg.getFrom();
        if (me && from.toLowerCase().includes(me)) return; // skip your own replies

        const subject = msg.getSubject();
        const body = msg.getPlainBody();
        const head = body.slice(0, 3000);

        if (AUTOMATED_REJECTION_REGEX.test(subject + ' ' + head)) return;   // handled by the automated sheet
        if (MANUAL_REJECTION_NOISE_REGEX.test(subject)) return;
        if (!MANUAL_REJECTION_REGEX.test(subject + ' ' + head)) return;     // must read like a rejection

        const hay = normalizeText(subject + ' ' + from + ' ' + body.slice(0, 4000));
        const match = matchApplication(apps, hay, msg.getDate());

        let company, title, source, label;
        if (match) {
          company = match.app.company;
          title = match.app.title;
          source = match.app.source;
          label = match.uncertain ? 'Matched - check title'
                : match.partial ? 'Matched - partial name'
                : 'Matched';
        } else {
          company = guessCompanyFromSender(from);
          title = '';
          source = '';
          label = 'Unmatched';
        }

        rows.push([msg.getDate(), company, title, source, label, from, subject, thread.getPermalink(), id]);
        existingIds.add(id);
      });
    });
    start += pageSize;
  } while (threads.length === pageSize);

  if (rows.length) {
    rows.sort((a, b) => a[0] - b[0]);
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, MANUAL_REJECTION_HEADERS.length).setValues(rows);
    sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');

    // Highlight rows that need a manual check
    rows.forEach((r, i) => {
      if (r[4] !== 'Matched') {
        sheet.getRange(startRow + i, 1, 1, MANUAL_REJECTION_HEADERS.length).setBackground('#fff2cc');
      }
    });
  }

  const unsure = rows.filter(r => r[4] !== 'Matched').length;
  notify(
    `Done! Added ${rows.length} new rejection(s). ${unsure} highlighted row(s) need a quick check.`
  );
}