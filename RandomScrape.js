const RANDOM_QUERY =
  'subject:("thank you for applying" OR "thanks for applying" OR "application received" OR ' +
  '"we received your application" OR "application confirmation" OR "your application") ' +
  '-from:jobstreet -from:indeed -from:linkedin -from:me after:2026/01/01';
const RANDOM_HEADERS = ['Date Applied', 'Job Title', 'Company', 'Sender', 'Subject', 'Email Link', 'Message ID'];
const ATS_DOMAINS = /greenhouse|lever|workable|smartrecruiters|myworkday|workday|icims|jobvite|ashby|bamboohr|breezy|recruitee|teamtailor/i;

function scrapeRandomEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('random') || ss.insertSheet('random');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(RANDOM_HEADERS);
    sheet.getRange(1, 1, 1, RANDOM_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Message IDs already in the sheet (column 7) to avoid duplicates
  const lastRow = sheet.getLastRow();
  const existingIds = new Set(
    lastRow > 1 ? sheet.getRange(2, 7, lastRow - 1, 1).getValues().flat() : []
  );

  const cutoff = new Date('2026-01-01T00:00:00+08:00');
  const rows = [];
  const pageSize = 100;
  let start = 0;
  let threads;

  do {
    threads = GmailApp.search(RANDOM_QUERY, start, pageSize);
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        const id = msg.getId();
        if (existingIds.has(id)) return;
        if (msg.getDate() < cutoff) return;

        const subject = msg.getSubject();
        const body = msg.getPlainBody();
        const from = msg.getFrom();

        // Skip your own emails and rejection / update emails
        if (/jobstreet|indeed|linkedin/i.test(from)) return;
        if (/unfortunately|regret to inform|not moving forward|other candidates|decided not to/i
              .test(subject + ' ' + body.slice(0, 1500))) return;

        // Must look like an application confirmation
        if (!/appl(y|ied|ication)/i.test(subject + ' ' + body.slice(0, 500))) return;

        const parsed = parseRandomEmail(subject, body, from);
        rows.push([msg.getDate(), parsed.title, parsed.company, from, subject, thread.getPermalink(), id]);
        existingIds.add(id);
      });
    });
    start += pageSize;
  } while (threads.length === pageSize);

  if (rows.length) {
    rows.sort((a, b) => a[0] - b[0]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, RANDOM_HEADERS.length).setValues(rows);
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd');
  }

  notify(`Done! Added ${rows.length} new random application(s). Please review the Job Title and Company columns.`);
}

// Best-effort guess of job title and company
function parseRandomEmail(subject, body, from) {
  let title = '';
  let company = '';
  const clean = s => s.replace(/[.!:,\-–]+$/, '').replace(/^["'“]|["'”]$/g, '').trim();

  // 1. Subject patterns: "Your application for X at Y", "Thank you for applying to Y"
  let m = subject.match(/application (?:for|to)\s+(?:the\s+)?(.+?)\s+(?:at|with)\s+(.+)$/i)
       || subject.match(/appl(?:ying|ied)\s+(?:for|to)\s+(?:the\s+)?(.+?)\s+(?:at|with)\s+(.+)$/i);
  if (m) {
    title = clean(m[1].replace(/\s+(position|role)$/i, ''));
    company = clean(m[2]);
  } else {
    m = subject.match(/appl(?:ying|ication)\s+(?:to|with)\s+(.+)$/i);
    if (m) company = clean(m[1]);
    m = subject.match(/application (?:received|confirmation)\s*[:\-–]\s*(.+)$/i)
     || subject.match(/application for\s+(?:the\s+)?(.+?)(?:\s+position|\s+role)?$/i);
    if (m && !title) title = clean(m[1]);
  }

  // 2. Body patterns for the title: "position of X", "Position: X", "applying for the X role"
  if (!title) {
    const head = body.slice(0, 2000);
    const b = head.match(/(?:position|role|job title|applying for)(?:\s+of)?\s*[:\-]?\s*(?:the\s+)?["“]?([A-Z][^\n.,"”]{2,60}?)["”]?(?:\s+(?:position|role))?\s*(?:\n|\.|,| at | with |$)/);
    if (b) title = clean(b[1]);
  }

  // 3. Company from the sender's display name or domain
  if (!company) {
    const nameMatch = from.match(/^"?([^"<]*?)"?\s*<([^>]+)>/);
    const name = nameMatch ? nameMatch[1] : '';
    const email = nameMatch ? nameMatch[2] : from;

    company = name
      .replace(/\b(careers?|recruiting|recruitment|talent( acquisition)?|hr|jobs|hiring|team|no-?reply|notifications?)\b/gi, '')
      .replace(/[-–|:]+\s*$/, '')
      .replace(/\s{2,}/g, ' ')
      .trim();

    if (!company) {
      const domain = (email.split('@')[1] || '');
      if (domain && !ATS_DOMAINS.test(domain)) {
        const base = domain.split('.')[0];
        company = base.charAt(0).toUpperCase() + base.slice(1);
      }
    }
  }

  return { title, company };
}