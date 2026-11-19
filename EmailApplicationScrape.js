// ---------- EMAIL APPLICATIONS (sent by you) ----------
const SENT_QUERY =
  'in:sent -to:me after:2026/01/01 ' +
  '(subject:(application OR applying OR "apply for" OR "applicant" OR resume OR CV OR "cover letter") ' +
  'OR filename:resume OR filename:cv OR filename:"cover letter")';

const SENT_HEADERS = ['Date Sent', 'Job Title', 'Company', 'Sent To', 'Subject', 'Attachments', 'Email Link', 'Message ID'];
const SENT_CONTENT_REGEX = /appl(y|ying|ied|ication|icant)|resume|\bcv\b|curriculum vitae|cover letter/i;
const FREE_MAIL_DOMAINS = /^(gmail|yahoo|outlook|hotmail|live|icloud|proton|protonmail|aol|msn)\./i;

function scrapeSentApplications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('email application') || ss.insertSheet('email application');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SENT_HEADERS);
    sheet.getRange(1, 1, 1, SENT_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Message IDs already in the sheet (column 8) to avoid duplicates
  const lastRow = sheet.getLastRow();
  const existingIds = new Set(
    lastRow > 1 ? sheet.getRange(2, 8, lastRow - 1, 1).getValues().flat() : []
  );

  // Your address plus any aliases, so we only keep emails YOU sent
  const myAddresses = [Session.getEffectiveUser().getEmail()]
    .concat(GmailApp.getAliases())
    .map(a => a.toLowerCase())
    .filter(Boolean);

  const cutoff = new Date('2026-01-01T00:00:00+08:00');
  const rows = [];
  const pageSize = 100;
  let start = 0;
  let threads;

  do {
    threads = GmailApp.search(SENT_QUERY, start, pageSize);
    threads.forEach(thread => {
      // The application is the first email in the thread. Later messages are follow-ups or replies.
      const msg = thread.getMessages()[0];
      const id = msg.getId();
      if (existingIds.has(id)) return;
      if (msg.getDate() < cutoff) return;

      const from = msg.getFrom().toLowerCase();
      if (!myAddresses.some(a => from.includes(a))) return; // not sent by you

      const subject = msg.getSubject();
      const body = msg.getPlainBody();
      if (!SENT_CONTENT_REGEX.test(subject + ' ' + body.slice(0, 1500))) return;

      const to = msg.getTo();
      const parsed = parseSentApplication(subject, body, to);
      const attachments = msg.getAttachments({ includeInlineImages: false })
        .map(a => a.getName())
        .join(', ');

      rows.push([msg.getDate(), parsed.title, parsed.company, to, subject, attachments, thread.getPermalink(), id]);
      existingIds.add(id);
    });
    start += pageSize;
  } while (threads.length === pageSize);

  if (rows.length) {
    rows.sort((a, b) => a[0] - b[0]);
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, SENT_HEADERS.length).setValues(rows);
    sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');

    // Highlight rows where the title or company couldn't be found
    rows.forEach((r, i) => {
      if (!r[1] || !r[2]) {
        sheet.getRange(startRow + i, 1, 1, SENT_HEADERS.length).setBackground('#fff2cc');
      }
    });
  }

  const unsure = rows.filter(r => !r[1] || !r[2]).length;
  notify(
    `Done! Added ${rows.length} new sent application(s). ${unsure} highlighted row(s) need a quick check.`
  );
}

// Best-effort guess of job title and company from a sent email
function parseSentApplication(subject, body, to) {
  let title = '';
  let company = '';
  const clean = s => s
    .replace(/^(re|fwd?):\s*/i, '')
    .replace(/\b(position|role|vacancy|opening)\b\s*$/i, '')
    .replace(/[.!:,\-–|]+$/, '')
    .replace(/^["'“]|["'”]$/g, '')
    .trim();

  // 1. Subject: "Application for X at Y", "Applying for the X position"
  let m = subject.match(/(?:application|applying|apply|applicant)\s+(?:for|as|to)\s+(?:the\s+|an?\s+)?(?:position of\s+|role of\s+)?(.+?)\s+(?:at|with|@)\s+(.+?)(?:\s+[-–|]\s+.*)?$/i);
  if (m) {
    title = clean(m[1]);
    company = clean(m[2]);
  } else {
    m = subject.match(/(?:application|applying|apply|applicant)\s*(?:for|as|to)?\s*[:\-–]?\s*(?:the\s+|an?\s+)?(?:position of\s+|role of\s+)?(.+?)(?:\s+[-–|]\s+.*)?$/i)
     || subject.match(/^(.+?)\s+(?:application|applicant)\b/i);
    if (m) title = clean(m[1]);
  }

  // Drop things that are clearly not a title
  if (/^(job|my|a|the)?\s*(application|resume|cv|cover letter)$/i.test(title) || title.length > 80) title = '';

  // 2. Body: "applying for the X position at Y"
  if (!title || !company) {
    const head = body.slice(0, 1500);
    const b = head.match(/(?:apply(?:ing)?|application)\s+for\s+(?:the\s+|an?\s+)?(?:position of\s+|role of\s+)?([A-Z][^\n.,]{2,60}?)(?:\s+(?:position|role))?\s+(?:at|with)\s+([A-Z][^\n.,]{2,50})/);
    if (b) {
      if (!title) title = clean(b[1]);
      if (!company) company = clean(b[2]);
    } else if (!company) {
      const c = head.match(/\b(?:at|with|to)\s+([A-Z][A-Za-z0-9&' ]{2,40}?)(?:[.,\n]|\s+(?:for|as)\b)/);
      if (c && !/^(Hiring|Human|The|Dear|Your|Whom)/.test(c[1])) company = clean(c[1]);
    }
  }

  // 3. Company from the recipient: display name first, then the email domain
  if (!company) {
    const first = to.split(',')[0];
    const r = first.match(/^"?([^"<]*?)"?\s*<([^>]+)>/);
    const name = r ? r[1] : '';
    const email = (r ? r[2] : first).trim();
    const domain = email.split('@')[1] || '';

    const nameClean = name
      .replace(/\b(careers?|recruiting|recruitment|talent( acquisition)?|hr|jobs|hiring|team|human resources)\b/gi, '')
      .replace(/[-–|:]+\s*$/, '')
      .replace(/\s{2,}/g, ' ')
      .trim();

    if (domain && !FREE_MAIL_DOMAINS.test(domain) && !ATS_DOMAINS.test(domain)) {
      const base = domain.split('.')[0];
      company = base.charAt(0).toUpperCase() + base.slice(1);
    } else if (nameClean && /\s/.test(nameClean) === false && nameClean.length > 2 && !domain) {
      company = nameClean;
    }
  }

  return { title, company };
}