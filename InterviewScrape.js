// ---------- INTERVIEW INVITATIONS ----------

// Gmail-side search: only emails that contain invitation-style phrases
const INTERVIEW_QUERY =
  '("interview invitation" OR "invitation to interview" OR "invitation for an interview" OR "Initial interview" OR ' +
  '"invite you to interview" OR "invite you for an interview" OR "invited to interview" OR ' +
  '"schedule an interview" OR "schedule your interview" OR "schedule a interview" OR ' +
  '"interview schedule" OR "interview request" OR "interview availability" OR ' +
  '"would like to interview you" OR "like to invite you") ' +
  '-category:promotions -category:social -from:me -iacademy -from:iacademy.edu.ph after:2026/01/01';

const INTERVIEW_HEADERS = ['Email Date', 'Company', 'Job Title', 'Source', 'Match', 'Sender', 'Subject', 'Email Link', 'Message ID'];
const APPLICATION_SHEETS = ['Job Street', 'Indeed', 'LinkedIn', 'random']; // tab names to match against

// Code-side checks (second layer after the Gmail search)
const INVITE_REGEX = new RegExp([
  'interview invit',
  'invit\\w*[^.\\n]{0,60}interview',
  'interview[^.\\n]{0,40}invit',
  '(schedule|arrange|set up|book|conduct)\\s+(an?\\s+|your\\s+|the\\s+)?(\\w+\\s+)?interview',
  'interview\\s+(schedule|request|availability|slot)',
  'would like to interview you',
  'like to invite you',
  'initial interview'
].join('|'), 'i');

const INTERVIEW_EXCLUDE_REGEX = /\bi[\s-]?academy/i;
const INTERVIEW_REJECT_REGEX = /unfortunately|regret to inform|not moving forward|other candidates|decided not to|position has been filled/i;
const INTERVIEW_NOISE_REGEX = /job alert|jobs? (similar|for you)|recommended|new jobs|interview tips|top picks|webinar|newsletter/i;
const SCREENING_EXCLUDE_REGEX = /initial[\s-]+screening/i;

function scrapeInterviewEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('interview') || ss.insertSheet('interview');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(INTERVIEW_HEADERS);
    sheet.getRange(1, 1, 1, INTERVIEW_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Message IDs already in the sheet (column 9) to avoid duplicates
  const lastRow = sheet.getLastRow();
  const existingIds = new Set(
    lastRow > 1 ? sheet.getRange(2, 9, lastRow - 1, 1).getValues().flat() : []
  );
  const existingCompanies = new Set(
  lastRow > 1
    ? sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().map(c => normalizeText(c)).filter(Boolean)
    : []
);

  const apps = loadApplications();
  const me = Session.getEffectiveUser().getEmail().toLowerCase();
  const cutoff = new Date('2026-01-01T00:00:00+08:00');
  let rows = [];
  const pageSize = 100;
  let start = 0;
  let threads;

  do {
    threads = GmailApp.search(INTERVIEW_QUERY, start, pageSize);
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

        if (INTERVIEW_EXCLUDE_REGEX.test(from + ' ' + subject + ' ' + body)) return; // iAcademy
        if (SCREENING_EXCLUDE_REGEX.test(subject + ' ' + head)) return; // initial screening, not an interview
        if (!INVITE_REGEX.test(subject + ' ' + head)) return;                         // must be an invitation
        if (INTERVIEW_NOISE_REGEX.test(subject)) return;
        if (INTERVIEW_REJECT_REGEX.test(subject + ' ' + head)) return;

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

  // Keep only the earliest invitation per company
  rows.sort((a, b) => a[0] - b[0]);
  rows = rows.filter(r => {
    const key = normalizeText(r[1]); // Company column
    if (!key) return true;           // can't tell the company, so keep it for manual review
    if (existingCompanies.has(key)) return false;
    existingCompanies.add(key);
    return true;
  });

  if (rows.length) {
    rows.sort((a, b) => a[0] - b[0]);
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, INTERVIEW_HEADERS.length).setValues(rows);
    sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');

    // Highlight rows that need a manual check
    rows.forEach((r, i) => {
      if (r[4] !== 'Matched') {
        sheet.getRange(startRow + i, 1, 1, INTERVIEW_HEADERS.length).setBackground('#fff2cc');
      }
    });
  }

  const unsure = rows.filter(r => r[4] !== 'Matched').length;
  notify(
    `Done! Added ${rows.length} new interview invitation(s). ${unsure} highlighted row(s) need a quick check.`
  );
}

// Reads Date / Title / Company (columns A-C) from all application sheets
function loadApplications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const apps = [];
  APPLICATION_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(([date, title, company]) => {
      if (!company) return;
      const v = companyVariants(String(company));
      if (!v.full) return;
      apps.push({
        source: name,
        date: date instanceof Date ? date : new Date(date),
        title: String(title || '').trim(),
        company: String(company).trim(),
        full: v.full,
        short: v.short
      });
    });
  });
  return apps;
}

// Lowercase, strip punctuation, collapse spaces
function normalizeText(s) {
  return String(s).toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, ' ').trim();
}

// "Kumon Philippines, Inc" -> full: "kumon philippines inc", short: "kumon"
function companyVariants(company) {
  const full = normalizeText(company);
  const short = full
    .replace(/\b(inc|incorporated|corp|corporation|co|ltd|llc|company|philippines|philippine|ph|group|holdings|the)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return { full, short: short.length >= 4 && short !== full ? short : '' };
}

// Finds which application an interview email most likely belongs to
function matchApplication(apps, hay, msgDate) {
  const padded = ' ' + hay + ' ';
  const hits = [];

  apps.forEach(a => {
    if (padded.includes(' ' + a.full + ' ')) {
      hits.push({ a, partial: false, score: 1000 + a.full.length });
    } else if (a.short && padded.includes(' ' + a.short + ' ')) {
      hits.push({ a, partial: true, score: a.short.length });
    }
  });
  if (!hits.length) return null;

  const top = hits.reduce((x, y) => (y.score > x.score ? y : x));
  const sameCompany = hits.filter(h => h.a.full === top.a.full).map(h => h.a);

  // 1. A job title from your applications is mentioned in the email
  const titleHits = sameCompany.filter(a => a.title && padded.includes(' ' + normalizeText(a.title) + ' '));
  if (titleHits.length) {
    titleHits.sort((x, y) => y.date - x.date);
    const chosen = titleHits.find(a => a.date <= msgDate) || titleHits[0];
    return { app: chosen, partial: top.partial, uncertain: false };
  }

  // 2. Only one distinct title applied for at that company
  const distinctTitles = new Set(sameCompany.map(a => normalizeText(a.title)));
  const byDateDesc = sameCompany.slice().sort((x, y) => y.date - x.date);
  if (distinctTitles.size === 1) {
    return { app: byDateDesc[0], partial: top.partial, uncertain: false };
  }

  // 3. Several titles: most recent application before the email, flagged for review
  const chosen = byDateDesc.find(a => a.date <= msgDate) || byDateDesc[0];
  return { app: chosen, partial: top.partial, uncertain: true };
}

// Fallback company guess when nothing matches your application sheets
function guessCompanyFromSender(from) {
  const m = from.match(/^"?([^"<]*?)"?\s*<([^>]+)>/);
  const name = m ? m[1] : '';
  const email = m ? m[2] : from;

  const cleaned = name
    .replace(/\b(careers?|recruiting|recruitment|talent( acquisition)?|hr|jobs|hiring|team|no-?reply|notifications?)\b/gi, '')
    .replace(/[-–|:]+\s*$/, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
  if (cleaned) return cleaned;

  const domain = email.split('@')[1] || '';
  if (domain && !/greenhouse|lever|workable|smartrecruiters|workday|icims|jobvite|ashby|bamboohr|gmail|yahoo|outlook/i.test(domain)) {
    const base = domain.split('.')[0];
    return base.charAt(0).toUpperCase() + base.slice(1);
  }
  return '';
}