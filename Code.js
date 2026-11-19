// ---------- RUN EVERYTHING ----------
function runAllScrapers() {
  // Order matters: applications first, because interviews/rejections match against them
  const steps = [
    ['JobStreet', scrapeJobStreetEmails],
    ['Indeed', scrapeIndeedEmails],
    ['LinkedIn', scrapeLinkedInEmails],
    ['Random', scrapeRandomEmails],
    ['Sent applications', scrapeSentApplications],
    ['Interviews', scrapeInterviewEmails],
    ['Automated rejections', scrapeRejectionEmails],
    ['Manual rejections', scrapeManualRejectionEmails],
    ['Closed applications', scrapeClosedEmails]
  ];

  const MAX_MS = 4.5 * 60 * 1000; // stop starting new steps after 4.5 min (Apps Script limit is 6)
  const startTime = Date.now();
  const skipped = [];
  RUN_LOG = [];

  let log;
  try {
    steps.forEach(([label, fn]) => {
      if (Date.now() - startTime > MAX_MS) {
        skipped.push(label);
        return;
      }
      try {
        RUN_LOG.push(`${label}: ` + '...');
        const idx = RUN_LOG.length - 1;
        fn();
        // fn() pushed its own summary after our placeholder; merge them
        const summary = RUN_LOG.splice(idx + 1).join(' ');
        RUN_LOG[idx] = `${label}: ${summary}`.replace(/: \.\.\.\s*/, ': ');
        SpreadsheetApp.flush(); // make sure new rows are saved before the next step reads them
      } catch (e) {
        RUN_LOG[RUN_LOG.length - 1] = `${label}: FAILED - ${e.message}`;
      }
    });
  } finally {
    log = RUN_LOG;
    RUN_LOG = null; // restore normal alerts
  }

  let message = log.join('\n');
  if (skipped.length) {
    message += `\n\nRan out of time before: ${skipped.join(', ')}. Run "Run ALL scrapers" again to continue. Nothing is duplicated.`;
  }
  SpreadsheetApp.getUi().alert('Scrape complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Job Tracker')
    .addItem('Run ALL scrapers', 'runAllScrapers')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Applications')
      .addItem('Scrape JobStreet emails', 'scrapeJobStreetEmails')
      .addItem('Scrape Indeed emails', 'scrapeIndeedEmails')
      .addItem('Scrape LinkedIn emails', 'scrapeLinkedInEmails')
      .addItem('Scrape random emails', 'scrapeRandomEmails')
      .addItem('Scrape sent applications', 'scrapeSentApplications'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Results')
      .addItem('Scrape interview emails', 'scrapeInterviewEmails')
      .addItem('Scrape automated rejections', 'scrapeRejectionEmails')
      .addItem('Scrape rejection emails', 'scrapeManualRejectionEmails')
      .addItem('Scrape closed applications', 'scrapeClosedEmails'))
    .addSeparator()
    .addItem('Set up dashboard', 'setupDashboard')
    .addItem('Set up monthly breakdown', 'setupMonthlyDashboard')
    .addToUi();
}