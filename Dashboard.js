// ---------- DASHBOARD ----------
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Make sure every sheet the formulas point to exists
  ['Job Street', 'Indeed', 'LinkedIn', 'random', 'email application',
   'interview', 'automated rejections', 'rejection', 'closed applications'].forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  const sh = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard', 0);
  sh.clear();

  const count = name => `COUNTA('${name}'!A2:A)`;

  const data = [
    ['Job Application Dashboard', '', ''],
    ['Updates automatically from your other sheets', '', ''],
    ['', '', ''],
    ['Metric', 'Count', '% of Applications'],
    ['Applications Sent', '=B20', '=IFERROR(B5/B5,0)'],
    ['Interview Opportunities', '=' + count('interview'), '=IFERROR(B6/B5,0)'],
    ['Rejections Received', '=B8+B9', '=IFERROR(B7/B5,0)'],
    ['   Automated rejections', '=' + count('automated rejections'), '=IFERROR(B8/B5,0)'],
    ['   Personal rejections', '=' + count('rejection'), '=IFERROR(B9/B5,0)'],
    ['Closed Applications', '=' + count('closed applications'), '=IFERROR(B10/B5,0)'],
    ['No response yet (estimate)', '=MAX(0,B5-B6-B7)', '=IFERROR(B11/B5,0)'],
    ['No response, excluding closed (estimate)', '=MAX(0,B5-B6-B7-B10)', '=IFERROR(B12/B5,0)'],
    ['', '', ''],
    ['Applications by Source', 'Count', '% of Applications'],
    ['JobStreet', '=' + count('Job Street'), '=IFERROR(B15/$B$20,0)'],
    ['Indeed', '=' + count('Indeed'), '=IFERROR(B16/$B$20,0)'],
    ['LinkedIn', '=' + count('LinkedIn'), '=IFERROR(B17/$B$20,0)'],
    ['Random / company sites', '=' + count('random'), '=IFERROR(B18/$B$20,0)'],
    ['Sent by email', '=' + count('email application'), '=IFERROR(B19/$B$20,0)'],
    ['Total', '=SUM(B15:B19)', '=IFERROR(B20/$B$20,0)']
  ];

  sh.getRange(1, 1, data.length, 3).setValues(data);

  // Formatting
  sh.getRange('A1').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setFontStyle('italic').setFontColor('#666666');
  ['A4:C4', 'A14:C14'].forEach(r =>
    sh.getRange(r).setFontWeight('bold').setBackground('#d9ead3'));
  sh.getRange('A5:C5').setFontWeight('bold');
  sh.getRange('A20:C20').setFontWeight('bold').setBorder(true, null, null, null, null, null);
  sh.getRange('C5:C12').setNumberFormat('0.0%');
  sh.getRange('C15:C20').setNumberFormat('0.0%');
  sh.getRange('B5:C20').setHorizontalAlignment('right');
  sh.setColumnWidth(1, 280);
  sh.setColumnWidths(2, 2, 130);
  sh.setHiddenGridlines(true);
}