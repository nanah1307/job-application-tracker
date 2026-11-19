// ---------- MONTHLY BREAKDOWN ----------
function setupMonthlyDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Make sure every sheet the formulas point to exists
  ['Job Street', 'Indeed', 'LinkedIn', 'random', 'email application',
   'interview', 'automated rejections', 'rejection', 'closed applications'].forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  const sh = ss.getSheetByName('Monthly') || ss.insertSheet('Monthly', 1);
  sh.clear();
  sh.getDataRange().clearFormat();

  // Counts rows in a sheet whose date (column A) falls in the month shown in column A of this row
  const cnt = (name, r) =>
    `COUNTIFS('${name}'!A:A,">="&$A${r},'${name}'!A:A,"<"&EDATE($A${r},1))`;

  const APP_SHEETS = ['Job Street', 'Indeed', 'LinkedIn', 'random', 'email application'];
  const COLS = 11;

  const data = [
    ['Monthly Breakdown', '', '', '', '', '', '', '', '', '', ''],
    ['Applications counted by the month you applied; results counted by the month the email arrived', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', ''],
    ['Month', 'Applications', 'Interviews', 'Automated Rejections', 'Manual Rejections', 'Total Rejections',
     'Closed', 'Interview % of Apps', 'Automated Rej. % of Apps', 'Manual Rej. % of Apps', 'Total Rej. % of Apps']
  ];

  const firstRow = 5;
  const months = 12;
  for (let i = 0; i < months; i++) {
    const r = firstRow + i;
    data.push([
      i === 0 ? '=DATE(2026,1,1)' : `=EDATE(A${r - 1},1)`,
      '=' + APP_SHEETS.map(n => cnt(n, r)).join('+'),
      '=' + cnt('interview', r),
      '=' + cnt('automated rejections', r),
      '=' + cnt('rejection', r),
      `=D${r}+E${r}`,
      '=' + cnt('closed applications', r),
      `=IFERROR(C${r}/B${r},0)`,
      `=IFERROR(D${r}/B${r},0)`,
      `=IFERROR(E${r}/B${r},0)`,
      `=IFERROR(F${r}/B${r},0)`
    ]);
  }

  const lastMonthRow = firstRow + months - 1;
  const totalRow = lastMonthRow + 1;
  const sum = col => `=SUM(${col}${firstRow}:${col}${lastMonthRow})`;
  data.push([
    'Total',
    sum('B'), sum('C'), sum('D'), sum('E'), sum('F'), sum('G'),
    `=IFERROR(C${totalRow}/B${totalRow},0)`,
    `=IFERROR(D${totalRow}/B${totalRow},0)`,
    `=IFERROR(E${totalRow}/B${totalRow},0)`,
    `=IFERROR(F${totalRow}/B${totalRow},0)`
  ]);

  sh.getRange(1, 1, data.length, COLS).setValues(data);
  sh.getRange(firstRow, 2, months + 1, 6).setNumberFormat('0'); // counts: columns B to G
  
  // Formatting
  sh.getRange('A1').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setFontStyle('italic').setFontColor('#666666');
  sh.getRange(4, 1, 1, COLS).setFontWeight('bold').setBackground('#d9ead3').setWrap(true);
  sh.getRange(firstRow, 1, months, 1).setNumberFormat('mmmm yyyy').setHorizontalAlignment('left');
  sh.getRange(firstRow, 8, months + 1, 4).setNumberFormat('0.0%');
  sh.getRange(totalRow, 1, 1, COLS).setFontWeight('bold').setBorder(true, null, null, null, null, null);
  sh.getRange(4, 2, months + 2, COLS - 1).setHorizontalAlignment('right');
  sh.setColumnWidth(1, 150);
  sh.setColumnWidths(2, COLS - 1, 110);
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(4);
}