# Job Application Tracker

A Google Apps Script project that scans your Gmail inbox for job application emails and logs them into Google Sheets. It tracks applications from JobStreet, Indeed, LinkedIn, and direct employers, classifies the replies (interview invitations, rejections, closed postings), and builds live dashboards with totals, response rates, and a monthly breakdown.

Everything runs inside your own Google account. No data is sent to any outside server.

## Features

- **Application scrapers** for JobStreet, Indeed, LinkedIn, applications sent directly by email, and "random" applications to company sites
- **Result scrapers** for interview invitations, automated rejections, personal rejections, and closed job postings
- **Automatic matching** of interview and rejection emails to your original applications by company name, filling in the job title and source
- **Duplicate protection** using each email's Message ID, so re-running never creates repeat rows or overwrites your manual edits
- **Live dashboards** built with formulas, so they update by themselves after each scrape
- **One-click pipeline** from a custom menu that runs every scraper in the correct order

## Sheets created

| Sheet | Contents |
|---|---|
| `Job Street` | Date applied, job title, company, email link, message ID |
| `Indeed` | Date applied, job title, company, location, email link, message ID |
| `LinkedIn` | Date applied, job title, company, location, email link, message ID |
| `random` | Applications to company sites (best-effort title and company guess) |
| `email application` | Applications you sent by email, with recipient and attachment names |
| `interview` | Interview invitations matched to company and job title |
| `automated rejections` | JobStreet automated "unlikely to progress further" emails |
| `rejection` | Personal rejection emails ("regret to inform", "not moving forward", etc.) |
| `closed applications` | JobStreet emails saying a posting you applied for has expired |
| `Dashboard` | Totals and percentages across all sheets |
| `Monthly` | Applications and results broken down by month |

## Setup

1. Open Google sheets template link 
<a href="https://docs.google.com/spreadsheets/d/1DvMc4VCm3c5PP9OcG901NkjGsvY0slPQlM0AyNOISmc/edit?usp=sharing" target="_blank">Click Here for Template</a>
2. Create your own copy
<img width="999" height="757" alt="image" src="https://github.com/user-attachments/assets/17e4cd78-dd79-45f4-b83e-4cea212ead92" />

3. Open Job Tracker Menu on your own copy
4. Run all Scrapers
<img width="967" height="664" alt="image" src="https://github.com/user-attachments/assets/de2b1ba8-c020-402a-82d9-f059b371f678" />
![Uploading image.png…]()

## Usage

Open the **Job Tracker** menu:

- **Run ALL scrapers** runs every scraper in order (applications first, then results) and shows one summary at the end.
- **Applications** submenu runs the JobStreet, Indeed, LinkedIn, random, and sent-email scrapers individually.
- **Results** submenu runs the interview, rejection, and closed-posting scrapers individually.
- **Set up dashboard** and **Set up monthly breakdown** build the summary sheets. Run these once. They use formulas, so they refresh on their own.

Suggested order on a first run: application scrapers, then result scrapers, then the dashboards.

## How it works

1. Each scraper uses `GmailApp.search()` with a query for its email type (for example `from:jobstreet "successfully submitted" after:2026/01/01`).
2. Each message's text is flattened and parsed with a regex to extract the job title and company. If the plain-text body fails, the script falls back to the HTML body.
3. Each message's ID is checked against the sheet's Message ID column, and only new emails are appended.
4. Interview and rejection scrapers load your application sheets and match the email to a company, trying the full name first and then a shortened name (for example "Kumon Philippines, Inc" and "Kumon").
5. Rows that need a manual check are highlighted yellow.

## Customizing

- **Date cutoff:** the queries use `after:2026/01/01`, and the code uses a matching `2026-01-01T00:00:00+08:00` cutoff. Change both to start from a different date.
- **Excluded senders:** the interview scraper excludes iAcademy through `INTERVIEW_EXCLUDE_REGEX`. Edit that regex to exclude other names.
- **Interview wording:** add phrases to `INTERVIEW_QUERY` and `INVITE_REGEX` if real invitations are missed.
- **Rejection wording:** add phrases to `MANUAL_REJECTION_QUERY` and `MANUAL_REJECTION_REGEX`.
- **Sheet names:** the dashboards and matching depend on the exact tab names, so don't rename them without updating the code.

## Known limitations

- **Templates change.** The parsers match the current wording of JobStreet, Indeed, and LinkedIn emails. If a platform changes its template, the pattern needs updating. Rows the pattern can't read are still logged, with blank cells.
- **Direct-employer emails are best effort.** The `random` and `email application` scrapers guess titles and companies from the subject, body, and sender, so review the first batch.
- **Matching can be uncertain.** If you applied for several roles at one company and the email doesn't name the role, the script picks the most recent application and highlights the row.
- **Monthly results are counted by the month the email arrived**, not the month you applied, so individual monthly percentages can look uneven.
- **Search is thread-based.** In rare cases, Gmail's thread results can lag and miss the newest message in a thread. Searching messages directly with the advanced Gmail service avoids this.
- **Apps Script has a runtime limit of about 6 minutes per run.** On a very large inbox, run the pipeline again. Duplicate protection lets it continue without repeating rows.
- **Deleting the Message ID column's contents** causes the next run to add duplicates. Hide the column instead.

## Privacy

The script reads email only to extract job-related details (date, title, company, and a link to the email) and writes them to the spreadsheet you own. It runs under your Google account's permissions and does not transmit data anywhere else.

## Tech stack

Google Apps Script, JavaScript (ES6), Gmail service, Google Sheets (formulas and formatting)

## License

Add your preferred license here (for example MIT).
