## CareerOS

Spreadsheet-driven job application dashboard built with React and Vite.

### Run locally

`npm install` then `npm run dev`.

### Import your tracker

Use Update Data to upload an Excel (.xlsx/.xls) or CSV file. The dashboard reads the workbook in your browser, lets you switch between worksheets, and maps common job-tracker columns such as company, job title, location, work mode, status, salary, skills and application URL. Your spreadsheet is not uploaded to a backend. Data stays in the current browser session and must be re-uploaded after refresh.

### Deploy

Vercel build command: `npm run build`; output directory: `dist`.
