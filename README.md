# Half Marathon Plan

A phone-friendly training plan and log for the half marathon on **Saturday, February 6, 2027**.

- **Today**: the day's session as a race bib, with target paces worked out from your current VDOT.
- **Plan**: all 15 weeks; tick off runs as you go.
- **Paces**: enter your 5K (Nov 21) and 10K (Dec 19) test times and every pace updates.
- **Progress**: planned vs logged km per week, backup and restore, CSV export.
- **Guide**: gym sessions, hill sprints, strides, heat and fuel.

Your ticks and logs are saved in the browser on your phone. Use *Back up data* on the Progress tab to keep a copy.

Add it to your home screen (Share > Add to Home Screen on iPhone, or the browser menu on Android) to open it like an app; it works offline once loaded.

`half-marathon-plan.xlsx` is the same plan as a spreadsheet you can open in Excel or Google Sheets.

Built from Daniels' Running Formula, Advanced Marathoning, Run Faster and Run Elite, plus the research cited in the full plan document.

Live app: https://pengkistan.github.io/marathonplan/

## Google Sheet sync

`apps-script/Code.gs` turns a Google Sheet into the app's online log:

1. In a new Google Sheet: Extensions > Apps Script, replace the code with `Code.gs`, save.
2. Run `setup` once (it builds the Start here, Plan, Weekly, Paces and Gym sheets) and allow access.
3. Deploy > New deployment > Web app, execute as **Me**, access **Anyone**. Copy the link ending in `/exec`.
4. In the app: Progress > Google Sheet > paste the link > Connect. "Copy link for another phone" connects other devices.

Every tick, log, test time and max heart rate is saved to the Plan and Start here sheets. Edits typed into the sheet's yellow columns come back to the app the next time it opens. Keep the `/exec` link private: anyone with it can read and change your log.
