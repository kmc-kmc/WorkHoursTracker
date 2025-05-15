📊 Monthly Work Hours & Progress Tracker (Google Calendar + Google Sheets)

This Google Apps Script project tracks total working hours per month by reading events from your Google Calendar and logging the data into a connected Google Sheet. It’s designed to help freelancers, contractors, or part-time workers monitor their workload and stay within their target hours.

By automatically logging daily and monthly totals and enforcing a weekly hour cap, it ensures better control over time management and project pacing.

✅ Features
Reads events from a specified Google Calendar (e.g., "CA Job").
Calculates total hours worked per month and logs it into a Google Sheet.
Optionally enforces a weekly hour cap (e.g., 30 hours/week).
Automatically updates or adds monthly data rows.
Ideal for remote workers needing transparency or time-tracking.

🛠 Setup Instructions
Create or Open a Google Sheet
Add a column for Month (e.g., May 2025) in Column A.

Open Apps Script
Go to Extensions → Apps Script in your Google Sheet.
Paste the full script (calculateAndLogCurrentMonthHours() and countCAJobHours() functions).

Customize Settings
Change calendarName = "CA Job" if your calendar has a different name.
Adjust maxWeeklyHours = 30 to change the weekly cap.
You can manually add months in Column A (e.g., May 2025) for backtracking.

Run the Functions
Run calculateAndLogCurrentMonthHours() to log total hours for the current month.
Run countCAJobHours() to calculate hours respecting the weekly cap and update Column D.

Automation
Use Triggers to run calculateAndLogCurrentMonthHours when the Google sheet is open.

📊 Example
https://docs.google.com/spreadsheets/d/1xyugAW-SriAawgp9hqZz_PAICCuMiDTLburPAjNmtuw/edit?gid=0#gid=0

📦 Functions Explained
calculateAndLogCurrentMonthHours()
Fetches all events in the current month from your calendar.
Sums total duration in hours.
Checks if the month exists in Column A; if not, it adds a new row.
Writes total hours into Column B.

countCAJobHours()
Reads each date in Column A.
For each month:
Fetches events.
Groups them into weeks.
Sums hours per week, capping each week at a maximum (e.g., 30 hours).
Writes the capped total to Column D.

📌 Notes
Events must have clear start and end times.
The script handles multi-day events and distributes time accordingly.
Weekly limits help avoid burnout or contractual overages.

🚀 Future Improvements
Add a bar for showing the the progress of work hours accumulation.
