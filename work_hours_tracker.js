function calculateAndLogCurrentMonthHours() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const calendarName = "CA Job"; // The name of your Google Calendar

  // 1. Get current date information
  const now = new Date(); // Uses the current date when the script runs
  const currentYear = now.getFullYear();
  const currentMonthIndex = now.getMonth(); // 0-indexed (January=0, May=4, etc.)
  const currentMonthName = now.toLocaleString('default', { month: 'long' }); // e.g., "May"
  // Ensure our script-generated string is also trimmed, though toLocaleString usually doesn't add extra spaces.
  const scriptGeneratedMonthYearString = (currentMonthName + " " + currentYear).trim();

  Logger.log('Script started. Processing for current month: "' + scriptGeneratedMonthYearString + '"');

  // 2. Get the target calendar
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length === 0) {
    SpreadsheetApp.getUi().alert('Calendar not found: ' + calendarName);
    Logger.log('Error: Calendar not found - ' + calendarName);
    return;
  }
  const calendar = calendars[0];
  Logger.log('Calendar "' + calendarName + '" found.');

  // 3. Calculate start and end of the current month
  const firstDayOfCurrentMonth = new Date(currentYear, currentMonthIndex, 1);
  const lastDayOfCurrentMonth = new Date(currentYear, currentMonthIndex + 1, 0, 23, 59, 59);
  Logger.log('Date range: ' + firstDayOfCurrentMonth.toString() + ' to ' + lastDayOfCurrentMonth.toString());

  // 4. Fetch events for the current month
  let totalHoursForCurrentMonth = 0;
  let eventCountThisMonth = 0;
  try {
    const events = calendar.getEvents(firstDayOfCurrentMonth, lastDayOfCurrentMonth);
    eventCountThisMonth = events.length;
    Logger.log('Found ' + eventCountThisMonth + ' events for ' + scriptGeneratedMonthYearString);

    events.forEach(function(event) {
      const eventStartTime = event.getStartTime().getTime();
      const eventEndTime = event.getEndTime().getTime();
      const durationHours = (eventEndTime - eventStartTime) / (1000 * 60 * 60);
      totalHoursForCurrentMonth += durationHours;
    });
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error fetching events for ' + scriptGeneratedMonthYearString + ': ' + e.toString());
    Logger.log('Error fetching events for ' + scriptGeneratedMonthYearString + ': ' + e.toString());
    return;
  }

  Logger.log('Total hours calculated for ' + scriptGeneratedMonthYearString + ': ' + totalHoursForCurrentMonth);

  // 5. Write to the sheet: Find the row for the current month or add it.
  // It's better to read the display values if we are comparing strings that are visually formatted in the sheet
  const dataRange = sheet.getRange("A1:A" + sheet.getLastRow()); // Get column A up to the last row with content
  const columnAValues = dataRange.getDisplayValues(); // Use getDisplayValues() for string comparison
  let targetRow = -1;

  // Prepare the script's target string for comparison (lowercase)
  const scriptMonthLower = scriptGeneratedMonthYearString.toLowerCase();

  for (let i = 0; i < columnAValues.length; i++) {
    // Ensure sheetCellValue is treated as a string, trim it, and convert to lowercase
    const sheetCellValue = (columnAValues[i][0] || "").toString().trim(); // (value || "") handles potential null/undefined
    const sheetCellLower = sheetCellValue.toLowerCase();

    // Detailed logging for comparison
    Logger.log('Comparing (Row ' + (i + 1) + '): Script Value (lower)="' + scriptMonthLower + '" <==> Sheet Value (lower)="' + sheetCellLower + '" (Original Sheet: "' + columnAValues[i][0] + '")');

    if (sheetCellLower === scriptMonthLower) {
      targetRow = i + 1; // Sheet rows are 1-indexed, array is 0-indexed
      Logger.log('MATCH FOUND: "' + scriptGeneratedMonthYearString + '" in column A at row ' + targetRow);
      break;
    }
  }

  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange("A" + targetRow).setValue(scriptGeneratedMonthYearString); // Use the original cased version for writing
    Logger.log('"' + scriptGeneratedMonthYearString + '" not found in column A. Adding it to row ' + targetRow);
  }

  sheet.getRange("B" + targetRow).setValue(totalHoursForCurrentMonth.toFixed(2));
  Logger.log('Wrote ' + totalHoursForCurrentMonth.toFixed(2) + ' hours to cell B' + targetRow);

  
}

function countCAJobHours() {
  const calendarName = "CA Job";
  const maxWeeklyHours = 30;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();

  const calendar = CalendarApp.getCalendarsByName(calendarName);
  if (calendar.length === 0) {
    Logger.log(`Calendar "${calendarName}" not found.`);
    return;
  }
  const caCalendar = calendar[0];

  for (let i = 1; i <= lastRow; i++) {
    const monthCell = sheet.getRange(i, 1).getValue();
    if (monthCell instanceof Date) {
      const year = monthCell.getFullYear();
      const month = monthCell.getMonth(); // 0-indexed

      const startDateOfMonth = new Date(year, month, 1);
      const endDateOfMonth = new Date(year, month + 1, 0, 23, 59, 59); // Last day of the month

      let totalHoursThisMonth = 0;
      let events = caCalendar.getEvents(startDateOfMonth, endDateOfMonth);

      // Group events by week (Sunday to Saturday)
      const eventsByWeek = {};
      events.forEach(event => {
        let start = new Date(event.getStartTime());
        let end = new Date(event.getEndTime());
        let durationMs = end.getTime() - start.getTime();
        let durationHours = durationMs / (1000 * 60 * 60);

        // Iterate through each day the event spans
        let currentDate = new Date(start);
        while (currentDate <= end) {
          const dayOfWeek = currentDate.getDay(); // 0 for Sunday, 6 for Saturday
          const weekStartDate = new Date(currentDate);
          weekStartDate.setDate(currentDate.getDate() - dayOfWeek); // Get the Sunday of the week
          weekStartDate.setHours(0, 0, 0, 0);

          const weekEndDate = new Date(weekStartDate);
          weekEndDate.setDate(weekStartDate.getDate() + 6); // Get the Saturday of the week
          weekEndDate.setHours(23, 59, 59, 999);

          const weekKey = weekStartDate.toDateString();

          if (!eventsByWeek[weekKey]) {
            eventsByWeek[weekKey] = 0;
          }

          // Check if the event falls within the current month
          if (currentDate.getMonth() === month && currentDate.getFullYear() === year) {
            // Calculate the portion of the event within the current day
            let overlapStart = new Date(Math.max(start.getTime(), currentDate.getTime()));
            let overlapEnd = new Date(Math.min(end.getTime(), new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate(), 23, 59, 59, 999).getTime()));
            let overlapDurationMs = overlapEnd.getTime() - overlapStart.getTime();
            let overlapDurationHours = overlapDurationMs / (1000 * 60 * 60);

            eventsByWeek[weekKey] += overlapDurationHours;
          }

          // Move to the next day
          currentDate.setDate(currentDate.getDate() + 1);
        }
      });

      // Sum the hours for the current month, respecting the weekly limit
      let weeklyHoursCounter = {};
      for (const weekKey in eventsByWeek) {
        const weekStartDate = new Date(weekKey);
        if (weekStartDate.getMonth() === month && weekStartDate.getFullYear() === year) {
          if (!weeklyHoursCounter[weekKey]) {
            weeklyHoursCounter[weekKey] = 0;
          }
          weeklyHoursCounter[weekKey] += eventsByWeek[weekKey];
          totalHoursThisMonth += Math.min(weeklyHoursCounter[weekKey], maxWeeklyHours);
        }
      }

      sheet.getRange(i, 4).setValue(totalHoursThisMonth);
    }
  }
}
