function calculateAndWriteFormattedSLAsWithLog() {
try{

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SLA Calculation");
  if (!sheet) {
    Logger.log("Sheet 'SLA Calculation' not found");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const holidaysColIndex = headers.indexOf("Holidays") + 1;
  if (holidaysColIndex === 0) {
    Logger.log("Header 'Holidays' not found");
    return;
  }
  Logger.log("Holidays column found at index: " + holidaysColIndex);

  const lastRow = sheet.getLastRow();
  const holidaysRange = sheet.getRange(2, holidaysColIndex, lastRow - 1);
  const holidayRawValues = holidaysRange.getValues().flat();
  const holidayDates = holidayRawValues
    .filter(d => d instanceof Date && !isNaN(d))
    .map(d => new Date(d.getFullYear(), d.getMonth(), d.getDate()));

  Logger.log("Holiday dates: " + holidayDates.map(d => d.toDateString()).join(", "));

  const raisedTime = sheet.getRange("B1").getValue();
  const startTime = sheet.getRange("B2").getValue();
  const endTime = sheet.getRange("B3").getValue();
  let workDaysStr = sheet.getRange("B4").getValue().toString() || "0000000";
  const slaRange = sheet.getRange("B5:B9").getValues();

  Logger.log("Raised Time: " + raisedTime);
//  Logger.log("Start Time: " + startTime);
  // Logger.log("End Time: " + endTime);
  Logger.log("Work Days String: " + workDaysStr);
  Logger.log("SLA Range (hours): " + JSON.stringify(slaRange));

  if (!(raisedTime instanceof Date)) {
    Logger.log("Invalid Raised Time");
    SpreadsheetApp.getUi().alert("Invalid Raised Time");

    return;
  }

  const slaHours = slaRange.map(row => (typeof row[0] === "number" ? row[0] : 0));
  Logger.log("Parsed SLA Hours: " + slaHours.join(", "));

  const timeToHour = t => {
    if (typeof t === "string") {
      let [h, m] = t.split(":").map(Number);
      return h + (m || 0) / 60;
    }
    if (t instanceof Date) return t.getHours() + t.getMinutes() / 60;
    return 0;
  };

  const workStartHour = timeToHour(startTime);
  const workEndHour = timeToHour(endTime);
  const workDayDuration = workEndHour - workStartHour;

  Logger.log("Work Start Hour: " + workStartHour);
  Logger.log("Work End Hour: " + workEndHour);
  Logger.log("Work Day Duration (hours): " + workDayDuration);

  workDaysStr = workDaysStr.split('').map(ch => (ch === '1' ? '1' : '0')).join('');
  const workDays = [];
  for (let i = 0; i < workDaysStr.length; i++) {
    if (workDaysStr[i] === "1") workDays.push(i + 1);
  }
  Logger.log("Working Days (numeric): " + workDays.join(", "));

  function isWorkingDay(date) {
    let day = date.getDay();
    day = day === 0 ? 7 : day;
    if (!workDays.includes(day)) return false;
    const checkDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    return !holidayDates.some(holiday => holiday.getTime() === checkDate.getTime());
  }

  function nextWorkingDayStart(date) {
    let d = new Date(date);
    d.setDate(d.getDate() + 1);
    while (!isWorkingDay(d)) {
      d.setDate(d.getDate() + 1);
    }
    d.setHours(Math.floor(workStartHour), (workStartHour % 1) * 60, 0, 0);
    return d;
  }

  function adjustToWorkingHours(date) {
    let d = new Date(date);
    if (!isWorkingDay(d)) return nextWorkingDayStart(d);

    const dayStart = new Date(d);
    dayStart.setHours(Math.floor(workStartHour), (workStartHour % 1) * 60, 0, 0);

    const dayEnd = new Date(d);
    dayEnd.setHours(Math.floor(workEndHour), (workEndHour % 1) * 60, 0, 0);

    if (d < dayStart) return dayStart;
    if (d >= dayEnd) return nextWorkingDayStart(d);
    return d;
  }

  function formatTimeOnly(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "h:mm a").replace(/^0/, "");
  }

  function formatDate(d) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM d, yyyy h:mm a");
  }

  // function formatDuration(ms) {
  //   let totalMinutes = Math.floor(ms / 60000);
  //   const days = Math.floor(totalMinutes / (60 * 24));
  //   totalMinutes -= days * 60 * 24;
  //   const hours = Math.floor(totalMinutes / 60);
  //   const minutes = totalMinutes % 60;
  //   Logger.log(`formatDuration: ms=${ms}, days=${days}, hours=${hours}, minutes=${minutes}`);
  //   return `${days}D ${hours}H ${minutes}M`;
  // }


  function padRow(row, length) {
    const newRow = [...row];
    while (newRow.length < length) newRow.push("");
    return newRow;
  }

// function addWorkingHoursWithLog(start, totalHours) {
//   let log = [];
//   let current = new Date(start);
//   let hoursRemaining = totalHours;
//   let dayCount = 1;
//   const loggedDates = new Set();

//   while (hoursRemaining > 0 || !loggedDates.has(current.toDateString())) {
//     let dayLabel = "Day " + dayCount;
//     let currentDateOnly = new Date(current.getFullYear(), current.getMonth(), current.getDate());
//     let dateKey = currentDateOnly.toDateString();

//     if (!loggedDates.has(dateKey)) {
//       loggedDates.add(dateKey);

//       const isHoliday = holidayDates.some(h => h.getTime() === currentDateOnly.getTime());
//       const isWeekend = !workDays.includes(current.getDay() === 0 ? 7 : current.getDay());

//       if (isHoliday || isWeekend) {
//         let reason = isHoliday ? "Holiday" : "Weekend";
//         log.push(padRow([
//           dayLabel,
//           formatDate(currentDateOnly) + ` (${reason})`,
//           "0.00 hours",
//           hoursRemaining.toFixed(2) + " hours left"
//         ], 4));

//         current = new Date(currentDateOnly.getTime() + 24 * 60 * 60 * 1000);
//         dayCount++;
//         continue;
//       }
      
//     }

//     const workStart = new Date(currentDateOnly);
//     workStart.setHours(Math.floor(workStartHour), (workStartHour % 1) * 60, 0, 0);
//     const workEnd = new Date(currentDateOnly);
//     workEnd.setHours(Math.floor(workEndHour), (workEndHour % 1) * 60, 0, 0);

//     if (current < workStart) current = new Date(workStart);
//     if (current >= workEnd) {
//       current = new Date(currentDateOnly.getTime() + 24 * 60 * 60 * 1000);
//       dayCount++;
//       continue;
//     }

//     let availableHours = (workEnd - current) / (1000 * 60 * 60);
//     let hoursUsed = Math.min(availableHours, hoursRemaining);

//     log.push(padRow([
//       dayLabel,
//       formatDate(current),
//       hoursUsed.toFixed(2) + " hours",
//       (hoursRemaining - hoursUsed).toFixed(2) + " hours left"
//     ], 4));

//     current = new Date(current.getTime() + hoursUsed * 60 * 60 * 1000);
//     hoursRemaining -= hoursUsed;
//     dayCount++;
//   }

//   return { end: current, log };
// }


function addWorkingHoursWithLog(start, totalHours) {
  let log = [];
  let current = new Date(start);
  let hoursRemaining = totalHours;
  let dayCount = 1;
  const loggedDates = new Set();

  while (hoursRemaining > 0 || !loggedDates.has(current.toDateString())) {
    let currentDateOnly = new Date(current.getFullYear(), current.getMonth(), current.getDate());
    let dateKey = currentDateOnly.toDateString();

    const isHoliday = holidayDates.some(h => h.getTime() === currentDateOnly.getTime());
    const isWeekend = !workDays.includes(current.getDay() === 0 ? 7 : current.getDay());

    if (isHoliday || isWeekend) {
      if (!loggedDates.has(dateKey)) {
        log.push(padRow([
          "Day " + dayCount,
          formatDate(currentDateOnly) + ` (${isHoliday ? "Holiday" : "Weekend"})`,
          "0.00 hours",
          hoursRemaining.toFixed(2) + " hours left"
        ], 4));
        loggedDates.add(dateKey);
        dayCount++;  // increment only once per new date
      }

      current = new Date(currentDateOnly.getTime() + 24 * 60 * 60 * 1000);
      continue;
    }

    const workStart = new Date(currentDateOnly);
    workStart.setHours(Math.floor(workStartHour), (workStartHour % 1) * 60, 0, 0);
    const workEnd = new Date(currentDateOnly);
    workEnd.setHours(Math.floor(workEndHour), (workEndHour % 1) * 60, 0, 0);

    if (current < workStart) current = new Date(workStart);
    if (current >= workEnd) {
      current = new Date(currentDateOnly.getTime() + 24 * 60 * 60 * 1000);
      continue;
    }

    let availableHours = (workEnd - current) / (1000 * 60 * 60);
    let hoursUsed = Math.min(availableHours, hoursRemaining);

    if (!loggedDates.has(dateKey)) {
      log.push(padRow([
        "Day " + dayCount,
        formatDate(current),
        hoursUsed.toFixed(2) + " hours",
        (hoursRemaining - hoursUsed).toFixed(2) + " hours left"
      ], 4));
      loggedDates.add(dateKey);
      dayCount++;
    } else {
      log.push(padRow([
        "",
        formatDate(current),
        hoursUsed.toFixed(2) + " hours",
        (hoursRemaining - hoursUsed).toFixed(2) + " hours left"
      ], 4));
    }

    current = new Date(current.getTime() + hoursUsed * 60 * 60 * 1000);
    hoursRemaining -= hoursUsed;
  }

  return { end: current, log };
}




  const actualStart = raisedTime;

  // Clear previous content
  sheet.getRange("A13:H90").clearContent();

  // Prepare breakdown rows
  const breakdownRows = [
    ["Working Hours Start:", formatTimeOnly(new Date(1970, 0, 1, Math.floor(workStartHour), (workStartHour % 1) * 60))],
    ["Working Hours End:", formatTimeOnly(new Date(1970, 0, 1, Math.floor(workEndHour), (workEndHour % 1) * 60))],
    ["Total Working Hours:", `${workDayDuration.toFixed(2)} hours`],
    ["Working Days Mask (Mon=1):", workDaysStr],
    ["Working Days Active:", workDays.join(", ")],
    ["Holiday Dates:", holidayDates.map(d => Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM d, yyyy")).join(", ")],
    ["Raised On:", formatDate(raisedTime)],
    ["Adjusted Start:", formatDate(actualStart)]
  ];

  // Write breakdown rows
  sheet.getRange(13, 2, breakdownRows.length, 2).setValues(breakdownRows);
  sheet.getRange(13, 2, breakdownRows.length, 1).setFontWeight("bold").setFontColor("#0B5394");

  // Headers for escalation
  sheet.getRange("B22:F22").setValues([["L1 Escalation", "L2", "L3", "L4", "L5"]]);
  sheet.getRange("B22:F22").setFontWeight("bold").setFontColor("white").setBackground("#134f5c");

  let escalationTimes = [];
  let durations = [];
  let logRows = [];
  let currentStart = actualStart;

  for (let i = 0; i < slaHours.length; i++) {
    const hours = slaHours[i];
    if (hours === 0) {
      escalationTimes.push("");
      durations.push("");
      continue;
    }
    const result = addWorkingHoursWithLog(currentStart, hours);
    escalationTimes.push(formatDate(result.end));
  //const now = new Date();
  const now = new Date();

//    durations.push(formatDuration(result.end - raisedTime));
  durations.push(formatDuration(result.end.getTime() - now.getTime()));

  
    logRows.push(padRow(["---", `L${i + 1} Escalation (+${hours}h)`, "---"], 4));
    logRows = logRows.concat(result.log);
    currentStart = result.end;
  }

  // Write escalation result
  sheet.getRange("B23:F23").setValues([escalationTimes]).setNumberFormat("@STRING@");
  sheet.getRange("B24:F24").setValues([durations]);

  // Write log
  sheet.getRange(25, 2, logRows.length, 4).setValues(logRows);
SpreadsheetApp.getActiveSpreadsheet().toast("SLA Calculation Completed!", "Done", 5);

}
catch (error)
{
    Logger.log(`An error occurred: ${error.message}\nStack: ${error.stack}`);

    // Display a user-friendly alert message in the spreadsheet
    SpreadsheetApp.getUi().alert(`An error occurred during the SLA calculation. Please check the script logs for details. Error: ${error.message}`);
 
  handleError(error);
}

}


function formatDuration(ms) {
  Logger.log(`formatDurationAbsolute called with ms=${ms}`);

  let totalMinutes = Math.floor(ms / 60000);
  Logger.log(`Total minutes calculated: ${totalMinutes}`);

  const days = Math.floor(totalMinutes / (24 * 60));
  totalMinutes -= days * 24 * 60;

  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;

  const result = `${days}D ${hours}H ${minutes}M`;
  Logger.log(`Returning formatted duration: ${result}`);

  return result;
}



function testFormatDuration() {
  const now = new Date("June6, 2025 12:32 PM");
  const sla = new Date("June 16, 2025 1:44 PM");

  const msDifference = sla.getTime() - now.getTime();
  const formattedDuration = formatDuration(msDifference);
  Logger.log("Duration to SLA: " + formattedDuration);
}

