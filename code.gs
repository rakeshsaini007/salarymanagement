function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Salary Management App")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 🔹 Ensure sheets exist with correct headers
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Employees Sheet
  let emp = ss.getSheetByName("Employees");
  if (!emp) {
    emp = ss.insertSheet("Employees");
    emp.appendRow([
      "Employee ID", "Employee Name", "Category", 
      "Adhar Number", "Mobile Number"
    ]);
  }

  // Journal Sheet
  let journal = ss.getSheetByName("Journal");
  if (!journal) {
    journal = ss.insertSheet("Journal");
    journal.appendRow([
      "Employee ID", "Employee Name", "Category",
      "Employee Adhar Number", "Mobile Number",
      "Date of Payment", "Transaction Money", "Transaction Mode", "Transaction Note Details", "Month-Year"
    ]);
  } else {
    // Check if we need to migrate old data structure
    const headers = journal.getRange(1, 1, 1, journal.getLastColumn()).getValues()[0];
    if (headers[7] === "Transaction Note" && headers.length === 9) {
      // Old structure detected, add new column
      journal.insertColumnAfter(7);
      journal.getRange(1, 8).setValue("Transaction Mode");
      journal.getRange(1, 9).setValue("Transaction Note Details");

      // Migrate existing data: copy old "Transaction Note" to "Transaction Mode"
      const lastRow = journal.getLastRow();
      if (lastRow > 1) {
        const oldNotes = journal.getRange(2, 8, lastRow - 1, 1).getValues();
        journal.getRange(2, 8, lastRow - 1, 1).setValues(oldNotes);
        for (let i = 2; i <= lastRow; i++) {
          journal.getRange(i, 9).setValue("");
        }
      }
    }
  }

  // Account Sheet
  let account = ss.getSheetByName("Account");
  if (!account) {
    account = ss.insertSheet("Account");
    account.appendRow([
      "Employee ID", "Employee Name", "Category",
      "Adhar Number", "Mobile Number", "Month-Year",
      "Total Duty", "Total OT", "Previous Balance", "Salary/month",
      "Gross Salary", "Advance", "Salary Paid", "Balance to be Paid"
    ]);
  }
}

// 🔹 Get list of Employee IDs
function getEmployeeIds() {
  try {
    Logger.log("getEmployeeIds called");
    ensureSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Employees");

    if (!sheet) {
      Logger.log("Employees sheet not found");
      return [];
    }

    const data = sheet.getDataRange().getValues();

    if (!data || data.length === 0) {
      Logger.log("No data in Employees sheet");
      return [];
    }

    const headers = data[0];
    const empIndex = headers.indexOf("Employee ID");

    if (empIndex === -1) {
      Logger.log("Employee ID column not found");
      return [];
    }

    let ids = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][empIndex]) {
        ids.push(data[i][empIndex]);
      }
    }

    Logger.log("Returning " + ids.length + " employee IDs");
    return ids;
  } catch (error) {
    Logger.log("Error in getEmployeeIds: " + error.toString());
    throw new Error("Failed to load employee IDs: " + error.message);
  }
}

// 🔹 Fetch details of employee by ID
function getEmployeeDetails(empId) {
  try {
    Logger.log("getEmployeeDetails called with empId: " + empId);
    ensureSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Employees");

    if (!sheet) {
      Logger.log("Employees sheet not found");
      return null;
    }

    const data = sheet.getDataRange().getValues();

    if (!data || data.length === 0) {
      Logger.log("No data in Employees sheet");
      return null;
    }

    const headers = data[0];

    const empIndex = headers.indexOf("Employee ID");
    const nameIndex = headers.indexOf("Employee Name");
    const catIndex = headers.indexOf("Category");
    const adharIndex = headers.indexOf("Adhar Number");
    const mobileIndex = headers.indexOf("Mobile Number");

    if (empIndex === -1) {
      Logger.log("Employee ID column not found");
      return null;
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][empIndex]) === String(empId)) {
        Logger.log("Found employee at row " + i);
        return {
          employeeId: data[i][empIndex],
          employeeName: data[i][nameIndex],
          category: data[i][catIndex],
          adhar: data[i][adharIndex],
          mobile: data[i][mobileIndex]
        };
      }
    }

    Logger.log("Employee not found: " + empId);
    return null;
  } catch (error) {
    Logger.log("Error in getEmployeeDetails: " + error.toString());
    throw new Error("Failed to fetch employee details: " + error.message);
  }
}

// 🔹 Save Journal Entry
function submitJournal(data) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Journal");
  sheet.appendRow([
    data.employeeId,
    data.employeeName,
    data.category,
    data.employeeAdhar,
    data.mobile,
    data.datePayment,
    data.transactionMoney,
    data.transactionMode,
    data.transactionNote || "",
    data.monthYear
  ]);
  return "✅ Journal entry saved successfully!";
}

// 🔹 Update Journal Entry
function updateJournal(data, rowIndex) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Journal");

  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const dateIndex = headers.indexOf("Date of Payment");
  const empIdIndex = headers.indexOf("Employee ID");

  let targetRow = -1;
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    let rowDate = row[dateIndex];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (rowDate) {
      rowDate = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    if (String(row[empIdIndex]) === String(data.employeeId) && rowDate === data.datePayment) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error("Record not found for update");
  }

  sheet.getRange(targetRow, 1, 1, 10).setValues([[
    data.employeeId,
    data.employeeName,
    data.category,
    data.employeeAdhar,
    data.mobile,
    data.datePayment,
    data.transactionMoney,
    data.transactionMode,
    data.transactionNote || "",
    data.monthYear
  ]]);

  return "✅ Journal entry updated successfully!";
}

// 🔹 Save or Update Account Entry
function submitAccount(data) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Account");

  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const empIdIndex = headers.indexOf("Employee ID");
  const monthYearIndex = headers.indexOf("Month-Year");

  let existingRowIndex = -1;

  for (let i = 1; i < allData.length; i++) {
    const rowEmpId = String(allData[i][empIdIndex]);
    const rowMonthYear = allData[i][monthYearIndex];

    let formattedMonthYear = "";
    if (rowMonthYear instanceof Date) {
      formattedMonthYear = Utilities.formatDate(rowMonthYear, Session.getScriptTimeZone(), "MMM-yyyy");
    } else if (rowMonthYear) {
      const parsed = new Date(rowMonthYear);
      if (!isNaN(parsed)) {
        formattedMonthYear = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM-yyyy");
      } else {
        formattedMonthYear = String(rowMonthYear);
      }
    }

    if (rowEmpId === String(data.employeeId) && formattedMonthYear === data.monthYear) {
      existingRowIndex = i + 1;
      break;
    }
  }

  const rowData = [
    data.employeeId,
    data.employeeName,
    data.category,
    data.adhar,
    data.mobile,
    data.monthYear,
    data.totalDuty,
    data.totalOT,
    data.prevBalance,
    data.salaryMonth,
    data.grossSalary,
    data.advance,
    data.salaryPaid,
    data.balanceToBePaid
  ];

  if (existingRowIndex !== -1) {
    sheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  recalculateSubsequentMonths(data.employeeId, data.monthYear);

  return existingRowIndex !== -1
    ? "✅ Account entry updated successfully!"
    : "✅ Account entry saved successfully!";
}

// 🔹 Recalculate all subsequent months after an update
function recalculateSubsequentMonths(empId, startMonthYear) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountSheet = ss.getSheetByName("Account");

  const allData = accountSheet.getDataRange().getValues();
  const headers = allData[0];
  const empIdIndex = headers.indexOf("Employee ID");
  const monthYearIndex = headers.indexOf("Month-Year");
  const prevBalanceIndex = headers.indexOf("Previous Balance");
  const totalDutyIndex = headers.indexOf("Total Duty");
  const totalOTIndex = headers.indexOf("Total OT");
  const salaryMonthIndex = headers.indexOf("Salary/month");
  const grossSalaryIndex = headers.indexOf("Gross Salary");
  const advanceIndex = headers.indexOf("Advance");
  const salaryPaidIndex = headers.indexOf("Salary Paid");
  const balanceToBePaidIndex = headers.indexOf("Balance to be Paid");

  const employeeRecords = [];

  // --- Collect employee rows ---
  for (let i = 1; i < allData.length; i++) {
    const rowEmpId = String(allData[i][empIdIndex]);
    const rowMonthYear = allData[i][monthYearIndex];
    if (rowEmpId !== String(empId)) continue;

    let formattedMonthYear = "";
    if (rowMonthYear instanceof Date) {
      formattedMonthYear = Utilities.formatDate(rowMonthYear, Session.getScriptTimeZone(), "MMM-yyyy");
    } else if (rowMonthYear) {
      const parsed = new Date(rowMonthYear);
      formattedMonthYear = !isNaN(parsed)
        ? Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM-yyyy")
        : String(rowMonthYear);
    }

    employeeRecords.push({
      rowIndex: i + 1,
      monthYear: formattedMonthYear,
      monthDate: new Date(formattedMonthYear),
      data: allData[i]
    });
  }

  // --- Sort by chronological order ---
  employeeRecords.sort((a, b) => a.monthDate - b.monthDate);

  const startDate = new Date(startMonthYear);
  const startIndex = employeeRecords.findIndex(rec => rec.monthDate >= startDate);
  if (startIndex === -1) return;

  // --- Main recalculation loop ---
  for (let i = startIndex; i < employeeRecords.length; i++) {
    const record = employeeRecords[i];
    const rowIndex = record.rowIndex;

    // Previous Balance
    let prevBalance = 0;
    if (i > 0) {
      const previousRecord = employeeRecords[i - 1];
      prevBalance = parseFloat(previousRecord.data[balanceToBePaidIndex]) || 0;
    }

    // Get Advance & Salary Paid from Journal
    const advancedData = getAdvancedData(empId, record.monthYear);
    const advance = Math.round(advancedData.advance);
    const salaryPaid = Math.round(advancedData.salaryPaid);

    // Core salary calculations
    const totalDuty = parseFloat(record.data[totalDutyIndex]) || 0;
    const totalOT = parseFloat(record.data[totalOTIndex]) || 0;
    const salaryMonth = parseFloat(record.data[salaryMonthIndex]) || 0;

    const grossSalary = Math.round((salaryMonth * totalDuty / 26) + (salaryMonth * totalOT / 208));
    const balanceToBePaid = Math.round(prevBalance + grossSalary - advance - salaryPaid);
    const prevBalanceRounded = Math.round(prevBalance);

    // --- Write back updated values ---
    accountSheet.getRange(rowIndex, prevBalanceIndex + 1).setValue(prevBalanceRounded);
    accountSheet.getRange(rowIndex, grossSalaryIndex + 1).setValue(grossSalary);
    accountSheet.getRange(rowIndex, advanceIndex + 1).setValue(advance);
    accountSheet.getRange(rowIndex, salaryPaidIndex + 1).setValue(salaryPaid);
    accountSheet.getRange(rowIndex, balanceToBePaidIndex + 1).setValue(balanceToBePaid);

    // Update in-memory data for chain calculation
    record.data[balanceToBePaidIndex] = balanceToBePaid;

    Logger.log(`🔄 Recalculated ${record.monthYear} for ${empId}: Gross=${grossSalary}, Advance=${advance}, Paid=${salaryPaid}, Balance=${balanceToBePaid}`);
  }
}

// 🔹 Update next month's Previous Balance with current month's Balance to be Paid (with cascading updates)
function updateNextMonthPreviousBalance(empId, currentMonthYear, balanceToBePaid) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Account");

  const [currentMonth, currentYear] = currentMonthYear.split('-');
  const currentDate = new Date(currentMonth + ' 1, ' + currentYear);
  const nextMonth = new Date(currentDate);
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  const nextMonthFormatted = Utilities.formatDate(nextMonth, Session.getScriptTimeZone(), "MMM-yyyy");

  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const empIdIndex = headers.indexOf("Employee ID");
  const monthYearIndex = headers.indexOf("Month-Year");
  const prevBalanceIndex = headers.indexOf("Previous Balance");
  const grossSalaryIndex = headers.indexOf("Gross Salary");
  const advanceIndex = headers.indexOf("Advance");
  const salaryPaidIndex = headers.indexOf("Salary Paid");
  const balanceToBePaidIndex = headers.indexOf("Balance to be Paid");

  for (let i = 1; i < allData.length; i++) {
    const rowEmpId = String(allData[i][empIdIndex]);
    const rowMonthYear = allData[i][monthYearIndex];

    let formattedMonthYear = "";
    if (rowMonthYear instanceof Date) {
      formattedMonthYear = Utilities.formatDate(rowMonthYear, Session.getScriptTimeZone(), "MMM-yyyy");
    } else if (rowMonthYear) {
      const parsed = new Date(rowMonthYear);
      if (!isNaN(parsed)) {
        formattedMonthYear = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM-yyyy");
      } else {
        formattedMonthYear = String(rowMonthYear);
      }
    }

    if (rowEmpId === String(empId) && formattedMonthYear === nextMonthFormatted) {
      sheet.getRange(i + 1, prevBalanceIndex + 1).setValue(balanceToBePaid);

      const grossSalary = parseFloat(allData[i][grossSalaryIndex]) || 0;
      const advance = parseFloat(allData[i][advanceIndex]) || 0;
      const salaryPaid = parseFloat(allData[i][salaryPaidIndex]) || 0;

      const newBalanceToBePaid = balanceToBePaid + grossSalary - advance - salaryPaid;
      sheet.getRange(i + 1, balanceToBePaidIndex + 1).setValue(newBalanceToBePaid);

      Logger.log("Updated next month (" + nextMonthFormatted + ") for employee " + empId + ": Previous Balance = " + balanceToBePaid + ", Balance to be Paid = " + newBalanceToBePaid);

      updateNextMonthPreviousBalance(empId, nextMonthFormatted, newBalanceToBePaid);
      break;
    }
  }
}

// 🔹 Get unique Month-Year values from Journal sheet
// 🔹 Get unique Month-Year values from Journal sheet in "MMM-YYYY" format
function getMonthYearOptions() {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Journal");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const monthIndex = headers.indexOf("Month-Year");

  let options = [];
  for (let i = 1; i < data.length; i++) {
    const cellValue = data[i][monthIndex];
    if (cellValue) {
      let formatted = "";
      if (cellValue instanceof Date) {
        formatted = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "MMM-yyyy");
      } else {
        const parsed = new Date(cellValue);
        if (!isNaN(parsed)) {
          formatted = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM-yyyy");
        } else {
          formatted = String(cellValue);
        }
      }
      options.push(formatted);
    }
  }

  options = [...new Set(options)];
  options.sort((a, b) => new Date(a) - new Date(b));
  console.log(options);
  return options;
}

function getAttendanceData(empId, monthYear) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let attendanceSheet = ss.getSheetByName("Attendance");

  if (!attendanceSheet) {
    attendanceSheet = ss.insertSheet("Attendance");
    attendanceSheet.appendRow([
      "Employee ID", "Employee Name", "Date", "Duty", "O.T.", "Month-Year"
    ]);
  }

  const data = attendanceSheet.getDataRange().getValues();
  const headers = data[0];
  const empIdIndex = headers.indexOf("Employee ID");
  const monthYearIndex = headers.indexOf("Month-Year");
  const dutyIndex = headers.indexOf("Duty");
  const otIndex = headers.indexOf("O.T.");

  let totalDuty = 0;
  let totalOT = 0;

  for (let i = 1; i < data.length; i++) {
    const rowEmpId = String(data[i][empIdIndex]);
    const rowMonthYear = data[i][monthYearIndex];

    let formattedMonthYear = "";
    if (rowMonthYear instanceof Date) {
      formattedMonthYear = Utilities.formatDate(rowMonthYear, Session.getScriptTimeZone(), "MMM-yyyy");
    } else if (rowMonthYear) {
      const parsed = new Date(rowMonthYear);
      if (!isNaN(parsed)) {
        formattedMonthYear = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM-yyyy");
      } else {
        formattedMonthYear = String(rowMonthYear);
      }
    }

    if (rowEmpId === String(empId) && formattedMonthYear === monthYear) {
      const duty = parseFloat(data[i][dutyIndex]) || 0;
      const ot = parseFloat(data[i][otIndex]) || 0;
      totalDuty += duty;
      totalOT += ot;
    }
  }

  return {
    totalDuty: totalDuty,
    totalOT: totalOT
  };
}

function getAdvancedData(empId, monthYear) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const journalSheet = ss.getSheetByName("Journal");

  if (!journalSheet) {
    return {
      advance: 0,
      salaryPaid: 0,
      prevBalance: 0
    };
  }

  const data = journalSheet.getDataRange().getValues();
  const headers = data[0];
  const empIdIndex = headers.indexOf("Employee ID");
  const datePaymentIndex = headers.indexOf("Date of Payment");
  const transactionMoneyIndex = headers.indexOf("Transaction Money");
  const monthYearIndex = headers.indexOf("Month-Year");

  let advance = 0;
  let salaryPaid = 0;

  const [targetMonth, targetYear] = monthYear.split('-');
  const targetDate = new Date(targetMonth + ' 1, ' + targetYear);
  const nextMonth = new Date(targetDate);
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  const nextMonthFormatted = Utilities.formatDate(nextMonth, Session.getScriptTimeZone(), "MMM-yyyy");

  for (let i = 1; i < data.length; i++) {
    const rowEmpId = String(data[i][empIdIndex]);
    const rowMonthYear = data[i][monthYearIndex];
    const datePayment = data[i][datePaymentIndex];
    const transactionMoney = parseFloat(data[i][transactionMoneyIndex]) || 0;

    let formattedMonthYear = "";
    if (rowMonthYear instanceof Date) {
      formattedMonthYear = Utilities.formatDate(rowMonthYear, Session.getScriptTimeZone(), "MMM-yyyy");
    } else if (rowMonthYear) {
      const parsed = new Date(rowMonthYear);
      if (!isNaN(parsed)) {
        formattedMonthYear = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM-yyyy");
      } else {
        formattedMonthYear = String(rowMonthYear);
      }
    }

    if (rowEmpId === String(empId)) {
      let paymentDate;
      if (datePayment instanceof Date) {
        paymentDate = datePayment;
      } else if (datePayment) {
        paymentDate = new Date(datePayment);
      }

      if (formattedMonthYear === monthYear) {
        if (paymentDate && paymentDate.getDate() !== 15) {
          advance += transactionMoney;
        }
      }

      if (formattedMonthYear === nextMonthFormatted && paymentDate && paymentDate.getDate() === 15) {
        salaryPaid += transactionMoney;
      }
    }
  }

  const prevBalance = getPreviousBalance(empId, monthYear);

  return {
    advance: advance,
    salaryPaid: salaryPaid,
    prevBalance: prevBalance
  };
}

function getPreviousBalance(empId, monthYear) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountSheet = ss.getSheetByName("Account");

  if (!accountSheet) {
    return 0;
  }

  const [targetMonth, targetYear] = monthYear.split('-');
  const targetDate = new Date(targetMonth + ' 1, ' + targetYear);
  const previousMonth = new Date(targetDate);
  previousMonth.setMonth(previousMonth.getMonth() - 1);
  const previousMonthFormatted = Utilities.formatDate(previousMonth, Session.getScriptTimeZone(), "MMM-yyyy");

  const data = accountSheet.getDataRange().getValues();
  const headers = data[0];
  const empIdIndex = headers.indexOf("Employee ID");
  const monthYearIndex = headers.indexOf("Month-Year");
  const balanceToBePaidIndex = headers.indexOf("Balance to be Paid");

  for (let i = data.length - 1; i >= 1; i--) {
    const rowEmpId = String(data[i][empIdIndex]);
    const rowMonthYear = data[i][monthYearIndex];

    let formattedMonthYear = "";
    if (rowMonthYear instanceof Date) {
      formattedMonthYear = Utilities.formatDate(rowMonthYear, Session.getScriptTimeZone(), "MMM-yyyy");
    } else if (rowMonthYear) {
      const parsed = new Date(rowMonthYear);
      if (!isNaN(parsed)) {
        formattedMonthYear = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM-yyyy");
      } else {
        formattedMonthYear = String(rowMonthYear);
      }
    }

    if (rowEmpId === String(empId) && formattedMonthYear === previousMonthFormatted) {
      const balanceToBePaid = parseFloat(data[i][balanceToBePaidIndex]) || 0;
      return balanceToBePaid;
    }
  }

  return 0;
}


function getAccountHistory(empId, monthYearFilter, isRange, fromMonthYear, toMonthYear) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountSheet = ss.getSheetByName("Account");
  const history = [];

  if (!accountSheet) {
    return history;
  }

  const accountData = accountSheet.getDataRange().getValues();
  const accountHeaders = accountData[0];
  const empIndex = accountHeaders.indexOf("Employee ID");
  const monthYearIndex = accountHeaders.indexOf("Month-Year");

  for (let i = 1; i < accountData.length; i++) {
    if (String(accountData[i][empIndex]) === String(empId)) {
      let entry = {};
      accountHeaders.forEach((h, j) => {
        entry[h] = accountData[i][j];
      });

      const rowMonthYear = accountData[i][monthYearIndex];
      let formattedMonthYear = "";
      if (rowMonthYear instanceof Date) {
        formattedMonthYear = Utilities.formatDate(rowMonthYear, Session.getScriptTimeZone(), "MMM-yyyy");
      } else if (rowMonthYear) {
        const parsed = new Date(rowMonthYear);
        if (!isNaN(parsed)) {
          formattedMonthYear = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM-yyyy");
        } else {
          formattedMonthYear = String(rowMonthYear);
        }
      }
      entry["Month-Year"] = formattedMonthYear;

      if (isRange && fromMonthYear && toMonthYear) {
        const fromDate = new Date(fromMonthYear);
        const toDate = new Date(toMonthYear);
        const entryDate = new Date(formattedMonthYear);
        if (entryDate >= fromDate && entryDate <= toDate) {
          history.push(entry);
        }
      } else if (monthYearFilter) {
        if (formattedMonthYear === monthYearFilter) {
          history.push(entry);
        }
      } else {
        history.push(entry);
      }
    }
  }

  history.sort((a, b) => {
    const dateA = new Date(a["Month-Year"]);
    const dateB = new Date(b["Month-Year"]);
    return dateB - dateA;
  });

  return history;
}


function getAccountMonthYearOptions() {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Account");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const monthIndex = headers.indexOf("Month-Year");

  let options = [];
  for (let i = 1; i < data.length; i++) {
    const cellValue = data[i][monthIndex];
    if (cellValue) {
      let formatted = "";
      if (cellValue instanceof Date) {
        formatted = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "MMM-yyyy");
      } else {
        const parsed = new Date(cellValue);
        if (!isNaN(parsed)) {
          formatted = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MMM-yyyy");
        } else {
          formatted = String(cellValue);
        }
      }
      options.push(formatted);
    }
  }

  options = [...new Set(options)];
  options.sort((a, b) => new Date(a) - new Date(b));
  return options;
}

function getJournalHistory(empId, specificDate, fromDate, toDate) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const journalSheet = ss.getSheetByName("Journal");
  const history = [];

  if (!journalSheet) {
    Logger.log("Journal sheet not found");
    return history;
  }

  const data = journalSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("No data in Journal sheet");
    return history;
  }

  const headers = data[0];
  const empIdIndex = headers.indexOf("Employee ID");
  const datePaymentIndex = headers.indexOf("Date of Payment");
  const monthYearIndex = headers.indexOf("Month-Year");

  Logger.log("Date Payment Index: " + datePaymentIndex);
  if (empId) {
    Logger.log("Searching for Employee ID: " + empId);
  } else {
    Logger.log("Fetching all employees");
  }

  for (let i = 1; i < data.length; i++) {
    const matchesEmployee = !empId || String(data[i][empIdIndex]) === String(empId);

    if (matchesEmployee) {
      if (empId) {
        Logger.log("Found matching employee at row " + i);
      }

      let entry = {};
      headers.forEach((h, j) => {
        let value = data[i][j];

        if (h === "Month-Year" && value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "MMM-yyyy");
        }

        entry[h] = value;
      });

      if (!entry["Transaction Mode"] && entry["Transaction Note"]) {
        entry["Transaction Mode"] = entry["Transaction Note"];
        entry["Transaction Note Details"] = "";
      }

      const paymentDate = data[i][datePaymentIndex];
      let entryDate;

      if (paymentDate instanceof Date) {
        entryDate = paymentDate;
      } else if (paymentDate) {
        entryDate = new Date(paymentDate);
      }

      if (entryDate && !isNaN(entryDate)) {
        const formattedDate = Utilities.formatDate(entryDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        entry["Date of Payment"] = formattedDate;

        Logger.log("Date formatted: " + formattedDate);

        if (specificDate) {
          Logger.log("Filtering by specific date: " + specificDate);
          if (formattedDate === specificDate) {
            history.push(entry);
            Logger.log("Match found for specific date");
          }
        } else if (fromDate && toDate) {
          Logger.log("Filtering by date range: " + fromDate + " to " + toDate);
          if (formattedDate >= fromDate && formattedDate <= toDate) {
            history.push(entry);
            Logger.log("Match found in date range");
          }
        } else {
          Logger.log("No filter, adding record");
          history.push(entry);
        }
      } else {
        Logger.log("Invalid date for row " + i);
      }
    }
  }

  Logger.log("Total records found: " + history.length);

  history.sort((a, b) => {
    const dateA = new Date(a["Date of Payment"]);
    const dateB = new Date(b["Date of Payment"]);
    return dateB - dateA;
  });

  return history;
}
