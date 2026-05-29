# Corporate-Debt-Obligation-Tracker
Automate Sheet to maintain the debt obligation of a company
### 📊 Production Project: High-Performance Commercial Loan & Amortization Engine

**Tech Stack:** JavaScript (ES6+), Google Apps Script V8 Engine, Google Sheets API
**Domain:** Commercial Banking & Wealth Management Operations

#### 🚀 Overview
An optimized, event-driven Google Apps Script `onEdit` automation built to process multi-row commercial banking data fields in real-time. The engine isolates modifications, performs fast batch memory-reads (`getValues`), calculates complex dynamic debt schedules, and batch-writes the financial values back in a single IO call to eliminate structural performance overhead.

#### 🛠️ Key Architectural Features
* **True Outstanding Balance Matrix:** Implements an internal iterative loop tracking interest-vs-principal degradation over variable month lengths.
* **Dynamic Calendar-Edge Mitigation:** Handles leap years and irregular calendar months (`28/29/30/31` variations) using JavaScript native `Date` constraints to calculate flawless repayment timelines.
* **Performance Optimization:** Eradicates the common `API-call-inside-loop` anti-pattern by caching spreadsheet ranges to memory objects before calculation loops.

---

#### 💻 Source Code

```javascript
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;

  const startRow = range.getRow();
  const endRow   = range.getLastRow();
  const startCol = range.getColumn();
  const endCol   = range.getLastColumn();

  // Monitored columns remain C to F (Columns 3 to 6)
  if (startRow > 26 || endRow < 2 || startCol > 6 || endCol < 3) return;

  const processStartRow  = Math.max(2, startRow);
  const processEndRow    = Math.min(26, endRow);
  const numRowsToProcess = processEndRow - processStartRow + 1;

  const dataValues = sheet.getRange(processStartRow, 3, numRowsToProcess, 4).getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Helper function to calculate true outstanding balance
  function calcOutstanding(principal, monthlyRate, emiExact, tenure, startDate) {
    let balance = principal;
    const startDay = startDate.getDate();
    
    for (let m = 1; m <= tenure; m++) {
      let targetMonth = startDate.getMonth() + m;
      let targetYear = startDate.getFullYear() + Math.floor(targetMonth / 12);
      targetMonth = targetMonth % 12;
      
      const maxDays = new Date(targetYear, targetMonth + 1, 0).getDate();
      const currentDueDay = Math.min(startDay, maxDays);
      const dueDate = new Date(targetYear, targetMonth, currentDueDay);
      
      const interest      = balance * monthlyRate;
      const principalPaid = emiExact - interest;
      balance             = Math.max(0, balance - principalPaid);
      
      if (dueDate >= today) {
        break;
      }
    }
    return Math.round(balance * 100) / 100;
  }

  const outputValues = [];

  for (let i = 0; i < dataValues.length; i++) {
    const row = dataValues[i];
    const C = Number(row[0]); // Principal
    const D = Number(row[1]); // Interest Rate
    const E = Number(row[2]); // Tenure (Months)
    let rawDate = row[3];     // Start Date

    if (typeof rawDate === 'string' && rawDate.trim() !== "") {
      rawDate = new Date(rawDate);
    }

    // If data is invalid, clear columns G through L
    if (isNaN(C) || !C || isNaN(D) || !D || isNaN(E) || !E || !(rawDate instanceof Date) || isNaN(rawDate.getTime())) {
      outputValues.push(["", "", "", "", "", ""]);
      continue;
    }

    const annualRate = D > 1 ? D / 100 : D;
    const r          = annualRate / 12;

    const emiExact = r === 0
      ? C / E
      : (C * r * Math.pow(1 + r, E)) / (Math.pow(1 + r, E) - 1);

    const startDate = new Date(rawDate.getTime());
    startDate.setHours(0, 0, 0, 0);
    const startDay = startDate.getDate();

    // 1. Column G: Today's Date
    const colG = today;

    // 2. Column H: Draft End Date
    let targetMonth = startDate.getMonth() + E;
    let targetYear = startDate.getFullYear() + Math.floor(targetMonth / 12);
    targetMonth = targetMonth % 12;
    const maxDaysInTargetMonth = new Date(targetYear, targetMonth + 1, 0).getDate();
    let finalDay = Math.min(startDay, maxDaysInTargetMonth);
    const endDate = new Date(targetYear, targetMonth, finalDay);
    const colH = endDate;

    // 3. Column I: Actual End Date (Draft End Date minus 1 Month)
    let actualMonth = startDate.getMonth() + (E - 1);
    let actualYear = startDate.getFullYear() + Math.floor(actualMonth / 12);
    actualMonth = actualMonth % 12;
    const maxDaysInActualMonth = new Date(actualYear, actualMonth + 1, 0).getDate();
    let actualFinalDay = Math.min(startDay, maxDaysInActualMonth);
    const colI = new Date(actualYear, actualMonth, actualFinalDay);

    // 4. Column J: Remaining Repayment Tenure
    let nRemaining = 0;
    for (let m = 1; m <= (E - 1); m++) {
      let mMonth = startDate.getMonth() + m;
      let mYear = startDate.getFullYear() + Math.floor(mMonth / 12);
      mMonth = mMonth % 12;
      
      const maxDays = new Date(mYear, mMonth + 1, 0).getDate();
      const currentDueDay = Math.min(startDay, maxDays);
      
      const dueDate = new Date(mYear, mMonth, currentDueDay);
      if (dueDate >= today) nRemaining++;
    }
    const colJ = nRemaining;

    // 5. Column K: Calculated Monthly EMI
    const colK = Math.round(emiExact * 100) / 100;

    // 6. Column L: Principal Outstanding
    const colL = calcOutstanding(C, r, emiExact, E, startDate);

    // Push row data for G, H, I, J, K, L
    outputValues.push([colG, colH, colI, colJ, colK, colL]);
  }

  if (outputValues.length > 0) {
    const targetRange = sheet.getRange(processStartRow, 7, numRowsToProcess, 6);
    targetRange.setValues(outputValues);

    // Auto-formatting for date outputs
    sheet.getRange(processStartRow, 7, numRowsToProcess, 3).setNumberFormat("mm/dd/yyyy"); 
  }
}
