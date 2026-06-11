var ID_GUESTS_LIST = "1Puw0OezY18OWFt8wtwzv5BFxcJw314Hfov5GZMUXCbk";
var ID_QID_VERIFIED_LIST = "1cmRFirWeg_tHFbZ9VS-E0Gz80SHHSsIU9lu5jV_GBKk";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Vinyasa Nilaya | Guest Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://ik.imagekit.io/h87o83ayxm/Icons/Icon03_png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Core Data Lifecycle Engine.
 * Fetches all valid year sheets and guarantees the active year's tracking sheet exists.
 */
function getAvailableYearsAndInitialize() {
  try {
    const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
    const currentYear = new Date().getFullYear().toString(); // e.g., "2026"

    // 1. DYNAMIC SYSTEM INITIALIZATION: Check/Create New Year Sheet on Jan 1st
    let targetSheet = ss.getSheetByName(currentYear);
    if (!targetSheet) {
      console.log(`>>> [SYSTEM] Jan 1st Lifecyle Trigger: Creating new tracking tab for [${currentYear}]`);

      // Attempt to find your most recent historical sheet to use as a structural template
      const allSheets = ss.getSheets();
      let templateSheet = allSheets[0]; // Fallback to first sheet

      // Try to find a sheet name that looks like a 4-digit number to copy layout from
      for (let s of allSheets) {
        if (/^\d{4}$/.test(s.getName())) {
          templateSheet = s;
          break;
        }
      }

      // Duplicate the template sheet structure to preserve column names, widths, and formatting models
      if (templateSheet) {
        let newSheet = templateSheet.copyTo(ss);
        newSheet.setName(currentYear);

        // Clean out old historical row cell values while maintaining structural formatting headers
        const lastRow = newSheet.getLastRow();
        if (lastRow > 1) {
          // Assuming row 1 has your "Month", "Name", "Amount" header values
          newSheet.getRange(2, 1, lastRow - 1, newSheet.getLastColumn()).clearContent();
        }
        // Move to front position for visibility
        ss.setActiveSheet(newSheet);
        ss.moveActiveSheet(1);
      }
    }

    // 2. RETRIEVAL SYSTEM: Gather all 4-digit numeric sheet names for the UI dropdown
    const sheets = ss.getSheets();
    let availableYears = [];

    sheets.forEach(sheet => {
      const name = sheet.getName().trim();
      if (/^\d{4}$/.test(name)) { // Matches exact 4-digit years like "2025", "2026"
        availableYears.push(name);
      }
    });

    // Sort years in descending order so the newest year is listed first (e.g., 2026, 2025)
    availableYears.sort((a, b) => Number(b) - Number(a));

    return {
      years: availableYears,
      activeYear: currentYear
    };

  } catch (err) {
    console.error("Initialization / Year discovery failure: " + err.message);
    return { years: [new Date().getFullYear().toString()], activeYear: new Date().getFullYear().toString() };
  }
}

/*
 * Fetches guest data and summaries based on the specific column headers:
 * Year, Month, Name, NoOfGuests, Amount, Check-in Date, Days, AirBnb\Personal, Floor, Mobile, Customer Ratings, Comments
 */
function getDashboardData(filterYear, filterMonth) {
  try {
    const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
    const sheets = ss.getSheets();

    // =========================================================================
    // REQUIRED SPEED CHANGE 1: PRE-FETCH ALL VERIFIED NUMBERS IN ONE ROUND-TRIP
    // =========================================================================
    let verifiedMobilesSet = new Set();
    try {
      const qidSpreadsheet = SpreadsheetApp.openById(ID_QID_VERIFIED_LIST);
      const verificationSheet = qidSpreadsheet.getSheetByName(SHEET_NAME_QID); 
      
      if (verificationSheet) {
        const verData = verificationSheet.getDataRange().getValues();
        const PHONE_COL_INDEX = 5; // Matches Column 6 (Phone / Whatsapp) from your registry
        
        for (let i = 1; i < verData.length; i++) {
          let cellVal = verData[i][PHONE_COL_INDEX];
          if (cellVal) {
            let cleanVerMobile = cellVal.toString().replace(/\D/g, "");
            if (cleanVerMobile.length > 10) cleanVerMobile = cleanVerMobile.slice(-10);
            if (cleanVerMobile.length === 10) {
              verifiedMobilesSet.add(cleanVerMobile); // Stored in ultrafast RAM cache
            }
          }
        }
      }
    } catch (qidErr) {
      console.error(">>> [SPEED ENGINE WARNING] Cache pre-fetch failed: " + qidErr.message);
    }
    // =========================================================================

    // -----------------------------------------------------------------
    // PHASE 1: COMPUTE CONSOLIDATED LIFETIME REVENUE ACROSS ALL YEARS
    // -----------------------------------------------------------------
    let lifetimeTotalRevenue = 0;
    let lifetimeTotalCheckIns = 0;

    sheets.forEach(sheet => {
      const sheetName = sheet.getName().trim();

      if (/^\d{4}$/.test(sheetName)) {
        const sheetData = sheet.getDataRange().getValues();
        const headerRowIdx = sheetData.findIndex(row => row.includes("Name") || row.includes("Amount"));

        if (headerRowIdx !== -1) {
          const sheetHeaders = sheetData[headerRowIdx];
          const amountIdx = sheetHeaders.indexOf("Amount");
          const nameIdx = sheetHeaders.indexOf("Name");

          if (amountIdx !== -1) {
            const sheetRows = sheetData.slice(headerRowIdx + 1);

            sheetRows.forEach(row => {
              const nameVal = row[nameIdx] ? row[nameIdx].toString().trim() : "";
              if (!nameVal || nameVal === "Total" || nameVal === "No Guests" || nameVal === "") return;

              let amtStr = (row[amountIdx] || "0").toString().replace(/[₹,]/g, "");
              let amtNum = Number(amtStr) || 0;
              lifetimeTotalRevenue += amtNum;
              lifetimeTotalCheckIns++; 
            });
          }
        }
      }
    });

    // -----------------------------------------------------------------
    // PHASE 2: PROCESSING SELECTED FOCUS TARGET SHEET DATA
    // -----------------------------------------------------------------
    const targetTab = filterYear;
    let sheet = ss.getSheetByName(targetTab) || ss.getSheets()[0];

    const data = sheet.getDataRange().getValues();
    const headerRowIndex = data.findIndex(row => row.includes("Name") || row.includes("Month"));

    if (headerRowIndex === -1) {
      return {
        guests: [],
        summary: {
          totalRevenue: "₹0",
          count: 0,
          period: "No Headers Found",
          lifetimeRevenue: lifetimeTotalRevenue.toLocaleString('en-IN')
        }
      };
    }

    const headers = data[headerRowIndex];
    const rows = data.slice(headerRowIndex + 1);

    let totalRevenue = 0;
    let guestCount = 0;

    const filteredData = rows.filter(row => {
      const name = row[headers.indexOf("Name")];
      const month = row[headers.indexOf("Month")];
      if (!name || name === "Total" || name === "No Guests" || name === "") return false;

      const rowMonth = month ? month.toString().trim() : "";
      return !filterMonth || rowMonth === filterMonth;
    }).map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        let key = header.toString();
        if (key === "AirBnb\\Personal" || key === "Source") {
          key = "Source";
        }
        else if (key === "Floor") {
          key = "Floor";
        }
        else {
          key = key.replace(/\\|\s|-/g, "_");
        }

        let value = row[i];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        obj[key] = value;
      });

      // 1. Revenue & Stats Calculation
      let amtStr = (obj.Amount || "0").toString().replace(/[₹,]/g, "");
      let amtNum = Number(amtStr) || 0;
      totalRevenue += amtNum;
      guestCount++;

      // 2. Mobile Sanitization for WhatsApp & Verification
      let mobileRaw = (obj.Mobile || "").toString().replace(/\D/g, "");
      if (mobileRaw.length > 10) {
        mobileRaw = mobileRaw.slice(-10); 
      }
      obj['WhatsApp_Num'] = mobileRaw;

      // =========================================================================
      // REQUIRED SPEED CHANGE 2: USE INSTANT MEMORY SET INSTEAD OF THE SLOW LOOP FUNCTION
      // =========================================================================
      obj['isVerified'] = (mobileRaw.length === 10 && verifiedMobilesSet.has(mobileRaw));
      // =========================================================================

      return obj;
    });

    return {
      guests: filteredData,
      summary: {
        totalRevenue: totalRevenue.toLocaleString('en-IN', { maximumFractionDigits: 0 }).replace("INR", "").trim(),
        count: guestCount,
        period: filterMonth ? `${filterMonth} ${targetTab}` : targetTab,
        lifetimeRevenue: lifetimeTotalRevenue.toLocaleString('en-IN', { maximumFractionDigits: 0 }),
        lifetimeCount: lifetimeTotalCheckIns 
      }
    };
  } catch (err) {
    console.error("Dashboard Sync Error: " + err.message);
    return { guests: [], summary: { totalRevenue: "Error", count: 0, period: "Sheet Error", lifetimeRevenue: "Error" } };
  }
}

/**
 * Verifies if a specific WhatsApp/Mobile number exists in the Master QID Verified registry.
 * Normalizes both inputs to ensure accurate 10-digit matching (ignoring prefixes like +91).
 * * @param {string|number} whatsappNo - The guest mobile number to look up.
 * @return {boolean} True if verified in the master ledger, false otherwise.
 */
function findGuestQIDVerified(whatsappNo) {
  if (!whatsappNo) return false;
  
  try {
    // 1. Normalize target number to raw last 10 digits
    let targetClean = whatsappNo.toString().replace(/\D/g, "");
    if (targetClean.length > 10) targetClean = targetClean.slice(-10);
    if (targetClean.length !== 10) return false; // Guard clause against invalid structures

    // 2. Access the Master QID Document
    const ss = SpreadsheetApp.openById(ID_QID_VERIFIED_LIST);
    const sheet = ss.getSheetByName(SHEET_NAME_QID);
    
    if (!sheet) {
      console.error(`>>> ❌ [QID LOOKUP ERROR] Tab "${SHEET_NAME_QID}" not found.`);
      return false;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false; // Empty sheet check

    // 3. Extract data range (starting from row 2 to bypass headers)
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    // Column Index 5 corresponds to Column 6 (Phone / Whatsapp) in your registry sheet mapping
    const PHONE_COL_INDEX = 5; 

    // 4. Scan column for a sanitized match
    for (let i = 0; i < data.length; i++) {
      let cellVal = data[i][PHONE_COL_INDEX];
      if (cellVal) {
        let currentClean = cellVal.toString().replace(/\D/g, "");
        if (currentClean.length > 10) currentClean = currentClean.slice(-10);
        
        if (currentClean === targetClean) {
          console.log(`>>> 🎯 [QID MATCH FOUND] Verified number match at registry index row: ${i + 2}`);
          return true; // Match found instantly, break loop early
        }
      }
    }
    
    return false; // No matches found across the iteration feed
    
  } catch (err) {
    console.error(">>> ❌ [findGuestQIDVerified Exception] Check failed: " + err.message);
    return false;
  }
}

/**** Sync with Airbnb bookings feature - Fresh Confirmations with Daily Timeline Sorting *****/
function syncAirbnbEmails(selectedYear) {
  try {
    const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);

    // 1. DYNAMIC YEAR FALLBACK: Use passed year, or calculate current year if empty
    const targetYear = selectedYear || new Date().getFullYear().toString();

    let sheet = ss.getSheetByName(targetYear) || ss.getSheets()[0];

    const data = sheet.getDataRange().getValues();
    const headers = data.find(row => row.includes("Name"));
    if (!headers) return "Error: Could not find header row in sheet.";

    // Setup the Gmail Sync tracking label
    const labelName = "Vinyasa-Synced";
    let syncLabel = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);

    // Dynamic Column Matrix Mapping Indexes for Cancellations & Reviews
    const nameColIdx = headers.indexOf("Name");
    const checkInColIdx = headers.indexOf("Check-in Date");
    const amountColIdx = headers.indexOf("Amount");
    const commentsColIdx = headers.indexOf("Comments");
    const ratingsColIdx = headers.indexOf("Customer Ratings") !== -1 ? headers.indexOf("Customer Ratings") : headers.indexOf("Ratings");

    // --- FULL MONTH NAME DICTIONARY ---
    const monthMap = {
      "jan": "January", "feb": "February", "mar": "March", "apr": "April",
      "may": "May", "jun": "June", "jul": "July", "aug": "August",
      "sep": "September", "oct": "October", "nov": "November", "dec": "December"
    };

    // Target fresh confirmations, cancellations, or reviews that haven't been processed yet
    const query = `from:automated@airbnb.com ("Reservation confirmed" OR "review" OR "Canceled:") -label:${labelName}`;
    const threads = GmailApp.search(query, 0, 15);

    let newBookingsCount = 0;
    let reviewsCount = 0;
    let cancellationCount = 0;

    threads.forEach(thread => {
      const messages = thread.getMessages();
      let processedThread = false;

      messages.forEach(message => {
        const subject = message.getSubject();
        const body = message.getPlainBody();
        const combinedTextToAnalyze = (subject + " " + body);

        // =========================================================================
        // SCENARIO A: GUEST LEFT A REVIEW (REPEATING STAR GRAPHIC FORMAT)
        // =========================================================================
        if (combinedTextToAnalyze.toLowerCase().includes("left a") && combinedTextToAnalyze.toLowerCase().includes("review")) {
          console.log(`\n=================================================================`);
          console.log(`>>> [REVIEW TRACE START] Processing Review Email`);

          const reviewMatch = combinedTextToAnalyze.match(/(.+?)\s+left\s+a\s+(\d+)-star\s+review/i);

          if (reviewMatch) {
            let rawName = reviewMatch[1].replace(/canceled:|cancelled:|reservation confirmed\s*-\s*/i, "").trim();
            const reviewerShortName = rawName.toLowerCase();

            // Extract the numeric rating and turn it into a row of emojis (e.g., 5 -> "⭐⭐⭐⭐⭐")
            const numericRating = parseInt(reviewMatch[2].trim());
            const starRating = "⭐".repeat(numericRating);

            let currentSheetData = sheet.getDataRange().getValues();
            const platformColIdx = headers.indexOf("AirBnb\\Personal");

            console.log(`>>> [REVIEW TARGET] Seeking Short Name: "${reviewerShortName}" | Formatted Rating: ${starRating}`);
            console.log(`>>> [REVIEW LOOP] Scanning spreadsheet rows for partial name match...`);

            let matchFound = false;

            for (let row = 1; row < currentSheetData.length; row++) {
              let rawSheetName = currentSheetData[row][nameColIdx] ? currentSheetData[row][nameColIdx].toString() : "";
              let existingFullName = rawSheetName.trim().toLowerCase();

              let rawPlatform = platformColIdx !== -1 && currentSheetData[row][platformColIdx] ? currentSheetData[row][platformColIdx].toString() : "";
              let platformType = rawPlatform.trim().toLowerCase();

              let nameMatches = existingFullName.includes(reviewerShortName);

              if (nameMatches && platformType === "airbnb" && ratingsColIdx !== -1) {
                sheet.getRange(row + 1, ratingsColIdx + 1).setValue(starRating);
                console.log(`\n>>> [REVIEW MATCH SUCCESS] Found row! Set row ${row + 1}, col ${ratingsColIdx + 1} to: ${starRating}`);
                reviewsCount++;
                processedThread = true;
                matchFound = true;
                break;
              }
            }

            if (!matchFound) {
              console.warn(`\n>>> [REVIEW FAILURE] Scanned entire loop matrix. Could not partially match "${reviewerShortName}" under channel category "airbnb".`);
            }
          } else {
            console.warn(`>>> [REVIEW REGEX ERROR] Found review text triggers, but match pattern execution dropped.`);
          }
          console.log(`=================================================================\n`);
          return;
        }

        // =========================================================================
        // SCENARIO B: CANCELLATION NOTICE RECEIVED (OBJECT-SAFE COMPONENT MATCH)
        // =========================================================================
        if (combinedTextToAnalyze.toLowerCase().includes("canceled:") || combinedTextToAnalyze.toLowerCase().includes("cancelled:")) {
          const cancelMatch = combinedTextToAnalyze.match(/Reservation\s+([A-Z0-9]{10})/i) || combinedTextToAnalyze.match(/code\s+([A-Z0-9]{10})/i);
          const dateMatch = subject.match(/for\s+([A-Z][a-z]{2}\s+\d+)/i);

          let targetYearNum = parseInt(targetYear);
          let targetMonthNum = -1;
          let targetDayNum = -1;

          if (dateMatch) {
            let emailDateObj = new Date(dateMatch[1].trim() + `, ${targetYear}`);
            if (!isNaN(emailDateObj.getTime())) {
              targetMonthNum = emailDateObj.getMonth() + 1;
              targetDayNum = emailDateObj.getDate();
            }
          }

          const targetConfirmationCode = cancelMatch ? cancelMatch[1].trim() : "NOT_FOUND";
          let currentSheetData = sheet.getDataRange().getValues();
          let matchRowIdx = -1;

          for (let row = 1; row < currentSheetData.length; row++) {
            let commentsValue = currentSheetData[row][commentsColIdx] ? currentSheetData[row][commentsColIdx].toString() : "";
            let rowCheckInValue = currentSheetData[row][checkInColIdx];

            // Strategy 1: Unique Confirmation Code Matching
            if (targetConfirmationCode !== "NOT_FOUND" && commentsValue.includes(targetConfirmationCode)) {
              matchRowIdx = row + 1;
              break;
            }

            // Strategy 2: Type-Safe Calendar Date Comparison
            if (targetMonthNum !== -1 && rowCheckInValue) {
              let rowMonth, rowDay, rowYear;

              if (rowCheckInValue instanceof Date) {
                rowMonth = rowCheckInValue.getMonth() + 1;
                rowDay = rowCheckInValue.getDate();
                rowYear = rowCheckInValue.getFullYear();
              } else {
                let dateParts = rowCheckInValue.toString().trim().split("/");
                if (dateParts.length === 3) {
                  rowMonth = parseInt(dateParts[0]);
                  rowDay = parseInt(dateParts[1]);
                  rowYear = parseInt(dateParts[2]);
                }
              }

              if (rowMonth === targetMonthNum && rowDay === targetDayNum && rowYear === targetYearNum) {
                matchRowIdx = row + 1;
                break;
              }
            }
          }

          if (matchRowIdx !== -1) {
            sheet.getRange(matchRowIdx, commentsColIdx + 1).setValue(`CANCELLED CODE: ${targetConfirmationCode}. Flagged via sync.`);

            // --- CLEAR AMOUNT COLUMN FOR CALCULATION EXCLUSION ---
            if (amountColIdx !== -1) {
              sheet.getRange(matchRowIdx, amountColIdx + 1).setValue("");
            }

            sheet.getRange(matchRowIdx, 1, 1, headers.length).setFontColor("#e53e3e");
            cancellationCount++;
            processedThread = true;
          }
          return;
        }

        // =========================================================================
        // SCENARIO C: RESERVATION CONFIRMED (YOUR ORIGINAL UNTOUCHED WORKFLOW)
        // =========================================================================
        const subjectMatch = subject.match(/Reservation confirmed\s*-\s*(.*?)\s+arrives\s+(.*)/i);

        let guestName = "";
        let checkInStr = "";

        if (subjectMatch) {
          guestName = subjectMatch[1].trim();
          checkInStr = subjectMatch[2].trim() + `, ${targetYear}`;
        } else {
          return;
        }

        if (!guestName) return;

        const nightsMatch = body.match(/(\d+)\s*nights\s*room\s*fee/i) || body.match(/(\d+)\s*night/i);
        const nights = nightsMatch ? Number(nightsMatch[1]) : 1;

        const amountMatch = body.match(/You earn[\s\S]*?₹?\s*([\d,]+\.?\d*)/i);
        let finalAmount = "0";
        if (amountMatch) {
          let rawAmt = amountMatch[1].replace(/,/g, "");
          finalAmount = Math.round(parseFloat(rawAmt)).toString();
        }

        // --- FIXED GUEST PARSING ENGINE ---
        // Looks specifically for an isolated digit followed by "adult" or "guest" to bypass header profiles
        const guestsMatch = body.match(/(\d+)\s*adult/i) || body.match(/(\d+)\s*guest(?!\s+will)/i);
        const totalGuests = guestsMatch ? Number(guestsMatch[1]) : 1; // Default fallback to 1 instead of 2

        //let currentMonthStr = checkInStr.split(" ")[0]; 
        // --- FIXED MONTH STR TRANSFORMATION ---
        let rawMonthAbbreviation = checkInStr.split(" ")[0].toLowerCase().replace(/[^a-z]/g, "");
        let fullMonthName = monthMap[rawMonthAbbreviation] || checkInStr.split(" ")[0];

        const incomingCheckInDate = new Date(checkInStr);
        const incomingCheckInTime = !isNaN(incomingCheckInDate.getTime()) ? incomingCheckInDate.getTime() : 0;

        // Parse verification code to anchor upcoming cancellation logic lookups
        const codeMatch = body.match(/Reservation\s*code\s*([A-Z0-9]{10})/i) || body.match(/Confirmation\s*code\s*([A-Z0-9]{10})/i);
        const confirmationCode = codeMatch ? codeMatch[1].trim() : "NOT_FOUND";
        const savedComment = confirmationCode !== "NOT_FOUND" ? `Code: ${confirmationCode}. Automated Gmail Sync Engine.` : "Automated Gmail Sync Engine.";

        let newRowData = new Array(headers.length).fill("");
        newRowData[headers.indexOf("Month")] = fullMonthName;
        newRowData[headers.indexOf("Name")] = guestName;
        newRowData[headers.indexOf("Guests")] = totalGuests;
        newRowData[headers.indexOf("Amount")] = finalAmount;
        newRowData[headers.indexOf("Check-in Date")] = checkInStr;
        newRowData[headers.indexOf("Days")] = nights;
        newRowData[headers.indexOf("AirBnb\\Personal")] = "AirBnb";
        newRowData[headers.indexOf("Floor")] = "Ground";
        newRowData[commentsColIdx] = savedComment;

        let currentSheetData = sheet.getDataRange().getValues();
        let lastRowWithContent = sheet.getLastRow();

        let insertionRowIndex = lastRowWithContent;
        let foundInsertionSpot = false;

        for (let i = currentSheetData.length - 1; i >= 1; i--) {
          let rowCheckInVal = currentSheetData[i][checkInColIdx] ? currentSheetData[i][checkInColIdx].toString().trim() : "";

          if (rowCheckInVal) {
            let rowDate = new Date(rowCheckInVal);
            let rowDateTime = rowDate.getTime();

            if (!isNaN(rowDateTime)) {
              if (rowDateTime <= incomingCheckInTime) {
                insertionRowIndex = i + 1;
                foundInsertionSpot = true;
                break;
              }
            }
          }
        }

        if (!foundInsertionSpot && lastRowWithContent > 1) {
          insertionRowIndex = 1;
        }

        console.log(`>>> [TIMELINE INSERT] Placing ${guestName} (${checkInStr}) directly after row: [${insertionRowIndex}]`);

        sheet.insertRowsAfter(insertionRowIndex, 1);

        let templateRow = (insertionRowIndex === 1) ? 2 : insertionRowIndex;
        let templateRange = sheet.getRange(templateRow, 1, 1, headers.length);
        let targetRange = sheet.getRange(insertionRowIndex + 1, 1, 1, headers.length);

        templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        targetRange.setValues([newRowData]);

        // --- CALENDAR SYNC INJECTION POINT ---
        // Fire sync immediately following successful cell mapping matrix row placement
        syncBookingToVinyasaCalendar(
          guestName, 
          incomingCheckInDate, 
          totalGuests, 
          "AirBnb", 
          "Ground", 
          nights, 
          savedComment
        );

        newBookingsCount++;
        processedThread = true;

      });

      if (processedThread) {
        thread.addLabel(syncLabel);
        thread.markRead();
      }
    });

    return `Sync Successfully Completed!\n\n` +
      `📥 New Bookings:\u2003\u2003\u2003\u2003${newBookingsCount} Added\n` +
      `⭐ Ratings/Reviews:\u2003\u2003${reviewsCount} Updated\n` +
      `❌ Cancellations:\u2003\u2003\u2003${cancellationCount} Processed`;

  } catch (err) {
    console.error("Parser tracking fail: " + err.message);
    throw new Error("Sync processing aborted: " + err.message);
  }
}

/*****************************Modal to ADD\EDIT Guest Details ***********************************/
/*****************************Modal to ADD\EDIT Guest Details ***********************************/
function writeGuestDataRow(mode, payload) {
  // --- FIRST-LINE GATEKEEPER VALIDATION ---
  if (!payload.checkIn || payload.checkIn.toString().trim() === "") {
    throw new Error("Transaction Denied: Check-in Date is a mandatory field and cannot be left blank.");
  }

  const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
  let sheet = ss.getSheetByName(payload.year.toString().trim());

  if (!sheet) {
    sheet = ss.insertSheet(payload.year.toString().trim());
    sheet.appendRow(["Month", "Name", "Guests", "Amount", "Check-in Date", "Days", "Source", "AirBnb\\Personal", "Floor", "Mobile", "Customer Ratings", "Comments"]);
    SpreadsheetApp.flush();
  }

  let dataRange = sheet.getDataRange().getValues();
  const headers = dataRange[0].map(h => h.toString().trim());

  const mapping = {
    Name: headers.indexOf("Name"),
    Amount: headers.indexOf("Amount"),
    Source: headers.indexOf("AirBnb\\Personal"),
    Floor: headers.indexOf("Floor"),
    Mobile: headers.indexOf("Mobile"),
    CheckIn: headers.indexOf("Check-in Date"),
    Days: headers.indexOf("Days"),
    Guests: headers.indexOf("Guests"),
    Ratings: headers.indexOf("Customer Ratings"),
    Comments: headers.indexOf("Comments")
  };

  const monthColIdx = headers.indexOf("Month");

  // --- 1. NORMALIZE MOBILE NUMBER FORMAT ---
  let cleanMobile = payload.mobile ? payload.mobile.toString().trim() : "";
  if (cleanMobile.startsWith("+91")) {
    cleanMobile = cleanMobile.substring(3);
  } else if (cleanMobile.startsWith("91") && cleanMobile.length > 10) {
    cleanMobile = cleanMobile.substring(2);
  }

  // --- 2. SECURE TIMESTAMP PARSING & UNIFIED STRINGS FORMATION ---
  let generatedMonthLabel = "January";
  let standardizedCheckInString = payload.checkIn;

  const monthsArray = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const shortMonthsArray = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  const dParsed = parseDateSecurely(payload.checkIn);

  if (!dParsed) {
    throw new Error("Transaction Denied: Invalid Date format provided for Check-in Date.");
  }

  generatedMonthLabel = monthsArray[dParsed.getMonth()];

  // Format the raw UI value (YYYY-MM-DD) into the standard string format: "May 23, 2026"
  standardizedCheckInString = `${shortMonthsArray[dParsed.getMonth()]} ${dParsed.getDate()}, ${dParsed.getFullYear()}`;

  // --- 3. BUILD DATA VECTOR ARRAY ---
  let newRow = new Array(headers.length).fill("");
  if (mapping.Name !== -1) newRow[mapping.Name] = payload.name;
  if (mapping.Amount !== -1) newRow[mapping.Amount] = payload.amount;
  if (monthColIdx !== -1) newRow[monthColIdx] = generatedMonthLabel;

  if (mapping.Source !== -1) {
    let cleanSource = payload.source.toString().trim().toLowerCase();
    newRow[mapping.Source] = cleanSource.includes("airbnb") ? "AirBnb" : "Personal";
  }

  if (mapping.Floor !== -1) {
    let cleanFloor = payload.floor.toString().trim().toLowerCase();
    newRow[mapping.Floor] = cleanFloor.includes("second") ? "Second" : "Ground";
  }

  if (mapping.Mobile !== -1) newRow[mapping.Mobile] = cleanMobile;
  if (mapping.CheckIn !== -1) newRow[mapping.CheckIn] = standardizedCheckInString;
  if (mapping.Days !== -1) newRow[mapping.Days] = payload.days;
  if (mapping.Guests !== -1) newRow[mapping.Guests] = payload.guests;
  if (mapping.Ratings !== -1) newRow[mapping.Ratings] = payload.ratings;
  if (mapping.Comments !== -1) newRow[mapping.Comments] = payload.comments;

  // =========================================================================
  // --- 4. EXECUTE ROUTING (IN-PLACE EDIT VS CHRONOLOGICAL APPEND) ---
  // =========================================================================
  if (mode === "EDIT") {
    let targetRowNumber = parseInt(payload.rowIndex);

    // Fallback: If row index is lost, search the values matrix for a name match
    if (isNaN(targetRowNumber) || targetRowNumber < 2) {
      console.log(">>> [EDIT] RowIndex missing. Running fallback search by Name...");
      const cleanTargetName = payload.name.toString().trim().toLowerCase();
      for (let i = 1; i < dataRange.length; i++) {
        const currentSheetName = dataRange[i][mapping.Name] ? dataRange[i][mapping.Name].toString().trim().toLowerCase() : "";
        if (currentSheetName === cleanTargetName) {
          targetRowNumber = i + 1;
          break;
        }
      }
    }

    if (targetRowNumber > 1 && targetRowNumber <= sheet.getLastRow()) {
      console.log(`>>> [EDIT SUCCESS] Modifying row [${targetRowNumber}] in-place.`);
      // Update row fields dynamically without breaking structural row boundaries
      sheet.getRange(targetRowNumber, 1, 1, headers.length).setValues([newRow]);
    } else {
      throw new Error("Transaction Aborted: Target spreadsheet row link data could not be verified.");
    }

  } else {
    // --- MODE: ADD NEW ENTRY (TIMELINE FIXED INSERTION) ---
    let currentSheetData = sheet.getDataRange().getValues();
    let lastRowWithContent = sheet.getLastRow();
    let incomingCheckInTime = dParsed.getTime();

    let insertionRowIndex = lastRowWithContent;
    let foundInsertionSpot = false;

    for (let i = currentSheetData.length - 1; i >= 1; i--) {
      let rowCheckInVal = currentSheetData[i][mapping.CheckIn];
      if (rowCheckInVal) {
        let rowDate = parseDateSecurely(rowCheckInVal);
        if (rowDate) {
          if (rowDate.getTime() <= incomingCheckInTime) {
            insertionRowIndex = i + 1;
            foundInsertionSpot = true;
            break;
          }
        }
      }
    }

    if (!foundInsertionSpot && lastRowWithContent > 1) {
      insertionRowIndex = 1;
    }

    console.log(`>>> [ADD SUCCESS] Inserting ${payload.name} safely after row: [${insertionRowIndex}]`);
    sheet.insertRowsAfter(insertionRowIndex, 1);

    let templateRow = (insertionRowIndex === 1) ? 2 : insertionRowIndex;
    let templateRange = sheet.getRange(templateRow, 1, 1, headers.length);
    let targetRange = sheet.getRange(insertionRowIndex + 1, 1, 1, headers.length);

    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    targetRange.setValues([newRow]);
  }

  // =========================================================================
  // --- 5. AUTOMATED SHEET RE-SORTATION ENGINE ---
  // =========================================================================
  // This step ensures your ledger stays perfectly chronological, whether adding or editing
  SpreadsheetApp.flush();
  const activeLastRow = sheet.getLastRow();
  if (activeLastRow > 1) {
    // Sorts range by the "Check-in Date" column index ascending, keeping headers intact
    const sortTargetRange = sheet.getRange(2, 1, activeLastRow - 1, headers.length);
    sortTargetRange.sort({ column: mapping.CheckIn + 1, ascending: true });
    console.log(">>> [SORT ENGINE] Sheet successfully sorted chronologically by Check-in Date.");
  }

  // --- CALENDAR SYNC INJECTION POINT ---
  // Captures updates from both new manual reservations and profile updates from the dashboard editor
  syncBookingToVinyasaCalendar(
    payload.name,
    dParsed,
    payload.guests,
    payload.source,
    payload.floor,
    payload.days,
    payload.comments
  );

  SpreadsheetApp.flush();
  return "SUCCESS";
}

/**
 * Core Database & Calendar Sync Writer: Purges an entire reservation entry row
 * from the sheets AND automatically wipes the matching Google Calendar event.
 */
/**
 * Core Database & Calendar Sync Writer: Purges an entire reservation entry row
 * from the sheets AND automatically wipes the matching Google Calendar event.
 */
function deleteGuestRowBackend(name, checkInStr) {
  try {
    if (!checkInStr || checkInStr === "undefined") {
      throw new Error("Missing mandatory parameter: checkInStr evaluation failed.");
    }

    // 1. Resolve target spreadsheet year natively on the server side
    const checkInDateObj = new Date(checkInStr);
    if (isNaN(checkInDateObj.getTime())) {
      throw new Error(`Invalid date string received by server: '${checkInStr}'`);
    }
    
    const targetSheetName = checkInDateObj.getFullYear().toString(); // Resolves safely to "2026"
    
    const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
    const sheet = ss.getSheetByName(targetSheetName);
    
    if (!sheet) {
      throw new Error(`Target year data registry tab '${targetSheetName}' could not be discovered.`);
    }
    
    const dataRangeValues = sheet.getDataRange().getValues();
    const headers = dataRangeValues[0];
    const nameColIdx = headers.indexOf("Name");
    const checkInColIdx = headers.indexOf("Check-in Date");
    
    if (nameColIdx === -1 || checkInColIdx === -1) {
      throw new Error("Ledger file corruption: Mandatory tracking columns are missing.");
    }
    
    const targetNameClean = name.toString().trim().toLowerCase();
    
    // 2. Traverse the registry rows backwards to avoid index shifting during extraction
    for (let row = dataRangeValues.length - 1; row >= 1; row--) {
      const rowName = dataRangeValues[row][nameColIdx] ? dataRangeValues[row][nameColIdx].toString().trim().toLowerCase() : "";
      
      // Parse row date to ensure accurate comparison matching format styles
      const rowCheckInRaw = dataRangeValues[row][checkInColIdx];
      let rowDateMatch = false;
      
      if (rowCheckInRaw) {
        const rowDateObj = new Date(rowCheckInRaw);
        if (!isNaN(rowDateObj.getTime())) {
          // Compare clean time integers rather than volatile string formats
          rowDateMatch = (rowDateObj.toDateString() === checkInDateObj.toDateString());
        }
      }
      
      // Match the row identifiers safely
      if (rowName === targetNameClean && rowDateMatch) {
        const targetRowPosition = row + 1;
        
        // 3. Automated Calendar Purge Module
        try {
          const calendarName = "Vinyasa Nilaya";
          const calendars = CalendarApp.getCalendarsByName(calendarName);
          const targetCalendar = (calendars.length > 0) ? calendars[0] : CalendarApp.getDefaultCalendar();
          
          // Set up a clean 24-hour block window to scan for this specific check-in day
          const startWindow = new Date(checkInDateObj.getFullYear(), checkInDateObj.getMonth(), checkInDateObj.getDate(), 0, 0, 0);
          const endWindow = new Date(checkInDateObj.getFullYear(), checkInDateObj.getMonth(), checkInDateObj.getDate(), 23, 59, 59);
          
          const events = targetCalendar.getEvents(startWindow, endWindow);
          let calendarDeletedCount = 0;
          
          events.forEach(ev => {
            const title = ev.getTitle();
            if (title.toLowerCase().includes(targetNameClean)) {
              ev.deleteEvent();
              calendarDeletedCount++;
            }
          });
          
          if (calendarDeletedCount > 0) {
            console.log(`✨ [CALENDAR SYNC] Removed ${calendarDeletedCount} timeline blocks for ${name}.`);
          }
        } catch (calErr) {
          console.error("❌ [CALENDAR PURGE ERROR] Sync skipped:", calErr.toString());
        }
        
        // 4. Delete the row out of the spreadsheet layout boundaries
        sheet.deleteRow(targetRowPosition);
        SpreadsheetApp.flush(); 
        
        console.log(`>>> [BACKEND REMOVAL SUCCESS] Deleted row index [${targetRowPosition}] for guest: ${name}`);
        return `SUCCESS: Row ${targetRowPosition} eliminated cleanly.`;
      }
    }
    
    throw new Error(`The requested guest profile context for '${name}' could not be matched within active row limits.`);
  } catch (err) {
    console.error(">>> [DELETE GUEST BACKEND FATAL ERROR]", err);
    throw new Error(err.message);
  }
}

/**
 * HELPER FUNCTION: Safely parses strings (both YYYY-MM-DD and Textual formats), Objects, or Serials
 */
function parseDateSecurely(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) {
    return !isNaN(dateVal.getTime()) ? dateVal : null;
  }

  let dateStr = dateVal.toString().trim();

  // Fixes HTML standard format layout detection (YYYY-MM-DD)
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    let parts = dateStr.split("-");
    // Explicit construction avoiding local timezone shifts
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }

  let parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) return parsed;

  const monthsMap = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
    january: 0, february: 1, march: 2, april: 3, may: 4, june: 5, july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
  };

  let tokens = dateStr.replace(/,/g, "").split(/\s+/);
  if (tokens.length >= 2) {
    let day = parseInt(tokens[1]);
    let monthStr = tokens[0].toLowerCase();
    let year = parseInt(tokens[2]) || new Date().getFullYear();

    if (isNaN(day)) {
      day = parseInt(tokens[0]);
      monthStr = tokens[1].toLowerCase();
    }

    if (!isNaN(day) && monthsMap[monthStr] !== undefined) {
      return new Date(year, monthsMap[monthStr], day);
    }
  }

  return null;
}
/**
 * HELPER FUNCTION: Safely parses dates from strings, objects, or numbers
 */
function parseDateSecurely(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) {
    return !isNaN(dateVal.getTime()) ? dateVal : null;
  }

  let dateStr = dateVal.toString().trim();
  let parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) return parsed;

  const monthsMap = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
    january: 0, february: 1, march: 2, april: 3, may: 4, june: 5, july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
  };

  let tokens = dateStr.replace(/,/g, "").split(/\s+/);
  if (tokens.length >= 2) {
    let day = parseInt(tokens[1]);
    let monthStr = tokens[0].toLowerCase();
    let year = parseInt(tokens[2]) || new Date().getFullYear();

    if (isNaN(day)) {
      day = parseInt(tokens[0]);
      monthStr = tokens[1].toLowerCase();
    }

    if (!isNaN(day) && monthsMap[monthStr] !== undefined) {
      return new Date(year, monthsMap[monthStr], day);
    }
  }

  return null;
}

/**
 * HELPER FUNCTION: Safely parses dates from strings, objects, or numbers
 */
function parseDateSecurely(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) {
    return !isNaN(dateVal.getTime()) ? dateVal : null;
  }

  let dateStr = dateVal.toString().trim();
  let parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) return parsed;

  const monthsMap = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
    january: 0, february: 1, march: 2, april: 3, may: 4, june: 5, july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
  };

  let tokens = dateStr.replace(/,/g, "").split(/\s+/);
  if (tokens.length >= 2) {
    let day = parseInt(tokens[1]);
    let monthStr = tokens[0].toLowerCase();
    let year = parseInt(tokens[2]) || new Date().getFullYear();

    if (isNaN(day)) {
      day = parseInt(tokens[0]);
      monthStr = tokens[1].toLowerCase();
    }

    if (!isNaN(day) && monthsMap[monthStr] !== undefined) {
      return new Date(year, monthsMap[monthStr], day);
    }
  }

  return null;
}

/**** QID Verified modal ****/
// Ensure this constant matches your application constants at the top of Code.gs
const SHEET_NAME_QID = 'QID-Verified';

/**
 * Pull and serialize raw database matrix rows from the verified QID spreadsheet layout
 * 
 */

/**
 * Pull, serialize, and diagnose raw database matrix rows from the verified QID spreadsheet layout
 */
function fetchQidVerifiedRegistry() {
 // console.log("=======================================================");
 // console.log(">>> 🔍 [QID DIAGNOSTIC START] Initializing Registry Engine...");
 // console.log("=======================================================");
  
  try {
    const ss = SpreadsheetApp.openById(ID_QID_VERIFIED_LIST);
   // console.log(">>> [DIAGNOSTIC] Active Spreadsheet Name: " + ss.getName());
   // console.log(">>> [DIAGNOSTIC] Active Spreadsheet ID: " + ss.getId());
    
    const sheet = ss.getSheetByName(SHEET_NAME_QID);
    
    if (!sheet) {
     // console.error(">>> ❌ [DIAGNOSTIC ERROR] Could not locate sheet tab named precisely: '" + SHEET_NAME_QID + "'");
      const sheets = ss.getSheets();
      let sheetNames = sheets.map(function(s) { return "'" + s.getName() + "'"; }).join(", ");
     // console.log(">>> [DIAGNOSTIC] Available tabs in this file are: [" + sheetNames + "]");
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    //console.log(`>>> 📊 [DIAGNOSTIC] Target Sheet Found! Dimensions: [Rows: ${lastRow} | Columns: ${lastCol}]`);
    
    if (lastRow <= 1) {
     // console.warn(">>> ⚠️ [DIAGNOSTIC WARN] Sheet exists but appears to have NO data rows below the header.");
      return [];
    }
    
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    //console.log(">>> 📋 [DIAGNOSTIC] Raw Headers Discovered: " + JSON.stringify(headers));
    
    let serializedRecords = [];
    let blankNameCount = 0;
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rawName = row[4]; // Col 5: Name
      
      if (i <= 3) {
        //console.log(`>>> 🔍 [ROW ${i+1} SAMPLE] Raw Array Data: ` + JSON.stringify(row));
        console.log(`>>> [ROW ${i+1} SAMPLE] Extracted Name field (Index 4): "${rawName}"`);
      }
      
      if (!rawName || rawName.toString().trim() === "") {
        blankNameCount++;
        continue; 
      }

      // --- CRITICAL NETWORK SERIALIZATION FIX ---
      // Converts raw JavaScript Date Objects to an ISO String to prevent network transfer dropping
      let safeTimestamp = "";
      if (row[1]) {
        safeTimestamp = (row[1] instanceof Date) ? row[1].toISOString() : row[1].toString();
      }
      
      serializedRecords.push({
        slNo: row[0],                                   // Col 1: SlNo
        timestamp: safeTimestamp,                       // Col 2: Timestamp
        idType: row[2] || "Govt ID",                    // Col 3: ID Type
        idNo: row[3] || "-",                            // Col 4: ID No
        name: rawName.toString().trim(),                // Col 5: Name
        phone: row[5] ? row[5].toString().trim() : "-", // Col 6: Phone / Whatsapp
        purpose: row[6] || "-",                         // Col 7: Purpose Of Travel
        arrivingCity: row[7] || "-",                    // Col 8: Ariving City
        emergencyName: row[8] || "-",                   // Col 9: Emergency Contact Name
        emergencyPhone: row[9] || "-",                  // Col 10: Emergency Contant No
        frontUrl: row[10] || "",                        // Col 11: Govt-ID-Front URL
        backUrl: row[11] || "",                         // Col 12: Govt-ID-Back URL
        selfieUrl: row[12] || "",                       // Col 14: Selfie URL (Index 13 matches your update logic)
        checkinStatus: row[13] || "Verified",           // Col 13: Checkin status (Index 12)
        address: row[14] || "-"                         // Col 15: Address
      });
    
    }

    
   // console.log(`=======================================================`);
   // console.log(`>>> 🏁 [QID DIAGNOSTIC END] Successfully parsed: [${serializedRecords.length}] records.`);
   // console.log(`>>> [DIAGNOSTIC] Skipped [${blankNameCount}] rows due to empty Name fields.`);
   // console.log(`=======================================================`);
    
    return serializedRecords.reverse();
    
  } catch (err) {
    console.error(">>> ❌ [DIAGNOSTIC CRITICAL EXCEPTION]", err);
    throw new Error(err.message);
  }
}

/**
 * Delete a QID entry safely by anchoring to its unique sequential Serial Number
 */
function deleteQidRowBackend(slNo) {
  try {
    // --- FIX: Switch from getActiveSpreadsheet to explicit ID matching ---
    const ss = SpreadsheetApp.openById(ID_QID_VERIFIED_LIST);
    const sheet = ss.getSheetByName(SHEET_NAME_QID);
    
    if (!sheet) {
      throw new Error("Target ledger sheet configuration '" + SHEET_NAME_QID + "' not found.");
    }
    
    const values = sheet.getDataRange().getValues();
    
    for (let i = 1; i < values.length; i++) {
      // Column A contains the SlNo map index
      if (parseInt(values[i][0]) === parseInt(slNo)) {
        const actualRowInSheet = i + 1;
        const targetRowData = values[i];

        // Extract URLs from indices 10, 11, and 12
        const frontUrl  = targetRowData[10] || "";
        const backUrl   = targetRowData[11] || "";
        const selfieUrl = targetRowData[12] || "";

        const filesToPurge = [frontUrl, backUrl, selfieUrl];
        let fileDeleteCount = 0;

        filesToPurge.forEach(url => {
          if (url && url.toString().trim() !== "") {
            // Call the robust extractor helper defined below
            const fileId = extractDriveIdSafely(url.toString().trim());
            
            if (fileId) {
              try {
                const file = DriveApp.getFileById(fileId);
                file.setTrashed(true);
                fileDeleteCount++;
                console.log(`>>> [STORAGE PURGE] Trashed associated Drive File ID: [${fileId}]`);
              } catch (fileErr) {
                console.warn(`>>> [STORAGE PURGE WARNING] File ID [${fileId}] could not be found or was already removed: ${fileErr.message}`);
              }
            }
          }
        });

        console.log(`>>> [STORAGE PURGE COMPLETE] Total asset objects moved to trash bin: [${fileDeleteCount}]`);

        // Delete the ledger row matrix completely from the spreadsheet
        sheet.deleteRow(actualRowInSheet);
        SpreadsheetApp.flush();
        console.log(">>> [BACKEND DELETE SUCCESS] Removed SlNo [" + slNo + "] at Sheet Row [" + actualRowInSheet + "]");
        return "SUCCESS";
      }
    }
    throw new Error("Record with Serial Number " + slNo + " was not found inside the ledger matrix.");
  } catch (err) {
    console.error(">>> [DELETE QID BACKEND CRITICAL ERROR]", err);
    throw new Error(err.message);
  }
}

/**
 * Robust helper function to extract a 33-character Google Drive File ID 
 * from various standard Google Drive link formats.
 */
function extractDriveIdSafely(url) {
  if (!url) return null;
  
  // Pattern 1: Look for standard /file/d/{ID}/view formats
  if (url.includes("/file/d/")) {
    const parts = url.split("/file/d/");
    if (parts.length > 1) {
      return parts[1].split("/")[0];
    }
  }
  
  // Pattern 2: Look for query parameter structures (?id=... or &id=...)
  const match = url.match(/[?&]id=([^&]+)/);
  if (match && match[1]) {
    return match[1];
  }
  
  return null;
}

/**
 * Perform custom inline edits for name or phone profiles securely
 */
function modifyQidRowBackend(slNo, updates) {
  try {
    // --- FIX: Switch from getActiveSpreadsheet to explicit ID matching ---
    const ss = SpreadsheetApp.openById(ID_QID_VERIFIED_LIST);
    const sheet = ss.getSheetByName(SHEET_NAME_QID);
    
    if (!sheet) {
      throw new Error("Target ledger sheet configuration '" + SHEET_NAME_QID + "' not found.");
    }
    
    const values = sheet.getDataRange().getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (parseInt(values[i][0]) === parseInt(slNo)) {
        const actualRowInSheet = i + 1;
        
        // Update Name (Col 5 -> index 4) & Phone (Col 6 -> index 5)
        sheet.getRange(actualRowInSheet, 5).setValue(updates.name);
        sheet.getRange(actualRowInSheet, 6).setValue(updates.phone);
        
        SpreadsheetApp.flush();
        console.log(">>> [BACKEND UPDATE SUCCESS] Modified fields for SlNo [" + slNo + "] at Row [" + actualRowInSheet + "]");
        return "SUCCESS";
      }
    }
    throw new Error("Record with Serial Number " + slNo + " could not be found.");
  } catch (err) {
    console.error(">>> [MODIFY QID BACKEND CRITICAL ERROR]", err);
    throw new Error(err.message);
  }
}


/**
 * Server-Side Proxy: Fetches a Google Drive asset internally and converts it to a safe inline Base64 stream
 */
function getDriveImageAsBase64(rawUrl) {
  try {
    if (!rawUrl || rawUrl.trim() === "" || rawUrl.indexOf("-") === 0) return "";
    
    let fileId = "";
    
    // Extract the raw file ID out of whatever format is stored in the sheet
    if (rawUrl.indexOf("id=") !== -1) {
      fileId = rawUrl.split("id=")[1].split("&")[0];
    } else {
      const matches = rawUrl.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
      fileId = (matches && matches[1]) ? matches[1] : "";
    }
    
    if (!fileId) throw new Error("Could not parse file identifier.");
    
    // Fetch the file bytes directly inside Google's cloud boundaries
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    const contentType = blob.getContentType();
    
    // Encode to a clean inline web asset string
    const base64String = Utilities.base64Encode(bytes);
    return `data:${contentType};base64,${base64String}`;
    
  } catch (err) {
    console.error(">>> [PROXY ERROR] Failed to stream image file bytes: ", err);
    return "ERROR";
  }
}

/**
 * Perform secure inline data mapping modifications for verified guest registry records
 */
function modifyQidRowBackend(slNo, updates) {
  try {
    const ss = SpreadsheetApp.openById(ID_QID_VERIFIED_LIST);
    const sheet = ss.getSheetByName(SHEET_NAME_QID);
    
    if (!sheet) {
      throw new Error("Target registration spreadsheet tab configuration could not be opened.");
    }
    
    const values = sheet.getDataRange().getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (parseInt(values[i][0]) === parseInt(slNo)) {
        const actualRowInSheet = i + 1;
        
        // --- SECURE COLUMN ORIENTATION RE-MAPPING ---
        // Col 4 (Index 3) -> ID Document Number
        // Col 5 (Index 4) -> Guest Full Name
        // Col 6 (Index 5) -> Phone/WhatsApp Mobile Identity
        sheet.getRange(actualRowInSheet, 4).setValue(updates.idNo);
        sheet.getRange(actualRowInSheet, 5).setValue(updates.name);
        sheet.getRange(actualRowInSheet, 6).setValue(updates.phone);
        
        SpreadsheetApp.flush(); // Flush updates out of internal caches straight into the file cell blocks
        console.log(">>> [BACKEND LEDGER SAVE COMPLETED] Modified SlNo [" + slNo + "] inside Row Matrix [" + actualRowInSheet + "]");
        return "SUCCESS";
      }
    }
    throw new Error("Record referencing identification sequence index " + slNo + " vanished unexpectedly.");
  } catch (err) {
    console.error(">>> [MODIFY QID BACKEND EXCEPTION]", err);
    throw new Error(err.message);
  }
}


/*** Google calender sync */
/**
 * Global Sync Engine: Creates or updates an all-day event on the Vinyasa Nilaya calendar
 * Attaches a professional metadata overview card and locks an alert reminder exactly 1 day prior.
 */
function syncBookingToVinyasaCalendar(guestName, checkInDateInput, totalGuests, platformType, floorName, nights, notes) {
  try {
    const calendarName = "Vinyasa Nilaya";
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    const targetCalendar = (calendars.length > 0) ? calendars[0] : CalendarApp.getDefaultCalendar();

    // 1. Resolve date object coordinates safely across input variants
    let checkInDate = (checkInDateInput instanceof Date) ? checkInDateInput : new Date(checkInDateInput);
    if (isNaN(checkInDate.getTime())) {
      checkInDate = parseDateSecurely(checkInDateInput) || new Date();
    }

    // 2. Establish uniform naming blueprints
    const eventTitle = `Guest Check-In: ${guestName} (${platformType || 'Booking'})`;

    // 3. Render a highly professional text block card layout
    const eventDescription = [
      `=========================================`,
      `       VINYASA NILAYA RESERVATION        `,
      `=========================================`,
      `Guest Name    : ${guestName}`,
      `Booking Channel: ${platformType || 'Direct Personal'}`,
      `Allocated Floor: ${floorName || 'Ground'} Floor`,
      `Total Guests  : ${totalGuests || 1} Pax`,
      `Duration      : ${nights || 1} Night(s)`,
      `Check-In Date : ${checkInDate.toDateString()}`,
      `-----------------------------------------`,
      `Reference Information / Notes:`,
      `${notes || 'No operational comments attached.'}`,
      `=========================================`,
      `Synced automatically via Vinyasa Workspace Integration Hub.`
    ].join('\n');

    // 4. Set up strict 24-hour day checking boundaries
    const searchStart = new Date(checkInDate.getTime());
    searchStart.setHours(0, 0, 0, 0);
    const searchEnd = new Date(checkInDate.getTime());
    searchEnd.setHours(23, 59, 59, 999);

    // Pull all active grid elements for that specific check-in day
    const dayEvents = targetCalendar.getEvents(searchStart, searchEnd);
    
    // Scan and purge duplicate entries matching the guest name sequence
    dayEvents.forEach(event => {
      const currentTitle = event.getTitle();
      if (currentTitle.toLowerCase().includes(guestName.trim().toLowerCase())) {
        console.log(`>>> [CALENDAR ENGINE] Dropping old matching layout block: "${currentTitle}"`);
        event.deleteEvent();
      }
    });

    // 5. Append a fresh all-day block into the calendar grid
    const targetEvent = targetCalendar.createAllDayEvent(eventTitle, checkInDate, {
      description: eventDescription
    });
    console.log(`>>> [CALENDAR ENGINE] Added fresh all-day entry block for: ${guestName}`);

    // 6. ENFORCE 24-HOUR ADVANCE REMINDERS
    targetEvent.removeAllReminders();
    targetEvent.addPopupReminder(1440);  
    targetEvent.addEmailReminder(1440);  

    return true;
  } catch (err) {
    console.warn(`>>> [CALENDAR INTERCEPT SYSTEM WARNING] Sync bypassed: ${err.toString()}`);
    return false;
  }
}


/**** QID fileter and bulk delte feature */
/**
 * Feature 2 Backend: Iterates over selected row vectors, wipes Drive file allocations,
 * and handles consecutive matrix indexing contractions from highest index down.
 *
 * @param {Array<string|number>} slNoArray - Unified collection of Serial mapping records.
 * @return {string} Confirmation token feed.
 */
function deleteQidRowsBatchBackend(slNoArray) {
  if (!slNoArray || !Array.isArray(slNoArray) || slNoArray.length === 0) {
    throw new Error("Invalid selection payload collection provided.");
  }

  try {
    const ss = SpreadsheetApp.openById(ID_QID_VERIFIED_LIST);
    const sheet = ss.getSheetByName(SHEET_NAME_QID);
    if (!sheet) throw new Error(`Target tab config "${SHEET_NAME_QID}" missing.`);

    const numericalSlNos = slNoArray.map(id => parseInt(id));
    const values = sheet.getDataRange().getValues();
    let deletedRowsRecordList = [];

    // 1. CLEAR ASSOCIATED DRIVE STORAGE ASSETS
    for (let i = 1; i < values.length; i++) {
      const currentSlNo = parseInt(values[i][0]);
      
      if (numericalSlNos.includes(currentSlNo)) {
        const targetRowData = values[i];
        const frontUrl  = targetRowData[10] || "";
        const backUrl   = targetRowData[11] || "";
        const selfieUrl = targetRowData[12] || "";
        const filesToPurge = [frontUrl, backUrl, selfieUrl];

        filesToPurge.forEach(url => {
          if (url && url.toString().trim() !== "") {
            const fileId = extractDriveIdSafely(url.toString().trim());
            if (fileId) {
              try {
                DriveApp.getFileById(fileId).setTrashed(true);
              } catch (fErr) {
                console.warn(`[BATCH PURGE SKIP] ID ${fileId} inaccessible: ${fErr.message}`);
              }
            }
          }
        });

        // Store the original sheet row index coordinate (1-indexed mapping adjustment)
        deletedRowsRecordList.push(i + 1);
      }
    }

    // 2. CRITICAL STEP: Sort row indices in DESCENDING order before deleting.
    // If you delete row 5 first, row 10 shifts up to row 9, causing data misalignment.
    // Deleting from the bottom up completely bypasses this indexing bug.
    deletedRowsRecordList.sort((a, b) => b - a);

    deletedRowsRecordList.forEach(rowIndex => {
      sheet.deleteRow(rowIndex);
    });

    SpreadsheetApp.flush();
    console.log(`>>> [BATCH DELETION SUCCESS] Successfully purged ${deletedRowsRecordList.length} records from ledger.`);
    return "SUCCESS";

  } catch (err) {
    console.error(">>> [BATCH DELETION EXCEPTION BLOCK] Action failed: ", err);
    throw new Error(err.message);
  }
}

/**
 * Backend API: Scans a parent Google Drive directory for folders named "QID-YYYY",
 * extracts unique year tokens, and returns them ordered with the current active year.
 * * @return {Object} An inventory of available years and the active calendar default year tag.
 */
function getDynamicQidYearsConfig() {
  // CONSTANT PARAMETER: Replace with the exact Folder ID where your QID folders live
  const PARENT_FOLDER_ID = "YOUR_MASTER_ROOT_DRIVE_FOLDER_ID"; 
  
  let detectedYears = [];
  const currentCalendarYear = new Date().getFullYear().toString(); // Default fallback "2026"
  
  try {
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    const subFolders = parentFolder.getFolders();
    
    // Regular Expression targeting patterns like "QID-2026" or "QID-2027"
    const pattern = /^QID-(\d{4})$/;
    
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      const match = folder.getName().trim().match(pattern);
      
      if (match && match[1]) {
        const yearValue = match[1];
        if (!detectedYears.includes(yearValue)) {
          detectedYears.push(yearValue);
        }
      }
    }
    
    // Sort years chronologically in descending order (newest years first)
    detectedYears.sort((a, b) => parseInt(b) - parseInt(a));
    
    // Failsafe condition: If no folders are matched, append current year to keep UI operational
    if (detectedYears.length === 0) {
      detectedYears.push(currentCalendarYear);
    }
    
    console.log(`>>> [SERVER DRIVE ARMED] Detected dynamic QID year sets: [${detectedYears.join(', ')}]`);
    
    return {
      years: detectedYears,
      activeYear: detectedYears.includes(currentCalendarYear) ? currentCalendarYear : detectedYears[0]
    };
    
  } catch (err) {
    console.error(">>> [SERVER EXCEPTION CRASH] Dynamic folder parsing dropped unexpected exception: ", err);
    // Hard fallback layout parameter return
    return {
      years: [currentCalendarYear],
      activeYear: currentCalendarYear
    };
  }
}


/********************* Test functions *************************/
/**
 * Run this function to see exactly what getDashboardData is producing
 * and where it might be failing.
 */
function debugDashboard() {
  const testYear = "2026"; // Change to a year that exists in your tabs
  const testMonth = "Febrauary"; // Change to a month that has data

  console.log(`--- DEBUG START: Year[${testYear}] Month[${testMonth}] ---`);

  try {
    const result = getDashboardData(testYear, testMonth);

    // 1. Check the Summary object
    console.log("Summary Result:", JSON.stringify(result.summary, null, 2));

    // 2. Check the Guest Count
    console.log("Number of guests found:", result.guests.length);

    // 3. Inspect the first guest's data structure
    if (result.guests.length > 0) {
      console.log("First Guest Data Mapping (Sample):");
      console.log(JSON.stringify(result.guests[0], null, 2));

      // Check specific keys that usually cause issues
      const sample = result.guests[0];
      console.log("Key Check - Name:", sample.Name);
      console.log("Key Check - Source (AirBnb_Personal):", sample.AirBnb_Personal);
      console.log("Key Check - Amount:", sample.Amount);
    } else {
      console.warn("⚠️ No guests found. Possible causes: Tab name mismatch, Month spelling mismatch, or Header row not found.");
    }

  } catch (e) {
    console.error("❌ CRITICAL ERROR during execution:");
    console.error("Message: " + e.message);
    console.error("Stack: " + e.stack);
  }

  console.log("--- DEBUG END ---");
}

/**
 * Helper to check all tab names in your spreadsheet
 * Run this if you get 'null' errors to verify your sheet naming.
 */
function listAllTabNames() {
  const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
  const sheets = ss.getSheets();
  console.log("Available Tabs in this Spreadsheet:");
  sheets.forEach(s => console.log("- " + s.getName()));
}

/**
 * Test function to verify the Airbnb email parser logic.
 * This runs locally in the Apps Script editor and logs the extracted data.
 */
function test_syncAirbnbEmails() {
  try {
    console.log(">>> [TEST] Starting Airbnb Sync Parser Test...");

    // 1. Target Sri Harsha's specific thread using the exact subject line from your PDF
    const searchQuery = 'from:automated@airbnb.com subject:"Reservation confirmed - Sri Harsha Kuchimanchi arrives May 30"';
    const threads = GmailApp.search(searchQuery, 0, 1);

    if (threads.length === 0) {
      console.warn(">>> [TEST] Could not find Sri Harsha's email thread. Ensure the email is in your inbox/trash and hasn't been permanently deleted.");
      return;
    }

    syncAirbnbEmails("2026");

    const message = threads[0].getMessages()[0];
    const subject = message.getSubject();
    const body = message.getPlainBody();

    console.log(">>> [TEST] Found Email Subject: " + subject);

    // 2. Execute the exact Regular Expressions used in your main function
    const currentYear = "2026";
    const subjectMatch = subject.match(/Reservation confirmed\s*-\s*(.*?)\s+arrives\s+(.*)/i);

    let guestName = "";
    let checkInStr = "";

    if (subjectMatch) {
      guestName = subjectMatch[1].trim();
      checkInStr = subjectMatch[2].trim() + `, ${currentYear}`;
    }

    const nightsMatch = body.match(/(\d+)\s*nights\s*room\s*fee/i) || body.match(/(\d+)\s*night/i);
    const nights = nightsMatch ? Number(nightsMatch[1]) : 1;

    const amountMatch = body.match(/You earn[\s\S]*?₹?\s*([\d,]+\.?\d*)/i);
    let finalAmount = "0";
    if (amountMatch) {
      let rawAmt = amountMatch[1].replace(/,/g, "");
      finalAmount = Math.round(parseFloat(rawAmt)).toString();
    }

    const guestsMatch = body.match(/(\d+)\s*adults/i) || body.match(/(\d+)\s*guest/i);
    const totalGuests = guestsMatch ? Number(guestsMatch[1]) : 2;

    // 3. Output the parsed results directly to the execution log
    console.log("----------------------------------------");
    console.log(">>> [TEST] PARSE RESULTS:");
    console.log("Parsed Name: ", guestName, (guestName === "Sri Harsha Kuchimanchi" ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("Parsed Check-in: ", checkInStr, (checkInStr === "May 30, 2026" ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("Parsed Nights: ", nights, (nights === 3 ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("Parsed Payout Amount: ", finalAmount, (finalAmount === "4586" ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("Parsed Guests: ", totalGuests, (totalGuests === 2 ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("----------------------------------------");

  } catch (err) {
    console.error(">>> [TEST] Test function execution crashed: " + err.message);
  }
}

/**
 * Utility function to clear the 'Vinyasa-Synced' label from Sri Harsha's email 
 * if you want to test the real button multiple times.
 */
function debug_removeSyncLabel() {
  const label = GmailApp.getUserLabelByName("Vinyasa-Synced");
  if (!label) {
    console.log("Label 'Vinyasa-Synced' does not exist yet.");
    return;
  }
  const threads = GmailApp.search('subject:"Sri Harsha Kuchimanchi arrives May 30" label:Vinyasa-Synced');
  threads.forEach(t => {
    t.removeLabel(label);
    console.log("Removed sync label from thread: " + t.getFirstMessageSubject());
  });
}

/**
 * TEST HARNESS: Run this function directly inside the Apps Script Editor 
 * to debug and inspect writeGuestDataRow behavior.
 */
function debug_writeGuestDataRow_Suite() {
  console.log("=== 🧪 STARTING writeGuestDataRow DEBUG SUITE 🧪 ===");

  // Choose a real or test year sheet tab present in your spreadsheet
  const testYear = "2026";

  // -----------------------------------------------------------------
  // TEST CASE 1: INSERT NEW RECORD (ADD MODE)
  // -----------------------------------------------------------------
  const addPayload = {
    name: "Test Guest Debugger",
    mobile: "9999988888",
    amount: 2500,
    guests: 3,
    checkIn: "2026-05-25", // Will convert to "May"
    days: 2,
    source: "Personal",
    floor: "Second Floor",
    ratings: "⭐⭐⭐⭐⭐",
    comments: "Created via automated GAS test runner suite execution.",
    year: testYear,
    rowIndex: "" // Blank for fresh additions
  };

  console.log("\n▶️ [TEST 1] Dispatching ADD payload for:", addPayload.name);
  try {
    const addResult = writeGuestDataRow("ADD", addPayload);
    console.log("✅ [TEST 1 SUCCESS] Backend returned response:", addResult);
  } catch (error) {
    console.error("❌ [TEST 1 FAILED] Execution crashed with error:", error.message);
  }

  // -----------------------------------------------------------------
  // TEST CASE 2: MODIFY EXISTING RECORD (EDIT MODE)
  // -----------------------------------------------------------------
  // We will pass the same name/mobile to update the entry we just made
  const editPayload = {
    name: "Test Guest Debugger",
    mobile: "9999988888",
    amount: 3200, // Modifying amount from 2500 to 3200
    guests: 3,
    checkIn: "2026-05-25",
    days: 3,      // Modifying nights from 2 to 3
    source: "Airbnb", // Modifying source from Personal to Airbnb
    floor: "Second Floor",
    ratings: "⭐⭐⭐⭐", // Modifying ratings
    comments: "Updated successfully via test runner execution script.",
    year: testYear,
    rowIndex: "" // Leaving blank to test our robust Name/Mobile fallback scanner
  };

  console.log("\n▶️ [TEST 2] Dispatching EDIT payload (Fallback Scan) for:", editPayload.name);
  try {
    const editResult = writeGuestDataRow("EDIT", editPayload);
    console.log("✅ [TEST 2 SUCCESS] Backend returned response:", editResult);
  } catch (error) {
    console.error("❌ [TEST 2 FAILED] Execution crashed with error:", error.message);
  }

  console.log("\n=== 🧪 DEBUG SUITE COMPLETION LOGS END ===");
}


/**
 * Test Harness to safely debug the Airbnb Sync Logic 
 * without modifying real Gmail threads or inbox state.
 */
function runDebugTests() {
  console.log("=== STARTING AIRBNB SYNC ENGINE DEBUG SUITE ===");

  // 1. Setup Mock Headers matching your actual sheet layout
  const mockHeaders = ["Month", "Name", "Guests", "Amount", "Check-in Date", "Days", "AirBnb\\Personal", "Floor", "Customer Ratings", "Comments"];

  // 2. TEST CASE 1: A Raw Cancellation Email (The one causing issues)
  const sampleCancelSubject = "Canceled: Reservation HMNXX8RCKX for Jun 15 – 17, 2026";
  const sampleCancelBody = "Hi Host, Reservation HMNXX8RCKX has been canceled by the guest. These dates are now open.";

  console.log("\n--- Testing Scenario B: Cancellation Parsing ---");
  debugIndividualPayload(sampleCancelSubject, sampleCancelBody, mockHeaders);

  // 3. TEST CASE 2: A Raw Review Email
  const sampleReviewSubject = "Sri Harsha left a 5-star review!";
  const sampleReviewBody = "Read on for a snapshot of what Sri Harsha loved about their stay.";

  console.log("\n--- Testing Scenario A: Review Parsing ---");
  debugIndividualPayload(sampleReviewSubject, sampleReviewBody, mockHeaders);

  console.log("\n=== DEBUG SUITE COMPLETE ===");
}

/**
 * Isolated logic tester to print exactly what your Regex matches
 */
function debugIndividualPayload(subject, body, headers) {
  const combinedTextToAnalyze = (subject + " " + body);
  const targetYear = "2026";

  // --- ISOLATED CANCELLATION TEST ---
  if (combinedTextToAnalyze.toLowerCase().includes("canceled:") || combinedTextToAnalyze.toLowerCase().includes("cancelled:")) {
    console.log("[CHECK] Detected Cancellation Trigger keyword.");

    const cancelMatch = combinedTextToAnalyze.match(/Reservation\s+([A-Z0-9]{10})/i) || combinedTextToAnalyze.match(/code\s+([A-Z0-9]{10})/i);
    const dateMatch = subject.match(/for\s+([A-Z][a-z]{2}\s+\d+)/i);

    let backupCheckInStr = "";
    if (dateMatch) {
      backupCheckInStr = dateMatch[1].trim() + `, ${targetYear}`;
    }

    console.log(`-> Extracted Code: ${cancelMatch ? cancelMatch[1] : "FAILED TO PARSE CODE"}`);
    console.log(`-> Extracted Backup Date: ${backupCheckInStr || "FAILED TO PARSE DATE"}`);
    return;
  }

  // --- ISOLATED REVIEW TEST ---
  if (combinedTextToAnalyze.toLowerCase().includes("left a") && combinedTextToAnalyze.toLowerCase().includes("review")) {
    console.log("[CHECK] Detected Review Trigger keywords.");

    const reviewMatch = combinedTextToAnalyze.match(/([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s+left\s+a\s+(\d+)-star\s+review/i);

    if (reviewMatch) {
      console.log(`-> Extracted Reviewer Name: ${reviewMatch[1]}`);
      console.log(`-> Extracted Rating: ${reviewMatch[2]} Stars`);
    } else {
      console.log("-> FAILED TO PARSE REVIEW REGEX");
    }
    return;
  }
}

/**
 * Test Harness: Executed manually in the editor to isolate and debug
 * calendar event creation, duplicate search rules, and 24-hour reminder triggers.
 */
function debugCalendarSyncWorkflow() {
  console.log("🚀 [DEBUG START] Initializing Calendar Synchronization Test...");
  
  // 1. Simulate a realistic booking payload bundle
  const mockPayload = {
    guestName: "Test Guest Vinay",
    checkInDate: "2026-06-15", // Simulates a future date execution string
    totalGuests: 3,
    platformType: "AirBnb",
    floorName: "Ground",
    nights: 2,
    notes: "Code: ABC123XYZ9. Automated debug check runner."
  };
  
  console.log("📋 [MOCK DATA] Payload configuration compiled:", JSON.stringify(mockPayload));
  
  // 2. Perform validation pre-checks inside the logs
  const parsedDateCheck = new Date(mockPayload.checkInDate);
  console.log(`📅 [DATE PARSING] Raw string '${mockPayload.checkInDate}' translated to Object: ${parsedDateCheck.toString()}`);
  
  if (isNaN(parsedDateCheck.getTime())) {
    console.error("❌ [DATE ERROR] System failed to resolve time coordinates for the incoming check-in date string.");
    return;
  }
  
  // 3. Verify target calendar accessibility
  const calendarName = "Vinyasa Nilaya";
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  console.log(`🔍 [CALENDAR ACCESSIBILITY] Searching for calendar named: '${calendarName}'`);
  
  if (calendars.length === 0) {
    console.warn(`⚠️ [CALENDAR WARNING] No calendar found named '${calendarName}'. The engine will use your primary default Google Account calendar instead.`);
  } else {
    console.log(`✅ [CALENDAR FOUND] Target calendar successfully bound. ID: ${calendars[0].getId()}`);
  }
  
  // 4. Fire the actual live function execution path
  console.log("⚙️ [EXECUTION] Dispatching payload variables straight to syncBookingToVinyasaCalendar...");
  
  try {
    const isSuccess = syncBookingToVinyasaCalendar(
      mockPayload.guestName,
      mockPayload.checkInDate,
      mockPayload.totalGuests,
      mockPayload.platformType,
      mockPayload.floorName,
      mockPayload.nights,
      mockPayload.notes
    );
    
    if (isSuccess) {
      console.log("🎉 [DEBUG SUCCESS] The sync engine executed flawlessly. Check your Google Calendar grid for June 15, 2026!");
    } else {
      console.error("❌ [DEBUG FAILED] Sync returned false. Read the execution logs above to trace structural bottlenecks.");
    }
    
  } catch (err) {
    console.error("💥 [CRITICAL CRASH] The calendar sync workflow thrown an unhandled exception:", err.toString());
  }
  
  console.log("🏁 [DEBUG END] Test sequence completed.");
}

