var ID_GUESTS_LIST = "1Puw0OezY18OWFt8wtwzv5BFxcJw314Hfov5GZMUXCbk";

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
    
    // -----------------------------------------------------------------
    // PHASE 1: COMPUTE CONSOLIDATED LIFETIME REVENUE ACROSS ALL YEARS
    // -----------------------------------------------------------------
    let lifetimeTotalRevenue = 0;
    let lifetimeTotalCheckIns = 0;

    sheets.forEach(sheet => {
      const sheetName = sheet.getName().trim();
      
      // Isolate sheets that match a 4-digit year pattern (e.g., "2024", "2026")
      if (/^\d{4}$/.test(sheetName)) {
        const sheetData = sheet.getDataRange().getValues();
        const headerRowIdx = sheetData.findIndex(row => row.includes("Name") || row.includes("Amount"));
        
        if (headerRowIdx !== -1) {
          const sheetHeaders = sheetData[headerRowIdx];
          const amountIdx = sheetHeaders.indexOf("Amount");
          const nameIdx = sheetHeaders.indexOf("Name");
          
          if (amountIdx !== -1) {
            // Unpack rows beneath the header row
            const sheetRows = sheetData.slice(headerRowIdx + 1);
            
            sheetRows.forEach(row => {
              const nameVal = row[nameIdx] ? row[nameIdx].toString().trim() : "";
              
              // Standard safety exclusions to match your row filter baseline
              if (!nameVal || nameVal === "Total" || nameVal === "No Guests" || nameVal === "") return;
              
              let amtStr = (row[amountIdx] || "0").toString().replace(/[₹,]/g, "");
              let amtNum = Number(amtStr) || 0;
              lifetimeTotalRevenue += amtNum;
              lifetimeTotalCheckIns++; // Increment global record counter
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
    // Locate header row using "Name" or "Month"
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

      // Filter out empty rows, summary rows, or "No Guests" placeholders
      if (!name || name === "Total" || name === "No Guests" || name === "") return false;

      const rowMonth = month ? month.toString().trim() : "";
      return !filterMonth || rowMonth === filterMonth;
    }).map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        let key = header.toString();

        // 1. Unify Source Column
        if (key === "AirBnb\\Personal" || key === "Source") {
          key = "Source";
        } 
        // 2. Unify Floor Column Name (Adjust "Floor" to match your exact spreadsheet column header text)
        else if (key === "Floor") {
          key = "Floor";
        } 
        else {
          key = key.replace(/\\|\s|-/g, "_");
        }

        let value = row[i];

        // Safety: Convert Date objects to strings for the frontend
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
        mobileRaw = mobileRaw.slice(-10); // Extract last 10 digits
      }
      obj['WhatsApp_Num'] = mobileRaw;
      obj['isVerified'] = mobileRaw.length === 10;

      return obj;
    });

    return {
      guests: filteredData,
      summary: {
        totalRevenue: totalRevenue.toLocaleString('en-IN', { maximumFractionDigits: 0 }).replace("INR", "").trim(),
        count: guestCount,
        period: filterMonth ? `${filterMonth} ${targetTab}` : targetTab,
        lifetimeRevenue: lifetimeTotalRevenue.toLocaleString('en-IN', { maximumFractionDigits: 0 }),
        lifetimeCount: lifetimeTotalCheckIns // 
      }
    };
  } catch (err) {
    console.error("Dashboard Sync Error: " + err.message);
    return { guests: [], summary: { totalRevenue: "Error", count: 0, period: "Sheet Error", lifetimeRevenue: "Error" } };
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

        let currentMonthStr = checkInStr.split(" ")[0]; 

        const incomingCheckInDate = new Date(checkInStr);
        const incomingCheckInTime = !isNaN(incomingCheckInDate.getTime()) ? incomingCheckInDate.getTime() : 0;

        // Parse verification code to anchor upcoming cancellation logic lookups
        const codeMatch = body.match(/Reservation\s*code\s*([A-Z0-9]{10})/i) || body.match(/Confirmation\s*code\s*([A-Z0-9]{10})/i);
        const confirmationCode = codeMatch ? codeMatch[1].trim() : "NOT_FOUND";
        const savedComment = confirmationCode !== "NOT_FOUND" ? `Code: ${confirmationCode}. Automated Gmail Sync Engine.` : "Automated Gmail Sync Engine.";

        let newRowData = new Array(headers.length).fill("");
        newRowData[headers.indexOf("Month")] = currentMonthStr;
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
function writeGuestDataRow(mode, payload) {
  // --- FIRST-LINE GATEKEEPER VALIDATION ---
  if (!payload.checkIn || payload.checkIn.toString().trim() === "") {
    throw new Error("Transaction Denied: Check-in Date is a mandatory field and cannot be left blank.");
  }

  const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
  let sheet = ss.getSheetByName(payload.year.trim());
  
  if (!sheet) {
    sheet = ss.insertSheet(payload.year.trim());
    sheet.appendRow(["Month","Name","Guests", "Amount","Check-in Date","Days", "Source","AirBnb\Personal","Floor", "Mobile",    "Customer Ratings", "Comments"]);
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
  let incomingCheckInTime = 0;
  let standardizedCheckInString = payload.checkIn; // Default fallback
  
  const monthsArray = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const shortMonthsArray = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  
  const dParsed = parseDateSecurely(payload.checkIn);
  
  if (!dParsed) {
    throw new Error("Transaction Denied: Invalid Date format provided for Check-in Date.");
  }
  
  generatedMonthLabel = monthsArray[dParsed.getMonth()];
  incomingCheckInTime = dParsed.getTime();
  
  // Format the raw UI value (YYYY-MM-DD) into the Airbnb string format: "May 23, 2026"
  standardizedCheckInString = `${shortMonthsArray[dParsed.getMonth()]} ${dParsed.getDate()}, ${dParsed.getFullYear()}`;

  // --- 3. EXECUTE EDIT CLEANUP (Remove row to allow re-sorting) ---
  if (mode === "EDIT") {
    let targetRowNumber = parseInt(payload.rowIndex);
    
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
      console.log(`>>> [EDIT RELOCATION] Dropping Row [${targetRowNumber}] to calculate its new sorted position.`);
      sheet.deleteRow(targetRowNumber);
      SpreadsheetApp.flush(); 
    } else {
      throw new Error("Could not find the target row to edit.");
    }
  }

  // --- 4. MAP DATA VECTOR ARRAY ---
 /* let newRow = new Array(headers.length).fill("");
  if (mapping.Name !== -1) newRow[mapping.Name] = payload.name;
  if (mapping.Amount !== -1) newRow[mapping.Amount] = payload.amount;
  if (monthColIdx !== -1) newRow[monthColIdx] = generatedMonthLabel;
  if (mapping.Source !== -1) newRow[mapping.Source] = payload.source;
  if (mapping.Floor !== -1) newRow[mapping.Floor] = payload.floor;
  if (mapping.Mobile !== -1) newRow[mapping.Mobile] = cleanMobile;
  
  // CRITICAL FIX: Save the standardized string layout "May 23, 2026" instead of "2026-05-23"
  if (mapping.CheckIn !== -1) newRow[mapping.CheckIn] = standardizedCheckInString;
  
  if (mapping.Days !== -1) newRow[mapping.Days] = payload.days;
  if (mapping.Guests !== -1) newRow[mapping.Guests] = payload.guests;
  if (mapping.Ratings !== -1) newRow[mapping.Ratings] = payload.ratings;
  if (mapping.Comments !== -1) newRow[mapping.Comments] = payload.comments;*/

  // --- 4. MAP DATA VECTOR ARRAY ---
  let newRow = new Array(headers.length).fill("");
  if (mapping.Name !== -1) newRow[mapping.Name] = payload.name;
  if (mapping.Amount !== -1) newRow[mapping.Amount] = payload.amount;
  if (monthColIdx !== -1) newRow[monthColIdx] = generatedMonthLabel;
  
  // Clean Source matching data validation constraint exactly ("AirBnb" or "Personal")
  if (mapping.Source !== -1) {
    let cleanSource = payload.source.toString().trim().toLowerCase();
    newRow[mapping.Source] = cleanSource.includes("airbnb") ? "AirBnb" : "Personal";
  }
  
  // Clean Floor matching data validation constraint exactly ("Ground" or "Second")
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

  // --- 5. TIMELINE FIXED INSERTION ROUTINE ---
  let currentSheetData = sheet.getDataRange().getValues();
  let lastRowWithContent = sheet.getLastRow();
  
  let insertionRowIndex = lastRowWithContent; 
  let foundInsertionSpot = false;

  for (let i = currentSheetData.length - 1; i >= 1; i--) {
    let rowCheckInVal = currentSheetData[i][mapping.CheckIn];
    
    if (rowCheckInVal) {
      let rowDate = parseDateSecurely(rowCheckInVal);
      
      if (rowDate) {
        let rowDateTime = rowDate.getTime();
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

  console.log(`>>> [FORM TARGET COMPLIANCE] Inserting ${payload.name} safely after row: [${insertionRowIndex}]`);

  sheet.insertRowsAfter(insertionRowIndex, 1);
  
  let templateRow = (insertionRowIndex === 1) ? 2 : insertionRowIndex; 
  let templateRange = sheet.getRange(templateRow, 1, 1, headers.length);
  let targetRange = sheet.getRange(insertionRowIndex + 1, 1, 1, headers.length);
  
  templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  targetRange.setValues([newRow]);
  SpreadsheetApp.flush();

  return "SUCCESS";
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
    jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11,
    january:0, february:1, march:2, april:3, may:4, june:5, july:6, august:7, september:8, october:9, november:10, december:11
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
    jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11,
    january:0, february:1, march:2, april:3, may:4, june:5, july:6, august:7, september:8, october:9, november:10, december:11
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
    jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11,
    january:0, february:1, march:2, april:3, may:4, june:5, july:6, august:7, september:8, october:9, november:10, december:11
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
