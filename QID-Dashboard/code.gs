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

        // --- KEY MAPPING LOGIC ---
        if (key === "AirBnb\\Personal") {
          key = "Source";
        } else {
          // Converts "Check-in Date" to "Check_in_Date", "Customer Ratings" to "Customer_Ratings"
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
        // --- ADDED NEW CORE CHARACTERISTICS ---
        lifetimeRevenue: lifetimeTotalRevenue.toLocaleString('en-IN', { maximumFractionDigits: 0 })
      }
    };
  } catch (err) {
    console.error("Dashboard Sync Error: " + err.message);
    return { guests: [], summary: { totalRevenue: "Error", count: 0, period: "Sheet Error", lifetimeRevenue: "Error" } };
  }
}

/**** Sync with Airbnb bookings feature *****/
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

    // Target the specific Airbnb email subject format
    const query = `from:automated@airbnb.com subject:"Reservation confirmed" -label:${labelName}`;
    const threads = GmailApp.search(query, 0, 10);
    let newBookingsCount = 0;

    threads.forEach(thread => {
      const messages = thread.getMessages();
      let processedThread = false;

      messages.forEach(message => {
        const subject = message.getSubject();
        const body = message.getPlainBody();
        
        // --- 1. PARSE SUBJECT FOR NAME & ARRIVAL ---
        // Matches: "Reservation confirmed - Sri Harsha Kuchimanchi arrives May 30"
        const subjectMatch = subject.match(/Reservation confirmed\s*-\s*(.*?)\s+arrives\s+(.*)/i);
        
        let guestName = "";
        let checkInStr = "";
        
        if (subjectMatch) {
          guestName = subjectMatch[1].trim(); 
          checkInStr = subjectMatch[2].trim() + `, ${currentYear}`; // Outputs: "May 30, 2026"
        } else {
          return; // Skip if subject doesn't match standard confirmation format
        }
        
        if (!guestName) return;

        // --- 2. PARSE BODY FOR METRICS (Based exactly on PDF text) ---
        
        // Target: "3 nights" or "3 nights room fee"
        const nightsMatch = body.match(/(\d+)\s*nights\s*room\s*fee/i) || body.match(/(\d+)\s*night/i);
        const nights = nightsMatch ? Number(nightsMatch[1]) : 1;
        
        // Target: "You earn" followed closely by the actual payout code line "₹4,586.16"
        // This regex skips "Total (INR)" and grabs your actual net payout
        const amountMatch = body.match(/You earn[\s\S]*?₹?\s*([\d,]+\.?\d*)/i);
        let finalAmount = "0";
        if (amountMatch) {
          // Removes commas and extracts whole number integer for clean math handling
          let rawAmt = amountMatch[1].replace(/,/g, ""); 
          finalAmount = Math.round(parseFloat(rawAmt)).toString(); // Outputs: "4586"
        }
        
        // Target: "2 adults" or "1 guest"
        const guestsMatch = body.match(/(\d+)\s*adults/i) || body.match(/(\d+)\s*guest/i);
        const totalGuests = guestsMatch ? Number(guestsMatch[1]) : 2;

        // --- 3. MAP TO SPREADSHEET ROW WITH FORMATTING PRESERVATION ---
        let newRowData = new Array(headers.length).fill("");
        let currentMonthStr = checkInStr.split(" ")[0]; // Extracts "May"
        
        newRowData[headers.indexOf("Month")] = currentMonthStr;
        newRowData[headers.indexOf("Name")] = guestName;
        newRowData[headers.indexOf("Guests")] = totalGuests;
        newRowData[headers.indexOf("Amount")] = finalAmount;
        newRowData[headers.indexOf("Check-in Date")] = checkInStr;
        newRowData[headers.indexOf("Days")] = nights;
        newRowData[headers.indexOf("AirBnb\\Personal")] = "AirBnb";
        newRowData[headers.indexOf("Floor")] = "Ground"; 
        newRowData[headers.indexOf("Comments")] = "Automated Gmail Sync Engine.";
        
        // --- FIXED FORWARD INSERTION LOGIC ---
        // Find the absolute last row that currently has text content
        let lastRowWithContent = sheet.getLastRow();
        let targetRow = lastRowWithContent + 1;
        
        // 1. Insert a clean row right below your last template row
        sheet.insertRowsAfter(lastRowWithContent, 1);
        
        // 2. Reference the template row above it to copy the styling styles
        let templateRange = sheet.getRange(lastRowWithContent, 1, 1, headers.length);
        let targetRange = sheet.getRange(targetRow, 1, 1, headers.length);
        
        // 3. Copy font family, background colors, alignments, borders, and number formats
        templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        
        // 4. Safely set the clean raw data values into the beautifully styled row
        targetRange.setValues([newRowData]);
        
        newBookingsCount++;
        processedThread = true;

      });

      // --- 4. SIGN OFF AND LABEL THREAD ---
      if (processedThread) {
        thread.addLabel(syncLabel); 
        thread.markRead();          
      }
    });
    
    return `Sync Complete! Successfully added ${newBookingsCount} booking(s).`;
    
  } catch (err) {
    console.error("Parser tracking fail: " + err.message);
    throw new Error("Sync processing aborted: " + err.message);
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
