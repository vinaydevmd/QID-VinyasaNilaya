var ID_GUESTS_LIST = "1Puw0OezY18OWFt8wtwzv5BFxcJw314Hfov5GZMUXCbk";

var TAB_2024 = "2024"

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
 * Fetches guest data and calculates totals based on filters
 */
/**
 * Fetches guest data and summaries based on the specific column headers:
 * Year, Month, Name, NoOfGuests, Amount, Check-in Date, Days, AirBnb\Personal, Floor, Mobile, Customer Ratings, Comments
 */
function getDashboardData(filterYear, filterMonth) {
  try {
    const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
    const targetTab = filterYear || "2026";
    let sheet = ss.getSheetByName(targetTab) || ss.getSheets()[0];

    const data = sheet.getDataRange().getValues();
    // Locate header row using "Name" or "Month"
    const headerRowIndex = data.findIndex(row => row.includes("Name") || row.includes("Month"));

    if (headerRowIndex === -1) {
      return { guests: [], summary: { totalRevenue: "₹0", count: 0, period: "No Headers Found" } };
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
        totalRevenue: totalRevenue.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 }),
        count: guestCount,
        period: filterMonth ? `${filterMonth} ${targetTab}` : targetTab
      }
    };
  } catch (err) {
    console.error("Dashboard Sync Error: " + err.message);
    return { guests: [], summary: { totalRevenue: "Error", count: 0, period: "Sheet Error" } };
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
