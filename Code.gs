/**
 * Global Configuration
 */
const FOLDER_ID = '1cmRFirWeg_tHFbZ9VS-E0Gz80SHHSsIU9lu5jV_GBKk'; 
const SHEET_NAME = 'QID-Verified';

const IMAGEKIT_PRIVATE_KEY = "private_mB1ln5VP44E0u/PFyJ8QEG8dGR8="; // Note: must end with a colon for Basic Auth
const IMAGEKIT_FOLDER = "/QID/"; // The folder in your ImageKit media library

// DigiLocker Credentials
const CLIENT_ID = 'YOUR_DIGILOCKER_CLIENT_ID';
const REDIRECT_URI = ScriptApp.getService().getUrl(); 

function doGet(e) {
  if (e.parameter.code) {
    return handleDigiLockerCallback(e.parameter.code);
  }
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Guest Verification | Secure Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl('https://ik.imagekit.io/h87o83ayxm/Icons/Icon03_png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDigiLockerLoginUrl() {
  if (CLIENT_ID === 'YOUR_DIGILOCKER_CLIENT_ID') return "#";
  const state = Utilities.getUuid();
  return "https://test.digitallocker.gov.in/public/oauth2/1/authorize" + 
          "?response_type=code&client_id=" + CLIENT_ID + 
          "&redirect_uri=" + encodeURIComponent(REDIRECT_URI) + "&state=" + state;
}

function handleDigiLockerCallback(authCode) {
  // Simply closes the popup or shows success
  return HtmlService.createHtmlOutput("<script>window.close();</script><h3 style='font-family:sans-serif; text-align:center; margin-top:20%; color:#059669;'>✓ Verified. You may close this window.</h3>");
}

function processSubmission(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found.");

  const now = new Date();
  
  // 1. Determine Status & Integer Row Number
  const isUpdate = !!data.rowNumber;
  const checkInStatus = isUpdate ? "Returning Guest - Updated" : "New Guest - Verified";
  const rowTarget = isUpdate ? parseInt(data.rowNumber, 10) : null;

  // 2. Calculate SlNo
  let finalSlNo = data.slNo; 
  if (!isUpdate) {
    const lastRow = sheet.getLastRow();
    finalSlNo = 1;
    if (lastRow > 1) {
      const lastVal = sheet.getRange(lastRow, 1).getValue();
      finalSlNo = (typeof lastVal === 'number') ? lastVal + 1 : lastRow;
    }
  }

  // 3. ImageKit Upload
  const datePrefix = `${now.getFullYear()}_${data.idNumber}_${String(finalSlNo).padStart(2, '0')}`;
  const selfieUrl = uploadToImageKit(data.selfieBase64, `${datePrefix}_Selfie`);

  // 4. Construct Full 20-Column Row (Must be exactly 20 elements)
  const rowData = new Array(20).fill(""); // Initialize empty array of 20
  
  rowData[0] = finalSlNo;                 // Col 1
  rowData[1] = now;                       // Col 2
  rowData[2] = data.name || "";           // Col 3
  rowData[3] = data.idType || "";         // Col 4
  rowData[4] = data.idNumber || "";       // Col 5
  rowData[5] = data.whatsapp || "";       // Col 6
  rowData[6] = data.purpose || "";        // Col 7
  rowData[7] = data.city || "";           // Col 8
  rowData[8] = data.emergencyName || "";  // Col 9
  rowData[9] = data.emergencyPhone || ""; // Col 10
  // Cols 11-15 (Stay Details) remain empty strings ""
  rowData[15] = data.idFrontUrl || "";    // Col 16
  rowData[16] = data.idBackUrl || "";     // Col 17
  rowData[17] = selfieUrl;                // Col 18
  rowData[18] = "Yes";                    // Col 19
  rowData[19] = checkInStatus;            // Col 20

  // 5. Save to Sheet
  if (isUpdate && rowTarget) {
    // PRESERVE OLD IDs: Fetch current values from sheet
    const existingValues = sheet.getRange(rowTarget, 1, 1, 20).getValues()[0];
    
    // If current submission has no new ID URL, use the one already in the sheet
    if (!rowData[15]) rowData[15] = existingValues[15]; 
    if (!rowData[16]) rowData[16] = existingValues[16]; 

    sheet.getRange(rowTarget, 1, 1, 20).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  
  return { 
    success: true, 
    slNo: finalSlNo, 
    status: checkInStatus 
  };
}

/*function processSubmission(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found.");

  // 1. Calculate Total Stay (Nights)
  let totalNights = "";
  if (data.checkIn && data.checkOut) {
    const d1 = new Date(data.checkIn);
    const d2 = new Date(data.checkOut);
    
    // Calculate difference in milliseconds and convert to days
    const diffTime = d2.getTime() - d1.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
    
    totalNights = diffDays > 0 ? diffDays + (diffDays === 1 ? " Night" : " Nights") : "0 Nights";
  }

  // 2. Calculate SlNo (Auto-increment)
  const lastRow = sheet.getLastRow();
  let finalSlNo = 1;
  if (lastRow > 1) {
    const lastVal = sheet.getRange(lastRow, 1).getValue();
    finalSlNo = (typeof lastVal === 'number') ? lastVal + 1 : lastRow;
  }

  // 3. Generate the dynamic filename prefix for ImageKit
  const now = new Date();
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const datePrefix = `${monthNames[now.getMonth()]}_${now.getFullYear()}_${String(finalSlNo).padStart(2, '0')}`;

  // 4. ImageKit Uploads
  const idFrontUrl = uploadToImageKit(data.idFrontBase64, `${datePrefix}_GovtID-Front`);
  const idBackUrl = data.idBackBase64 ? uploadToImageKit(data.idBackBase64, `${datePrefix}_GovtID-Back`) : "";
  const selfieUrl = uploadToImageKit(data.selfieBase64, `${datePrefix}_Selfie`);

  // 5. Map to your 19 Columns
  const formData = [
    finalSlNo,                      // Col 1: SlNo
    now,                            // Col 2: Timestamp
    data.name || "",                // Col 3: Name
    data.idType || "",              // Col 4: IdType
    data.idNumber || "",            // Col 5: IDNo
    data.whatsapp || "",            // Col 6: Phone
    data.purpose || "",             // Col 7: Purpose Of Travel
    data.city || "",                // Col 8: Arriving City
    data.emergencyName || "",       // Col 9: Emergency Contact Name
    data.emergencyPhone || "",      // Col 10: Emergency Contact No
    data.checkIn || "",             // Col 11: CheckIn
    data.checkOut || "",            // Col 12: CheckOut
    data.adults || 0,               // Col 13: Adults
    data.kids || 0,                 // Col 14: Kids
    totalNights,                    // Col 15: Total Stay (Calculated Above)
    idFrontUrl,                     // Col 16: Govt-ID-Front URL
    idBackUrl,                      // Col 17: Govt-ID-Back URL
    selfieUrl,                      // Col 18: Selfie URL
    data.consent ? "Yes" : "No"     // Col 19: Terms Accepted
  ];

  sheet.appendRow(formData);
  
  return { success: true, slNo: finalSlNo };
}*/

/**
 * Helper: Uploads Base64 string to ImageKit
 */
/**
 * Uploads a base64 image to ImageKit with overwrite permissions.
 * @param {string} base64Data - The full data URI (e.g. data:image/jpeg;base64,...)
 * @param {string} fileName - The desired name (Month_Year_SlNo_Type)
 * @return {string} The URL of the uploaded file
 */
function uploadToImageKit(base64Data, fileName) {
  // 1. Extract the raw base64 content
  const base64Content = base64Data.split(',')[1];

  // 2. Prepare the payload (similar to your JS FormData example)
  // UrlFetchApp handles multipart/form-data automatically when the payload is an object
  const payload = {
    'file': base64Content,
    'fileName': fileName,
    'folder': IMAGEKIT_FOLDER, // Or your specific ImageKit path
    'useUniqueFileName': "false",
    'overwriteFile': "true" // Ensures that if you re-upload, it replaces the old one
  };

  const authHeader = "Basic " + Utilities.base64Encode(IMAGEKIT_PRIVATE_KEY + ":");

  // 3. Configure the Request
  // Note: IMAGEKIT_PRIVATE_KEY should be your private key. 
  // ImageKit requires the colon ":" at the end for Basic Auth.
  const options = {
    'method': 'post',
    'headers': {
      'Authorization': authHeader,
      'Accept': 'application/json'
    },
    'payload': payload,
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch("https://upload.imagekit.io/api/v1/files/upload", options);
    const result = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      throw new Error("ImageKit Error: " + (result.message || response.getContentText()));
    }

    return result.url; // Returns the optimized CDN URL
  } catch (err) {
    console.error("Upload failed for " + fileName + ": " + err.toString());
    throw err;
  }
}
/**
 * Uses Google Cloud Vision API to extract text from image
 */

function extractTextFromImage(base64Data) {
  const API_KEY = 'AIzaSyBJu27q1YjARED4wFPCYV1h2AyZNRLygVo'; // Your key
  const url = `https://vision.googleapis.com/v1/images:annotate?key=${API_KEY}`;

  let base64Image = base64Data;
  if (base64Image.indexOf(",") !== -1) {
    base64Image = base64Image.split(",")[1];
  }

  const payload = {
    requests: [{
      image: { content: base64Image },
      features: [{ type: "TEXT_DETECTION" }]
    }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  // DEBUG: Check the full response if it fails
  if (result.responses && result.responses[0].error) {
    throw new Error("Vision API Error: " + result.responses[0].error.message);
  }

  if (result.responses && result.responses[0].fullTextAnnotation) {
    return result.responses[0].fullTextAnnotation.text;
  } else {
    throw new Error("Vision API could not detect any text. The image might be too blurry or dark.");
  }
}

/**
 * Enhanced parsing logic to handle variations in OCR text
 */
function parseAadhaarData(rawText) {
  const lines = rawText.split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 2);
  
  // Improved Regex: Handles 12 digits with spaces or no spaces
  const idMatch = rawText.match(/(\d{4}\s\d{4}\s\d{4})|(\d{12})/);
  
  let detectedName = "Not found";
  // Consistency fix: Use 'noiseKeywords' everywhere
  const noiseKeywords = ["GOVERNMENT", "INDIA", "FATHER", "DOB", "MALE", "FEMALE", "ADDRESS", "ENROLLMENT", "UNIQUE", "HELP", "YEAR"];

  for (let line of lines) {
    // Strip non-English characters
    let englishOnlyLine = line.replace(/[^\x00-\x7F]/g, "").trim();
    const upperLine = englishOnlyLine.toUpperCase();
    
    // Check if line contains any noise words
    const isNoise = noiseKeywords.some(word => upperLine.includes(word));
    const hasNumbers = /\d/.test(englishOnlyLine);

    // Logic: First line with >5 chars, at least 2 words, no numbers, no noise
    if (englishOnlyLine.length > 5 && !isNoise && !hasNumbers && englishOnlyLine.split(/\s+/).length >= 2) {
      detectedName = englishOnlyLine;
      break; 
    }
  }

  return {
    name: detectedName,
    idNumber: idMatch ? idMatch[0] : "",
    raw: rawText 
  };
}
/******** Parse Voter ID *********************/
function parseVoterIDData(rawText) {
  const lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 2);
  
  // EPIC Pattern: 3 Letters followed by 7 Digits (e.g., WBH4159588)
  const idMatch = rawText.match(/[A-Z]{3}\d{7}/i);
  
  let detectedName = "Not found";
  
  // 1. PRIMARY SEARCH: Look specifically for the "Elector's Name" label
  for (let i = 0; i < lines.length; i++) {
    const upperLine = lines[i].toUpperCase();
    if (upperLine.includes("ELECTOR'S NAME") || upperLine.includes("ELECTORS NAME")) {
      // Extract everything after the colon or dash
      let namePart = lines[i].split(/[:|-]/).pop().trim();
      
      // If the name is on the NEXT line instead of the same line
      if (namePart.length < 3 && i + 1 < lines.length) {
        namePart = lines[i+1].trim();
      }
      
      detectedName = namePart.replace(/[^\x00-\x7F]/g, "").trim();
      if (detectedName.length > 3) break;
    }
  }

  // 2. FALLBACK SEARCH: Use the standard noise-keyword loop if label search failed
  if (detectedName === "Not found") {
    const noiseKeywords = ["ELECTION", "COMMISSION", "INDIA", "IDENTITY", "CARD", "ELECTOR", "FATHER", "MOTHER", "HUSBAND", "SEX", "MALE", "FEMALE", "DATE", "BIRTH", "ADDRESS", "CONSTITUENCY"];

    for (let line of lines) {
      let englishOnlyLine = line.replace(/[^\x00-\x7F]/g, "").trim();
      const upperLine = englishOnlyLine.toUpperCase();
      
      const isNoise = noiseKeywords.some(word => upperLine.includes(word));
      const hasNumbers = /\d/.test(englishOnlyLine);

      if (englishOnlyLine.length > 3 && !isNoise && !hasNumbers && englishOnlyLine.split(/\s+/).length >= 2) {
        detectedName = englishOnlyLine;
        break; 
      }
    }
  }

  return {
    name: detectedName,
    idNumber: idMatch ? idMatch[0].toUpperCase() : "",
    raw: rawText 
  };
}

/**
 * Universal Parser for Indian Driving Licenses (All States)
 * Targets MoRTH standard DL formats and common English labels.
 */

function parseDrivingLicenseData(combinedText) {
  const lines = combinedText.split('\n').map(l => l.trim()).filter(l => l.length > 1);
  
  // 1. Enhanced DL Number Regex
  const dlPattern = /([A-Z0-4]{2}\d{2})[\s\-]?(\d{4})[\s\-]?(\d{5,7})/i;
  const dlMatch = combinedText.match(dlPattern);
  let idNumber = dlMatch ? (dlMatch[1] + " " + dlMatch[2] + dlMatch[3]).toUpperCase() : "";

  // 2. High-Reliability Name Extraction
  let detectedName = "Not found";
  const noiseKeywords = ["TRANSPORT", "DATE", "BIRTH", "D.O.B", "ISSUE", "EXPIRY", "VALID", "ADDRESS", "S/O", "D/O", "W/O", "FATHER", "HUSBAND", "COV", "DOI", "INDIA", "CARD"];

  for (let i = 0; i < lines.length; i++) {
    const upperLine = lines[i].toUpperCase();
    
    // FUZZY ANCHOR CHECK: Catch "NAME", "N4ME", "NAM E", or "HOLDER"
    const isNameLine = /N[A-Z0-4\s]{2,3}E|HOLDER/i.test(upperLine);
    
    if (isNameLine) {
      let potentialName = "";
      
      // Strategy A: Check after a colon on the same line
      if (lines[i].includes(":")) {
        potentialName = lines[i].split(":").pop().trim();
      }

      // Strategy B: If Strategy A failed, check the next 3 lines (Deep Search)
      // This is crucial for vertical cards like Yashaswini's
      let searchOffset = 1;
      while (potentialName.length < 3 && searchOffset <= 3 && (i + searchOffset) < lines.length) {
        const candidate = lines[i + searchOffset].trim();
        
        // Validation: Ensure the candidate isn't just a date or noise
        const isDate = /\d{2}[\/\-]\d{2}[\/\-]\d{4}/.test(candidate);
        const isNoise = noiseKeywords.some(word => candidate.toUpperCase().includes(word));
        const hasNumbers = /\d/.test(candidate);

        if (candidate.length > 3 && !isDate && !isNoise && !hasNumbers) {
          potentialName = candidate;
        }
        searchOffset++;
      }

      // Final Clean-up
      const cleanName = potentialName.replace(/[^\x00-\x7F]/g, "").replace(/^[:\s\-]+/, "").trim();

      if (cleanName.length > 3) {
        detectedName = cleanName.toUpperCase();
        break; 
      }
    }
  }

  return { name: detectedName, idNumber: idNumber, raw: combinedText };
}

/**
 * Specialized parser for Indian Passports using regex for labels and MRZ patterns
 */
function parsePassportData(rawText) {
  const lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 1);
  
  // Passport Pattern: 1 Letter + 7 Digits (Matches C3075733 and R9091871)
  const idMatch = rawText.match(/[A-Z]\d{7}/i);
  
  let surname = "";
  let givenName = "";
  
  // 1. TARGETED EXTRACTION: Search for Surname and Given Name labels
  const headers = ["SURNAME", "GIVEN NAME", "NAME", "दिया गया नाम", "उपनाम"];

  for (let i = 0; i < lines.length; i++) {
    const upperLine = lines[i].toUpperCase();

    // Find Surname
    if (upperLine.includes("SURNAME") || upperLine.includes("उपनाम")) {
      let val = lines[i].split(/[:/|-]/).pop().trim();
      if (val.length < 3 && i + 1 < lines.length) val = lines[i + 1].trim();
      
      if (!headers.some(h => val.toUpperCase().includes(h))) {
        surname = val.replace(/[^\x00-\x7F]/g, "").trim();
      }
    }

    // Find Given Name
    if (upperLine.includes("GIVEN NAME") || upperLine.includes("दिया गया नाम")) {
      let val = lines[i].split(/[:/|-]/).pop().trim();
      if (val.length < 3 && i + 1 < lines.length) val = lines[i + 1].trim();
      
      if (!headers.some(h => val.toUpperCase().includes(h))) {
        givenName = val.replace(/[^\x00-\x7F]/g, "").trim();
      }
    }
  }

  // Initial Full Name assembly
  let fullName = (givenName + " " + surname).trim();

  // 2. MRZ FALLBACK: Reconstructs name from the machine-readable line at the bottom
  // Specifically updated to fix the "PKIND" bug by aggressively stripping prefixes
  if (!fullName || fullName.length < 5 || fullName.toUpperCase().includes("SURNAME")) {
    const mrzLine = lines.find(l => l.startsWith("P<") || l.includes("<<"));
    
    if (mrzLine) {
      // FIX: Aggressively strip the Passport type and Country code (e.g., P<IND, PKIND, P<)
      const cleanMRZ = mrzLine.replace(/^P.[A-Z]{3}/i, "").replace(/^P</i, "");
      const parts = cleanMRZ.split("<<");
      
      if (parts.length >= 2) {
        // MRZ Format: SURNAME << GIVEN < NAME
        const mrzSurname = parts[0].replace(/</g, " ").trim();
        const mrzGiven = parts[1].replace(/</g, " ").trim();
        
        // Final sanity check to remove any lingering prefix fragments surviving the split
        const finalSurname = mrzSurname.replace(/^[P|I|N|D|K]{1,5}\s+/i, "").trim();
        
        fullName = (mrzGiven + " " + finalSurname).trim();
      }
    }
  }

  return {
    name: fullName.toUpperCase() || "Not found",
    idNumber: idMatch ? idMatch[0].toUpperCase() : "",
    raw: rawText 
  };
}

/**
 * Unified OCR Flow to handle single or dual image inputs
 */
function executeOcrFlow(frontBase64, backBase64, idType) {
  try {
    // 1. Extract text from both images
    const frontText = frontBase64 ? extractTextFromImage(frontBase64) : "";
    const backText = backBase64 ? extractTextFromImage(backBase64) : "";
    
    // 2. Merge text into a single search pool
    const combinedRawText = `${frontText}\n${backText}`.trim();
    
    if (combinedRawText.length === 0) {
      throw new Error("No text detected. Please ensure the photos are clear and well-lit.");
    }

    const upperText = combinedRawText.toUpperCase();
    
    // 3. Broad Validation
    const idKeywords = ["GOVERNMENT", "INDIA", "INCOME TAX", "ELECTION", "DRIVING", "LICENSE", "ID", "CARD", "UNIQUE", "PASSPORT", "REPUBLIC"];
    const hasIdKeywords = idKeywords.some(keyword => upperText.includes(keyword));

    if (!hasIdKeywords) {
      throw new Error("This doesn't look like a valid Government ID. Please upload clear photos of your original card.");
    }

    // 4. Routing to Parsers
    // All parsers now receive the combined text from both sides
    switch(idType) {
      case "Aadhaar":
        return parseAadhaarData(combinedRawText);
      case "VoterID":
        return parseVoterIDData(combinedRawText);
      case "DL":
        return parseDrivingLicenseData(combinedRawText);
      case "Passport":
        return parsePassportData(combinedRawText);
      default:
        throw new Error("Unsupported ID type selected.");
    }

  } catch (e) {
    console.error("OCR Flow Error: " + e.message);
    throw new Error(e.message);
  }
}

/**
 * Searches guest by mobile number
 * @param {string} mobile - The mobile number to search
 */
function searchGuestByMobile(mobile) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // Mobile is in Column 6 (index 5)
  const mobileColIndex = 5; 
  const searchMobile = String(mobile).replace(/[\s-+]/g, '');

  for (let i = 1; i < data.length; i++) {
    const existingMobile = String(data[i][mobileColIndex]).replace(/[\s-+]/g, '');
    
    if (existingMobile.endsWith(searchMobile) && searchMobile.length >= 10) {
      return {
        exists: true,
        rowNumber: i + 1,
        name: data[i][2],
        idType: data[i][3],
        idNumber: data[i][4],
        emergencyName: data[i][8],
        emergencyPhone: data[i][9],
        city: data[i][7]
      };
    }
  }
  return { exists: false };
}

/**
 * Cross-references an extracted ID with existing database to prevent duplicate IDs 
 * across different mobile numbers.
 */
function checkIdMobileAssociation(extractedId, currentMobile) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  const idColIndex = 4; // Column 5 (IDNo)
  const mobileColIndex = 5; // Column 6 (Phone)
  
  const searchId = String(extractedId).replace(/[\s-]/g, '').toUpperCase();
  const searchMobile = String(currentMobile).replace(/[\s-+]/g, '');

  for (let i = 1; i < data.length; i++) {
    const existingId = String(data[i][idColIndex]).replace(/[\s-]/g, '').toUpperCase();
    const existingMobile = String(data[i][mobileColIndex]).replace(/[\s-+]/g, '');

    if (existingId === searchId && searchId !== "") {
      // If ID matches but Mobile is different, we found a conflict
      if (existingMobile !== searchMobile) {
        return {
          conflict: true,
          existingName: data[i][2], // Column 3 (Name)
          existingMobile: data[i][mobileColIndex]
        };
      }
    }
  }
  return { conflict: false };
}


/******************** Test Functions **********************************/

/**
 * Test function to verify the parsing logic for Aadhaar IDs
 */
function testAadhaarParsing() {
  // Simulated OCR output from an Aadhaar Front
  const mockOcrText = `
    Government of India
    Rahul Sharma
    DOB: 01/01/1985
    Male
    1234 5678 9012
    HELP LINE: 1947
  `;

  console.log("--- Starting Aadhaar Parse Test ---");
  
  const result = parseAadhaarData(mockOcrText);
  
  console.log("Detected Name: " + result.name);
  console.log("Detected ID: " + result.idNumber);

  // Assertions
  if (result.name === "Rahul Sharma" && result.idNumber === "1234 5678 9012") {
    console.log("✅ TEST PASSED: Name and ID correctly identified.");
  } else {
    console.warn("❌ TEST FAILED: Check parsing logic.");
  }
}

function testAadhaarBackParsing() {
  const mockBackText = `
    Address: 
    S/O: Ramesh Sharma, House No 101, 
    Sector 5, Bengaluru, Karnataka, 560001
  `;
  
  // Regex to find 6-digit Indian Pincode
  const pincodeMatch = mockBackText.match(/\d{6}/);
  console.log("Detected Pincode: " + (pincodeMatch ? pincodeMatch[0] : "Not found"));
}

/**
 * Test function to verify OCR on a REAL image file from Google Drive
 */
function testRealImageOCR() {
  // 1. Replace with an actual File ID from your Google Drive for testing
  const TEST_FILE_ID = '1fEibDIh9utvo01Holg73bDh-5CsQ3GaX'; 
  
  try {
    const file = DriveApp.getFileById(TEST_FILE_ID);
    const blob = file.getBlob();
    
    // Convert blob to Base64 to simulate the frontend behavior
    const base64Data = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
    
    console.log("--- Starting Real Image OCR Test ---");
    
    // Step 1: Extract raw text
    const rawText = extractTextFromImage(base64Data);
    console.log("RAW TEXT FROM DRIVE OCR:\n" + rawText);
    
    // Step 2: Parse the raw text
    const result = parseAadhaarData(rawText);
    
    console.log("--- FINAL RESULTS ---");
    console.log("Mapped Name: " + result.name);
    console.log("Mapped ID: " + result.idNumber);
    
  } catch (e) {
    console.error("Test Failed: " + e.toString());
  }
}

/**
 * Test function to verify OCR on a REAL image file from Google Drive
 */
function testRealImageExtraction() {
  // 1. Upload a clear photo of an Aadhaar card to your Google Drive.
  // 2. Right-click the file -> Get Link -> Copy the ID (the long string of characters).
  const TEST_FILE_ID = '1fEibDIh9utvo01Holg73bDh-5CsQ3GaX'; 
  
  try {
    const file = DriveApp.getFileById(TEST_FILE_ID);
    const blob = file.getBlob();
    
    // Simulate the Base64 data that comes from the browser
    const base64Data = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
    
    console.log("--- Starting Real Image OCR Test ---");
    
    // Step 1: Run the OCR engine
    const rawText = extractTextFromImage(base64Data);
    console.log("RAW TEXT DETECTED:\n" + rawText);
    
    // Step 2: Run the parsing logic
    const result = parseAadhaarData(rawText);
    
    console.log("--- MAPPED DATA ---");
    console.log("Mapped Name: " + result.name);
    console.log("Mapped ID:   " + result.idNumber);
    
    if (result.idNumber) {
      console.log("✅ Success: Data extracted successfully.");
    } else {
      console.warn("⚠️ Partial Success: Text was read, but ID number wasn't found. Check image clarity.");
    }
    
  } catch (e) {
    console.error("Test Failed: " + e.toString());
  }
}

/**
 * Run this function from the Apps Script Editor to test the logic.
 * Check your Execution Logs and Google Sheet after running.
 */
function testProcessSubmission() {
  console.log("Starting Test: processSubmission");
  
  // 1. Mock Data Object (Simulating the frontend 'formData')
  // Using a sample Base64 string for a 1x1 transparent pixel to test ImageKit
  const mockBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";
  
  const testData = {
    name: "Test Guest",
    idType: "Aadhar",
    idNumber: "[Aadhaar Redacted]", // Placeholder for testing
    whatsapp: "9876543210",
    purpose: "Business",
    city: "Bengaluru",
    emergencyName: "Emergency Contact",
    emergencyPhone: "9000000000",
    checkIn: "2026-05-10",
    checkOut: "2026-05-15",
    adults: 2,
    kids: 1,
    totalStay: "5 Nights",
    idFrontBase64: mockBase64,
    idBackBase64: mockBase64,
    selfieBase64: mockBase64,
    consent: true
  };

  try {
    // 2. Execute the function
    const result = processSubmission(testData);
    
    // 3. Log results for verification
    console.log("Success!");
    console.log("Assigned Serial Number: " + result.slNo);
    console.log("Check your Google Sheet and ImageKit folder.");
    
  } catch (e) {
    console.error("Test Failed: " + e.toString());
  }
}

/**
 * TEST CASE: Karnataka Driving License (13-digit)
 * Image Reference: 20260512_101221_2.jpg
 */
function TestDL() {
  const mockRawText = `
    DL No. : KA02 20040011775
    NAME : YASHASWINI H J
    D.O.B : 30/06/1985
    VALID TILL : 26/09/2034(NT)
    S/O : JAVARAJA
    ADDRESS : 912 16TH MAIN ROAD
    BANGALORE NORTH, KA 560010
  `;

  console.log("--- Starting DL Debug Test ---");
  const result = parseDrivingLicenseData(mockRawText);

  // Assertions
  const expectedID = "KA02 20040011775";
  const expectedName = "YASHASWINI H J";

  console.log("Testing ID Match:", result.idNumber === expectedID ? "✅ PASS" : "❌ FAIL (Got: " + result.idNumber + ")");
  console.log("Testing Name Match:", result.name === expectedName ? "✅ PASS" : "❌ FAIL (Got: " + result.name + ")");
  console.log("Full Result Object:", result);
  console.log("------------------------------");
}

