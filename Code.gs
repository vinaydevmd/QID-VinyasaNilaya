/**
 * Global Configuration
 */
const FOLDER_ID = '1cmRFirWeg_tHFbZ9VS-E0Gz80SHHSsIU9lu5jV_GBKk'; 
const SHEET_NAME = 'QID-Booking';

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
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

/*
function processSubmission(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Guest Name', 'Phone', 'Arrival Date', 'ID Type', 'ID Number', 'Selfie URL', 'Terms Accepted']);
      sheet.setFrozenRows(1);
    }

    const contentType = data.selfie.split(';')[0].split(':')[1];
    const bytes = Utilities.base64Decode(data.selfie.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, `Selfie_${data.name.replace(/\s+/g, '_')}_${Date.now()}.jpg`);
    const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);

    let displayId = (data.idType === "Aadhaar" && data.idNumber.length >= 4) ? 
                    "XXXX-XXXX-" + data.idNumber.slice(-4) : data.idNumber;

    sheet.appendRow([new Date(), data.name, data.phone, data.arrival, data.idType, displayId, file.getUrl(), data.terms ? "Yes" : "No"]);

    return { success: true, message: "Booking confirmed! Your details are stored securely." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}*/
function processSubmission(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('QID-Booking');
  
  // Headers match: Timestamp, Name, WhatsApp, ID Type, ID Num, Purpose, City, Emergency Contact, Emergency Phone, ID Front URL, ID Back URL
  sheet.appendRow([
    new Date(),
    data.name,
    data.whatsapp,
    data.idType,
    data.idNumber,
    data.purpose,
    data.city,
    data.emergencyName,
    data.emergencyPhone,
    data.idFrontUrl, // Save URL from Drive
    data.idBackUrl,  // Save URL from Drive
    data.consent ? "Confirmed" : "No"
  ]);
  
  return { success: true, message: "Registration successful!" };
}


/**
 * Uses Google Drive OCR to extract text from image
 */
/*function extractTextFromImage(base64Data) {
  const contentType = base64Data.split(';')[0].split(':')[1];
  const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
  const blob = Utilities.newBlob(bytes, contentType, "temp_ocr.jpg");

  // Create temp file with OCR enabled
  const fileResource = {
    title: 'Temp_OCR_Scan',
    mimeType: contentType
  };
  
  // Requires Drive API Service enabled
  const tempFile = Drive.Files.insert(fileResource, blob, { ocr: true });
  
  // Open as Doc to get text
  const doc = DocumentApp.openById(tempFile.id);
  const text = doc.getBody().getText();
  
  // Cleanup
  Drive.Files.remove(tempFile.id);
  
  return text;
}*/

/*function extractTextFromImage(base64Data) {
  const contentType = base64Data.split(';')[0].split(':')[1];
  const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
  const blob = Utilities.newBlob(bytes, contentType, "scan.jpg");

  // V2 Syntax
  const resource = {
    title: 'OCR_TEMP',
    mimeType: contentType
  };
  
  // Create file with OCR
  const tempFile = Drive.Files.insert(resource, blob, { ocr: true });
  
  // Open and Read
  const doc = DocumentApp.openById(tempFile.id);
  const text = doc.getBody().getText();
  
  // Cleanup
  Drive.Files.remove(tempFile.id);
  
  return text;
}*/

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
 * Logic to clean and find specific Aadhaar data
 */
/*function parseAadhaarData(rawText) {
  const lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 5);
  
  // Aadhaar Number Pattern: 12 digits with spaces
  const idMatch = rawText.match(/\d{4}\s\d{4}\s\d{4}/);
  
  // Name logic: Usually the first or second line of significant text 
  // skipping "Government of India" header
  let detectedName = "Not found";
  for (let line of lines) {
    if (!line.includes("India") && !line.includes("Government") && !line.match(/\d/)) {
      detectedName = line;
      break;
    }
  }

  return {
    name: detectedName,
    idNumber: idMatch ? idMatch[0] : ""
  };
}*/
/*
function parseAadhaarData(rawText) {
  const lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 3);
  const idMatch = rawText.match(/(\d{4}\s\d{4}\s\d{4})|(\d{12})/);
  
  let detectedName = "Not found";
  const noise = ["GOVERNMENT", "INDIA", "FATHER", "DOB", "YEAR", "MALE", "FEMALE", "ADDRESS", "UNIQUE"];

  for (let line of lines) {
    // Keep English characters only
    let englishLine = line.replace(/[^\x00-\x7F]/g, "").trim();
    const upperLine = englishLine.toUpperCase();
    
    const isNoise = noise.some(word => upperLine.includes(word));
    const hasNumbers = /\d/.test(englishLine);

    if (englishLine.length > 5 && !isNoise && !hasNumbers && englishLine.split(" ").length >= 2) {
      detectedName = englishLine;
      break; 
    }
  }

  return {
    name: detectedName,
    idNumber: idMatch ? idMatch[0] : ""
  };
}
*/

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
/**
 * Combined function to handle the flow from the UI
 */
function executeOcrFlow(base64Data) {
  try {
    // 1. Get the raw text from the image
    const rawText = extractTextFromImage(base64Data);
    
    // 2. Parse that text into Name and ID
    const result = parseAadhaarData(rawText);
    
    return result; // Returns {name: "...", idNumber: "..."}
  } catch (e) {
    throw new Error("OCR Processing failed: " + e.message);
  }
}


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


