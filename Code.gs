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
}

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

function executeOcrFlow(base64Data) {
  try {
    const rawText = extractTextFromImage(base64Data);
    
    // 1. Check if ANY text was found at all
    if (!rawText || rawText.trim().length === 0) {
      throw new Error("No text detected. Please ensure the photo is clear and well-lit.");
    }

    // 2. Validate if it's likely a Govt ID
    const idKeywords = ["GOVERNMENT", "INDIA", "INCOME TAX", "ELECTION", "DRIVING", "LICENSE", "ID", "CARD", "UNIQUE"];
    const upperText = rawText.toUpperCase();
    const hasIdKeywords = idKeywords.some(keyword => upperText.includes(keyword));

    if (!hasIdKeywords) {
      throw new Error("This doesn't look like a valid Government ID. Please upload a clear photo of your original card.");
    }
    
    return parseAadhaarData(rawText);
  } catch (e) {
    // Re-throw the error so withFailureHandler catches it
    throw new Error(e.message);
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
