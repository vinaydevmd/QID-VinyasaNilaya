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

  // 1. Determine the Target Row
  const isUpdate = data.isUpdate && data.rowNumber;
  const targetRow = isUpdate ? parseInt(data.rowNumber) : sheet.getLastRow() + 1;

  // 2. Determine Status & Integer Row Number
  const checkInStatus = isUpdate ? "Returning Guest - Updated" : "New Guest - Verified";

  // 3. Handle Serial Number (SlNo)
  let finalSlNo;
  if (isUpdate) {
    // Keep the existing SlNo from Column A (Index 1)
    finalSlNo = sheet.getRange(targetRow, 1).getValue();
  } else {
    // New Guest: Calculate new SlNo
    const lastRow = sheet.getLastRow();
    finalSlNo = 1;
    if (lastRow > 1) {
      const lastVal = sheet.getRange(lastRow, 1).getValue();
      finalSlNo = (typeof lastVal === 'number') ? lastVal + 1 : lastRow;
    }
  }

  // 6. Save to Sheet
  if (isUpdate && targetRow) {

    // 4. Google drive Uploads
    const selfieUrl = uploadToDrive(data.selfieBase64, finalSlNo, "Selfie", "");

    // Define your column indices (Update these to match your actual Sheet layout)
    const COL_TIMESTAMP = 2;
    const COL_PURPOSE = 7;
    const COL_ARRIVINGCITY = 8;
    const COL_EMERGENCYCONTNAME = 9;
    const COL_EMERGENCYCONTNO = 10;
    const COL_SELFIE = 13;
    const COL_CHECKINSTATUS = 14;


    sheet.getRange(targetRow, COL_TIMESTAMP).setValue(now);
    sheet.getRange(targetRow, COL_PURPOSE).setValue(data.purpose);
    sheet.getRange(targetRow, COL_ARRIVINGCITY).setValue(data.city);
    sheet.getRange(targetRow, COL_EMERGENCYCONTNAME).setValue(data.emergencyName);
    sheet.getRange(targetRow, COL_EMERGENCYCONTNO).setValue(data.emergencyPhone);
    sheet.getRange(targetRow, COL_SELFIE).setValue(selfieUrl);
    sheet.getRange(targetRow, COL_CHECKINSTATUS).setValue(checkInStatus);

  } else {
    // 4. Google drive Uploads
    const idFrontUrl = uploadToDrive(data.idFrontBase64, finalSlNo, data.idType, "Front");
    const idBackUrl = uploadToDrive(data.idBackBase64, finalSlNo, data.idType, "Back");
    const selfieUrl = uploadToDrive(data.selfieBase64, finalSlNo, "Selfie", "");

    // 5. Construct Full 20-Column Row (Must be exactly 20 elements)
    const rowData = new Array(19).fill(""); // Initialize empty array of 20

    rowData[0] = finalSlNo;                 // Col 1
    rowData[1] = now;                       // Col 2
    rowData[2] = data.idType || "";         // Col 3
    rowData[3] = data.idNumber || "";       // Col 4
    rowData[4] = data.name || "";           // Col 5
    rowData[5] = data.whatsapp || "";       // Col 6
    rowData[6] = data.purpose || "";        // Col 7
    rowData[7] = data.city || "";           // Col 8
    rowData[8] = data.emergencyName || "";  // Col 9
    rowData[9] = data.emergencyPhone || ""; // Col 10
    rowData[10] = idFrontUrl || "";         // Col 16
    rowData[11] = idBackUrl || "";          // Col 17
    rowData[12] = selfieUrl;                // Col 18
    rowData[13] = checkInStatus;            // Col 19
    rowData[14] = data.address || "";             // Col 20

    sheet.appendRow(rowData);
  }

  return {
    success: true,
    slNo: finalSlNo,
    status: checkInStatus
  };
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
  const lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 1);
  const standardizedText = rawText.replace(/[x*×K]/g, 'X');

  // --- 1. ID NUMBER EXTRACTION (Anti-VID Priority Logic) ---
  const idRegex = /(\b[X\d]{4}\s[X\d]{4}\s\d{4}\b)|(\b\d{4}\b$)/gm;
  const matches = standardizedText.match(idRegex) || [];
  let idNumber = "";
  let fallbackId = "";

  const blacklisted = ["1947", "2021", "2022", "2023", "2024", "2025", "2026"];

  for (let i = 0; i < matches.length; i++) {
    let candidate = matches[i].trim();
    let cleanDigits = candidate.replace(/\s/g, "");

    // PRIORITY 1: Full 12-character patterns (e.g., XXXX XXXX 9761)
    if (cleanDigits.length === 12) {
      // Check if it's labeled as VID in the surrounding text
      const matchIndex = standardizedText.indexOf(candidate);
      const contextBefore = standardizedText.substring(Math.max(0, matchIndex - 15), matchIndex).toUpperCase();

      if (!contextBefore.includes("VID")) {
        idNumber = candidate.toUpperCase();
        break; // Stop immediately - we found the primary ID
      }
    }

    // PRIORITY 2: Standalone 4-digit blocks (Fallback for box-redacted cards)
    else if (cleanDigits.length === 4 && !idNumber) {
      const matchIndex = standardizedText.indexOf(candidate);
      const contextBefore = standardizedText.substring(Math.max(0, matchIndex - 15), matchIndex).toUpperCase();

      if (!blacklisted.includes(candidate) && !contextBefore.includes("VID")) {
        fallbackId = "XXXX XXXX " + candidate;
      }
    }
  }

  // Use the 12-character match if found, otherwise use the best fallback
  idNumber = idNumber || fallbackId;

  // --- 2. DATA EXTRACTION LOOP ---
  let detectedName = "Not found";
  let detectedAddress = "Not found";
  let capturingAddress = false;
  let addressLines = [];

  const noiseKeywords = [
    "GOVERNMENT", "INDIA", "FATHER", "DOB", "MALE", "FEMALE",
    "ENROLLMENT", "UNIQUE", "HELP", "YEAR", "VID", "INDA",
    "WWW.", "HELP@", "ELITEBOOK", "LATITUDE", "THINKPAD", "MACBOOK", "HP", "DELL"
  ];
  const searchLimit = Math.floor(lines.length * 0.4);

  for (let i = 0; i < lines.length; i++) {
    let englishOnlyLine = lines[i].replace(/[^\x00-\x7F]/g, "").trim();
    const upperLine = englishOnlyLine.toUpperCase();

    // A. NAME EXTRACTION (Limited to top 40% of card)
    if (detectedName === "Not found" && i < searchLimit) {
      // 1. Watermark Shield: Detect repeating patterns from ghost images (e.g., "UIDAIUIDAI")
      const isWatermarkGarbage = /(UIDAI|GOI|IDAI|OIG|G0I){2,}/.test(upperLine);

      // 2. Structural Checks
      const isRelation = /S\/O|D\/O|W\/O|SON OF|DAUGHTER OF|WIFE OF/i.test(upperLine);
      const isNoise = noiseKeywords.some(word => upperLine.includes(word));
      const hasNumbers = /\d/.test(englishOnlyLine);
      const hasVowels = /[AEIOUY]/.test(upperLine); // Real names must have vowels

      if (englishOnlyLine.length > 3 && !isRelation && !isNoise && !hasNumbers && !isWatermarkGarbage && hasVowels) {

        let potentialName = englishOnlyLine.replace(/^[:\s,-]+/, "").trim();

        // 3. Look-Ahead Logic: Check if the next line is a continuation (like "Ram" or "M D")
        if (i + 1 < searchLimit) {
          let nextLine = lines[i + 1].replace(/[^\x00-\x7F]/g, "").trim();
          const nextUpper = nextLine.toUpperCase();
          const nextIsNoise = noiseKeywords.some(word => nextUpper.includes(word));

          // Append if next line is short, clean, and not a relation/number
          if (nextLine.length > 0 && nextLine.length < 15 && !nextIsNoise && !/\d/.test(nextLine) && !/S\/O|D\/O|W\/O/i.test(nextUpper)) {
            potentialName += " " + nextLine;
            i++; // Consume the next line
          }
        }

        detectedName = potentialName;
      }
    }

    // B. ADDRESS LOGIC
    const isAddressLabel = upperLine.includes("ADDRESS");
    const isRelationTrigger = upperLine.includes("S/O") || upperLine.includes("D/O") || upperLine.includes("W/O");

    if (isAddressLabel || isRelationTrigger) {
      if (capturingAddress) { addressLines = []; } // Reset if we find English version after local
      capturingAddress = true;

      let startText = englishOnlyLine.replace(/Address[:\s]*/i, "").trim();
      startText = startText.replace(/^[:,\s\d]+/, "").trim();

      if (startText.replace(/[^a-zA-Z]/g, "").length > 3) {
        addressLines.push(startText);
      }
      continue;
    }

    if (capturingAddress) {
      const isFooter = ["WWW.", "UNIQUE", "HELP", "1947", "UIDAI"].some(word => upperLine.includes(word));
      const isIdRepeat = idNumber && englishOnlyLine.replace(/\s/g, '').includes(idNumber.replace(/\s/g, '').slice(-4));

      if (isFooter || isIdRepeat) {
        capturingAddress = false;
      } else {
        if (englishOnlyLine.replace(/[^a-zA-Z]/g, "").length > 3) {
          addressLines.push(englishOnlyLine);
        }
      }
    }
  }

  // --- 3. FINAL CLEANUP ---
  if (addressLines.length > 0) {
    detectedAddress = addressLines.join(", ")
      .replace(/,\s*,/g, ",")
      .trim();
    detectedAddress = detectedAddress.replace(/^(Address|S\/O|D\/O|W\/O)\s+\1/i, "$1");
  }

  return {
    name: detectedName,
    idNumber: idNumber || "",
    address: detectedAddress,
    raw: rawText
  };
}

/******** Parse Voter ID *********************/
function parseVoterIDData(rawText) {
  const lines = rawText.split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 2);

  // EPIC Pattern: 3 Letters followed by 7 Digits (e.g., WBH4159588)
  const idMatch = rawText.match(/[A-Z]{3}\d{7}/i);

  let detectedName = "Not found";
  let detectedAddress = "Not found";

  // 1. NAME EXTRACTION (Using your existing label + fallback logic)
  for (let i = 0; i < lines.length; i++) {
    const upperLine = lines[i].toUpperCase();
    if (upperLine.includes("ELECTOR'S NAME") || upperLine.includes("ELECTORS NAME") || upperLine.includes("NAME")) {
      let namePart = lines[i].split(/[:|-]/).pop().trim();
      if (namePart.length < 3 && i + 1 < lines.length) {
        namePart = lines[i + 1].trim();
      }
      detectedName = namePart.replace(/[^\x00-\x7F]/g, "").trim();
      if (detectedName.length > 3) break;
    }
  }

  // --- 2. ADDRESS EXTRACTION (Voter ID Specific) ---
  let capturingAddress = false;
  let addressLines = [];

  // Keywords that definitely mean the address has ended
  const stopKeywords = ["DATE", "PLACE", "ELECTORAL", "REGISTRATION", "OFFICER", "FACSIMILE", "CHANGE", "OBTAIN"];

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i];
    let englishOnlyLine = line.replace(/[^\x00-\x7F]/g, "").trim();
    const upperLine = englishOnlyLine.toUpperCase();

    // TRIGGER: Voter IDs often start English address with a '#' or a number
    // We look for a line that starts with # or has "Address"
    const isAddressStart = upperLine.includes("ADDRESS") || englishOnlyLine.startsWith("#") || /^\d{1,4}[/-]\d+/.test(englishOnlyLine);

    if (isAddressStart && !capturingAddress) {
      capturingAddress = true;
      let startText = englishOnlyLine.replace(/Address[:\s]*/i, "").trim();
      startText = startText.replace(/^[:,\s]+/, "").trim();

      if (startText.length > 3) {
        addressLines.push(startText);
      }
      continue;
    }

    if (capturingAddress) {
      // STOP LOGIC: If we hit a stop keyword, or a line that is purely a date/website
      const shouldStop = stopKeywords.some(word => upperLine.includes(word)) ||
        /\d{2}\/\d{2}\/\d{4}/.test(englishOnlyLine); // Matches dates like 12/05/2026

      if (shouldStop) {
        capturingAddress = false;
        break; // Exit the loop entirely once we hit the footer
      } else {
        // Only add if it contains alphabetic characters (avoids stray symbols)
        if (englishOnlyLine.replace(/[^a-zA-Z]/g, "").length > 3) {
          addressLines.push(englishOnlyLine);
        }
      }
    }
  }

  if (addressLines.length > 0) {
    detectedAddress = addressLines.join(", ").replace(/,\s*,/g, ",").trim();
  }

  return {
    name: detectedName,
    idNumber: idMatch ? idMatch[0].toUpperCase() : "",
    address: detectedAddress,
    raw: rawText
  };
}

/**
 * Universal Parser for Indian Driving Licenses (All States)
 * Targets MoRTH standard DL formats and common English labels.
 */
function parseDrivingLicenseData(combinedText) {
  const lines = combinedText.split('\n').map(l => l.trim()).filter(l => l.length > 1);

  // 1. DL Number Regex
  const dlPattern = /([A-Z]{2}\d{2})[\s\-]?(\d{4})[\s\-]?(\d{5,7})/i;
  const dlMatch = combinedText.match(dlPattern);
  let idNumber = dlMatch ? (dlMatch[1] + " " + dlMatch[2] + " " + dlMatch[3]).toUpperCase() : "";

  let detectedName = "Not found";
  let detectedAddress = "Not found";
  let capturingAddress = false;
  let addressLines = [];

  const noiseKeywords = ["TRANSPORT", "DATE", "BIRTH", "D.O.B", "ISSUE", "EXPIRY", "VALID", "ADDRESS", "S/O", "D/O", "W/O", "FATHER", "HUSBAND", "COV", "DOI", "INDIA", "CARD"];
  const stopKeywords = ["VALID", "TILL", "SIGN", "DOI", "COV", "AUTHORITY", "BLOOD", "B.G"];

  // THE LOOP STARTS HERE
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const upperLine = line.toUpperCase();
    const cleanLine = line.replace(/[^\x00-\x7F]/g, "").trim(); // Define cleanLine for each iteration

    // --- A. NAME EXTRACTION ---
    if (detectedName === "Not found") {
      const isNameLine = /N[A-Z0-4\s]{2,3}E|HOLDER/i.test(upperLine);
      if (isNameLine) {
        let potentialName = line.includes(":") ? line.split(":").pop().trim() : "";

        let searchOffset = 1;
        while (potentialName.length < 3 && searchOffset <= 3 && (i + searchOffset) < lines.length) {
          const candidate = lines[i + searchOffset].trim();
          const isDate = /\d{2}[\/\-]\d{2}[\/\-]\d{4}/.test(candidate);
          const isNoise = noiseKeywords.some(word => candidate.toUpperCase().includes(word));
          const hasNumbers = /\d/.test(candidate);

          if (candidate.length > 3 && !isDate && !isNoise && !hasNumbers) {
            potentialName = candidate;
          }
          searchOffset++;
        }
        const cleanNameResult = potentialName.replace(/[^\x00-\x7F]/g, "").replace(/^[:\s\-]+/, "").trim();
        if (cleanNameResult.length > 3) detectedName = cleanNameResult.toUpperCase();
      }
    }

    // --- B. ADDRESS EXTRACTION (Now correctly inside the loop) ---
    if (upperLine.includes("ADDRESS")) {
      capturingAddress = true;
      // Get text after "ADDRESS" label
      let startText = cleanLine.split(/[:|-]/).pop().trim();

      // If "ADDRESS" is a lone label, move to capture next line
      if (startText.toUpperCase() === "ADDRESS" || startText.length < 2) {
        continue;
      }
      addressLines.push(startText);
      continue;
    }

    if (capturingAddress) {
      const shouldStop = stopKeywords.some(word => upperLine.includes(word)) ||
        /\d{2}[\/\-]\d{2}[\/\-]\d{4}/.test(cleanLine);

      if (shouldStop) {
        capturingAddress = false; // Stop capturing when we hit the footer
      } else {
        // Validation: Ignore noisy words but keep capturing address details
        const isNoise = noiseKeywords.some(word => upperLine.includes(word) && word !== "ADDRESS");
        if (!isNoise && cleanLine.length > 2) {
          addressLines.push(cleanLine);
        }
      }
    }
  } // THE LOOP ENDS HERE

  if (addressLines.length > 0) {
    detectedAddress = addressLines.join(", ")
      .replace(/[:]/g, "")
      .replace(/,\s*,/g, ",")
      .trim();
  }

  return {
    name: detectedName,
    idNumber: idNumber,
    address: detectedAddress,
    raw: combinedText
  };
}

/**
 * Specialized parser for Indian Passports using regex for labels and MRZ patterns
 */
function parsePassportData(rawText) {
  const lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 1);
  const idMatch = rawText.match(/[A-Z]\d{7}/i);

  let surname = "";
  let givenName = "";
  let detectedAddress = "Not found";
  let capturingAddress = false;
  let addressLines = [];

  const headers = ["SURNAME", "GIVEN NAME", "NAME", "दिया गया नाम", "उपनाम"];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const upperLine = line.toUpperCase();

    // 1. NAME EXTRACTION
    if (upperLine.includes("SURNAME") || upperLine.includes("उपनाम")) {
      let val = line.split(/[:/|-]/).pop().trim();
      if (val.length < 3 && i + 1 < lines.length) val = lines[i + 1].trim();
      if (!headers.some(h => val.toUpperCase().includes(h))) {
        surname = val.replace(/[^\x00-\x7F]/g, "").trim();
      }
    }

    if (upperLine.includes("GIVEN NAME") || upperLine.includes("दिया गया नाम")) {
      let val = line.split(/[:/|-]/).pop().trim();
      if (val.length < 3 && i + 1 < lines.length) val = lines[i + 1].trim();
      if (!headers.some(h => val.toUpperCase().includes(h))) {
        givenName = val.replace(/[^\x00-\x7F]/g, "").trim();
      }
    }

    // 2. ADDRESS EXTRACTION (The New Logic)
    // Trigger on "Address" label
    if (upperLine.includes("ADDRESS") || upperLine.includes("पता")) {
      capturingAddress = true;
      let startText = line.split(/[:/|-]/).pop().trim();
      // If the label and value are on separate lines, skip the label-only line
      if (startText.toUpperCase() !== "ADDRESS" && startText.length > 2) {
        addressLines.push(startText);
      }
      continue;
    }

    if (capturingAddress) {
      // Termination: Stop if we hit PIN, File No, or date patterns
      const isStopWord = ["PIN:", "FILE NO", "PHTO", "OLD PASSPORT", "DATE"].some(word => upperLine.includes(word));
      const isDate = /\d{2}\/\d{2}\/\d{4}/.test(line);
      let englishOnlyLine = line.replace(/[^\x00-\x7F]/g, "").trim();

      if (isStopWord || isDate) {
        capturingAddress = false;
      } else {
        // --- The Improved Noise Cleanup Logic ---
        // 1. Remove non-alphanumeric junk at the start
        const cleanedLine = englishOnlyLine.replace(/^[^a-zA-Z0-9#]+/, "").trim();

        // 2. Artifact Filter: Skip lines that are mostly symbols (like "MI) | | DIY")
        // A valid address line should have a decent ratio of letters/numbers
        const letterCount = (cleanedLine.match(/[a-zA-Z0-9]/g) || []).length;
        const totalCount = cleanedLine.length;

        if (letterCount > 5 && (letterCount / totalCount) > 0.5) {
          addressLines.push(cleanedLine);
        } else {
          // If we haven't found a single valid line yet, we just keep looking.
          // If we are in the middle of an address, a single bad line might just be a separator.
          continue;
        }
      }
    }


  }

  // --- FINAL ASSEMBLY ---
  let fullName = (givenName + " " + surname).trim();

  // MRZ Fallback for Name
  if (!fullName || fullName.length < 5 || fullName.toUpperCase().includes("SURNAME")) {
    const mrzLine = lines.find(l => l.startsWith("P<") || l.includes("<<"));
    if (mrzLine) {
      const cleanMRZ = mrzLine.replace(/^P.[A-Z]{3}/i, "").replace(/^P</i, "");
      const parts = cleanMRZ.split("<<");
      if (parts.length >= 2) {
        const mrzSurname = parts[0].replace(/</g, " ").trim();
        const mrzGiven = parts[1].replace(/</g, " ").trim();
        const finalSurname = mrzSurname.replace(/^[P|I|N|D|K]{1,5}\s+/i, "").trim();
        fullName = (mrzGiven + " " + finalSurname).trim();
      }
    }
  }

  // Final Address Assembly
  if (addressLines.length > 0) {
    detectedAddress = addressLines.join(", ").replace(/,\s*,/g, ",").trim();
  }

  return {
    name: fullName.toUpperCase() || "Not found",
    idNumber: idMatch ? idMatch[0].toUpperCase() : "",
    address: detectedAddress,
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
    switch (idType) {
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
        name: data[i][4],
        idType: data[i][2],
        idNumber: data[i][3],
        emergencyName: data[i][8],
        emergencyPhone: data[i][9],
        city: data[i][7]
      };
    }
  }
  return { exists: false };
}

/**
 * Verifies if an ID is already registered under a different mobile number.
 */
function checkIdMobileAssociation(extractedId, currentMobile) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME); // Ensure this matches your Sheet Name
  const data = sheet.getDataRange().getValues();

  // 1. Correct Column Indices (0-based) based on your Vinyasa Nilaya Ledger:
  // Col A=0 (SlNo), Col B=1 (Name), Col C=2 (ID Type), Col D=3 (ID No), Col E=4 (Phone)
  const idColIndex = 3;
  const mobileColIndex = 5;
  const nameColIndex = 4;

  // 2. Normalize input data
  const searchId = String(extractedId).replace(/[\s-]/g, '').toUpperCase();
  const searchMobile = String(currentMobile).replace(/[\s-+\d]{0,2}/, '').trim();

  console.log(`Checking Conflict - ID: ${searchId}, Mobile: ${searchMobile}`);

  // 3. CRITICAL: Skip if ID is redacted or empty
  // Without this, all guests with "[Aadhaar Redacted]" will conflict with each other.
  if (!searchId || searchId.includes("REDACTED") || searchId === "") {
    console.log("Check Skipped: ID is redacted or empty.");
    return { conflict: false };
  }

  for (let i = 1; i < data.length; i++) {
    let existingId = String(data[i][idColIndex]).replace(/[\s-]/g, '').toUpperCase();
    let existingMobile = String(data[i][mobileColIndex]).replace(/[\s-+\d]{0,2}/, '').trim();

    // 4. Match Logic
    if (existingId === searchId) {
      console.log(`Match Found on Row ${i + 1}. Checking mobile compatibility...`);

      // If ID matches but Mobile is different
      if (existingMobile !== searchMobile && searchMobile !== "") {
        console.warn(`CONFLICT: ID ${searchId} belongs to ${data[i][nameColIndex]} (${existingMobile})`);

        return {
          conflict: true,
          existingName: data[i][nameColIndex],
          existingMobile: data[i][mobileColIndex]
        };
      } else {
        console.log("No conflict: ID and Mobile match existing record.");
      }
    }
  }

  console.log("No conflicts found in ledger.");
  return { conflict: false };
}

/***** Upload images to Google Drive ************/
function uploadToDrive(base64Data, serialNo, idType, side) {
  try {
    // 1. Handle Parent Folder: Identity Proofs
    const parentName = "Identity Proofs";
    let parentFolder, parentFolders = DriveApp.getFoldersByName(parentName);
    parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.createFolder(parentName);

    // 2. Handle Year-based Subfolder: QID-YYYY
    const currentYear = new Date().getFullYear();
    const subFolderName = "QID-" + currentYear;
    let targetFolder, subFolders = parentFolder.getFoldersByName(subFolderName);
    targetFolder = subFolders.hasNext() ? subFolders.next() : parentFolder.createFolder(subFolderName);

    // 3. Format the File Name
    const now = new Date();
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const formattedDate = `${now.getFullYear()}-${monthNames[now.getMonth()]}-${String(now.getDate()).padStart(2, '0')}`;

    let fileName = (idType === "Selfie")
      ? `${formattedDate}-${serialNo}-${idType}`
      : `${formattedDate}-${serialNo}-${idType}-${side}`;

    // --- 4. UPDATED: REPLACEMENT LOGIC FOR SELFIES ---
    const existingFiles = targetFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();

      if (idType === "Selfie") {
        // Delete the old selfie to make room for the new one
        console.log("Replacing existing selfie for repeating guest: " + fileName);
        existingFile.setTrashed(true);
      } else {
        // For ID Cards (Front/Back), we might want to keep the existing one to save time
        console.log("Existing ID found, skipping upload: " + fileName);
        return existingFile.getUrl();
      }
    }

    // 5. Proceed with Upload (New file or Replacement selfie)
    const contentType = base64Data.substring(base64Data.indexOf(":") + 1, base64Data.indexOf(";"));
    const bytes = Utilities.base64Decode(base64Data.split(",")[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);

    const file = targetFolder.createFile(blob);

    // Set permissions so the image can be viewed
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    console.log("New file created/updated: " + fileName);
    return file.getUrl();

  } catch (e) {
    console.error("Drive Upload Error: " + e.toString());
    return null;
  }
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

/**
 * TEST CASE: Google Drive Hierarchical Upload
 * Purpose: Verifies folder creation logic and filename formatting.
 */
function runDriveUploadTest() {
  console.log("--- Starting Google Drive Upload Test ---");

  // 1. Create a dummy 1x1 pixel transparent PNG in Base64
  const dummyBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";
  const testIdType = "Aadhaar";
  const testSide = "Front";

  try {
    // 2. Call your production function
    const fileUrl = uploadToDrive(dummyBase64, 1, testIdType, testSide);

    if (fileUrl && fileUrl.includes("drive.google.com")) {
      console.log("✅ TEST PASSED: File uploaded successfully.");
      console.log("Generated URL: " + fileUrl);

      // 3. Verify Folder Structure
      verifyFolderStructure();
    } else {
      console.error("❌ TEST FAILED: URL was not generated correctly.");
    }
  } catch (e) {
    console.error("❌ TEST FAILED with Error: " + e.message);
  }

  console.log("--- End of Drive Upload Test ---");
}
/**
 * Verification Helper: Checks if the folders were created correctly
 */
function verifyFolderStructure() {
  const parentName = "Identity Proofs";
  const currentYear = new Date().getFullYear();
  const subFolderName = "QID-" + currentYear;

  const parentFolders = DriveApp.getFoldersByName(parentName);
  if (parentFolders.hasNext()) {
    const parent = parentFolders.next();
    console.log(`✅ Parent folder "${parentName}" exists.`);

    const subFolders = parent.getFoldersByName(subFolderName);
    if (subFolders.hasNext()) {
      console.log(`✅ Subfolder "${subFolderName}" exists inside parent.`);

      const fileSearch = subFolders.next().getFiles();
      if (fileSearch.hasNext()) {
        const latestFile = fileSearch.next();
        console.log(`✅ Latest File Name: "${latestFile.getName()}"`);
        console.log(`✅ Sharing Permission: ${latestFile.getSharingAccess()}`);
      }
    } else {
      console.warn(`⚠️ Subfolder "${subFolderName}" was not found.`);
    }
  } else {
    console.warn(`⚠️ Parent folder "${parentName}" was not found.`);
  }
}
/**
 * TEST CASE: ID and Mobile Association
 * Purpose: Verifies conflict detection and Aadhaar redaction handling.
 */
function debugIdMobileAssociation() {
  console.log("--- Starting Conflict Detection Debugger ---");

  // Mock Data mimicking your Guest_Ledger structure
  // Col 0: SlNo | Col 1: Name | Col 2: IDType | Col 3: IDNo | Col 4: Phone
  const mockSheetData = [
    ["SlNo", "Name", "IDType", "IDNo", "Phone"],
    [1, "Yashaswini H J", "Aadhaar", "123456789012", "9876543210"],
    [2, "Previous Guest", "DL", "KA012023001", "9999988888"],
    [3, "Redacted User", "Aadhaar", "[Aadhaar Redacted]", "8888877777"]
  ];

  // Test Scenarios
  const testCases = [
    {
      desc: "CONFLICT: Same ID, Different Mobile",
      id: "1234-5678-9012",
      mobile: "1112223333",
      expected: true
    },
    {
      desc: "VALID: Same ID, Same Mobile (Returning Guest)",
      id: "123456789012",
      mobile: "+91 98765 43210",
      expected: false
    },
    {
      desc: "BYPASS: Redacted ID (Should never conflict)",
      id: "[Aadhaar Redacted]",
      mobile: "7776665555",
      expected: false
    },
    {
      desc: "NEW GUEST: Unique ID and Mobile",
      id: "NEWID999",
      mobile: "5554443333",
      expected: false
    }
  ];

  testCases.forEach((t, index) => {
    const result = runMockCheck(t.id, t.mobile, mockSheetData);

    const status = (result.conflict === t.expected) ? "✅ PASSED" : "❌ FAILED";
    console.log(`${status} - Case ${index + 1}: ${t.desc}`);
    if (result.conflict) {
      console.log(`   Result: Found conflict with ${result.existingName} (${result.existingMobile})`);
    }
  });

  console.log("--- Debugger Complete ---");
}

/**
 * Modified version of your function for isolated testing
 */
function runMockCheck(extractedId, currentMobile, data) {
  const idColIndex = 3;
  const mobileColIndex = 4;
  const nameColIndex = 1;

  const searchId = String(extractedId).replace(/[\s-]/g, '').toUpperCase();
  const searchMobile = String(currentMobile).replace(/[\s-+\d]{0,2}/, '').trim();

  // Redaction Bypass Logic
  if (!searchId || searchId.includes("REDACTED")) return { conflict: false };

  for (let i = 1; i < data.length; i++) {
    let existingId = String(data[i][idColIndex]).replace(/[\s-]/g, '').toUpperCase();
    let existingMobile = String(data[i][mobileColIndex]).replace(/[\s-+\d]{0,2}/, '').trim();

    if (existingId === searchId && !existingId.includes("REDACTED")) {
      if (existingMobile !== searchMobile && searchMobile !== "") {
        return {
          conflict: true,
          existingName: data[i][nameColIndex],
          existingMobile: data[i][mobileColIndex]
        };
      }
    }
  }
  return { conflict: false };
}
