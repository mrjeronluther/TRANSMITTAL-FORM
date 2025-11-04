/**
 * @fileoverview Server-side logic for the Transmittal Form Web App.
 * Manages data fetching, saving, and PDF generation for transmittals.
 */

// --- GLOBAL CONSTANTS ---
const SPREADSHEET_ID = "1iihUBB3vXloBq_j5dp_B31nrth1vpkMh1Vfe4P-YJU8";
const SHEET_NAME = "TransmittalRecord";
const DRIVE_FOLDER_ID = "1MHrBuuJPUeNaRn8C4rkhDy6F154OycQh";
const CONFIG_SHEET_NAME = "SourceFiles";
const PDF_TEMPLATE_FILENAME = "pdf_template.html";

const departmentDetails = {
    COG: { title: "COG TRANSMITTAL FORM", address: "2nd Floor 8 IBM Bldg. Eastwood Ave. Eastwood City Cyberpark Bagumbayan, Quezon City Metro Manila.", phone: "-" },
    PCU: { title: "PCU TRANSMITTAL FORM", address: "B2 (near Tower 3 Elevator) Uptown Mall, 38th Street corner 11th Avenue Uptown Bonifacio City, Philippines, 1634", phone: "(02) 8809 3405 local 7566" },
    DEFAULT: { title: "TRANSMITTAL FORM", address: "30/F Allaince Global Tower, 36th Street cor. 11th Avenue. Uptown Bonifacio, Taguig City 1630", phone: "(632)898-5999 Fax No. 857-9899 - www.megaworldcorp.com)" }
};

// --- MAIN WEB APP FUNCTIONS ---

/**
 * Serves the main HTML page of the web app.
 * @param {object} e The event parameter for a GET request.
 * @returns {HtmlOutput} The HTML service output.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle("COG-PCU Transmittal Form")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Includes the content of another server-side HTML file.
 * Used for modularizing CSS and JavaScript.
 * @param {string} filename The name of the file to include.
 * @returns {string} The HTML content of the file.
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// --- DATA FETCHING & CRUD OPERATIONS ---

/**
 * BEST PRACTICE: Uses TextFinder for optimized searching.
 * Fetches all rows from a source spreadsheet that match a given RFP/PEF number.
 * @param {string} rfpNumber The RFP/PEF number to search for.
 * @param {string} sourceId The Google Spreadsheet ID of the source file.
 * @returns {Array<Object>} An array of objects representing matched rows.
 */
function fetchRfpData(rfpNumber, sourceId) {
    if (!rfpNumber || !sourceId) return [];

    try {
        const mainSs = SpreadsheetApp.openById(SPREADSHEET_ID);
        const configSheet = mainSs.getSheetByName(CONFIG_SHEET_NAME);
        if (!configSheet) {
            throw new Error(`FATAL: Configuration sheet named "${CONFIG_SHEET_NAME}" could not be found.`);
        }

        const allSourceIds = configSheet.getRange("A2:A" + configSheet.getLastRow()).getValues().flat();
        const idIndex = allSourceIds.findIndex(id => id.toString().trim() === sourceId.toString().trim());
        if (idIndex === -1) {
            throw new Error(`The selected file ID "${sourceId}" was not found in the configuration sheet.`);
        }
        const configRow = idIndex + 2;
        const tabsCellValue = configSheet.getRange(configRow, 3).getValue();
        
        const sourceSs = SpreadsheetApp.openById(sourceId);
        let sheetsToSearch;
        const allowedTabs = (tabsCellValue && typeof tabsCellValue === 'string') ? tabsCellValue.split(',').map(tab => tab.trim()).filter(String) : [];

        if (allowedTabs.length > 0) {
            sheetsToSearch = allowedTabs.map(tabName => sourceSs.getSheetByName(tabName)).filter(Boolean);
            if (sheetsToSearch.length === 0) {
                throw new Error(`None of the specified tabs (${allowedTabs.join(', ')}) were found.`);
            }
        } else {
            sheetsToSearch = sourceSs.getSheets();
        }

        const headersToMap = {
            'DOCUMENT TRANSMITTED/ DETAILS': 'docDetails',
            'SUPPLIER/ VENDOR/ AGENCY': 'supplier',
            'PAYOR COMPANY': 'payorCompany',
            'PROPERTY': 'property',
            'LOCATION': 'location',
            'SECTOR': 'sector',
            'TYPE OF SERVICE': 'serviceType',
            'PERIOD COVERED': 'periodCovered',
            'PARTICULARS': 'particulars',
            'RFP/ PEF AMOUNT': 'rfpAmount'
        };
        
        const allMatches = [];
        // Convert the search term to a clean, standardized string ONCE.
        const normalizedRfpNumber = rfpNumber.toString().trim();

        for (const sheet of sheetsToSearch) {
            const headerRowIndex = 5;
            if (sheet.getLastRow() < headerRowIndex) continue;

            const sourceHeaders = sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
            const rfpColIndex = sourceHeaders.indexOf('RFP/ PEF #');
            if (rfpColIndex === -1) continue;

            // ---- MODIFICATION START: Replace createTextFinder with a more robust manual search ----

            const allDataValues = sheet.getDataRange().getValues(); // Get all data at once for efficiency
            const rfpColumnValues = allDataValues.slice(headerRowIndex); // Get only data rows

            // Loop through each row to find our reference number
            for (let i = 0; i < rfpColumnValues.length; i++) {
                const cellValue = rfpColumnValues[i][rfpColIndex];
                
                // Normalize the cell data: convert to string and trim whitespace
                const normalizedCellValue = cellValue.toString().trim();

                // Compare the normalized values
                if (normalizedCellValue === normalizedRfpNumber) {
                    // If we found a match, process the row
                    const rowIndex = headerRowIndex + i; // The actual row index in the full dataset
                    const rowData = allDataValues[rowIndex];
                    const headerColumnIndexMap = {};
                    sourceHeaders.forEach((header, index) => headerColumnIndexMap[header.trim()] = index);
                    
                    const result = { rfpPef: rfpNumber };
                    for (const sourceHeader in headersToMap) {
                        const resultKey = headersToMap[sourceHeader];
                        const columnIndex = headerColumnIndexMap[sourceHeader];
                        result[resultKey] = (columnIndex !== undefined) ? rowData[columnIndex] : '';
                    }
                    allMatches.push(result);
                }
            }
            // ---- MODIFICATION END ----
        }
        return allMatches;
    } catch (error) {
        Logger.log(`Error in fetchRfpData: ${error.toString()}`);
        throw new Error(`Server error while fetching data: ${error.message}`);
    }
}

/**
 * BEST PRACTICE: Reads all data in a single call.
 * Gets the list of available source files from the configuration sheet.
 * @returns {Array<{id: string, label: string}>} A list of source files.
 */
function getSourceFiles() {
  try {
    const configSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CONFIG_SHEET_NAME);
    if (!configSheet) throw new Error(`Configuration sheet "${CONFIG_SHEET_NAME}" not found.`);
    
    const lastRow = configSheet.getLastRow();
    if (lastRow < 2) return [];
    
    const sourceData = configSheet.getRange("A2:B" + lastRow).getValues();
    return sourceData
      .map(row => (row[0] && row[1]) ? { id: row[0].toString().trim(), label: row[1].toString().trim() } : null)
      .filter(Boolean);
  } catch (error) {
    Logger.log(`Error in getSourceFiles: ${error.message}`);
    throw new Error("Failed to fetch the source file list.");
  }
}

/**
 * BEST PRACTICE: Uses LockService to prevent race conditions.
 * Generates a unique transmittal number.
 * @returns {string} The unique transmittal number.
 */
function generateTransmittalNumber() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Server is busy generating another number. Please try again.");
  }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet named "${SHEET_NAME}" could not be found.`);
    
    const lastRow = sheet.getLastRow();
    const range = (lastRow > 1) ? sheet.getRange("B2:B" + lastRow) : null;
    const existingNumbers = new Set(range ? range.getValues().flat().filter(String) : []);
    
    let newNumber;
    let isUnique = false;
    const today = new Date();
    const datePart = Utilities.formatDate(today, "GMT+8", "yyyyMMdd");
    
    for (let attempts = 0; attempts < 100 && !isUnique; attempts++) {
      const randomPart = Math.floor(1000 + Math.random() * 9000).toString();
      newNumber = `${datePart}-${randomPart}`;
      if (!existingNumbers.has(newNumber)) isUnique = true;
    }

    if (!isUnique) throw new Error("Could not generate a unique number after 100 attempts.");
    return newNumber;
  } catch (error) {
    Logger.log(`Error in generateTransmittalNumber: ${error.message}`);
    throw new Error("Failed to generate a unique number.");
  } finally {
    lock.releaseLock();
  }
}

/**
 * BEST PRACTICES: Uses LockService and bulk setValues().
 * Saves the entire transmittal form data to the spreadsheet.
 * @param {object} formData The form data submitted from the client.
 * @returns {string} A success message.
 */
function saveTransmittalData(formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Server is busy saving another transmittal. Please try again.");
  }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

    const timestamp = new Date();
    const rowsToAppend = formData.items.map(item => [
        timestamp, formData.transmittalNo, formData.fromName, formData.fromDepartment,
        formData.dateTransmitted, formData.toName, formData.toDepartment, formData.toAddress,
        item.docDetails, item.rfpPef, item.supplier, item.payorCompany, item.property,
        item.location, item.sector, item.serviceType, item.periodCovered,
        item.particulars, item.rfpAmount, 'Generating...'
    ]);

    if (rowsToAppend.length === 0) throw new Error("No items were submitted.");

    // --- MODIFIED SECTION START ---
    // Find the last row with data specifically within columns A to T
    const rangeAtoT = sheet.getRange("A:T");
    const valuesAtoT = rangeAtoT.getValues();
    let lastRowInAtoT = 0;
    for (let i = valuesAtoT.length - 1; i >= 0; i--) {
      if (valuesAtoT[i].join("") !== "") {
        lastRowInAtoT = i + 1;
        break;
      }
    }
    const firstRowIndex = lastRowInAtoT + 1;
    // --- MODIFIED SECTION END ---

    sheet.getRange(firstRowIndex, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

    const pdfFileUrl = createTransmittalPdf(formData);
    const pdfUrlsForSheet = Array(rowsToAppend.length).fill([pdfFileUrl]);
    sheet.getRange(firstRowIndex, 20, pdfUrlsForSheet.length, 1).setValues(pdfUrlsForSheet);
    
    return `Transmittal ${formData.transmittalNo} submitted successfully.`;
  } catch (error) {
    Logger.log(`Error in saveTransmittalData: ${error.toString()}`);
    throw new Error(`Submission failed: ${error.message}`);
  } finally {
    lock.releaseLock();
  }
}

// --- PDF AND UTILITY FUNCTIONS ---

/**
 * **CORRECTED**
 * Creates a PDF file from the form data and saves it to Google Drive.
 * This version assigns variables individually to match the original template.
 * @param {object} formData The form data.
 * @returns {string} The URL of the generated PDF file.
 */
function createTransmittalPdf(formData) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const template = HtmlService.createTemplateFromFile(PDF_TEMPLATE_FILENAME);
    
    const dept = formData.fromDepartment ? formData.fromDepartment.toUpperCase() : "DEFAULT";
    const details = departmentDetails[dept] || departmentDetails.DEFAULT;

    // Assign template variables individually
    template.formTitle = details.title;
    template.addressLine = details.address;
    template.phoneLine = details.phone;
    
    template.transmittalNo = formData.transmittalNo || '';
    template.fromName = formData.fromName || '';
    template.fromDepartment = formData.fromDepartment || '';
    template.toName = formData.toName || '';
    template.toDepartment = formData.toDepartment || '';
    template.toAddress = formData.toAddress || '';
    template.dateTransmitted = formData.dateTransmitted || '';
    template.items = formData.items || [];
    template.logoHtml = getLogoHtml();
    
    const htmlContent = template.evaluate().getContent();
    const pdfBlob = Utilities.newBlob(htmlContent, MimeType.HTML).getAs(MimeType.PDF);
    const pdfName = `Transmittal_${formData.transmittalNo}.pdf`;
    const pdfFile = folder.createFile(pdfBlob.setName(pdfName));
    
    return pdfFile.getUrl();
  } catch(e) {
    Logger.log(`Error creating PDF: ${e.message}`);
    throw new Error("PDF generation failed.");
  }
}

/**
 * **CORRECTED**
 * Generates the HTML content for a print preview.
 * This version assigns variables individually to match the original template.
 * @param {object} formData The form data.
 * @returns {string} The evaluated HTML content.
 */
function getPrintPreviewHtml(formData) {
  try {
    const template = HtmlService.createTemplateFromFile(PDF_TEMPLATE_FILENAME);
    
    const dept = formData.fromDepartment ? formData.fromDepartment.toUpperCase() : "DEFAULT";
    const details = departmentDetails[dept] || departmentDetails.DEFAULT;
    
    // Assign template variables individually
    template.formTitle = details.title;
    template.addressLine = details.address;
    template.phoneLine = details.phone;

    template.transmittalNo = formData.transmittalNo;
    template.fromName = formData.fromName;
    template.fromDepartment = formData.fromDepartment;
    template.toName = formData.toName;
    template.toDepartment = formData.toDepartment;
    template.toAddress = formData.toAddress;
    template.dateTransmitted = formData.dateTransmitted;
    template.items = formData.items;
    template.logoHtml = getLogoHtml();

    return template.evaluate().getContent();
  } catch(error) {
    Logger.log(`Error in getPrintPreviewHtml: ${error.toString()}`);
    throw new Error('Could not generate the print preview HTML.');
  }
}

/**
 * Gets the HTML content for the logo.
 * @returns {string} The logo's HTML content.
 */
function getLogoHtml() {
  return HtmlService.createHtmlOutputFromFile('logo_template.html').getContent();
}
