// Utility functions for Cold Email Generator

/**
 * Simulates a Claude API response for testing without using actual API credits
 * Set CONFIG.USE_SIMULATED_API = true to use this
 */
function simulateClaudeApiResponse(name, email, context) {
  // Create template email parts
  const greetings = [
    `Hi ${name},`,
    `Hello ${name},`,
    `Dear ${name},`
  ];
  
  const openings = [
    `I noticed your work at [Company] and your focus on ${context}.`,
    `After learning about your experience with ${context}, I wanted to reach out.`,
    `Your background in ${context} caught my attention.`
  ];
  
  const values = [
    `Our solution has helped similar professionals increase productivity by 30%.`,
    `We've developed a specialized approach for people in your position.`,
    `Many in your industry have found our methodology particularly effective.`
  ];
  
  const callToActions = [
    `Would you be open to a 15-minute call next week to discuss how we might help?`,
    `I'd love to schedule a brief demo tailored to your specific needs. How does your calendar look next Tuesday?`,
    `Could we connect for a quick conversation about how this might benefit your work?`
  ];
  
  const closings = [
    `Looking forward to your response,\n\nBest regards,\nYour Name\nYour Position\nPhone: (555) 123-4567`,
    `Thank you for your consideration,\n\nWarm regards,\nYour Name\nYour Company\nEmail: your.email@company.com`,
    `I appreciate your time,\n\nSincerely,\nYour Name\nYour Title\nWebsite: www.yourcompany.com`
  ];
  
  // Select random parts to create variety
  const greeting = greetings[Math.floor(Math.random() * greetings.length)];
  const opening = openings[Math.floor(Math.random() * openings.length)];
  const value = values[Math.floor(Math.random() * values.length)];
  const callToAction = callToActions[Math.floor(Math.random() * callToActions.length)];
  const closing = closings[Math.floor(Math.random() * closings.length)];
  
  // Assemble the email
  const emailBody = `Subject: Opportunity to Enhance Your ${context} Approach

${greeting}

${opening}

${value}

${callToAction}

${closing}`;

  // Simulate API delay
  Utilities.sleep(1000);
  
  return emailBody;
}

/**
 * Creates a template spreadsheet with the required columns
 */
function setupTemplateSpreadsheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Define column headers
  const headers = [
    'Recipient Name',
    'Recipient Email',
    'Context',
    'Generated Email',
    'Status'
  ];
  
  // Set column headers
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(1, i + 1).setValue(headers[i]);
  }
  
  // Format the header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Set column widths for text fields
  sheet.setColumnWidth(CONFIG.COLUMNS.CONTEXT + 1, 300);
  sheet.setColumnWidth(CONFIG.COLUMNS.GENERATED_EMAIL + 1, 500);
  
  // Add sample data
  const sampleData = [
    ['John Smith', 'john@example.com', 'CTO at ABC Tech, interested in AI solutions, previously mentioned challenges with data processing'],
    ['Jane Doe', 'jane@example.com', 'Marketing Director at XYZ Corp, looking to improve customer engagement, recently expanded to European market']
  ];
  
  // Set sample data
  sheet.getRange(2, 1, sampleData.length, 3).setValues(sampleData);
  
  SpreadsheetApp.getUi().alert('Template setup complete! Fill in the recipient information and use the "Generate Emails" option to create cold emails.');
}

/**
 * Validates the structure of the current sheet
 * Returns true if valid, false otherwise
 */
function validateSheetStructure() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  // Check if sheet has any data
  if (data.length === 0) {
    return false;
  }
  
  // If we expect headers, validate them
  if (CONFIG.HAS_HEADER) {
    const expectedHeaders = [
      'Recipient Name',
      'Recipient Email',
      'Context',
      'Generated Email',
      'Status'
    ];
    
    // Check if we have enough columns
    if (data[0].length < 3) {
      return false;
    }
    
    // Check header names (only required columns)
    for (let i = 0; i < 3; i++) {
      if (data[0][i] !== expectedHeaders[i]) {
        return false;
      }
    }
  }
  
  return true;
}

/**
 * Backs up the generated emails to a new sheet
 */
function backupGeneratedEmails() {
  const sourceSheet = SpreadsheetApp.getActiveSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH:mm:ss");
  const backupSheetName = "Backup_" + timestamp;
  
  // Create a new sheet for backup
  const backupSheet = ss.insertSheet(backupSheetName);
  
  // Get all data from source sheet
  const sourceRange = sourceSheet.getDataRange();
  const sourceValues = sourceRange.getValues();
  
  // Copy data to backup sheet
  backupSheet.getRange(1, 1, sourceValues.length, sourceValues[0].length).setValues(sourceValues);
  
  // Format backup sheet
  if (CONFIG.HAS_HEADER) {
    const headerRange = backupSheet.getRange(1, 1, 1, sourceValues[0].length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');
  }
  
  // Auto-resize columns
  for (let i = 1; i <= sourceValues[0].length; i++) {
    backupSheet.autoResizeColumn(i);
  }
  
  SpreadsheetApp.getUi().alert(`Backup created in sheet "${backupSheetName}"`);
} 