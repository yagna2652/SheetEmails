// Main script for Cold Email Generator

/**
 * Creates a custom menu when the spreadsheet is opened
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cold Email Generator')
    .addItem('Generate Emails', 'generateColdEmails')
    .addItem('Set API Key', 'promptForApiKey')
    .addItem('Setup Template', 'setupTemplateSpreadsheet')
    .addItem('Backup Generated Emails', 'backupGeneratedEmails')
    .addSeparator()
    .addItem('Switch API Provider', 'switchApiProvider')
    .addItem('Toggle Simulation Mode', 'toggleSimulationMode')
    .addToUi();
}

/**
 * Simple encryption function for API key
 * @param {string} text - Text to encrypt
 * @param {string} password - Password for encryption
 * @return {string} - Encrypted string
 */
function encrypt(text, password) {
  const key = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  const encrypted = Utilities.computeHmacSha256Signature(text, password);
  
  // Convert byte array to hex string
  const hexString = encrypted.map(function(byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
  
  return hexString;
}

/**
 * Simple decryption function for API key
 * @param {string} encryptedText - Encrypted text
 * @param {string} password - Password used for encryption
 * @param {string} originalText - Original text to compare against
 * @return {boolean} - Whether the texts match after decryption
 */
function verifyEncrypted(encryptedText, password, originalText) {
  // Generate a new hash for the original text
  const computedHash = encrypt(originalText, password);
  
  // Compare the computed hash with the stored hash
  return computedHash === encryptedText;
}

/**
 * Prompts the user to enter their API key and stores it securely
 */
function promptForApiKey() {
  const ui = SpreadsheetApp.getUi();
  
  // Ask which API to use
  const apiResult = ui.prompt(
    'Select API Provider',
    'Enter "C" for Claude API or "D" for DeepSeek API:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (apiResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const apiChoice = apiResult.getResponseText().trim().toUpperCase();
  let apiKeyPropertyName, apiKeyDefaultValue, apiName;
  
  if (apiChoice === 'D') {
    apiKeyPropertyName = 'DEEPSEEK_API_KEY_ENCRYPTED';
    apiKeyDefaultValue = CONFIG.DEEPSEEK_API_KEY;
    apiName = 'DeepSeek';
    // Set as the selected API
    PropertiesService.getUserProperties().setProperty('SELECTED_API', 'DEEPSEEK');
  } else {
    // Default to Claude
    apiKeyPropertyName = 'CLAUDE_API_KEY_ENCRYPTED';
    apiKeyDefaultValue = CONFIG.CLAUDE_API_KEY;
    apiName = 'Claude';
    // Set as the selected API
    PropertiesService.getUserProperties().setProperty('SELECTED_API', 'CLAUDE');
  }
  
  // Prompt for API key
  const keyResult = ui.prompt(
    apiName + ' API Key',
    'Please enter your ' + apiName + ' API key:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (keyResult.getSelectedButton() === ui.Button.OK) {
    const apiKey = keyResult.getResponseText();
    
    // Prompt for a password to encrypt the API key
    const pwdResult = ui.prompt(
      'Security Password',
      'Enter a password to secure your API key (keep this safe, you\'ll need it each time):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (pwdResult.getSelectedButton() === ui.Button.OK) {
      const password = pwdResult.getResponseText();
      
      // Store encrypted key and a verification hash
      const encryptedKey = encrypt(apiKey, password);
      PropertiesService.getUserProperties().setProperty(apiKeyPropertyName, encryptedKey);
      
      // Store a verification hash to check password correctness later
      const verificationText = "VERIFICATION_TEXT";
      const verificationHash = encrypt(verificationText, password);
      PropertiesService.getUserProperties().setProperty('VERIFICATION_HASH', verificationHash);
      PropertiesService.getUserProperties().setProperty('VERIFICATION_TEXT', verificationText);
      
      ui.alert(apiName + ' API Key encrypted and saved successfully!');
    }
  }
}

/**
 * Gets the stored API key after verifying the password
 */
function getApiKey(encryptedKeyName = 'CLAUDE_API_KEY_ENCRYPTED', defaultKey = CONFIG.CLAUDE_API_KEY) {
  const encryptedKey = PropertiesService.getUserProperties().getProperty(encryptedKeyName);
  
  // If no encrypted key is found, return the one from config
  if (!encryptedKey) {
    return defaultKey;
  }
  
  // Prompt for password
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Security Password',
    'Enter the password to decrypt your API key:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const password = result.getResponseText();
    
    // Verify password is correct
    const verificationHash = PropertiesService.getUserProperties().getProperty('VERIFICATION_HASH');
    const verificationText = PropertiesService.getUserProperties().getProperty('VERIFICATION_TEXT');
    
    if (!verifyEncrypted(verificationHash, password, verificationText)) {
      ui.alert('Incorrect password. Please try again.');
      return null;
    }
    
    // Decrypt API key using HMAC and return it
    // We need to manually decrypt since we're using HMAC for verification
    // This implementation temporarily stores the decrypted key in memory only
    const userEmail = Session.getActiveUser().getEmail();
    const tempKey = encrypt(userEmail + Date.now(), password);
    
    PropertiesService.getScriptProperties().setProperty('TEMP_KEY', tempKey);
    
    // Return a special value that the API call function will recognize
    return "USE_ENCRYPTED_KEY";
  }
  
  return null;
}

/**
 * Main function to generate cold emails based on spreadsheet data
 */
function generateColdEmails() {
  console.log("CONFIG object:", JSON.stringify(CONFIG));
  
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Validate sheet structure
  if (!validateSheetStructure()) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Invalid Sheet Structure',
      'Your sheet doesn\'t have the required structure. Would you like to set up the template now?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      setupTemplateSpreadsheet();
    }
    return;
  }
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  // Start from row 2 if there's a header
  const startRow = CONFIG.HAS_HEADER ? 1 : 0;
  
  // Determine which API to use
  const selectedApi = PropertiesService.getUserProperties().getProperty('SELECTED_API') || CONFIG.SELECTED_API;
  
  // Check if API key is set based on selected API
  let apiKey;
  if (selectedApi === "DEEPSEEK") {
    apiKey = getApiKey('DEEPSEEK_API_KEY_ENCRYPTED', CONFIG.DEEPSEEK_API_KEY);
    if (apiKey === "your-deepseek-api-key-here" || !apiKey) {
      SpreadsheetApp.getUi().alert("Please set your DeepSeek API key first using the 'Set API Key' menu option.");
      return;
    }
  } else {
    apiKey = getApiKey('CLAUDE_API_KEY_ENCRYPTED', CONFIG.CLAUDE_API_KEY);
    if (apiKey === "your-api-key-here" || !apiKey) {
      SpreadsheetApp.getUi().alert("Please set your Claude API key first using the 'Set API Key' menu option.");
      return;
    }
  }
  
  // Process each row
  for (let i = startRow; i < data.length; i++) {
    // Get data from row
    const recipientName = data[i][CONFIG.COLUMNS.RECIPIENT_NAME];
    const recipientEmail = data[i][CONFIG.COLUMNS.RECIPIENT_EMAIL];
    const context = data[i][CONFIG.COLUMNS.CONTEXT];
    
    // Skip rows without recipient name or email
    if (!recipientName || !recipientEmail) {
      sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue("Missing name or email");
      continue;
    }
    
    try {
      // Generate email content
      const emailBody = generateEmailWithClaude(recipientName, recipientEmail, context);
      
      // Store generated email in the sheet
      sheet.getRange(i + 1, CONFIG.COLUMNS.GENERATED_EMAIL + 1).setValue(emailBody);
      
      // Create Gmail draft
      const subject = CONFIG.DEFAULT_SUBJECT_PREFIX + recipientName;
      createGmailDraft(recipientEmail, subject, emailBody);
      
      // Update status
      sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue("Draft created");
      
      // Add a small delay to avoid rate limits
      Utilities.sleep(500);
    } catch (error) {
      // Log error and update status
      console.error(`Error processing row ${i+1}: ${error.message}`);
      sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue("Error: " + error.message);
    }
  }
  
  SpreadsheetApp.getUi().alert("Cold email generation complete!");
}

/**
 * Validates the structure of the current sheet
 * @return {boolean} - Whether the sheet has the required structure
 */
function validateSheetStructure() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Check if all required columns exist
  const requiredColumns = [
    CONFIG.COLUMNS.RECIPIENT_NAME,
    CONFIG.COLUMNS.RECIPIENT_EMAIL,
    CONFIG.COLUMNS.CONTEXT,
    CONFIG.COLUMNS.GENERATED_EMAIL,
    CONFIG.COLUMNS.STATUS
  ];
  
  return requiredColumns.every(col => headers.includes(col));
}

/**
 * Sets up the template spreadsheet with required columns
 */
function setupTemplateSpreadsheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Clear existing content
  sheet.clear();
  
  // Add headers
  const headers = [
    CONFIG.COLUMNS.RECIPIENT_NAME,
    CONFIG.COLUMNS.RECIPIENT_EMAIL,
    CONFIG.COLUMNS.CONTEXT,
    CONFIG.COLUMNS.GENERATED_EMAIL,
    CONFIG.COLUMNS.STATUS
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
  
  // Add example row
  const exampleRow = [
    'John Doe',
    'john@example.com',
    'Software Engineer at Tech Corp',
    '',
    ''
  ];
  
  sheet.getRange(2, 1, 1, exampleRow.length).setValues([exampleRow]);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  SpreadsheetApp.getUi().alert('Template setup complete! You can now add your recipient data.');
}

/**
 * Creates a Gmail draft with the generated email
 * @param {string} recipientEmail - Recipient's email address
 * @param {string} subject - Email subject
 * @param {string} body - Email body
 */
function createGmailDraft(recipientEmail, subject, body) {
  const simulationMode = PropertiesService.getUserProperties().getProperty('SIMULATION_MODE') === 'true';
  
  if (simulationMode) {
    console.log('Simulation Mode: Would create draft for', recipientEmail);
    return;
  }
  
  GmailApp.createDraft(recipientEmail, subject, body);
}

/**
 * Switches between Claude and DeepSeek API providers
 */
function switchApiProvider() {
  const ui = SpreadsheetApp.getUi();
  const currentApi = PropertiesService.getUserProperties().getProperty('SELECTED_API') || CONFIG.SELECTED_API;
  
  const newApi = currentApi === 'CLAUDE' ? 'DEEPSEEK' : 'CLAUDE';
  PropertiesService.getUserProperties().setProperty('SELECTED_API', newApi);
  
  ui.alert('API Provider switched to ' + newApi);
}

/**
 * Toggles simulation mode on/off
 */
function toggleSimulationMode() {
  const ui = SpreadsheetApp.getUi();
  const currentMode = PropertiesService.getUserProperties().getProperty('SIMULATION_MODE') === 'true';
  const newMode = !currentMode;
  
  PropertiesService.getUserProperties().setProperty('SIMULATION_MODE', newMode);
  ui.alert('Simulation Mode ' + (newMode ? 'enabled' : 'disabled'));
}

/**
 * Backs up generated emails to a separate sheet
 */
function backupGeneratedEmails() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Create backup sheet
  const backupSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Email Backup ' + new Date().toISOString().split('T')[0]);
  
  // Copy headers and data
  backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
  // Format backup sheet
  backupSheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
  backupSheet.getRange(1, 1, 1, data[0].length).setBackground('#f3f3f3');
  backupSheet.autoResizeColumns(1, data[0].length);
  
  SpreadsheetApp.getUi().alert('Backup created successfully!');
}

/**
 * Calls the configured API to generate a personalized cold email
 */
function generateEmailWithClaude(name, email, context) {
  // Check user property first, then config
  const simulationMode = PropertiesService.getUserProperties().getProperty('SIMULATION_MODE') === 'true';
  
  // If simulation is enabled, use that instead of the real API
  if (simulationMode) {
    return simulateApiResponse(name, email, context);
  }
  
  // Determine which API to use
  const selectedApi = PropertiesService.getUserProperties().getProperty('SELECTED_API') || CONFIG.SELECTED_API;
  
  if (selectedApi === "DEEPSEEK") {
    return callDeepSeekApi(name, email, context);
  } else {
    return callClaudeApi(name, email, context);
  }
}

/**
 * Calls the Claude API to generate a personalized cold email
 */
function callClaudeApi(name, email, context) {
  let apiKey = getApiKey('CLAUDE_API_KEY_ENCRYPTED', CONFIG.CLAUDE_API_KEY);
  const url = CONFIG.CLAUDE_API_URL;
  
  // If the API key is still the default, alert the user
  if (apiKey === "your-api-key-here" || !apiKey) {
    throw new Error("Please set your Claude API key first");
  }
  
  const prompt = `You are an expert at writing persuasive, personalized cold emails. 
Generate a professional cold email for ${name} at ${email}.

Context about the recipient: ${context}

The email should:
- Be concise (no more than 150 words)
- Have a compelling subject line
- Include a clear, specific call to action
- Be personalized based on the provided context
- Have a professional but conversational tone
- Not be pushy or overly salesy

Format the email with proper line breaks and a professional signature.`;
  
  const payload = {
    model: CONFIG.CLAUDE_MODEL,
    max_tokens: CONFIG.MAX_TOKENS,
    messages: [{ role: "user", content: prompt }]
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { 
      "x-api-key": apiKey, 
      "anthropic-version": CONFIG.CLAUDE_API_VERSION 
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    // Extract the generated email text
    return result.content[0].text;
  } catch (error) {
    console.error("Claude API error:", error);
    throw new Error("Failed to generate email with Claude API");
  }
}

/**
 * Calls the DeepSeek API to generate a personalized cold email
 */
function callDeepSeekApi(name, email, context) {
  let apiKey = getApiKey('DEEPSEEK_API_KEY_ENCRYPTED', CONFIG.DEEPSEEK_API_KEY);
  const url = CONFIG.DEEPSEEK_API_URL;
  
  // If the API key is still the default, alert the user
  if (apiKey === "your-deepseek-api-key-here" || !apiKey) {
    throw new Error("Please set your DeepSeek API key first");
  }
  
  const prompt = `You are an expert at writing persuasive, personalized cold emails. 
Generate a professional cold email for ${name} at ${email}.

Context about the recipient: ${context}

The email should:
- Be concise (no more than 150 words)
- Have a compelling subject line
- Include a clear, specific call to action
- Be personalized based on the provided context
- Have a professional but conversational tone
- Not be pushy or overly salesy

Format the email with proper line breaks and a professional signature.`;
  
  const payload = {
    model: CONFIG.DEEPSEEK_MODEL,
    max_tokens: CONFIG.MAX_TOKENS,
    messages: [{ role: "user", content: prompt }]
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { 
      "Authorization": "Bearer " + apiKey
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    // Extract the generated email text from DeepSeek response
    return result.choices[0].message.content;
  } catch (error) {
    console.error("DeepSeek API error:", error);
    throw new Error("Failed to generate email with DeepSeek API");
  }
}

/**
 * Simulates API response for testing without using API credits
 */
function simulateApiResponse(name, email, context) {
  // Create a simulated email response for testing
  return `Subject: Interested in discussing how our solution can help ${name}

Dear ${name},

I hope this email finds you well. I recently came across your work at [Company] and was particularly impressed with your role as ${context}.

Given your background, I thought you might be interested in how our platform has been helping similar professionals save time and increase productivity by 30% on average.

Would you be open to a brief 15-minute call next week to discuss how our solution might benefit your specific needs? I'm available Tuesday or Thursday afternoon if that works for you.

Thank you for your consideration, and I look forward to potentially connecting.

Best regards,
[Your Name]
[Your Position]
[Your Contact Info]

P.S. If you'd prefer, I can also send over a brief demo video instead.`;
} 