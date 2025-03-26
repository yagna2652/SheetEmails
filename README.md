# Cold Email Generator

A Google Sheets extension that uses AI (Claude or DeepSeek) to generate personalized cold emails based on spreadsheet data and saves them as Gmail drafts.

## Features

- Reads recipient information (name, email, and context) from Google Sheets
- Uses AI (Claude 3.7 Sonnet or DeepSeek) to generate personalized, persuasive cold emails
- Supports multiple AI providers (Claude and DeepSeek) with easy switching
- Saves generated emails to both the spreadsheet and as Gmail drafts
- Encrypted API key storage for enhanced security
- Simulation mode to test without using API credits
- Easy setup with menu options for configuration
- Includes error handling and status reporting

## Setup Instructions

### 1. Create a New Google Sheet

Start by creating a new Google Sheet where you'll store your recipient data.

### 2. Open the Apps Script Editor

1. In your Google Sheet, go to **Extensions > Apps Script**.
2. This will open the Apps Script editor in a new tab.

### 3. Add the Code Files

1. In the Apps Script editor, rename the default `Code.gs` file to `main.js` and replace its content with the content of `main.js` from this repository.
2. Create additional script files by clicking the "+" button next to "Files":
   - Create `config.js` and paste the content from the config file.
   - Create `utils.js` and paste the content from the utils file.
   - Create `appsscript.json` by clicking on "Project Settings" (gear icon) and then "Show "appsscript.json" manifest file in editor". Paste the content from the appsscript.json file.

### 4. Get an API Key

You can choose between Claude API or DeepSeek API based on your preference and budget:

#### Claude API
1. Go to [Anthropic's Console](https://console.anthropic.com/) and sign up for an account.
2. Navigate to the API section and generate an API key.
3. You'll get a free trial with $5 worth of credits (approximately 125,000 input tokens and 25,000 output tokens).
4. After using your free credits, you'll need to add payment information to continue using the API.

#### DeepSeek API (More Cost-Effective)
1. Go to [DeepSeek's website](https://platform.deepseek.com/) and sign up for an account.
2. Navigate to the API section and generate an API key.
3. DeepSeek typically offers more competitive pricing than Claude, making it a more economical choice for generating cold emails at scale.

**Note:** API pricing may change, so check the providers' websites for the most current information.

### 5. Save and Deploy

1. Save all your files (Ctrl+S or Cmd+S).
2. Click on "Deploy" and select "New deployment" or "Test deployment".
3. Choose "Web app" as the deployment type.
4. Set "Who has access" to "Only myself" or appropriate for your organization.
5. Click "Deploy" and authorize the app when prompted.
6. Return to your Google Sheet and refresh the page.
7. You should now see a new menu item called "Cold Email Generator".

### 6. Configure the Sheet and Set API Key

1. From the "Cold Email Generator" menu, select "Set API Key".
2. Choose your preferred API provider (Claude or DeepSeek).
3. Enter your API key when prompted.
4. Enter a secure password to encrypt your API key. This password will be required each time you generate emails.
5. Use the "Setup Template" option to create the required columns with sample data.
6. (Optional) Enable "Simulation Mode" to test without using API credits.

## Switching Between API Providers

You can easily switch between Claude and DeepSeek:

1. From the "Cold Email Generator" menu, select "Switch API Provider".
2. The system will switch to the other provider.
3. If you haven't set up the API key for the selected provider yet, use the "Set API Key" option.

This allows you to compare results or switch to the more cost-effective option as needed.

## Using Simulation Mode

To preserve your Claude API credits, the extension includes a simulation mode:

1. From the "Cold Email Generator" menu, select "Toggle Simulation Mode".
2. When enabled, the system will generate simulated emails instead of calling the Claude API.
3. This is useful for:
   - Testing the extension's functionality
   - Developing and customizing the extension
   - Generating templates when you're out of API credits
   - Classroom demonstrations

The simulated emails use templates with randomized elements and the recipient data from your sheet, but won't have the same quality or personalization as real Claude-generated emails.

## Security Features

This extension includes enhanced security for your API key:

- Your API key is encrypted using HMAC-SHA256 before storage
- A password is required each time you generate emails
- The decrypted key is never stored persistently
- The encryption uses a combination of your password and Google's built-in security
- Even if someone gains access to your Google account, they would still need your encryption password to use your API key

## Usage Instructions

### Preparing Your Data

1. Fill in the required columns in your Google Sheet:
   - **Recipient Name**: The name of the person you're emailing (Column A)
   - **Recipient Email**: Their email address (Column B)
   - **Context**: Information about the recipient for personalization (Column C)

### Generating Emails

1. From the "Cold Email Generator" menu, select "Generate Emails".
2. Enter your encryption password when prompted.
3. The extension will process each row and:
   - Generate a personalized cold email using Claude
   - Save the generated email in the "Generated Email" column (Column D)
   - Create a draft in Gmail with the recipient's email address
   - Update the "Status" column (Column E)

### Managing Results

- You can review the generated emails in the spreadsheet before they are sent.
- All emails are saved as drafts in Gmail, so you can review, edit, or schedule them before sending.
- Use the "Backup Generated Emails" function to create a backup sheet with your results.

## Example Prompt Used

The system uses the following prompt template to generate cold emails:

```
You are an expert at writing persuasive, personalized cold emails. 
Generate a professional cold email for [Recipient Name] at [Recipient Email].

Context about the recipient: [Context]

The email should:
- Be concise (no more than 150 words)
- Have a compelling subject line
- Include a clear, specific call to action
- Be personalized based on the provided context
- Have a professional but conversational tone
- Not be pushy or overly salesy

Format the email with proper line breaks and a professional signature.
```

You can customize this prompt in the `generateEmailWithClaude()` function in `main.js`.

## Troubleshooting

- **API Key Issues**: Make sure you've set your Claude API key correctly.
- **Rate Limits**: If you hit rate limits, try generating fewer emails at once.
- **OAuth Permissions**: If you're getting permission errors, make sure you've accepted all the required permissions when prompted.
- **Script Errors**: Check the Apps Script logs for detailed error messages.

## Security Considerations

- Your Claude API key is stored in the script properties and is not visible to others with access to the sheet.
- All data remains within your Google Workspace and Anthropic's API.
- Only you can access the generated Gmail drafts, even if the sheet is shared.

## Customization

You can customize:
- The email generation prompt in `main.js`
- The column structure in `config.js`
- The formatting and template in `utils.js` # SheetEmails
