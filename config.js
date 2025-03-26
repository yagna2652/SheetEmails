// Configuration settings for the Cold Email Generator
const CONFIG = {
  // API settings
  CLAUDE_API_KEY: "your-api-key-here", // Replace with your actual API key
  CLAUDE_API_URL: "https://api.anthropic.com/v1/messages",
  CLAUDE_API_VERSION: "2023-06-01",
  CLAUDE_MODEL: "claude-3-sonnet-20240229", // Update to the latest model as needed
  
  // DeepSeek API settings
  DEEPSEEK_API_KEY: "your-deepseek-api-key-here", // You'll enter this through the UI for security
  DEEPSEEK_API_URL: "https://api.deepseek.com/v1/chat/completions",
  DEEPSEEK_MODEL: "deepseek-chat", // Or other available model
  
  // API Selection (Options: 'CLAUDE', 'DEEPSEEK')
  SELECTED_API: "DEEPSEEK", // Change to select which API to use
  
  // Use simulated API instead of real API (saves credits)
  USE_SIMULATED_API: false, // Set to true to use the simulator
  
  // Email settings
  DEFAULT_SUBJECT_PREFIX: "Cold Email: ",
  MAX_TOKENS: 500,
  
  // Sheet column indices (0-based)
  COLUMNS: {
    RECIPIENT_NAME: 0,    // Column A
    RECIPIENT_EMAIL: 1,   // Column B
    CONTEXT: 2,           // Column C
    GENERATED_EMAIL: 3,   // Column D
    STATUS: 4             // Column E
  },
  
  // Sheet has header row
  HAS_HEADER: true
}; 