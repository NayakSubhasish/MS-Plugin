// BotAtWork API utility for email generation and suggestions
const BOTATWORK_API_KEY = "e80f5458c550f5b85ef56175b789468a";
const BOTATWORK_API_URL = "https://api.botatwork.com/trigger-task/b6f44edd-8140-4084-881e-2c11c403c082";
const DEBUG_LOGS_ENABLED = false; // Toggle verbose logs

// Helper function to get logged-in user's email
const getLoggedInUserEmail = () => {
  try {
    if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) {
      return Office.context.mailbox.userProfile.emailAddress;
    }
  } catch (error) {
    if (DEBUG_LOGS_ENABLED) {
      console.warn('Failed to get user email:', error);
    }
  }
  return "unknown@unknown.com";
};

// Normalize and clamp prompt length to reduce payload and latency
const clampText = (text, maxLen = 4000) => {
  if (!text) return "";
  return text.length > maxLen ? text.slice(0, maxLen) + "..." : text;
};

// Simple in-memory cache to avoid duplicate roundtrips during a session
const responseCache = new Map();

// Track last in-flight request to cancel when a new one starts
let lastController = null;

// Helper function to add delay
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Helper function to check if error is retryable
const isRetryableError = (error) => {
  if (!error) return false;
  
  const errorMessage = (error.message || error.toString() || '').toLowerCase();
  
  // Always retry HTTP 5xx errors (server errors) as they are usually transient
  if (errorMessage.includes('500') || errorMessage.includes('502') || errorMessage.includes('503') || 
      errorMessage.includes('504') || errorMessage.includes('http 5')) {
    return true;
  }
  
  const retryableMessages = [
    'overloaded',
    'rate limit',
    'quota exceeded',
    'service unavailable',
    'internal error',
    'timeout',
    'network error',
    'connection',
    'server error',
    '5'  // HTTP 5xx errors
  ];
  
  return retryableMessages.some(msg => errorMessage.includes(msg));
};

// Helper function to determine task type and format payload
const formatPayload = (prompt, taskType = 'emailWrite', apiParams = {}) => {
  // Extract dynamic parameters with defaults
  const {
    additionalInstructions = "",
    tone = "formal",
    pointOfView = "organizationPerspective",
    description: apiDescription = null,
    emailContent: apiEmailContent = null
  } = apiParams;

  // Get logged-in user's email for hiddenValue
  const userEmail = getLoggedInUserEmail();
  const hiddenValue = `{From person: ${userEmail}}`;

  if (DEBUG_LOGS_ENABLED) {
    console.log('BotAtWork API - formatPayload received apiParams:', apiParams);
    console.log('BotAtWork API - Extracted parameters:', { tone, pointOfView, additionalInstructions });
    console.log('BotAtWork API - User email for hiddenValue:', userEmail);
  }

  // Clean and structure the prompt for better API understanding
  const cleanPrompt = clampText(prompt.replace(/^(Write a professional email with the following details:|Suggest a professional reply to this email considering the specified tone and point of view:|Please provide a helpful response to this follow-up question or request:)/i, '').trim());

  // For email writing tasks (new emails)
  if (taskType === 'emailWrite') {
    // Extract description from the prompt, but use dynamic tone and pointOfView
    // Updated regex to capture multi-line descriptions by using [\s\S]*? (non-greedy match of any character including newlines)
    const descriptionMatch = prompt.match(/Description:\s*([\s\S]*?)(?=\nTone:|$)/i);
    
    // Always use the dynamic parameters for tone and pointOfView, not extracted from prompt
    return {
      chooseATask: "emailWrite",
      description: clampText(descriptionMatch ? descriptionMatch[1].trim() : (apiDescription != null ? apiDescription : cleanPrompt)),
      //additionalInstructions: additionalInstructions || "Write a professional email that matches the specified tone and perspective exactly. Use 'we', 'our', 'us' for organization perspective and 'I', 'my', 'me' for individual perspective.",
     // tone: tone, // Always use the dynamic tone parameter
     // pointOfView: pointOfView === "individualPerspective" ? "individualPerspective" : "organizationPerspective", // Map to API expected values
      hiddenValue: hiddenValue,
      "run_on_behalf": userEmail,
      "emailPreferencePrompt": true
    };
  }

  // For email response tasks (replies, suggestions, enhancements)
  if (taskType === 'emailResponse') {
    return {
      chooseATask: "emailResponse",
      emailContent: clampText(apiEmailContent != null ? apiEmailContent : cleanPrompt),
      additionalInstructions: additionalInstructions || "",
     // tone,
     // pointOfView: pointOfView === "individualPerspective" ? "individualPerspective" : "organizationPerspective",
      hiddenValue: hiddenValue,
      "run_on_behalf": userEmail,
      "emailPreferencePrompt": true
    };
  }

  // For email rewrite tasks (editing existing emails)
  if (taskType === 'emailRewrite') {
    return {
      chooseATask: "emailRewrite",
      emailContent: clampText(apiEmailContent != null ? apiEmailContent : cleanPrompt),
      additionalInstructions: additionalInstructions || "",
     // tone,
     // pointOfView: pointOfView === "individualPerspective" ? "individualPerspective" : "organizationPerspective",
      hiddenValue: hiddenValue,
      "run_on_behalf": userEmail,
      "emailPreferencePrompt": true
    };
  }

  // For chat/conversation tasks
  if (taskType === 'chat') {
    return {
      chooseATask: "emailWrite", // Use emailWrite for chat as it's more flexible
      description: clampText(cleanPrompt),
     // additionalInstructions: additionalInstructions || "Provide a helpful and conversational response. Use 'we', 'our', 'us' for organization perspective and 'I', 'my', 'me' for individual perspective.",
     // tone,
     // pointOfView: pointOfView === "individualPerspective" ? "individualPerspective" : "organizationPerspective",
      hiddenValue: hiddenValue,
      "run_on_behalf": userEmail,
      "emailPreferencePrompt": true
    };
  }

  // Default fallback
  return {
    chooseATask: "emailWrite",
    description: clampText(apiDescription != null ? apiDescription : cleanPrompt),
   // additionalInstructions: additionalInstructions || "Write a professional email that matches the specified tone and perspective. Use 'we', 'our', 'us' for organization perspective and 'I', 'my', 'me' for individual perspective.",
   // tone,
    //pointOfView: pointOfView === "individualPerspective" ? "individualPerspective" : "organizationPerspective",
    hiddenValue: hiddenValue,
    "run_on_behalf": userEmail,
    "emailPreferencePrompt": true
  };
};

export async function getSuggestedReply(prompt, maxRetries = 10, apiParams = {}) {
  // Extract dynamic parameters with defaults
  const {
    chooseATask = "emailWrite",
    description = prompt,
    emailContent = prompt, // For emailResponse tasks
    additionalInstructions = "",
    tone = "formal",
    pointOfView = "organizationPerspective",
    anonymize = null,
    incognito = false,
    default_language = "en-US",
    should_stream = false
  } = apiParams;

  // Get logged-in user's email for run_on_behalf
  const userEmail = getLoggedInUserEmail();

  // Create payload based on task type using formatPayload
  const payload = formatPayload(clampText(prompt), chooseATask, apiParams);
  
  const requestBody = {
    data: {
      payload: payload
    },
    anonymize,
    incognito,
    default_language,
    should_stream,
    "run_on_behalf": userEmail,
    "emailPreferencePrompt": true
  };

  // Return cached response if available
  const cacheKey = JSON.stringify({ payload, anonymize, incognito, default_language });
  if (responseCache.has(cacheKey)) {
    if (DEBUG_LOGS_ENABLED) console.log('BotAtWork API - returning cached response');
    return responseCache.get(cacheKey);
  }

  if (DEBUG_LOGS_ENABLED) {
    console.log('BotAtWork API - Using dynamic parameters:', {
      chooseATask,
      tone,
      pointOfView,
      additionalInstructions,
      default_language
    });
    console.log('BotAtWork API - Final payload being sent:', JSON.stringify(payload, null, 2));
  }

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      if (DEBUG_LOGS_ENABLED) {
        console.log(`BotAtWork API attempt ${attempt}/${maxRetries}`);
        console.log("Request payload:", JSON.stringify(requestBody, null, 2));
      }

      // Cancel any previous in-flight request (only when a NEW request starts)
      if (lastController) {
        try { lastController.abort('new-request'); } catch (_) {}
      }
      const controller = new AbortController();
      lastController = controller;
      // IMPORTANT: NO TIMEOUT - Let the plugin wait as long as needed for the bot to generate
      // This is especially important for long emails (5000+ words) that can take 30+ seconds
      // We want to wait until the server responds, no matter how long it takes

      const response = await fetch(BOTATWORK_API_URL, {
        method: "POST",
        headers: { 
          "Content-Type": "application/json",
          "x-api-key": BOTATWORK_API_KEY
        },
        body: JSON.stringify(requestBody),
        signal: controller.signal
      });
      // No timeout to clear since we allow long-running requests
      
      if (!response.ok) {
        // For HTTP 5xx errors, treat them as retryable server errors
        if (response.status >= 500 && response.status < 600) {
          const error = new Error(`HTTP ${response.status}: ${response.statusText}`);
          error.isServerError = true; // Mark as server error for retry logic
          throw error;
        }
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
      
      const data = await response.json();
      if (DEBUG_LOGS_ENABLED) {
        console.log("BotAtWork API raw response:", data);
      }
      
      // Check for successful response - BotAtWork API format
      if (data && data.status === "SUCCESS" && data.data && data.data.content) {
        responseCache.set(cacheKey, data.data.content);
        return data.data.content;
      }
      
      // Fallback: Check for other possible response formats
      if (data && (data.result || data.response || data.output || data.content)) {
        const result = data.result || data.response || data.output || data.content;
        const normalized = typeof result === 'string' ? result : JSON.stringify(result);
        responseCache.set(cacheKey, normalized);
        return normalized;
      }
      
      // If data has a message field, use that
      if (data && data.message) {
        responseCache.set(cacheKey, data.message);
        return data.message;
      }
      
      // If data is a string, return it directly
      if (typeof data === 'string') {
        responseCache.set(cacheKey, data);
        return data;
      }
      
      // Check for API errors - BotAtWork API format
      if (data && data.status !== "SUCCESS") {
        const errorMessage = data.message || data.status || "Unknown error";
        if (DEBUG_LOGS_ENABLED) {
          console.log(`BotAtWork API error on attempt ${attempt}:`, errorMessage);
        }
        
        // Check if this is a retryable error
        if (isRetryableError({ message: errorMessage }) && attempt < maxRetries) {
          const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000; // Exponential backoff with jitter
          if (DEBUG_LOGS_ENABLED) {
            console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
          }
          await delay(backoffDelay);
          continue;
        }
        
        // If not retryable or max retries reached, return error
        return `API error: ${errorMessage}${attempt > 1 ? ` (after ${attempt} attempts)` : ''}`;
      }
      
      // Check for generic API errors
      if (data && data.error) {
        const errorMessage = data.error.message || data.error.toString() || JSON.stringify(data.error);
        if (DEBUG_LOGS_ENABLED) {
          console.log(`BotAtWork API error on attempt ${attempt}:`, errorMessage);
        }
        
        // Check if this is a retryable error
        if (isRetryableError(data.error) && attempt < maxRetries) {
          const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000; // Exponential backoff with jitter
          if (DEBUG_LOGS_ENABLED) {
            console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
          }
          await delay(backoffDelay);
          continue;
        }
        
        // If not retryable or max retries reached, return error
        return `API error: ${errorMessage}${attempt > 1 ? ` (after ${attempt} attempts)` : ''}`;
      }
      
      // If we get here, we have data but couldn't parse it
      if (DEBUG_LOGS_ENABLED) {
        console.log("Unexpected response format:", data);
      }
      return "Response received but format unexpected. Raw response: " + JSON.stringify(data);
      
    } catch (e) {
      if (DEBUG_LOGS_ENABLED) {
        console.log(`Network error on attempt ${attempt}:`, e.message);
      }
      
      // Always retry server errors (HTTP 5xx) as they are usually transient
      if (e.isServerError || e.message.includes('500') || e.message.includes('502') || 
          e.message.includes('503') || e.message.includes('504')) {
        if (attempt < maxRetries) {
          const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000;
          if (DEBUG_LOGS_ENABLED) {
            console.log(`Server error detected, retrying in ${Math.round(backoffDelay)}ms... (attempt ${attempt}/${maxRetries})`);
          }
          await delay(backoffDelay);
          continue;
        }
      }
      
      // Check if this is a retryable error (includes network issues, timeouts, or other transient errors)
      if (attempt < maxRetries && isRetryableError(e)) {
        const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000;
        if (DEBUG_LOGS_ENABLED) {
          console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
        }
        await delay(backoffDelay);
        continue;
      }
      
      // Max retries reached or non-retryable error
      return `Error calling BotAtWork API: ${e.message} (after ${attempt} attempts)`;
    }
  }
  
  // This should never be reached, but just in case
  return "Maximum retry attempts reached. Please try again later.";
}
