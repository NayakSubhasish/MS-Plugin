// BotAtWork API utility for email generation and suggestions
const BOTATWORK_API_KEY = "e80f5458c550f5b85ef56175b789468a";
const BOTATWORK_API_URL = "https://api.botatwork.com/trigger-task/b6f44edd-8140-4084-881e-2c11c403c082";

// Helper function to add delay
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Helper function to check if error is retryable
const isRetryableError = (error) => {
  if (!error) return false;
  
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
  
  const errorMessage = (error.message || error.toString() || '').toLowerCase();
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

  console.log('BotAtWork API - formatPayload received apiParams:', apiParams);
  console.log('BotAtWork API - Extracted parameters:', { tone, pointOfView, additionalInstructions });

  // Clean and structure the prompt for better API understanding
  const cleanPrompt = prompt.replace(/^(Write a professional email with the following details:|Suggest a professional reply to this email considering the specified tone and point of view:|Please provide a helpful response to this follow-up question or request:)/i, '').trim();

  // For email writing tasks (new emails)
  if (taskType === 'emailWrite') {
    // Extract description from the prompt, but use dynamic tone and pointOfView
    // Updated regex to capture multi-line descriptions by using [\s\S]*? (non-greedy match of any character including newlines)
    const descriptionMatch = prompt.match(/Description:\s*([\s\S]*?)(?=\nTone:|$)/i);
    
    // Always use the dynamic parameters for tone and pointOfView, not extracted from prompt
    return {
      chooseATask: "emailWrite",
      description: descriptionMatch ? descriptionMatch[1].trim() : (apiDescription != null ? apiDescription : cleanPrompt),
      additionalInstructions: additionalInstructions || "Write a professional email that matches the specified tone and perspective exactly. Use 'we', 'our', 'us' for organization perspective and 'I', 'my', 'me' for individual perspective.",
      tone: tone, // Always use the dynamic tone parameter
      pointOfView: pointOfView === "individualPerspective" ? "firstPerson" : "organizationPerspective" // Map to API expected values
    };
  }

  // For email response tasks (replies, suggestions, enhancements)
  if (taskType === 'emailResponse') {
    return {
      chooseATask: "emailResponse",
      emailContent: apiEmailContent != null ? apiEmailContent : cleanPrompt,
      additionalInstructions: additionalInstructions || "Provide a relevant and contextual response that matches the specified tone and perspective. Use 'we', 'our', 'us' for organization perspective and 'I', 'my', 'me' for individual perspective.",
      tone,
      pointOfView: pointOfView === "individualPerspective" ? "firstPerson" : "organizationPerspective"
    };
  }

  // For chat/conversation tasks
  if (taskType === 'chat') {
    return {
      chooseATask: "emailWrite", // Use emailWrite for chat as it's more flexible
      description: cleanPrompt,
      additionalInstructions: additionalInstructions || "Provide a helpful and conversational response. Use 'we', 'our', 'us' for organization perspective and 'I', 'my', 'me' for individual perspective.",
      tone,
      pointOfView: pointOfView === "individualPerspective" ? "firstPerson" : "organizationPerspective"
    };
  }

  // Default fallback
  return {
    chooseATask: "emailWrite",
    description: apiDescription != null ? apiDescription : cleanPrompt,
    additionalInstructions: additionalInstructions || "Write a professional email that matches the specified tone and perspective. Use 'we', 'our', 'us' for organization perspective and 'I', 'my', 'me' for individual perspective.",
    tone,
    pointOfView: pointOfView === "individualPerspective" ? "firstPerson" : "organizationPerspective"
  };
};

export async function getSuggestedReply(prompt, maxRetries = 3, apiParams = {}) {
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

  // Create payload based on task type using formatPayload
  const payload = formatPayload(prompt, chooseATask, apiParams);
  
  const requestBody = {
    data: {
      payload: payload
    },
    anonymize,
    incognito,
    default_language,
    should_stream
  };

  console.log('BotAtWork API - Using dynamic parameters:', {
    chooseATask,
    tone,
    pointOfView,
    additionalInstructions,
    default_language
  });
  
  console.log('BotAtWork API - Final payload being sent:', JSON.stringify(payload, null, 2));

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`BotAtWork API attempt ${attempt}/${maxRetries}`);
      console.log("Request payload:", JSON.stringify(requestBody, null, 2));
      
      const response = await fetch(BOTATWORK_API_URL, {
        method: "POST",
        headers: { 
          "Content-Type": "application/json",
          "x-api-key": BOTATWORK_API_KEY
        },
        body: JSON.stringify(requestBody),
      });
      
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
      
      const data = await response.json();
      console.log("BotAtWork API raw response:", data);
      
      // Check for successful response - BotAtWork API format
      if (data && data.status === "SUCCESS" && data.data && data.data.content) {
        return data.data.content;
      }
      
      // Fallback: Check for other possible response formats
      if (data && (data.result || data.response || data.output || data.content)) {
        const result = data.result || data.response || data.output || data.content;
        return typeof result === 'string' ? result : JSON.stringify(result);
      }
      
      // If data has a message field, use that
      if (data && data.message) {
        return data.message;
      }
      
      // If data is a string, return it directly
      if (typeof data === 'string') {
        return data;
      }
      
      // Check for API errors - BotAtWork API format
      if (data && data.status !== "SUCCESS") {
        const errorMessage = data.message || data.status || "Unknown error";
        console.log(`BotAtWork API error on attempt ${attempt}:`, errorMessage);
        
        // Check if this is a retryable error
        if (isRetryableError({ message: errorMessage }) && attempt < maxRetries) {
          const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000; // Exponential backoff with jitter
          console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
          await delay(backoffDelay);
          continue;
        }
        
        // If not retryable or max retries reached, return error
        return `API error: ${errorMessage}${attempt > 1 ? ` (after ${attempt} attempts)` : ''}`;
      }
      
      // Check for generic API errors
      if (data && data.error) {
        const errorMessage = data.error.message || data.error.toString() || JSON.stringify(data.error);
        console.log(`BotAtWork API error on attempt ${attempt}:`, errorMessage);
        
        // Check if this is a retryable error
        if (isRetryableError(data.error) && attempt < maxRetries) {
          const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000; // Exponential backoff with jitter
          console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
          await delay(backoffDelay);
          continue;
        }
        
        // If not retryable or max retries reached, return error
        return `API error: ${errorMessage}${attempt > 1 ? ` (after ${attempt} attempts)` : ''}`;
      }
      
      // If we get here, we have data but couldn't parse it
      console.log("Unexpected response format:", data);
      return "Response received but format unexpected. Raw response: " + JSON.stringify(data);
      
    } catch (e) {
      console.log(`Network error on attempt ${attempt}:`, e.message);
      
      // Check if this is a retryable network error
      if (attempt < maxRetries && (e.message.includes('fetch') || e.message.includes('network') || e.message.includes('timeout'))) {
        const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000;
        console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
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
