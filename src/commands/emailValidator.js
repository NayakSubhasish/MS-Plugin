// Email validation logic for checking missing keywords on send
Office.onReady(() => {
  console.log('Email validator loaded');
});

// Define required keywords and their categories
const REQUIRED_KEYWORDS = {
  agenda: {
    keywords: ['agenda', 'meeting agenda', 'topics to discuss', 'discussion points'],
    category: 'Meeting Structure',
    suggestions: ['agenda', 'meeting topics', 'discussion items']
  },
  action: {
    keywords: ['action item', 'action required', 'next steps', 'follow up', 'todo'],
    category: 'Action Items',
    suggestions: ['action items', 'next steps', 'follow-up required']
  },
  timeline: {
    keywords: ['deadline', 'due date', 'timeline', 'schedule', 'by when'],
    category: 'Timeline',
    suggestions: ['deadline', 'timeline', 'completion date']
  },
  participants: {
    keywords: ['attendees', 'participants', 'who should attend', 'invitees'],
    category: 'Participants',
    suggestions: ['attendees', 'participants', 'invitees']
  },
  greeting: {
    keywords: ['dear', 'hello', 'hi', 'good morning', 'good afternoon'],
    category: 'Greeting',
    suggestions: ['Dear', 'Hello', 'Hi']
  },
  closing: {
    keywords: ['regards', 'best regards', 'sincerely', 'thank you', 'thanks'],
    category: 'Closing',
    suggestions: ['Best regards', 'Thank you', 'Sincerely']
  }
};

// Helper function to convert markdown to proper HTML for Outlook
function convertMarkdownToOutlookHtml(markdownText) {
  if (!markdownText || typeof markdownText !== 'string') {
    return markdownText || '';
  }

  let html = markdownText
    // Convert bullet points (both - and * formats)
    .replace(/^[\s]*[-*][\s]+(.+)$/gm, '<li>$1</li>')
    // Convert numbered lists
    .replace(/^[\s]*\d+\.\s+(.+)$/gm, '<li>$1</li>')
    // Wrap lists in proper ul/ol tags
    .replace(/(<li>.*<\/li>)/gs, '<ul>$1</ul>')
    // Convert bold text (**text** or __text__)
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/__(.*?)__/g, '<strong>$1</strong>')
    // Convert italic text (*text* or _text_)
    .replace(/\*(.*?)\*/g, '<em>$1</em>')
    .replace(/_(.*?)_/g, '<em>$1</em>')
    // Convert line breaks to proper HTML
    .replace(/\n\n+/g, '</p><p>')
    .replace(/\n/g, '<br/>')
    // Wrap in paragraph tags
    .replace(/^(.+)$/s, '<p>$1</p>')
    // Clean up empty paragraphs and double tags
    .replace(/<p><\/p>/g, '')
    .replace(/<p><p>/g, '<p>')
    .replace(/<\/p><\/p>/g, '</p>')
    // Fix list formatting - ensure proper spacing
    .replace(/<\/ul>\s*<ul>/g, '</ul><ul>')
    .replace(/<\/p>\s*<ul>/g, '</p><ul>')
    .replace(/<\/ul>\s*<p>/g, '</ul><p>');

  return html;
}

// Main validation function called on email send
function validateEmail(event) {
  console.log('Email validation triggered');
  
  const item = Office.context.mailbox.item;
  
  // Get email body content
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get email body:', result.error);
      event.completed({ allowEvent: true });
      return;
    }
    
    const bodyText = result.value.toLowerCase();
    const missing = [];
    
    // Check for missing keywords
    Object.keys(REQUIRED_KEYWORDS).forEach(key => {
      const keywordData = REQUIRED_KEYWORDS[key];
      const hasKeyword = keywordData.keywords.some(keyword => 
        bodyText.includes(keyword.toLowerCase())
      );
      
      if (!hasKeyword) {
        missing.push({
          category: keywordData.category,
          suggestions: keywordData.suggestions,
          key: key
        });
      }
    });
    
    if (missing.length > 0) {
      // Show validation dialog
      showValidationDialog(missing, event);
    } else {
      // All keywords present, allow send
      event.completed({ allowEvent: true });
    }
  });
}

// Show validation dialog with missing keywords
function showValidationDialog(missing, event) {
  const dialogUrl = getDialogUrl();
  
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { 
      height: 60, 
      width: 50,
      displayInIframe: true
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error('Failed to open dialog:', asyncResult.error);
        event.completed({ allowEvent: true });
        return;
      }
      
      const dialog = asyncResult.value;
      
      // Send missing keywords data to dialog
      setTimeout(() => {
        dialog.messageChild(JSON.stringify({
          type: 'missing-keywords',
          data: missing
        }));
      }, 1000);
      
      // Handle dialog responses
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const response = JSON.parse(arg.message);
        
        if (response.action === 'send') {
          // User chose to send anyway
          event.completed({ allowEvent: true });
        } else if (response.action === 'add-keywords') {
          // User wants to add keywords
          addKeywordsToEmail(response.keywords, event);
        } else {
          // User cancelled
          event.completed({ allowEvent: false });
        }
        
        dialog.close();
      });
      
      // Handle dialog errors
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        console.error('Dialog error:', arg.error);
        event.completed({ allowEvent: true });
      });
    }
  );
}

// Add selected keywords to email and regenerate content
function addKeywordsToEmail(selectedKeywords, event) {
  const item = Office.context.mailbox.item;
  
  // Get current email body
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get email body for keyword addition:', result.error);
      event.completed({ allowEvent: false });
      return;
    }
    
    const currentBody = result.value;
    
    // Determine if this is a new email or reply
    const isReply = item.conversationId && item.conversationId.length > 0;
    
    // Prepare dynamic API parameters for email validator
    const emailValidatorApiParams = {
      chooseATask: "emailWrite",
      tone: "professional", // Can be made configurable based on email context
      pointOfView: "organizationPerspective", // Can be made configurable
      additionalInstructions: "Naturally incorporate missing keywords while maintaining the original email's intent and tone",
      anonymize: null,
      incognito: false,
      default_language: "en-US",
      should_stream: false
    };

    // Generate enhanced content with keywords
    generateEnhancedEmail(currentBody, selectedKeywords, isReply, (enhancedContent) => {
      if (enhancedContent && !enhancedContent.startsWith('API error:') && !enhancedContent.startsWith('Failed to generate enhanced email:') && !enhancedContent.startsWith('Error calling BotAtWork API:')) {
        // Update email body with enhanced content - convert markdown to HTML for proper formatting
        const formattedHtml = convertMarkdownToOutlookHtml(enhancedContent);
        console.log('🔧 Converting markdown to HTML for email validation...');
        
        item.body.setAsync(
          formattedHtml,
          { coercionType: Office.CoercionType.Html },
          (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('Email enhanced with keywords');
              // Don't send automatically - let user review and send manually
              event.completed({ allowEvent: false });
            } else {
              console.error('Failed to update email body:', setResult.error);
              event.completed({ allowEvent: false });
            }
          }
        );
      } else {
        // Show the actual API error message to the user
        const errorMessage = enhancedContent || 'Unknown error occurred';
        console.error('Enhancement failed:', errorMessage);
        event.completed({ allowEvent: false });
      }
    }, emailValidatorApiParams);
  });
}

// Generate enhanced email content with missing keywords
function generateEnhancedEmail(originalContent, selectedKeywords, isReply, callback, apiParams = {}) {
  // Create prompt for AI to enhance the email
  const keywordsList = selectedKeywords.join(', ');
  const emailType = isReply ? 'reply' : 'new email';
  
  const enhancementPrompt = `Please enhance this ${emailType} by naturally incorporating these missing elements: ${keywordsList}

Original content:
${originalContent}

Instructions:
- Keep the original tone and intent
- Naturally integrate the missing elements
- Make it professional and coherent
- Don't change the core message, just enhance it with the missing elements

Enhanced ${emailType}:`;

  // Extract dynamic parameters with defaults - use emailResponse for enhancing existing content
  const {
    chooseATask = "emailResponse",
    description = enhancementPrompt,
    emailContent = enhancementPrompt, // For emailResponse tasks
    additionalInstructions = "",
    tone = "professional",
    pointOfView = "organizationPerspective",
    anonymize = null,
    incognito = false,
    default_language = "en-US",
    should_stream = false
  } = apiParams;

  console.log('Email Validator API Parameters:', {
    chooseATask,
    tone,
    pointOfView,
    additionalInstructions,
    default_language
  });

  // Use the existing BotAtWork API to generate enhanced content
  if (typeof getSuggestedReply !== 'undefined') {
    // Pass the prompt with dynamic parameters
    getSuggestedReply(description, 3, apiParams)
      .then(enhancedContent => {
        callback(enhancedContent);
      })
      .catch(error => {
        console.error('Failed to generate enhanced email:', error);
        // Pass the actual API error message instead of null
        const errorMessage = error.toString();
        callback(errorMessage.startsWith('API error:') || errorMessage.startsWith('Error calling BotAtWork API:') ? errorMessage : `Failed to generate enhanced email: ${errorMessage}`);
      });
  } else {
    console.error('getSuggestedReply function not available');
    callback('Error: getSuggestedReply function not available');
  }
}

// Get dialog URL based on environment
function getDialogUrl() {
  const isLocalhost = window.location.hostname === 'localhost';
  const baseUrl = isLocalhost 
    ? 'https://localhost:3000' 
    : 'https://nayaksubhasish.github.io/MS-Plugin';
  
  return `${baseUrl}/validation-dialog.html`;
}

// Export for manifest registration
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { validateEmail };
} else {
  // Global registration for Office
  window.validateEmail = validateEmail;
} 