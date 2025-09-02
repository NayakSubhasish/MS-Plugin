/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Log that commands.js is being loaded
console.log('🚀 COMMANDS.JS LOADED - Email validation system initializing...');
console.log('📅 Load time:', new Date().toISOString());

// Import the BotAtWork API for generating enhanced emails
const BOTATWORK_API_KEY = "e80f5458c550f5b85ef56175b789468a";
const BOTATWORK_API_URL = "https://api.botatwork.com/trigger-task/b6f44edd-8140-4084-881e-2c11c403c082";

// Import getSuggestedReply function for consistent API calls
let getSuggestedReply;

// Try to import getSuggestedReply from the taskpane
try {
  if (typeof window !== 'undefined' && window.getSuggestedReply) {
    getSuggestedReply = window.getSuggestedReply;
    console.log('✅ getSuggestedReply imported from window');
  } else {
    console.log('⚠️ getSuggestedReply not found on window, will define locally');
  }
} catch (error) {
  console.error('❌ Error importing getSuggestedReply:', error);
}

// Local implementation of getSuggestedReply if import fails
if (!getSuggestedReply) {
  getSuggestedReply = async function(params) {
    console.log('🔧 Using local getSuggestedReply implementation');
    console.log('📤 API params:', params);
    
    try {
      const response = await fetch(BOTATWORK_API_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${BOTATWORK_API_KEY}`
        },
        body: JSON.stringify(params)
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      console.log('📥 API response:', data);
      
      return data.result || data.response || data.content || 'Enhanced email content generated';
    } catch (error) {
      console.error('❌ Local API call failed:', error);
      throw error;
    }
  };
  console.log('✅ Local getSuggestedReply implementation ready');
}
try {
  // Try to import the function (this will work in module environments)
  import('../taskpane/botAtWorkApi.js').then(module => {
    getSuggestedReply = module.getSuggestedReply;
  });
} catch (e) {
  console.log('Could not import getSuggestedReply, will use direct API calls');
}

// Define required keywords and their categories
// TODO: Make this configurable via API for different industries/companies
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

// Track validation state to ensure it triggers every time
let validationState = {
  lastValidationTime: 0,
  lastEmailBody: '',
  validationInProgress: false
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
  console.log('🚀🚀🚀 EMAIL VALIDATION TRIGGERED - SEND BUTTON INTERCEPTED! 🚀🚀🚀');
  console.log('🔥 SUCCESS: Send interception is working! 🔥');
  console.log('Event object:', event);
  console.log('📅 Timestamp:', new Date().toISOString());
  console.log('🔧 validateEmail function loaded and called');
  
  // Show alert to confirm interception is working
  console.log('🎯 SHOWING VALIDATION DIALOG...');
  console.log('🎯 Event type:', typeof event);
  console.log('🎯 Event completed function:', typeof event?.completed);
  
  // CRITICAL: Show alert to confirm validation is being called
  console.log('🚨 SHOWING ALERT TO CONFIRM VALIDATION IS CALLED');
  
  // CRITICAL: Store the event globally for debugging
  window.lastValidationEvent = event;
  
  // Reset validation state for new validation
  validationState.validationInProgress = false;
  
  // Add debugging information
  if (!event) {
    console.error('❌ No event object provided to validateEmail');
    return;
  }
  
  if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
    console.error('❌ Office context not available');
    if (event && typeof event.completed === 'function') {
      event.completed({ allowEvent: true });
    }
    return;
  }
  
  const item = Office.context.mailbox.item;
  
  // Get email body content
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get email body:', result.error);
      event.completed({ allowEvent: true });
      return;
    }
    
    const bodyText = result.value.toLowerCase();
    console.log('📧 Email body length:', bodyText.length);
    console.log('📧 Email body preview:', bodyText.substring(0, 200) + '...');
    
    // Update validation state
    validationState.lastEmailBody = bodyText;
    validationState.lastValidationTime = Date.now();
    
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
    
    console.log('🔍 Missing keywords found:', missing.length);
    if (missing.length > 0) {
      console.log('❌ Missing categories:', missing.map(m => m.category));
      console.log('🎯 SHOWING VALIDATION DIALOG FOR INCOMPLETE EMAIL');
      // Show validation dialog
      validationState.validationInProgress = true;
      showValidationDialog(missing, event);
    } else {
      console.log('✅ All keywords present, allowing send');
      // All keywords present, allow send
      if (event && typeof event.completed === 'function') {
        event.completed({ allowEvent: true });
      }
    }
    
    // Add a timeout fallback to ensure event is always completed
    setTimeout(() => {
      console.log('⏰ Validation timeout fallback - ensuring event completion');
      if (event && typeof event.completed === 'function' && !validationState.validationInProgress) {
        event.completed({ allowEvent: true });
      }
    }, 30000); // 30 second timeout
  });
}

// Show validation dialog with missing keywords
function showValidationDialog(missing, event) {
  const dialogUrl = getDialogUrl();
  console.log('🔍 Opening validation dialog:', dialogUrl);
  console.log('📊 Missing keywords data:', missing);
  console.log('🎯 Event object available:', !!event);
  
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { 
      height: 40, 
      width: 35,
      displayInIframe: true
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error('❌ Failed to open dialog:', asyncResult.error);
        event.completed({ allowEvent: true });
        return;
      }
      
      console.log('✅ Dialog opened successfully');
      const dialog = asyncResult.value;
      
      // Send missing keywords data to dialog with multiple attempts
      let attempts = 0;
      const maxAttempts = 5;
      
      const sendDataToDialog = () => {
        attempts++;
        console.log(`📤 Sending data to dialog (attempt ${attempts}):`, missing);
        try {
          dialog.messageChild(JSON.stringify({
            type: 'missing-keywords',
            data: missing
          }));
        } catch (error) {
          console.error('❌ Error sending message to dialog:', error);
          if (attempts < maxAttempts) {
            setTimeout(sendDataToDialog, 1000);
          }
        }
      };
      
      // Send initial message after dialog loads
      setTimeout(sendDataToDialog, 500);
      
      // Send backup messages
      setTimeout(sendDataToDialog, 1500);
      setTimeout(sendDataToDialog, 3000);
      
      // Handle dialog responses
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        console.log('📥 Dialog response received:', arg.message);
        const response = JSON.parse(arg.message);
        
        // Reset validation state
        validationState.validationInProgress = false;
        
        if (response.action === 'send') {
          console.log('✅ User chose to send anyway - ALLOWING SEND');
          if (event && typeof event.completed === 'function') {
            try {
              event.completed({ allowEvent: true });
              console.log('✅ event.completed({ allowEvent: true }) called successfully');
            } catch (error) {
              console.error('❌ Error calling event.completed for send:', error);
            }
          }
        } else if (response.action === 'add-keywords') {
          console.log('🔧 User wants to add keywords:', response.keywords);
          addKeywordsToEmail(response.keywords, event);
        } else if (response.action === 'cancel') {
          console.log('❌ User cancelled - BLOCKING SEND');
          console.log('🚫 Event completion with allowEvent: false');
          console.log('🎯 About to call event.completed({ allowEvent: false })');
          
          if (event && typeof event.completed === 'function') {
            try {
              event.completed({ allowEvent: false });
              console.log('✅ event.completed({ allowEvent: false }) called successfully');
            } catch (error) {
              console.error('❌ Error calling event.completed:', error);
            }
          }
        } else {
          console.log('❓ Unknown action:', response.action, '- BLOCKING SEND');
          console.log('🚫 Event completion with allowEvent: false for unknown action');
          
          if (event && typeof event.completed === 'function') {
            try {
              event.completed({ allowEvent: false });
              console.log('✅ event.completed({ allowEvent: false }) called successfully for unknown action');
            } catch (error) {
              console.error('❌ Error calling event.completed for unknown action:', error);
            }
          }
        }
        
        console.log('🔒 Closing dialog');
        dialog.close();
      });
      
      // Handle dialog errors
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        console.error('❌ Dialog error:', arg.error);
        // Reset validation state
        validationState.validationInProgress = false;
        event.completed({ allowEvent: true });
      });
    }
  );
}

// Add selected keywords to email and regenerate content
function addKeywordsToEmail(selectedKeywords, event) {
  console.log('🔧 ADD KEYWORDS TO EMAIL TRIGGERED');
  console.log('📝 Selected keywords:', selectedKeywords);
  console.log('🎯 Event object available:', !!event);
  
  const item = Office.context.mailbox.item;
  
  // Get current email body
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get email body for keyword addition:', result.error);
      if (event && typeof event.completed === 'function') {
        event.completed({ allowEvent: false });
      }
      return;
    }
    
    const currentBody = result.value;
    console.log('📧 Current email body:', currentBody.substring(0, 200) + '...');
    
    // Determine if this is a new email or reply
    const isReply = item.conversationId && item.conversationId.length > 0;
    console.log('📨 Is reply:', isReply);
    
    // Create enhanced prompt with all missing keywords
    const keywordsList = selectedKeywords.join(', ');
    
    // Get email subject using async method
    console.log('📧 Item object available:', !!item);
    console.log('📧 Getting subject asynchronously...');
    
    // Get subject asynchronously
    item.subject.getAsync((subjectResult) => {
      let emailSubject = '';
      
      if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
        emailSubject = subjectResult.value || '';
        console.log('📧 Subject retrieved successfully:', emailSubject);
      } else {
        console.log('📧 Failed to get subject:', subjectResult.error);
        emailSubject = '';
      }
      
      // Now create the enhancement prompt with the retrieved subject
      const enhancementPrompt = `Please enhance this email by adding the missing keywords while maintaining the original context and intent.

Subject: ${emailSubject}

Original Email:
${currentBody}

Missing Keywords to Include: ${keywordsList}

Please rewrite this email to include all the missing keywords naturally while preserving the original purpose and context from the subject and content.`;

      console.log('🤖 Enhancement prompt created');
      console.log('📧 Subject being passed:', emailSubject);
      console.log('📝 Body being passed:', currentBody.substring(0, 100) + '...');
      console.log('🏷️ Keywords being passed:', keywordsList);
      
      // Prepare API parameters - let BotAPI handle the enhancement
      const validationApiParams = {
        chooseATask: "emailResponse",
        emailContent: enhancementPrompt,
        tone: "professional",
        pointOfView: "organizationPerspective", 
        additionalInstructions: `Include these keywords: ${keywordsList}`,
        anonymize: null,
        incognito: false,
        default_language: "en-US",
        should_stream: false
      };

      // Call BotAtWork API directly for better reliability
      console.log('🚀 Calling BotAtWork API for email enhancement...');
      console.log('📧 Email subject:', emailSubject);
      console.log('📝 Original content:', currentBody);
      
      if (typeof getSuggestedReply === 'function') {
        console.log('✅ getSuggestedReply function available');
        getSuggestedReply(validationApiParams)
        .then(enhancedContent => {
          console.log('🎉 Enhanced content received (full):', enhancedContent);
          console.log('🔍 Content type:', typeof enhancedContent);
          console.log('🔍 Content length:', enhancedContent ? enhancedContent.length : 'null');
          console.log('🔍 Starts with API error?', enhancedContent ? enhancedContent.startsWith('API error:') : 'null');
          console.log('🔍 Starts with Error calling?', enhancedContent ? enhancedContent.startsWith('Error calling BotAtWork API:') : 'null');
          
          if (enhancedContent && !enhancedContent.startsWith('API error:') && !enhancedContent.startsWith('Error calling BotAtWork API:')) {
            // Update email body with enhanced content
            console.log('📝 ATTEMPTING TO UPDATE EMAIL BODY...');
            console.log('📧 Item object available:', !!item);
            console.log('📧 Item.body available:', !!item.body);
            console.log('📧 Item.body.setAsync available:', !!(item.body && item.body.setAsync));
            console.log('📧 Enhanced content:', enhancedContent);

            // Update email body with enhanced content - convert markdown to HTML for proper formatting
            const formattedHtml = convertMarkdownToOutlookHtml(enhancedContent);
            console.log('🔧 Converting markdown to HTML for Outlook...');
            console.log('📝 Original content:', enhancedContent.substring(0, 200) + '...');
            console.log('🔧 Formatted HTML:', formattedHtml.substring(0, 200) + '...');
            
            item.body.setAsync(
              formattedHtml,
              { coercionType: Office.CoercionType.Html },
              (setResult) => {
                console.log('📝 setAsync callback called');
                console.log('📝 setResult status:', setResult.status);
                console.log('📝 Expected success status:', Office.AsyncResultStatus.Succeeded);
                
                if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log('✅ EMAIL BODY UPDATED SUCCESSFULLY!');
                  console.log('🎯 Enhanced content was:', enhancedContent.substring(0, 300) + '...');

                  // Don't send automatically - let user review and send manually
                  if (event && typeof event.completed === 'function') {
                    event.completed({ allowEvent: false });
                  }
                } else {
                  console.error('❌ FAILED TO UPDATE EMAIL BODY');
                  console.error('❌ Error details:', setResult.error);
                  if (event && typeof event.completed === 'function') {
                    event.completed({ allowEvent: false });
                  }
                }
              }
            );
          } else {
            // Show the actual API error message to the user
            const errorMessage = enhancedContent || 'Unknown error occurred';
            console.error('❌ Enhancement failed - content rejected:', errorMessage);
            console.error('❌ Content was treated as error because it starts with error prefix');
            
            if (event && typeof event.completed === 'function') {
              event.completed({ allowEvent: false });
            }
          }
        })
        .catch(error => {
          console.error('❌ Error calling BotAtWork API:', error);
          if (event && typeof event.completed === 'function') {
            event.completed({ allowEvent: false });
          }
        });
    } else {
      console.error('❌ getSuggestedReply function not available - using simple fallback');
      
      // Simple dynamic fallback - just add basic structure to existing content
      const fallbackEnhanced = `Dear [Recipient],

${currentBody}

Missing elements that should be added: ${keywordsList}

Best regards,
[Your Name]`;

      console.log('🔄 ATTEMPTING FALLBACK EMAIL UPDATE...');
      console.log('📝 Fallback content:', fallbackEnhanced.substring(0, 200) + '...');
      
      // Convert fallback content to HTML for proper formatting
      const formattedFallbackHtml = convertMarkdownToOutlookHtml(fallbackEnhanced);
      console.log('🔧 Converting fallback content to HTML...');
      
      item.body.setAsync(
        formattedFallbackHtml,
        { coercionType: Office.CoercionType.Html },
        (setResult) => {
          console.log('📝 Fallback setAsync callback called');
          if (setResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✅ EMAIL ENHANCED WITH FALLBACK CONTENT!');
            if (event && typeof event.completed === 'function') {
              event.completed({ allowEvent: false });
            }
          } else {
            console.error('❌ FAILED TO UPDATE EMAIL BODY WITH FALLBACK');
            console.error('❌ Fallback error details:', setResult.error);
            if (event && typeof event.completed === 'function') {
              event.completed({ allowEvent: false });
            }
          }
        }
      );
      } // End of getSuggestedReply function check
    }); // End of subject.getAsync callback
  }); // End of body.getAsync callback
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

  // Call BotAtWork API to generate enhanced content
  const requestBody = {
    data: {
      payload: {
        chooseATask,
        description,
        additionalInstructions,
        tone,
        pointOfView
      }
    },
    anonymize,
    incognito,
    default_language,
    should_stream
  };

  // Use the centralized getSuggestedReply function from botAtWorkApi.js
  if (typeof getSuggestedReply !== 'undefined') {
    console.log('Using getSuggestedReply with apiParams:', apiParams);
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
    console.error('getSuggestedReply function not available - falling back to direct API call');
    
    console.log('BotAtWork API Request Body:', JSON.stringify(requestBody, null, 2));

    fetch(BOTATWORK_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': BOTATWORK_API_KEY
      },
      body: JSON.stringify(requestBody)
    })
    .then(response => response.json())
    .then(data => {
      console.log('BotAtWork API Response:', data);
      if (data && data.status === "SUCCESS" && data.data && data.data.content) {
        callback(data.data.content);
      } else if (data && data.response) {
        callback(data.response);
      } else {
        console.error('Invalid response from BotAtWork API:', data);
        callback('API error: Invalid response format from BotAtWork API');
      }
    })
    .catch(error => {
      console.error('Failed to generate enhanced email:', error);
      // Pass the actual API error message instead of null  
      const errorMessage = error.toString();
      callback(errorMessage.startsWith('API error:') || errorMessage.startsWith('Error calling BotAtWork API:') ? errorMessage : `Failed to generate enhanced email: ${errorMessage}`);
    });
  }
}

// Get dialog URL based on environment
function getDialogUrl() {
  const isLocalhost = window.location.hostname === 'localhost';
  const dialogUrl = isLocalhost 
    ? 'https://localhost:3000/validation-dialog.html'
    : 'https://nayaksubhasish.github.io/MS-Plugin/validation-dialog.html';
  
  console.log('🔗 Dialog URL:', dialogUrl);
  console.log('🌐 Current hostname:', window.location.hostname);
  console.log('🔧 Is localhost:', isLocalhost);
  return dialogUrl;
}

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
  console.log('Office.js is ready - registering email validation');
  
  // Make functions globally available
  window.validateEmail = validateEmail;
  console.log('validateEmail function registered globally');
  
  // CRITICAL: Register ItemSend event handler for send interception - DISABLED
  console.log('🔧 ItemSend event handler registration DISABLED');
  
  /*
  // Add delay to ensure Office context is fully initialized
  setTimeout(() => {
    try {
      if (Office.context && Office.context.mailbox && Office.context.mailbox.addHandlerAsync) {
        console.log('📧 Office context is available, registering ItemSend handler...');
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.ItemSend,
          validateEmail,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✅✅✅ ItemSend event handler registered successfully! Send button will be intercepted! ✅✅✅');
            } else {
              console.error('❌ Failed to register ItemSend event handler:', result.error);
            }
          }
        );
      } else {
        console.log('⚠️ Office context not fully available, trying again...');
        // Try again after another delay
        setTimeout(() => {
          if (Office.context && Office.context.mailbox && Office.context.mailbox.addHandlerAsync) {
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemSend, validateEmail);
            console.log('✅ Second attempt: ItemSend handler registered');
          }
        }, 2000);
      }
    } catch (error) {
      console.error('❌ Error registering ItemSend:', error);
    }
  }, 1000);
  */
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Make functions globally available
window.validateEmail = validateEmail;
window.action = action;

// Manual validation function that can be triggered by button
function validateEmailManual(event) {
  console.log('Manual email validation triggered');
  
  if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
    console.error('Office context not available');
    event.completed();
    return;
  }
  
  const item = Office.context.mailbox.item;
  
  // Get email body content
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get email body:', result.error);
      event.completed();
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
      showValidationDialogManual(missing, event);
    } else {
      // All keywords present
      const message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "✅ Email validation passed! All important elements are present.",
        icon: "Icon.80x80",
        persistent: true,
      };
      
      Office.context.mailbox.item?.notificationMessages.replaceAsync(
        "ValidationNotification",
        message
      );
      
      event.completed();
    }
  });
}

// Show validation dialog for manual validation
function showValidationDialogManual(missing, event) {
  const dialogUrl = getDialogUrl();
  
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { 
      height: 40, 
      width: 35,
      displayInIframe: true
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error('Failed to open dialog:', asyncResult.error);
        event.completed();
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
          // User chose to proceed
          const message = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Email validation completed. You can proceed with sending.",
            icon: "Icon.80x80",
            persistent: true,
          };
          
          Office.context.mailbox.item?.notificationMessages.replaceAsync(
            "ValidationNotification",
            message
          );
          
          event.completed();
        } else if (response.action === 'add-keywords') {
          // User wants to add keywords
          addKeywordsToEmailManual(response.keywords, event);
        } else {
          // User cancelled
          event.completed();
        }
        
        dialog.close();
      });
      
      // Handle dialog errors
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        console.error('Dialog error:', arg.error);
        event.completed();
      });
    }
  );
}

// Add keywords for manual validation
function addKeywordsToEmailManual(selectedKeywords, event) {
  const item = Office.context.mailbox.item;
  
  // Get current email body
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get email body for keyword addition:', result.error);
      event.completed();
      return;
    }
    
    const currentBody = result.value;
    
    // Determine if this is a new email or reply
    const isReply = item.conversationId && item.conversationId.length > 0;
    
    // Prepare dynamic API parameters for manual email validation
    const manualValidationApiParams = {
      chooseATask: "emailWrite",
      tone: "professional", // Can be made configurable based on user preference
      pointOfView: "organizationPerspective", // Can be made configurable
      additionalInstructions: "Enhance the email by naturally incorporating the missing keywords while preserving the original tone and message",
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
        console.log('🔧 Converting markdown to HTML for manual validation...');
        
        item.body.setAsync(
          formattedHtml,
          { coercionType: Office.CoercionType.Html },
          (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('Email enhanced with keywords');
              
              const message = {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: "✅ Email enhanced with selected keywords!",
                icon: "Icon.80x80",
                persistent: true,
              };
              
              Office.context.mailbox.item?.notificationMessages.replaceAsync(
                "ValidationNotification",
                message
              );
              
              event.completed();
            } else {
              console.error('Failed to update email body:', setResult.error);
              event.completed();
            }
          }
        );
      } else {
        // Show the actual API error message to the user
        const errorMessage = enhancedContent || 'Unknown error occurred';
        console.error('Enhancement failed:', errorMessage);
        
        const message = {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: errorMessage.startsWith('API error:') || errorMessage.startsWith('Error calling BotAtWork API:') ? errorMessage : `Failed to enhance email: ${errorMessage}`,
          icon: "Icon.80x80",
          persistent: true,
        };
        
        Office.context.mailbox.item?.notificationMessages.replaceAsync(
          "ValidationNotification",
          message
        );
        
        event.completed();
      }
    }, manualValidationApiParams);
  });
}

// Send with validation function - validates and sends if all good, or shows dialog if missing keywords
function sendWithValidation(event) {
  console.log('Send with validation triggered');
  
  if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
    console.error('Office context not available');
    event.completed();
    return;
  }
  
  const item = Office.context.mailbox.item;
  
  // Get email body content
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get email body:', result.error);
      event.completed();
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
      // Show validation dialog with send option
      showSendValidationDialog(missing, event);
    } else {
      // All keywords present, send directly
      sendEmail(event);
    }
  });
}

// Show validation dialog for send with validation
function showSendValidationDialog(missing, event) {
  const dialogUrl = getDialogUrl();
  
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { 
      height: 40, 
      width: 35,
      displayInIframe: true
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error('Failed to open dialog:', asyncResult.error);
        event.completed();
        return;
      }
      
      const dialog = asyncResult.value;
      
      // Send missing keywords data to dialog
      setTimeout(() => {
        dialog.messageChild(JSON.stringify({
          type: 'missing-keywords',
          data: missing,
          mode: 'send' // Indicate this is for send validation
        }));
      }, 1000);
      
      // Handle dialog responses
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const response = JSON.parse(arg.message);
        
        if (response.action === 'send') {
          // User chose to send anyway
          sendEmail(event);
        } else if (response.action === 'add-keywords') {
          // User wants to add keywords then send
          addKeywordsAndSend(response.keywords, event);
        } else {
          // User cancelled
          event.completed();
        }
        
        dialog.close();
      });
      
      // Handle dialog errors
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        console.error('Dialog error:', arg.error);
        event.completed();
      });
    }
  );
}

// Add keywords and then send email
function addKeywordsAndSend(selectedKeywords, event) {
  const item = Office.context.mailbox.item;
  
  // Get current email body
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get email body for keyword addition:', result.error);
      event.completed();
      return;
    }
    
    const currentBody = result.value;
    
    // Determine if this is a new email or reply
    const isReply = item.conversationId && item.conversationId.length > 0;
    
    // Prepare dynamic API parameters for send with validation
    const sendValidationApiParams = {
      chooseATask: "emailWrite",
      tone: "professional", // Can be made configurable
      pointOfView: "organizationPerspective", // Can be made configurable
      additionalInstructions: "Quickly enhance the email with missing keywords and prepare it for sending while maintaining professional tone",
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
        console.log('🔧 Converting markdown to HTML for send validation...');
        
        item.body.setAsync(
          formattedHtml,
          { coercionType: Office.CoercionType.Html },
          (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('Email enhanced with keywords, now sending...');
              // Send the enhanced email
              sendEmail(event);
            } else {
              console.error('Failed to update email body:', setResult.error);
              event.completed();
            }
          }
        );
      } else {
        // Show the actual API error message to the user
        const errorMessage = enhancedContent || 'Unknown error occurred';
        console.error('Enhancement failed:', errorMessage);
        
        // Failed to enhance, ask user if they want to send anyway
        const message = {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: errorMessage.startsWith('API error:') || errorMessage.startsWith('Error calling BotAtWork API:') ? `${errorMessage} Send anyway?` : `Failed to enhance email: ${errorMessage}. Send anyway?`,
          icon: "Icon.80x80",
          persistent: true,
        };
        
        Office.context.mailbox.item?.notificationMessages.replaceAsync(
          "ValidationNotification",
          message
        );
        
        event.completed();
      }
    }, sendValidationApiParams);
  });
}

// Send the email
function sendEmail(event) {
  if (Office.context && Office.context.mailbox && Office.context.mailbox.item) {
    // Show success message
    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "✅ Email validation completed. Sending email...",
      icon: "Icon.80x80",
      persistent: false,
    };
    
    Office.context.mailbox.item?.notificationMessages.replaceAsync(
      "ValidationNotification",
      message
    );
    
    // Send the email using Outlook's send method
    Office.context.mailbox.item.saveAsync((saveResult) => {
      if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
        // For compose mode, we can't directly send, but we can save and notify user
        const sendMessage = {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: "✅ Email ready to send! Please click the Send button in Outlook.",
          icon: "Icon.80x80",
          persistent: true,
        };
        
        Office.context.mailbox.item?.notificationMessages.replaceAsync(
          "SendNotification",
          sendMessage
        );
      }
    });
  }
  
  event.completed();
}

// Register the functions with Office
console.log('🔧 Registering Office.actions functions...');
Office.actions.associate("action", action);
console.log('✅ action function registered');
Office.actions.associate("validateEmail", validateEmail);
console.log('✅ validateEmail function registered');
Office.actions.associate("validateEmailManual", validateEmailManual);
console.log('✅ validateEmailManual function registered');
Office.actions.associate("sendWithValidation", sendWithValidation);
console.log('✅ sendWithValidation function registered');

// CRITICAL: Add alternative event registration for ItemSend - DISABLED
console.log('🔧 Alternative ItemSend registration DISABLED');
/*
try {
  // Try to register the event handler directly
  if (Office.context && Office.context.mailbox) {
    console.log('✅ Office context available for alternative registration');
    
    // Register ItemSend event handler
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemSend,
      validateEmail,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('✅ ItemSend event handler registered successfully');
        } else {
          console.error('❌ Failed to register ItemSend event handler:', result.error);
        }
      }
    );
  }
} catch (error) {
  console.error('❌ Error in alternative registration:', error);
}
*/

// Test function to manually trigger validation
function testValidation() {
  console.log('🧪 Testing validation manually');
  const testEvent = {
    completed: (result) => {
      console.log('✅ Test validation completed:', result);
    }
  };
  validateEmail(testEvent);
}

// Function to reset validation state
function resetValidationState() {
  console.log('🔄 Resetting validation state');
  validationState = {
    lastValidationTime: 0,
    lastEmailBody: '',
    validationInProgress: false
  };
  console.log('✅ Validation state reset:', validationState);
}

// Function to check validation state
function checkValidationState() {
  console.log('📊 Current validation state:', validationState);
  return validationState;
}

// Function to check if validation is properly registered
function checkValidationRegistration() {
  console.log('🔍 Checking validation registration...');
  console.log('✅ validateEmail function exists:', typeof validateEmail === 'function');
  console.log('✅ Office context available:', !!Office.context);
  console.log('✅ Mailbox available:', !!Office.context?.mailbox);
  console.log('✅ Current validation state:', validationState);
  console.log('🎯 Last validation event:', window.lastValidationEvent);
  
  // Test if we can access the email item
  if (Office.context?.mailbox?.item) {
    console.log('✅ Email item accessible');
    Office.context.mailbox.item.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log('✅ Email body accessible, length:', result.value.length);
      } else {
        console.log('❌ Email body not accessible:', result.error);
      }
    });
  } else {
    console.log('❌ Email item not accessible');
  }
}

// Function to test event completion manually
function testEventCompletion(allowEvent = false) {
  console.log('🧪 Testing event completion with allowEvent:', allowEvent);
  
  if (window.lastValidationEvent) {
    console.log('🎯 Using last validation event');
    try {
      window.lastValidationEvent.completed({ allowEvent });
      console.log('✅ Manual event.completed called successfully');
    } catch (error) {
      console.error('❌ Error in manual event.completed:', error);
    }
  } else {
    console.log('❌ No validation event available for testing');
  }
}

// Make functions globally available
window.testValidation = testValidation;
window.resetValidationState = resetValidationState;
window.checkValidationState = checkValidationState;
window.checkValidationRegistration = checkValidationRegistration;
window.testEventCompletion = testEventCompletion;
