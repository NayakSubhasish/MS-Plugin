import * as React from "react";
import PropTypes from "prop-types";
import PromptConfig from "./PromptConfig";
import { Button, makeStyles, tokens, FluentProvider, teamsLightTheme, teamsDarkTheme, Switch, Label, Tab, TabList } from "@fluentui/react-components";
import { getSuggestedReply } from "../botAtWorkApi";

const useStyles = makeStyles({
  root: {
    height: "100vh",
    background: "linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "8px",
    boxSizing: "border-box",
    overflow: "hidden",
    scrollbarWidth: "none",
    msOverflowStyle: "none",
    "&::-webkit-scrollbar": {
      display: "none",
    },
  },
  // Loading spinner styles - dynamic rotating loader
  spinner: {
    position: "relative",
    width: "80px",
    height: "80px",
  },
  spinnerRing1: {
    position: "absolute",
    width: "100%",
    height: "100%",
    borderRadius: "50%",
    border: "3px solid transparent",
    borderTop: "3px solid #0078d4",
    borderRight: "3px solid rgba(0, 120, 212, 0.3)",
    animation: "$rotate 1.5s linear infinite",
  },
  spinnerRing2: {
    position: "absolute",
    width: "70%",
    height: "70%",
    top: "15%",
    left: "15%",
    borderRadius: "50%",
    border: "2px solid transparent",
    borderTop: "2px solid #106ebe",
    borderLeft: "2px solid rgba(16, 110, 190, 0.3)",
    animation: "$rotate 1s linear infinite reverse",
  },
  spinnerRing3: {
    position: "absolute",
    width: "40%",
    height: "40%",
    top: "30%",
    left: "30%",
    borderRadius: "50%",
    border: "2px solid transparent",
    borderBottom: "2px solid #005a9e",
    borderRight: "2px solid rgba(0, 90, 158, 0.3)",
    animation: "$rotate 0.8s linear infinite",
  },
  spinnerCenter: {
    position: "absolute",
    width: "20%",
    height: "20%",
    top: "40%",
    left: "40%",
    borderRadius: "50%",
    backgroundColor: "#0078d4",
    animation: "$pulse 1.2s ease-in-out infinite",
  },
  "@keyframes rotate": {
    "0%": {
      transform: "rotate(0deg)",
    },
    "100%": {
      transform: "rotate(360deg)",
    },
  },
  "@keyframes pulse": {
    "0%, 100%": {
      transform: "scale(1)",
      opacity: "1",
    },
    "50%": {
      transform: "scale(1.2)",
      opacity: "0.7",
    },
  },
  // Global spin animation for button spinners
  "@global": {
    "@keyframes spin": {
      "0%": {
        transform: "rotate(0deg)",
      },
      "100%": {
        transform: "rotate(360deg)",
      },
    },
  },
  loadingContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    color: "#605e5c",
    textAlign: "center",
    padding: "30px",
  },
  loadingText: {
    marginTop: "16px",
    fontSize: "14px",
    fontWeight: "500",
    color: "#323130",
  },
  headerContainer: {
    width: "100%",
    display: "flex",
    justifyContent: "flex-end",
    alignItems: "center",
    padding: "0 8px",
    boxSizing: "border-box",
    marginBottom: "4px",
  },
  tabContainer: {
    width: "100%",
    maxWidth: "500px",
    marginBottom: "4px",
  },
  contentArea: {
    width: "100%",
    minHeight: "200px",
    maxHeight: "60vh", // Use viewport height for better responsiveness
    background: "linear-gradient(145deg, #e8e8e8, #d4d4d4)",
    borderRadius: "8px",
    padding: "16px",
    color: "#2d2d2d",
    fontSize: "15px",
    lineHeight: "1.7",
    boxShadow: "0 4px 20px rgba(0,0,0,0.12), 0 1px 3px rgba(0,0,0,0.15)",
    wordBreak: "break-word",
    overflowY: "auto", // Enable vertical scrolling
    overflowX: "hidden", // Hide horizontal overflow
    flex: "1 1 auto",
    display: "block", // Changed from flex to block for better text flow
    border: "1px solid rgba(0,0,0,0.12)",
    margin: "4px 8px 4px 8px",
    transition: "all 0.3s ease",
    maxWidth: "none",
    // Custom scrollbar styling for better UX
    scrollbarWidth: "thin",
    scrollbarColor: "rgba(0,0,0,0.3) transparent",
    "&::-webkit-scrollbar": {
      width: "6px",
    },
    "&::-webkit-scrollbar-track": {
      background: "transparent",
    },
    "&::-webkit-scrollbar-thumb": {
      backgroundColor: "rgba(0,0,0,0.3)",
      borderRadius: "3px",
    },
    "&::-webkit-scrollbar-thumb:hover": {
      backgroundColor: "rgba(0,0,0,0.5)",
    },
    "&:hover": {
      boxShadow: "0 8px 30px rgba(0,0,0,0.18), 0 2px 6px rgba(0,0,0,0.15)",
    },
  },
  gridButton: {
    minHeight: "44px",
    fontSize: "14px",
    fontWeight: "600",
    borderRadius: "4px",
    border: "1px solid rgba(0,120,212,0.2)",
    backgroundColor: "rgba(255,255,255,0.9)",
    color: "#323130",
    transition: "all 0.3s ease",
    cursor: "pointer",
    backdropFilter: "blur(10px)",
    boxShadow: "0 2px 8px rgba(0,0,0,0.05)",
    "&:hover": {
      backgroundColor: "rgba(255,255,255,1)",
      borderColor: "rgba(0,120,212,0.4)",
      transform: "translateY(-1px)",
      boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
    },
    "&:active": {
      transform: "translateY(0px)",
    },
  },
  activeButton: {
    background: "linear-gradient(135deg, #0078d4, #106ebe)",
    borderColor: "rgba(0,120,212,0.3)",
    color: "#ffffff",
    boxShadow: "0 4px 15px rgba(0,120,212,0.3)",
    "&:hover": {
      background: "linear-gradient(135deg, #106ebe, #005a9e)",
      transform: "translateY(-1px)",
      boxShadow: "0 6px 20px rgba(0,120,212,0.4)",
    },
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [generatedContent, setGeneratedContent] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [templates, setTemplates] = React.useState([]);
  const [activeTab, setActiveTab] = React.useState("writeEmail");
  const [showWriteEmailForm, setShowWriteEmailForm] = React.useState(true);
  const [emailForm, setEmailForm] = React.useState({
    description: "",
    // additionalInstructions: "", // Commented out as per user request
    tone: "formal",
    pointOfView: "Organization perspective"
  });

  // const [isDarkMode, setIsDarkMode] = React.useState(false); // dark mode temporarily disabled
  const [customPrompts, setCustomPrompts] = React.useState({
    suggestReply: "Email to reply to:\n{emailBody}\n\nWrite a professional reply with:\nTone: {tone}\nPoint of View: {pointOfView}",
    summarize: "Summarize this email in 2 sentences:\n{emailBody}",
    writeEmail: "Write a professional email:\nDescription: {description}\nTone: {tone}\nPoint of View: {pointOfView}",
  });
  const [chatInput, setChatInput] = React.useState("");
  const [chatHistory, setChatHistory] = React.useState([]);
  const [isFirstResponse, setIsFirstResponse] = React.useState(true);
  const [showEnhanceButton, setShowEnhanceButton] = React.useState(false);
  const [missingKeywords, setMissingKeywords] = React.useState([]);

  // Responsive header style
  const headerTitle = "SalesGenie AI";
  const headerLogo = "assets/logo-filled.webp";
  

  // Helper to get email body from Outlook
  let getEmailBody = () => {
    return new Promise((resolve, reject) => {
      if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        const item = Office.context.mailbox.item;
        
        // Check if we're in compose mode or read mode
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
          // For reading emails, get the current email body
          item.body.getAsync("text", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value);
            } else {
              reject("Failed to get email body.");
            }
          });
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
          // For appointments, get the body
          item.body.getAsync("text", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value);
            } else {
              reject("Failed to get appointment body.");
            }
          });
        } else {
          reject("Unsupported item type.");
        }
      } else {
        reject("Office.js not available or not in Outlook context.");
      }
    });
  };

  // Helper to get email subject
  let getEmailSubject = () => {
    return new Promise((resolve, reject) => {
      if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        const item = Office.context.mailbox.item;
        
        // For compose mode, use subject.getAsync
        if (item.itemType === Office.MailboxEnums.ItemType.Message && item.subject && item.subject.getAsync) {
          item.subject.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || "No subject");
            } else {
              resolve("No subject");
            }
          });
        } 
        // For read mode, subject might be directly available
        else if (typeof item.subject === 'string') {
          resolve(item.subject);
        }
        // Fallback - try to get from normalizedSubject
        else if (item.normalizedSubject) {
          resolve(item.normalizedSubject);
        }
        else {
          resolve("No subject");
        }
      } else {
        reject("Office.js not available or not in Outlook context.");
      }
    });
  };

  // Helper to get conversation thread (previous emails)
  let getConversationThread = () => {
    return new Promise((resolve, reject) => {
      if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        const item = Office.context.mailbox.item;
        
        // Get conversation thread if available
        if (item.conversationId) {
          // For now, we'll use the current item's subject to identify the thread
          // In a full implementation, you'd query the conversation
          resolve(`Thread: ${item.subject || 'No subject'}\n\nCurrent email content will be processed.`);
        } else {
          resolve("No conversation thread available.");
        }
      } else {
        reject("Office.js not available or not in Outlook context.");
      }
    });
  };

  // Fire-and-forget usage logger for plugin actions
  const logPluginUsage = (action) => {
    try {
      const userEmail = (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress)
        ? Office.context.mailbox.userProfile.emailAddress
        : "unknown@unknown.com";
      const url = `https://corporate-crm-data-enrichment-azbqwqgd6a-uc.a.run.app/DataEnrichment/SaveOutlookPluginLogsToDatabase?Action=${encodeURIComponent(action)}&UserEmail=${encodeURIComponent(userEmail)}`;
      console.log(`[PluginLog] Sending usage log`, { action, userEmail, url });
      // Use no-cors to avoid blocking on CORS; this is fire-and-forget
      fetch(url, { method: 'GET', mode: 'no-cors' })
        .then(() => {
          console.log(`[PluginLog] Usage log sent (opaque/no-cors)`, { action, userEmail });
        })
        .catch((err) => {
          console.warn(`[PluginLog] Usage log failed`, { action, userEmail, error: String(err) });
        });
    } catch (err) {
      console.warn(`[PluginLog] Usage log threw before fetch`, { action, error: String(err) });
    }
  };
  
  // getEmailBody= () => "I am a sales manager and I am sending this email to you";
  // Helper to call Gemini API with a custom prompt
  const callGemini = async (promptTemplate) => {
    // Usage tracking for Suggest Reply action
    logPluginUsage('Suggest Reply');
    setLoading(true);
    setGeneratedContent("Generating...");
    try {
      const emailBody = await getEmailBody();
      const conversationThread = await getConversationThread();
      
      // Include conversation thread context in the prompt
      const contextWithThread = `Conversation Context:\n${conversationThread}\n\nCurrent Email:\n${emailBody}`;
      const prompt = promptTemplate
        .replace("{emailBody}", contextWithThread)
        .replace("{tone}", emailForm.tone)
        .replace("{pointOfView}", emailForm.pointOfView);
      
      console.log("prompt", prompt);
      
      // Monitor for retry attempts (silent)
      let retryCount = 0;
      const originalLog = console.log;
      console.log = (...args) => {
        if (args[0] && args[0].includes && args[0].includes('BotAtWork API attempt')) {
          retryCount++;
          // Keep the loading message clean - no retry status shown to user
        }
        originalLog.apply(console, args);
      };
      
      // Use emailResponse for reply suggestions with proper structured parameters
      const apiParams = {
        chooseATask: "emailResponse",
        emailContent: contextWithThread, // Pass the actual email content, not the formatted prompt
        tone: emailForm.tone.toLowerCase() || "formal",
        pointOfView: emailForm.pointOfView === "Individual perspective" ? "individualPerspective" : "organizationPerspective",
        additionalInstructions: ""
      };
      
      const reply = await getSuggestedReply(contextWithThread, 3, apiParams);
      
      // Restore original console.log
      console.log = originalLog;
      
      setGeneratedContent(reply);
    } catch (e) {
      // Show API error messages as-is, they're already properly formatted
      const errorMessage = e.toString();
      setGeneratedContent(errorMessage.startsWith('API error:') || errorMessage.startsWith('Error calling BotAtWork API:') ? errorMessage : `Error: ${errorMessage}`);
    }
    setLoading(false);
  };

  // Email validation function that can be called from taskpane
  const validateCurrentEmail = async () => {
    setLoading(true);
    setGeneratedContent("Validating email content...");
    
    try {
      const emailBody = await getEmailBody();
      const emailSubject = await getEmailSubject();
      const missing = await checkMissingKeywords(emailBody, emailSubject);
      
      if (missing.length > 0) {
        // Show validation results in the taskpane
        setGeneratedContent(`
          <div style="color: #d13438; font-weight: bold; margin-bottom: 16px;">
            ⚠️ Missing Important Elements Detected
          </div>
          <div style="background: #fff4ce; border: 1px solid #ffb900; padding: 12px; border-radius: 4px; margin-bottom: 16px;">
            <div style="margin-bottom: 8px;"><strong>Email Subject:</strong> ${emailSubject}</div>
            <div style="margin-bottom: 12px;"><strong>Missing Elements:</strong></div>
            <ul style="margin: 8px 0; padding-left: 20px;">
              ${missing.map(item => `<li><strong>${item.category}</strong>: Consider adding ${item.suggestions.join(', ')}</li>`).join('')}
            </ul>
          </div>
          <div style="color: #605e5c; font-size: 14px;">
            <strong>Recommendations:</strong><br/>
            • Click "Generate Enhanced Response" below to automatically add missing elements<br/>
            • Or manually add these elements to your email before sending
          </div>
        `);
        
        // Show enhance button
        setShowEnhanceButton(true);
        setMissingKeywords(missing);
      } else {
        setGeneratedContent(`
          <div style="color: #107c10; font-weight: bold; margin-bottom: 16px;">
            ✅ Email Validation Passed!
          </div>
          <div style="background: #f3f9f1; border: 1px solid #107c10; padding: 12px; border-radius: 4px;">
            <div style="margin-bottom: 8px;"><strong>Email Subject:</strong> ${emailSubject}</div>
            <div>Your email contains all important elements and is ready to send.</div>
          </div>
        `);
        setShowEnhanceButton(false);
      }
    } catch (error) {
      setGeneratedContent(`Error validating email: ${error.message}`);
    }
    
    setLoading(false);
  };

  // Check for missing keywords in both subject and body
  const checkMissingKeywords = async (emailBody, emailSubject = "") => {
    const REQUIRED_KEYWORDS = {
      urgency: {
        keywords: ['urgent', 'asap', 'immediate', 'critical', 'priority', 'important'],
        category: 'Urgency Indicators',
        suggestions: ['urgent', 'important', 'priority', 'asap'],
        checkSubject: true
      },
      action: {
        keywords: ['action item', 'action required', 'next steps', 'follow up', 'todo', 'please', 'request'],
        category: 'Action Items',
        suggestions: ['action items', 'next steps', 'follow-up required', 'please review'],
        checkSubject: true
      },
      timeline: {
        keywords: ['deadline', 'due date', 'timeline', 'schedule', 'by when', 'when', 'date', 'time'],
        category: 'Timeline',
        suggestions: ['deadline', 'timeline', 'completion date', 'by [date]'],
        checkSubject: true
      },
      participants: {
        keywords: ['attendees', 'participants', 'who should attend', 'invitees', 'team', 'members'],
        category: 'Participants',
        suggestions: ['attendees', 'participants', 'invitees', 'team members'],
        checkSubject: false
      },
      greeting: {
        keywords: ['dear', 'hello', 'hi', 'good morning', 'good afternoon', 'greetings'],
        category: 'Greeting',
        suggestions: ['Dear', 'Hello', 'Hi', 'Good morning'],
        checkSubject: false
      },
      closing: {
        keywords: ['regards', 'best regards', 'sincerely', 'thank you', 'thanks', 'appreciate'],
        category: 'Closing',
        suggestions: ['Best regards', 'Thank you', 'Sincerely', 'Appreciate your time'],
        checkSubject: false
      },
      context: {
        keywords: ['regarding', 'about', 'concerning', 're:', 'subject', 'topic'],
        category: 'Context Clarity',
        suggestions: ['regarding', 'about', 'concerning'],
        checkSubject: true
      },
      meeting: {
        keywords: ['meeting', 'agenda', 'discussion', 'call', 'conference', 'presentation'],
        category: 'Meeting Elements',
        suggestions: ['meeting agenda', 'discussion points', 'call details'],
        checkSubject: true
      }
    };

    const bodyText = emailBody.toLowerCase();
    const subjectText = String(emailSubject || ""); // Ensure subject is a string
    const missing = [];
    
    Object.keys(REQUIRED_KEYWORDS).forEach(key => {
      const keywordData = REQUIRED_KEYWORDS[key];
      let hasKeyword = false;
      
      // Check body text
      hasKeyword = keywordData.keywords.some(keyword => 
        bodyText.includes(keyword.toLowerCase())
      );
      
      // If not found in body and should check subject, check subject
      if (!hasKeyword && keywordData.checkSubject) {
        hasKeyword = keywordData.keywords.some(keyword => 
          subjectText.toLowerCase().includes(keyword.toLowerCase())
        );
      }
      
      if (!hasKeyword) {
        missing.push({
          category: keywordData.category,
          suggestions: keywordData.suggestions,
          key: key
        });
      }
    });
    
    return missing;
  };

  // Generate enhanced email with missing keywords
  const generateEnhancedEmail = async () => {
    if (!missingKeywords || missingKeywords.length === 0) return;
    
    setLoading(true);
    setGeneratedContent("Enhancing email with missing keywords...");
    
    try {
      const currentBody = await getEmailBody();
      const emailSubject = await getEmailSubject();
      
      // Get all missing keywords and their categories
      const missingElements = missingKeywords.map(item => ({
        category: item.category,
        keywords: item.suggestions.slice(0, 2) // Take first 2 suggestions from each category
      }));
      
      const missingElementsText = missingElements.map(item => 
        `${item.category}: ${item.keywords.join(', ')}`
      ).join('\n');
      
      const enhancementPrompt = `Please enhance this email by naturally incorporating these missing elements:

Email Subject: ${emailSubject}

Missing Elements to Add:
${missingElementsText}

Original Email Content:
${currentBody}

Instructions:
- Keep the original tone and intent
- Naturally integrate the missing elements throughout the email
- Make it professional and coherent
- Don't change the core message, just enhance it with the missing elements
- Ensure the enhanced email flows naturally and maintains readability
- Use proper paragraph breaks and formatting
- Include appropriate spacing between sections
- Format the email with clear structure (greeting, body, closing)

Enhanced email:`;

      // Use emailWrite for generating a new enhanced email
      const apiParams = {
        chooseATask: "emailWrite",
        description: enhancementPrompt,
        tone: emailForm.tone.toLowerCase() || "formal",
        pointOfView: emailForm.pointOfView === "Individual perspective" ? "individualPerspective" : "organizationPerspective",
        additionalInstructions: ""
      };
      
      const enhancedContent = await getSuggestedReply(enhancementPrompt, 3, apiParams);
      
      // Update the email body in Outlook with proper HTML formatting
      if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        // Convert plain text to HTML with proper formatting
        const formattedContent = enhancedContent
          .trim() // Remove leading/trailing whitespace
          .replace(/\n\n+/g, '</p><p>') // Multiple line breaks become paragraph breaks
          .replace(/\n/g, '<br/>') // Single line breaks become <br/>
          .replace(/^(.+)$/s, '<p>$1</p>') // Wrap entire content in paragraph tags
          .replace(/<p><\/p>/g, '') // Remove empty paragraphs
          .replace(/<p><p>/g, '<p>') // Fix double paragraph tags
          .replace(/<\/p><\/p>/g, '</p>'); // Fix double closing paragraph tags
        
        Office.context.mailbox.item.body.setAsync(
          formattedContent,
          { coercionType: Office.CoercionType.Html },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              setGeneratedContent(`
                <div style="color: #107c10; font-weight: bold; margin-bottom: 16px;">
                  ✅ Email Enhanced Successfully!
                </div>
                <div style="background: #f3f9f1; border: 1px solid #107c10; padding: 12px; border-radius: 4px; margin-bottom: 16px;">
                  <div style="margin-bottom: 8px;"><strong>Email Subject:</strong> ${emailSubject}</div>
                  <div>Your email has been updated with the missing keywords. Please review the changes in your email composer and send when ready.</div>
                </div>
                <div style="background: #f8f9fa; border: 1px solid #d1d1d1; padding: 12px; border-radius: 4px;">
                  <strong>Enhanced Content Preview:</strong><br/>
                  ${enhancedContent.replace(/\n/g, '<br/>')}
                </div>
              `);
              setShowEnhanceButton(false);
            } else {
              setGeneratedContent("Failed to update email. Please copy the enhanced content manually.");
            }
          }
        );
      } else {
        setGeneratedContent(`
          <div style="color: #107c10; font-weight: bold; margin-bottom: 16px;">
            ✅ Enhanced Email Generated!
          </div>
          <div style="background: #f8f9fa; border: 1px solid #d1d1d1; padding: 12px; border-radius: 4px; margin-bottom: 16px;">
            <div style="margin-bottom: 8px;"><strong>Email Subject:</strong> ${emailSubject}</div>
            <strong>Enhanced Content:</strong><br/>
            ${enhancedContent.replace(/\n/g, '<br/>')}
          </div>
          <div style="margin-top: 12px; color: #605e5c; font-size: 14px;">
            Please copy this enhanced content to your email.
          </div>
        `);
      }
    } catch (error) {
      setGeneratedContent(`Error enhancing email: ${error.message}`);
    }
    
    setLoading(false);
  };



  // Chat input send handler: send direct prompt to LLM
  // Chat input send handler: send direct prompt to LLM
  const handleChatSend = async () => {
    if (!chatInput.trim()) return;
    
    const userMessage = chatInput.trim();
    setChatInput("");
    
    // Add user message to chat history
    const newUserMessage = { type: 'user', content: userMessage };
    setChatHistory(prev => [...prev, newUserMessage]);
    
    setLoading(true);
    
    try {
      let prompt;
      if (isFirstResponse && generatedContent && generatedContent !== "Generating..." && !generatedContent.includes("Validating") && !generatedContent.includes("✨ Your generated content will appear here")) {
        // If there's existing content, include it in context
        prompt = `Context: I previously provided this content/response: "${generatedContent.replace(/<[^>]*>/g, '').substring(0, 500)}..."\n\nNow the user is asking: ${userMessage}\n\nPlease respond helpfully to their follow-up question.`;
        setIsFirstResponse(false);
      } else {
        // Build context from chat history
        const context = chatHistory.filter(msg => msg.content && msg.content.trim()).map(msg => 
          msg.type === 'user' ? `User: ${msg.content}` : `Assistant: ${msg.content.replace(/<[^>]*>/g, '').substring(0, 300)}`
        ).join('\n');
        prompt = context ? `Previous conversation:\n${context}\n\nUser's new question: ${userMessage}\n\nPlease provide a helpful response.` : `User question: ${userMessage}\n\nPlease provide a helpful response.`;
      }
      
      // Monitor for retry attempts (silent)
      let retryCount = 0;
      const originalLog = console.log;
      console.log = (...args) => {
        if (args[0] && args[0].includes && args[0].includes('BotAtWork API attempt')) {
          retryCount++;
          // Keep the loading message clean - no retry status shown to user
        }
        originalLog.apply(console, args);
      };
      
      // Use chat task type for conversational interactions
      const contextualPrompt = `${prompt}`;
      
      const apiParams = {
        chooseATask: "chat",
        description: contextualPrompt,
        tone: emailForm.tone.toLowerCase() || "formal",
        pointOfView: emailForm.pointOfView === "Individual perspective" ? "individualPerspective" : "organizationPerspective",
        additionalInstructions: "Provide a helpful and conversational response to the user's question or request."
      };
      
      const reply = await getSuggestedReply(contextualPrompt, 3, apiParams);
      
      // Restore original console.log
      console.log = originalLog;
      
      // Add AI response to chat history
      const newAIMessage = { type: 'ai', content: reply };
      setChatHistory(prev => [...prev, newAIMessage]);
      
      // Update the main content area with latest response
      setGeneratedContent(reply);
    } catch (e) {
      // Show API error messages as-is, they're already properly formatted
      const errorString = e.toString();
      const errorMessage = errorString.startsWith('API error:') || errorString.startsWith('Error calling BotAtWork API:') ? errorString : `Error: ${errorString}`;
      setChatHistory(prev => [...prev, { type: 'ai', content: errorMessage }]);
      setGeneratedContent(errorMessage);
    }
    setLoading(false);
  };

  // Tab handler
  const handleTabSelect = (event, data) => {
    setActiveTab(data.value);
    if (data.value === 'writeEmail') {
      setShowWriteEmailForm(true);
      setChatHistory([]);
      setIsFirstResponse(true);
      setGeneratedContent("");
      setShowEnhanceButton(false);
    } else if (data.value === 'suggestReply') {
      setShowWriteEmailForm(false);
      setChatHistory([]);
      setIsFirstResponse(true);
      setShowEnhanceButton(false);
      // Removed automatic API call - let users manually trigger when ready
      setGeneratedContent("");
    } 
    /* else if (data.value === 'validate') {
      setShowWriteEmailForm(false);
      setChatHistory([]);
      setIsFirstResponse(true);
      validateCurrentEmail();
    } */
  };
  const handleGenerateEmail = async () => {
    if (!emailForm.description.trim()) {
      setGeneratedContent("Please enter a description for the email.");
      return;
    }
    // Usage tracking for Write Email action
    logPluginUsage('Write Email');
    
    const prompt = customPrompts.writeEmail
      .replace("{description}", emailForm.description)
      // .replace("{additionalInstructions}", emailForm.additionalInstructions || "None") // Commented out as per user request
      .replace("{tone}", emailForm.tone)
      .replace("{pointOfView}", emailForm.pointOfView);
    
    setLoading(true);
    setGeneratedContent("Generating email...");
    
    try {
      // Monitor for retry attempts (silent)
      let retryCount = 0;
      const originalLog = console.log;
      console.log = (...args) => {
        if (args[0] && args[0].includes && args[0].includes('BotAtWork API attempt')) {
          retryCount++;
          // Keep the loading message clean - no retry status shown to user
        }
        originalLog.apply(console, args);
      };
      
      // Use emailWrite for creating new emails
      const apiParams = {
        chooseATask: "emailWrite",
        description: prompt,
        tone: emailForm.tone.toLowerCase() || "formal",
        pointOfView: emailForm.pointOfView === "Individual perspective" ? "individualPerspective" : "organizationPerspective",
        additionalInstructions: ""
      };
      
      const reply = await getSuggestedReply(prompt, 3, apiParams);
      
      // Restore original console.log
      console.log = originalLog;
      
      setGeneratedContent(reply);
      setLoading(false);
    } catch (e) {
      // Show API error messages as-is, they're already properly formatted
      const errorMessage = e.toString();
      setGeneratedContent(errorMessage.startsWith('API error:') || errorMessage.startsWith('Error calling BotAtWork API:') ? errorMessage : `Error: ${errorMessage}`);
      setLoading(false);
    }
  };
  const handleSaveTemplate = () => {
    setTemplates((prev) => [...prev, generatedContent]);
    setGeneratedContent("Template saved.");
  };
  const handleViewTemplates = () => {
    if (templates.length === 0) {
      setGeneratedContent("No templates saved.");
    } else {
      setGeneratedContent(templates.map((t, i) => `Template ${i + 1}:\n${t}`).join("\n\n---\n\n"));
    }
  };
  const handleClear = () => {
    setGeneratedContent("");
  };

  // const handleSavePrompts = (newPrompts) => {
  //   setCustomPrompts(newPrompts);
  // };

  // const toggleDarkMode = () => {
  //   setIsDarkMode(!isDarkMode);
  // };

  return (
    <FluentProvider theme={teamsLightTheme}> {/* dark mode disabled for now */}
      <div className={styles.root}>
        <div className={styles.tabContainer}>
          <TabList selectedValue={activeTab} onTabSelect={handleTabSelect} style={{ width: '100%', display: 'flex' }}>
                          <Tab value="writeEmail" style={{ flex: '1 1 0', fontSize: '12px', padding: '8px 4px', textAlign: 'center', whiteSpace: 'nowrap', overflow: 'hidden' }}>Write Email</Tab>
              <Tab value="suggestReply" style={{ flex: '1 1 0', fontSize: '12px', padding: '8px 4px', textAlign: 'center', whiteSpace: 'nowrap', overflow: 'hidden' }}>Suggest Reply</Tab>
              {/* <Tab value="validate" style={{ flex: '1 1 0', fontSize: '12px', padding: '8px 4px', textAlign: 'center', whiteSpace: 'nowrap', overflow: 'hidden' }}>Validate</Tab> */}
          </TabList>
        </div>
        
        {showWriteEmailForm && (
          <div style={{ 
            padding: '8px', 
            width: '100%',
            boxSizing: 'border-box',
            background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
            borderRadius: '8px',
            marginBottom: '8px'
          }}>
            <div style={{ marginBottom: '8px' }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '3px', 
                fontWeight: '600',
                fontSize: '14px',
                color: '#323130'
              }}>Description *</label>
              <textarea
                placeholder="Describe what you want to write about and who the email is for."
                value={emailForm.description}
                onChange={(e) => setEmailForm({...emailForm, description: e.target.value})}
                style={{
                  width: '100%',
                  minHeight: '60px',
                  padding: '6px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
                  fontSize: '14px',
                  fontFamily: 'inherit',
                  resize: 'vertical',
                  boxSizing: 'border-box',
                  backgroundColor: '#ffffff',
                  transition: 'border-color 0.2s ease',
                  outline: 'none'
                }}
                onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
              />
            </div>
            
            {/* Additional Instructions section - Hidden as per user request */}
            {/* <div style={{ marginBottom: '16px' }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '6px', 
                fontWeight: '600',
                fontSize: '14px',
                color: '#323130'
              }}>Additional Instructions</label>
              <textarea
                placeholder="Any additional instructions or specific requirements..."
                value={emailForm.additionalInstructions}
                onChange={(e) => setEmailForm({...emailForm, additionalInstructions: e.target.value})}
                style={{
                  width: '100%',
                  minHeight: '60px',
                  padding: '10px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
                  fontSize: '14px',
                  fontFamily: 'inherit',
                  resize: 'vertical',
                  boxSizing: 'border-box',
                  backgroundColor: '#ffffff',
                  transition: 'border-color 0.2s ease',
                  outline: 'none'
                }}
                onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
              />
            </div> */}
            
            <div style={{ 
              display: 'flex', 
              gap: '8px', 
              marginBottom: '10px',
              alignItems: 'flex-end'
            }}>
              <div style={{ flex: 1 }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '3px', 
                fontWeight: '600',
                fontSize: '14px',
                color: '#323130'
              }}>Tone *</label>
              <select
                value={emailForm.tone}
                onChange={(e) => setEmailForm({...emailForm, tone: e.target.value})}
                style={{
                  width: '100%',
                  padding: '6px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
                  fontSize: '14px',
                  fontFamily: 'inherit',
                  backgroundColor: '#ffffff',
                  boxSizing: 'border-box',
                  outline: 'none',
                  cursor: 'pointer'
                }}
                onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
              >
                <option value="formal">Formal</option>
                <option value="casual">Casual</option>
                <option value="professional">Professional</option>
                <option value="empathetic">Empathetic</option>
              </select>
            </div>
            
              <div style={{ flex: 1 }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '3px', 
                fontWeight: '600',
                fontSize: '14px',
                color: '#323130'
              }}>Point of View *</label>
              <select
                value={emailForm.pointOfView}
                onChange={(e) => setEmailForm({...emailForm, pointOfView: e.target.value})}
                style={{
                  width: '100%',
                  padding: '6px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
                  fontSize: '14px',
                  fontFamily: 'inherit',
                  backgroundColor: '#ffffff',
                  boxSizing: 'border-box',
                  outline: 'none',
                  cursor: 'pointer'
                }}
                onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
              >
                <option value="Organization perspective">Organization perspective</option>
                <option value="Individual perspective">Individual perspective</option>
              </select>
              </div>
            </div>
            
            <Button
              appearance={emailForm.description.trim() ? "primary" : "secondary"}
              onClick={handleGenerateEmail}
              disabled={loading || !emailForm.description.trim()}
              style={{ 
                width: '100%',
                padding: '8px 16px',
                fontSize: '15px',
                fontWeight: '600',
                borderRadius: '4px',
                minHeight: '36px',
                backgroundColor: emailForm.description.trim() ? '#0078d4' : '#f3f2f1',
                color: emailForm.description.trim() ? '#ffffff' : '#323130',
                border: emailForm.description.trim() ? 'none' : '1px solid #d1d1d1'
              }}
            >
              {loading ? (
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}>
                  <div style={{
                    width: '16px',
                    height: '16px',
                    border: '2px solid rgba(255,255,255,0.3)',
                    borderRadius: '50%',
                    borderTop: '2px solid #ffffff',
                    animation: 'spin 1s linear infinite',
                    transformOrigin: 'center',
                    willChange: 'transform'
                  }}></div>
                  <span>Generating...</span>
                </div>
              ) : 'Generate Email'}
            </Button>
          </div>
        )}
        {activeTab === 'suggestReply' && (
          <div style={{
            padding: '8px',
            width: '100%',
            boxSizing: 'border-box',
            background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
            borderRadius: '8px',
            marginBottom: '8px'
          }}>
            <div style={{ 
              display: 'flex', 
              gap: '8px', 
              marginBottom: '10px',
              alignItems: 'flex-end'
            }}>
              <div style={{ flex: 1 }}>
                <label style={{ 
                  display: 'block', 
                  marginBottom: '3px', 
                  fontWeight: '600', 
                  fontSize: '14px',
                  color: '#323130' 
                }}>
                Tone
              </label>
              <select
                value={emailForm.tone}
                onChange={(e) => setEmailForm({ ...emailForm, tone: e.target.value })}
                style={{
                  width: '100%',
                    padding: '6px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
                  fontSize: '14px',
                  backgroundColor: '#ffffff',
                  boxSizing: 'border-box',
                    cursor: 'pointer',
                    outline: 'none'
                }}
                  onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                  onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
              >
                <option value="formal">Formal</option>
                <option value="casual">Casual</option>
                <option value="professional">Professional</option>
                <option value="empathetic">Empathetic</option>
              </select>
            </div>
              
              <div style={{ flex: 1 }}>
                <label style={{ 
                  display: 'block', 
                  marginBottom: '3px', 
                  fontWeight: '600', 
                  fontSize: '14px', 
                  color: '#323130' 
                }}>
                Point of View
              </label>
              <select
                value={emailForm.pointOfView}
                onChange={(e) => setEmailForm({ ...emailForm, pointOfView: e.target.value })}
                style={{
                  width: '100%',
                    padding: '6px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
                  fontSize: '14px',
                  backgroundColor: '#ffffff',
                  boxSizing: 'border-box',
                    cursor: 'pointer',
                    outline: 'none'
                }}
                  onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                  onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
              >
                <option value="Organization perspective">Organization perspective</option>
                <option value="Individual perspective">Individual perspective</option>
              </select>
              </div>
            </div>
            <Button
              appearance="primary"
              onClick={() => callGemini(customPrompts.suggestReply)}
              disabled={loading}
              style={{ 
                width: '100%',
                padding: '8px 16px',
                fontSize: '15px',
                fontWeight: '600',
                borderRadius: '4px',
                minHeight: '36px',
                backgroundColor: '#0078d4',
                color: '#ffffff',
                border: 'none',
                opacity: loading ? 0.5 : 1
              }}
            >
              {loading ? (
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}>
                  <div style={{
                    width: '16px',
                    height: '16px',
                    border: '2px solid rgba(255,255,255,0.3)',
                    borderRadius: '50%',
                    borderTop: '2px solid #ffffff',
                    animation: 'spin 1s linear infinite'
                  }}></div>
                  <span>Generating...</span>
                </div>
              ) : 'Generate New Reply'}
            </Button>
          </div>
        )}

        {/* Validate tab content - temporarily hidden
        {activeTab === 'validate' && (
          <div style={{
            padding: '8px',
            width: '100%',
            boxSizing: 'border-box',
            background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
            borderRadius: '8px',
            marginBottom: '8px'
          }}>
            <div style={{
              textAlign: 'center',
              marginBottom: '12px'
            }}>
              <h3 style={{
                margin: '0 0 8px 0',
                fontSize: '16px',
                fontWeight: '600',
                color: '#323130'
              }}>Email Validation</h3>
              <p style={{
                margin: '0',
                fontSize: '14px',
                color: '#605e5c',
                lineHeight: '1.4'
              }}>Check if your email contains all important elements</p>
            </div>
            <Button
              appearance="primary"
              onClick={validateCurrentEmail}
              disabled={loading}
              style={{ 
                width: '100%',
                padding: '8px 16px',
                fontSize: '15px',
                fontWeight: '600',
                borderRadius: '4px',
                minHeight: '36px',
                backgroundColor: '#0078d4',
                color: '#ffffff',
                border: 'none',
                opacity: loading ? 0.5 : 1
              }}
            >
              {loading ? 'Validating...' : '🔍 Validate Current Email'}
            </Button>
          </div>
        )}
        */}

        <div className={styles.contentArea}>
          {loading ? (
            <div className={styles.loadingContainer}>
              <div style={{
                position: 'relative',
                width: '60px',
                height: '60px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>
                <div style={{
                  position: 'absolute',
                  width: '100%',
                  height: '100%',
                  borderRadius: '50%',
                  border: '4px solid rgba(0, 120, 212, 0.2)',
                  borderTop: '4px solid #0078d4',
                  animation: 'spin 1s linear infinite',
                  transformOrigin: 'center',
                  willChange: 'transform'
                }}></div>
                <div style={{
                  position: 'absolute',
                  width: '70%',
                  height: '70%',
                  borderRadius: '50%',
                  border: '3px solid rgba(0, 120, 212, 0.15)',
                  borderRight: '3px solid #106ebe',
                  animation: 'spin 0.8s linear infinite reverse',
                  transformOrigin: 'center',
                  willChange: 'transform'
                }}></div>
                <div style={{
                  position: 'absolute',
                  width: '40%',
                  height: '40%',
                  borderRadius: '50%',
                  border: '2px solid rgba(0, 120, 212, 0.1)',
                  borderBottom: '2px solid #005a9e',
                  animation: 'spin 1.2s linear infinite',
                  transformOrigin: 'center',
                  willChange: 'transform'
                }}></div>
              </div>
              <div className={styles.loadingText}>
                {generatedContent === "Generating..." ? "Generating your content..." :
                 generatedContent === "Generating email..." ? "Creating your email..." :
                 generatedContent === "Enhancing email with missing keywords..." ? "Enhancing your email..." :
                 generatedContent === "Validating email content..." ? "Validating your email..." :
                 "Processing your request..."}
              </div>
            </div>
          ) : (
            <div
              dangerouslySetInnerHTML={{
                __html: generatedContent
                  ? generatedContent
                      .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noopener noreferrer">$1</a>') // links clickable
                      .replace(/\n/g, '<br>') // preserve line breaks
                      .replace(/\*\*(.*?)\*\*/g, '<b>$1</b>') // bold for **text**
                      .replace(/\*(.*?)\*/g, '<i>$1</i>') // italics for *text*
                  : '<div style="display: flex; align-items: center; justify-content: center; height: 100%; color: #605e5c; font-style: italic; text-align: center; padding: 30px;"><div><div style="font-size: 16px; margin-bottom: 6px;">✨ Your generated content will appear here</div><div style="font-size: 12px; opacity: 0.8;">Click a button above to get started</div></div></div>'
              }}
            />
          )}
        </div>
        
        {/* Enhanced Response Button - Only show when there's generated content and enhance button is enabled */}
        {showEnhanceButton && (
          <div style={{
            display: 'flex',
            gap: '8px',
            margin: '8px',
            flexShrink: 0
          }}>
            <Button
              appearance="primary"
              onClick={generateEnhancedEmail}
              disabled={loading}
              style={{ 
                flex: 1,
                padding: '8px 16px',
                fontSize: '14px',
                fontWeight: '600',
                borderRadius: '4px',
                minHeight: '36px',
                backgroundColor: '#0078d4',
                color: '#ffffff',
                border: 'none',
                opacity: loading ? 0.5 : 1
              }}
            >
              {loading ? (
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}>
                  <div style={{
                    width: '16px',
                    height: '16px',
                    border: '2px solid rgba(255,255,255,0.3)',
                    borderRadius: '50%',
                    borderTop: '2px solid #ffffff',
                    animation: 'spin 1s linear infinite'
                  }}></div>
                  <span>Enhancing...</span>
                </div>
              ) : '✨ Generate Enhanced Response'}
            </Button>
          </div>
        )}
        
        {(() => {
          const shouldShow = (activeTab === 'suggestReply' || (activeTab === 'writeEmail' && generatedContent && !generatedContent.includes("Generating")) /* || (activeTab === 'validate' && generatedContent && !generatedContent.includes("Validating")) */);
          console.log('Chat input should show:', shouldShow, 'activeTab:', activeTab, 'generatedContent:', generatedContent);
          return shouldShow;
        })() && (
          <div style={{
            position: 'relative',
            display: 'flex',
            alignItems: 'center',
            margin: '2px 0px',
            flexShrink: 0,
            width: '100%',
            paddingLeft: '4px',
            paddingRight: '4px',
            boxSizing: 'border-box'
          }}>
            <input
              type="text"
              placeholder="Type follow-up question here"
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              onKeyPress={(e) => e.key === 'Enter' && handleChatSend()}
              style={{
                width: '100%',
                padding: '8px 35px 8px 12px',
                fontSize: '14px',
                borderRadius: '20px',
                border: '1px solid #d1d1d1',
                outline: 'none',
                backgroundColor: '#ffffff',
                transition: 'all 0.3s ease',
                color: '#323130',
                boxSizing: 'border-box',
                height: '36px'
              }}
              onFocus={(e) => {
                e.target.style.borderColor = '#0078d4';
                e.target.style.boxShadow = '0 0 0 2px rgba(0,120,212,0.2)';
              }}
              onBlur={(e) => {
                e.target.style.borderColor = '#d1d1d1';
                e.target.style.boxShadow = 'none';
              }}
            />
            <button
              disabled={loading || !chatInput.trim()}
              onClick={handleChatSend}
              style={{
                position: 'absolute',
                right: '6px',
                top: '50%',
                transform: 'translateY(-50%)',
                borderRadius: '50%',
                width: '28px',
                height: '28px',
                backgroundColor: '#0078d4',
                color: '#ffffff',
                border: 'none',
                fontSize: '12px',
                fontWeight: '600',
                cursor: 'pointer',
                opacity: (loading || !chatInput.trim()) ? 0.5 : 1,
                transition: 'all 0.3s ease',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                boxShadow: '0 1px 3px rgba(0,120,212,0.3)'
              }}
              onMouseEnter={(e) => {
                if (!loading && chatInput.trim()) {
                  e.target.style.backgroundColor = '#106ebe';
                  e.target.style.transform = 'translateY(-50%) scale(1.05)';
                }
              }}
              onMouseLeave={(e) => {
                if (!loading && chatInput.trim()) {
                  e.target.style.backgroundColor = '#0078d4';
                  e.target.style.transform = 'translateY(-50%) scale(1)';
                }
              }}
            >
              {loading ? (
                <div style={{
                  width: '12px',
                  height: '12px',
                  border: '2px solid rgba(255,255,255,0.3)',
                  borderRadius: '50%',
                  borderTopColor: '#ffffff',
                  animation: 'spin 1s ease-in-out infinite'
                }}></div>
              ) : '↑'}
            </button>
          </div>
        )}
      </div>
    </FluentProvider>
  );
}

App.propTypes = {
  title: PropTypes.string,
};

export default App;

console.log("Arnav");
