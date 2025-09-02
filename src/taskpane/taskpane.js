/* global Office console */

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

export async function insertText(text) {
  // Write text to the cursor point in the compose surface with proper HTML formatting
  try {
    // Convert markdown to HTML for proper formatting in Outlook
    const formattedHtml = convertMarkdownToOutlookHtml(text);
    console.log('🔧 Converting markdown to HTML for text insertion...');
    
    Office.context.mailbox.item?.body.setSelectedDataAsync(
      formattedHtml,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          throw asyncResult.error.message;
        }
      }
    );
  } catch (error) {
    console.log("Error: " + error);
  }
}
