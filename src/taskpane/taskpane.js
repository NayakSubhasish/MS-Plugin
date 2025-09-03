/* global Office console */

// Helper function to convert markdown to proper HTML for Outlook
function convertMarkdownToOutlookHtml(markdownText) {
  if (!markdownText || typeof markdownText !== 'string') {
    return markdownText || '';
  }

  // First, handle tables before other conversions
  let html = markdownText;
  
  // Convert markdown tables to HTML tables
  html = html.replace(/^([^\n]*\|[^\n]*\n)([^\n]*\|[^\n]*\n)([^\n]*\|[^\n]*\n)+/gm, (match) => {
    const lines = match.trim().split('\n');
    if (lines.length < 3) return match; // Need at least header, separator, and one data row
    
    const header = lines[0];
    const separator = lines[1];
    const dataRows = lines.slice(2);
    
    // Check if separator line contains dashes and pipes (table format)
    if (!/^\s*\|?\s*:?[-|:\s]+\s*\|?\s*$/.test(separator)) {
      return match; // Not a valid table separator
    }
    
    // Parse header
    const headerCells = header.split('|').map(cell => cell.trim()).filter(Boolean);
    
    // Parse data rows
    const tableRows = dataRows.map(row => {
      const cells = row.split('|').map(cell => cell.trim()).filter(Boolean);
      return cells;
    });
    
    // Build HTML table
    let tableHtml = '<table style="border-collapse: collapse; width: 100%; margin: 16px 0; border: 1px solid #d1d1d1;">';
    
    // Add header
    if (headerCells.length > 0) {
      tableHtml += '<thead><tr>';
      headerCells.forEach(cell => {
        tableHtml += `<th style="border: 1px solid #d1d1d1; padding: 8px 12px; background-color: #f8f9fa; font-weight: 600; text-align: left;">${escapeHtml(cell)}</th>`;
      });
      tableHtml += '</tr></thead>';
    }
    
    // Add data rows
    if (tableRows.length > 0) {
      tableHtml += '<tbody>';
      tableRows.forEach(row => {
        tableHtml += '<tr>';
        row.forEach(cell => {
          tableHtml += `<td style="border: 1px solid #d1d1d1; padding: 8px 12px; vertical-align: top;">${escapeHtml(cell)}</td>`;
        });
        // Fill any missing cells if row is shorter than header
        for (let i = row.length; i < headerCells.length; i++) {
          tableHtml += '<td style="border: 1px solid #d1d1d1; padding: 8px 12px; vertical-align: top;"></td>';
        }
        tableHtml += '</tr>';
      });
      tableHtml += '</tbody>';
    }
    
    tableHtml += '</table>';
    return tableHtml;
  });

  // Now handle other markdown elements
  html = html
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
    .replace(/_(.*?)__/g, '<em>$1</em>')
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

// Helper function to escape HTML special characters
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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
