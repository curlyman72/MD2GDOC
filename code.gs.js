function onOpen() {
  DocumentApp.getUi()
    .createMenu('GenAI Markdown')
    .addItem('Open Converter', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Markdown Converter')
    .setWidth(380);
  DocumentApp.getUi().showSidebar(html);
}

function getSettings() {
  const props = PropertiesService.getUserProperties();
  return {
    mermaidProvider: props.getProperty('MERMAID_PROVIDER') || 'mermaid.ink',
    apiKey: props.getProperty('MERMAID_API_KEY') || '',
    codeRetention: props.getProperty('CODE_RETENTION') || 'alt',
    theme: props.getProperty('MERMAID_THEME') || 'neutral'
  };
}

function saveSettings(settings) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('MERMAID_PROVIDER', settings.provider);
  props.setProperty('MERMAID_API_KEY', settings.apiKey);
  props.setProperty('CODE_RETENTION', settings.retention);
  props.setProperty('MERMAID_THEME', settings.theme);
  return { success: true };
}

function getSelectedImageSource() {
  try {
    const selection = DocumentApp.getActiveDocument().getSelection();
    if (!selection) return { error: 'No image selected. Click on a diagram first.' };
    
    const elements = selection.getRangeElements();
    if (elements.length === 0) return { error: 'No element selected' };
    
    const element = elements[0].getElement();
    
    if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      const image = element.asInlineImage();
      return {
        type: 'image',
        altText: image.getAltDescription() || 'No source code stored',
        title: image.getAltTitle() || 'Mermaid Diagram',
        hasSource: !!image.getAltDescription()
      };
    }
    
    if (element.getType() === DocumentApp.ElementType.TABLE) {
      const table = element.asTable();
      if (table.getNumRows() > 0) {
        const cell = table.getCell(0, 0);
        const text = cell.getText();
        return {
          type: 'code',
          content: text,
          language: 'detected'
        };
      }
    }
    
    return { error: 'Selected element is not an image or code block' };
  } catch (e) {
    return { error: e.toString() };
  }
}

function convertMarkdown(markdown) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const cursor = doc.getCursor();
    const docId = doc.getId();
    
    let insertIndex = body.getNumChildren();
    if (cursor) {
      const element = cursor.getElement();
      insertIndex = body.getChildIndex(element);
    }
    
    processContent(body, markdown, insertIndex, docId);
    return { success: true, message: 'Document formatted successfully!' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function decodeHtmlEntities(text) {
  if (!text) return text;
  
  // Handle double-encoding (e.g., &amp;gt; -> &gt; -> >)
  let decoded = text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#x2F;/g, '/')
    .replace(/&#x27;/g, "'")
    .replace(/&#[0-9]+;/g, function(match) {
      return String.fromCharCode(parseInt(match.replace(/&#|;/g, '')));
    });
  
  // Handle potential double-encoding recursively (once more)
  if (decoded.includes('&amp;') || decoded.includes('&lt;') || decoded.includes('&gt;')) {
    decoded = decoded
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>');
  }
  
  return decoded;
}

function processContent(body, markdown, startIndex, docId) {
  // Decode HTML entities first (critical for copied HTML content)
  markdown = decodeHtmlEntities(markdown);
  
  const lines = markdown.split('\n');
  let i = 0;
  let currentIndex = startIndex;
  const settings = getSettings();
  
  while (i < lines.length) {
    let line = lines[i] || '';
    
    // Code blocks & Mermaid
    if (line.match(/^```/)) {
      const lang = line.match(/^```(\w*)/)?.[1] || '';
      const codeContent = [];
      i++;
      while (i < lines.length && !lines[i].match(/^```/)) {
        codeContent.push(lines[i]);
        i++;
      }
      
      const code = codeContent.join('\n');
      if (lang.toLowerCase() === 'mermaid') {
        currentIndex = insertMermaidDiagram(body, code, currentIndex, docId, settings);
      } else {
        currentIndex = insertCodeBlock(body, code, lang, currentIndex);
      }
      i++;
      continue;
    }
    
    // Tables - Fixed to handle inline formatting
    if (line.match(/^\s*\|.*\|\s*$/)) {
      const tableLines = [];
      let alignments = [];
      
      while (i < lines.length && lines[i].match(/^\s*\|.*\|\s*$/)) {
        const currentLine = lines[i].trim();
        // Check if it's an alignment row |:---|---:|
        if (currentLine.replace(/\|/g, '').trim().match(/^[\s:\-]+$/)) {
          alignments = parseAlignments(currentLine);
        } else {
          tableLines.push(currentLine);
        }
        i++;
      }
      
      if (tableLines.length > 0) {
        currentIndex = insertTable(body, tableLines, alignments, currentIndex);
      }
      continue;
    }
    
    // Headers
    const headerMatch = line.match(/^(#{1,6})\s+(.+)$/);
    if (headerMatch) {
      const level = headerMatch[1].length;
      const text = headerMatch[2];
      currentIndex = insertHeading(body, text, level, currentIndex);
      i++;
      continue;
    }
    
    // Horizontal rule
    if (line.match(/^(-{3,}|\*{3,}|_{3,})$/)) {
      currentIndex = insertHorizontalRule(body, currentIndex);
      i++;
      continue;
    }
    
    // Blockquotes - IMPROVED: handles leading whitespace and >
    const bqMatch = line.match(/^(\s*)>\s?(.*)$/);
    if (bqMatch) {
      const quoteLines = [bqMatch[2]];  // Get content after >
      i++;
      while (i < lines.length) {
        const nextLine = lines[i];
        const nextMatch = nextLine.match(/^(\s*)>\s?(.*)$/);
        if (nextMatch) {
          quoteLines.push(nextMatch[2]);
          i++;
        } else if (nextLine.trim() === '') {
          // Empty line ends blockquote
          break;
        } else {
          break;
        }
      }
      currentIndex = insertBlockquote(body, quoteLines.join('\n'), currentIndex);
      continue;
    }
    
    // Lists
    const listMatch = line.match(/^(\s*)([-*+]|\d+\.)\s+(.*)$/);
    if (listMatch) {
      const isOrdered = /^\d+\./.test(listMatch[2]);
      const baseIndent = listMatch[1].length;
      const items = [];
      
      while (i < lines.length) {
        const match = lines[i]?.match(/^(\s*)([-*+]|\d+\.)\s+(.*)$/);
        if (!match) break;
        
        const indent = match[1].length;
        const level = Math.max(0, Math.floor((indent - baseIndent) / 2));
        items.push({ level, text: match[3] });
        i++;
        
        // Continue this list item on next lines (if indented or continuation)
        // BUT stop if we hit a table, code block, blockquote, or header
        while (i < lines.length) {
          const nextLine = lines[i];
          
          // Check for patterns that should BREAK out of list processing
          if (nextLine.match(/^```/)) break; // Code block
          if (nextLine.match(/^\s*\|.*\|/)) break; // Table
          if (nextLine.match(/^>\s?/)) break; // Blockquote
          if (nextLine.match(/^#{1,6}\s/)) break; // Header
          if (nextLine.match(/^(-{3,}|\*{3,}|_{3,})$/)) break; // HR
          if (nextLine.trim() === '') break; // Empty line
          if (nextLine.match(/^(\s*)([-*+]|\d+\.)\s+/)) break; // New list item
          
          // Otherwise, continue this list item
          items[items.length - 1].text += ' ' + nextLine.trim();
          i++;
        }
      }
      
      currentIndex = insertNestedList(body, items, isOrdered, currentIndex);
      continue;
    }
    
    // Regular paragraph with inline formatting
    if (line.trim() !== '') {
      const para = body.insertParagraph(currentIndex, '');
      formatInlineText(para, line);
      para.setSpacingAfter(6);
      currentIndex++;
    }
    
    i++;
  }
}

function insertMermaidDiagram(body, code, index, docId, settings) {
  settings = settings || getSettings();
  
  // Always try Kroki first (most reliable), then fallback to selected provider
  const providers = [
    { name: 'kroki', fn: () => fetchKrokiDiagram(code, 'mermaid', settings.theme || 'neutral') },
    { name: settings.mermaidProvider, fn: () => fetchMermaidProvider(code, settings) }
  ];
  
  let lastError = '';
  
  for (const provider of providers) {
    try {
      console.log(`Trying mermaid provider: ${provider.name}`);
      const imageBlob = provider.fn();
      
      if (imageBlob) {
        const image = body.insertImage(index, imageBlob);
        
        // Calculate height maintaining aspect ratio
        const originalHeight = image.getHeight();
        const originalWidth = image.getWidth();
        const newWidth = 600;
        const newHeight = (originalHeight * newWidth) / originalWidth;
        
        image.setWidth(newWidth);
        image.setHeight(newHeight);
        
        // Center the image
        const para = image.getParent().asParagraph();
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        
        handleCodeRetention(body, image, code, index, docId, settings.codeRetention);
        return index + 1;
      }
    } catch (e) {
      console.error(`Provider ${provider.name} failed:`, e);
      lastError = e.message;
    }
  }
  
  // All failed - insert as code block with error comment
  console.error('All mermaid providers failed:', lastError);
  const errorCode = `%% Error: ${lastError}\n%% Fallback: Please check diagram syntax at mermaid.live\n\n${code}`;
  return insertCodeBlock(body, errorCode, 'mermaid', index);
}

function fetchMermaidProvider(diagramCode, settings) {
  switch(settings.mermaidProvider) {
    case 'mermaidchart':
      if (!settings.apiKey) throw new Error('MermaidChart API key not configured');
      return fetchMermaidChart(diagramCode, settings.apiKey, settings.theme);
      
    case 'custom':
      if (!settings.apiKey) throw new Error('Custom API URL not configured');
      return fetchCustomDiagram(diagramCode, settings.apiKey, settings.theme);
      
    case 'mermaid.ink':
    default:
      const encoded = Utilities.base64Encode(diagramCode, Utilities.Charset.UTF_8);
      const url = `https://mermaid.ink/img/${encoded}?theme=${settings.theme}&width=800`;
      
      console.log('Mermaid.ink URL length:', url.length);
      
      if (url.length > 8000) {
        throw new Error('Diagram too complex for mermaid.ink');
      }
      
      const response = UrlFetchApp.fetch(url, { 
        muteHttpExceptions: true,
        method: 'GET'
      });
      
      const responseCode = response.getResponseCode(); // FIXED: renamed from 'code'
      console.log('Mermaid.ink response:', responseCode);
      
      if (responseCode === 200) {
        return response.getBlob();
      } else if (responseCode === 400) {
        throw new Error('Invalid mermaid syntax (400)');
      } else if (responseCode === 414) {
        throw new Error('Diagram too large (414)');
      } else {
        throw new Error(`HTTP ${responseCode}`);
      }
  }
}

function handleCodeRetention(body, image, code, index, docId, method) {
  switch(method) {
    case 'alt':
      image.setAltDescription(code);
      image.setAltTitle('Mermaid Source Code');
      break;
      
    case 'caption':
      const caption = body.insertParagraph(index + 1, '');
      caption.setFontSize(9);
      caption.setForegroundColor('#5f6368');
      caption.setItalic(true);
      caption.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      const preview = code.length > 200 ? code.substring(0, 200) + '...' : code;
      caption.appendText('📊 Source: ' + preview.replace(/\n/g, ' '));
      caption.setSpacingBefore(4);
      caption.setSpacingAfter(12);
      break;
      
    case 'comment':
      addCommentToImage(docId, image, code);
      break;
      
    case 'none':
    default:
      break;
  }
}

function fetchMermaidChart(code, apiKey, theme) {
  const url = 'https://api.mermaidchart.com/diagrams';
  const response = UrlFetchApp.fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      code: code,
      theme: theme,
      format: 'png'
    }),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 200) {
    return response.getBlob();
  }
  throw new Error('MermaidChart API: ' + response.getContentText());
}

function fetchKrokiDiagram(code, diagramType, theme) {
  const encoded = Utilities.base64Encode(code, Utilities.Charset.UTF_8);
  const url = `https://kroki.io/${diagramType}/png/${encoded}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  
  if (response.getResponseCode() === 200) {
    return response.getBlob();
  }
  throw new Error('Kroki.io service error');
}

function fetchCustomDiagram(code, apiUrl, theme) {
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    payload: {
      code: code,
      theme: theme
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 200) {
    return response.getBlob();
  }
  throw new Error('Custom API error');
}

function addCommentToImage(docId, image, code) {
  try {
    const commentText = `🧜‍♀️ Mermaid Source Code:\n\n${code}\n\n---\nYou can copy this code to edit the diagram.`;
    
    if (typeof Drive !== 'undefined' && Drive.Comments) {
      const resource = {
        content: commentText,
        context: {
          type: 'text',
          value: 'Mermaid Diagram'
        }
      };
      Drive.Comments.insert(resource, docId);
    } else {
      image.setAltDescription(code);
    }
  } catch (e) {
    console.error('Comment addition failed:', e);
    image.setAltDescription(code);
  }
}

// FIXED: Removed invalid BorderStyle references
function insertHeading(body, text, level, index) {
  const para = body.insertParagraph(index, text);
  const style = DocumentApp.ParagraphHeading['HEADING' + level];
  para.setHeading(style);
  para.setSpacingBefore(12);
  para.setSpacingAfter(6);
  return index + 1;
}

function insertCodeBlock(body, code, lang, index) {
  const table = body.insertTable(index, [['']]);
  table.setBorderWidth(0);
  table.setColumnWidth(0, 520);
  
  const cell = table.getCell(0, 0);
  cell.setBackgroundColor('#f6f8fa');
  cell.setPaddingTop(12);
  cell.setPaddingBottom(12);
  cell.setPaddingLeft(16);
  cell.setPaddingRight(16);
  
  const para = cell.getChild(0).asParagraph();
  para.setFontFamily('Courier New');
  para.setFontSize(10);
  
  const lines = code.split('\n');
  if (lines.length > 0) {
    para.setText(lines[0]);
    for (let i = 1; i < lines.length; i++) {
      const newPara = cell.appendParagraph(lines[i]);
      newPara.setFontFamily('Courier New');
      newPara.setFontSize(10);
      newPara.setForegroundColor('#24292e');
    }
  }
  
  return index + 1;
}

function insertTable(body, lines, alignments, index) {
  // Parse cells preserving content with formatting markers
  const cells = lines.map(line => {
    // Remove leading/trailing pipes, keep internal content intact
    const cleaned = line.replace(/^\s*\|/, '').replace(/\|\s*$/, '');
    return cleaned.split('|').map(cell => cell.trim());
  });
  
  if (cells.length === 0 || cells[0].length === 0) return index;
  
  const table = body.insertTable(index, cells);
  table.setBorderWidth(1);
  table.setBorderColor('#d0d7de');
  
  // Style header row
  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    const cell = headerRow.getCell(i);
    cell.setBackgroundColor('#f6f8fa');
    
    // Process inline formatting in header cells
    const para = cell.getChild(0).asParagraph();
    const cellText = cell.getText();  // Get current text
    para.clear();  // Clear to re-format
    formatInlineText(para, cellText);  // Re-insert with formatting
    para.setBold(true);  // Headers are bold
    para.setFontSize(11);
    
    if (alignments[i] === 'center') {
      para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } else if (alignments[i] === 'right') {
      para.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    }
  }
  
  // Style body cells with inline formatting
  for (let r = 1; r < table.getNumRows(); r++) {
    for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
      const cell = table.getRow(r).getCell(c);
      cell.setPaddingTop(8);
      cell.setPaddingBottom(8);
      
      const para = cell.getChild(0).asParagraph();
      const cellText = cell.getText();
      para.clear();
      formatInlineText(para, cellText);  // Process bold/code/etc in cells
      
      if (alignments[c] === 'center') {
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      } else if (alignments[c] === 'right') {
        para.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      }
    }
  }
  
  return index + 1;
}

function parseAlignments(line) {
  const cells = line.split('|').filter(c => c.trim() !== '');
  return cells.map(cell => {
    cell = cell.trim();
    if (cell.startsWith(':') && cell.endsWith(':')) return 'center';
    if (cell.endsWith(':')) return 'right';
    return 'left';
  });
}

function insertBlockquote(body, text, index) {
  // Insert paragraph with text first
  const para = body.insertParagraph(index, '');
  
  // Process inline formatting (bold, italic, links, code)
  formatInlineText(para, text);
  
  // Apply blockquote styling to the entire paragraph
  para.setIndentStart(36);
  para.setIndentFirstLine(36);
  para.setItalic(true);  // Base italic for blockquote
  para.setForegroundColor('#57606a');
  para.setSpacingBefore(6);
  para.setSpacingAfter(6);
  
  return index + 1;
}

// FIXED: Removed invalid Table border methods, using simple paragraph with dashes instead
function insertHorizontalRule(body, index) {
  const para = body.insertParagraph(index, '');
  para.setSpacingBefore(12);
  para.setSpacingAfter(12);
  
  // Use em-dashes as a visual horizontal rule
  const text = para.appendText('────────────────────────────────────────');
  text.setForegroundColor('#d0d7de');
  text.setFontSize(11);
  
  return index + 1;
}

// FIXED: Removed invalid setIndentEnd method from ListItem
function insertNestedList(body, items, isOrdered, index) {
  items.forEach((item, i) => {
    const listItem = body.insertListItem(index + i, item.text);
    listItem.setGlyphType(isOrdered ? DocumentApp.GlyphType.NUMBER : DocumentApp.GlyphType.BULLET);
    
    if (item.level > 0) {
      listItem.setIndentStart(36 * item.level);
      // Removed: listItem.setIndentEnd(36 * item.level); // This method doesn't exist
    }
    
    formatInlineText(listItem, item.text);
  });
  
  return index + items.length;
}

function formatInlineText(element, text) {
  element.clear();
  
  const linkRegex = /\[([^\]]+)\]\(([^)]+)\)/g;
  let lastIndex = 0;
  let match;
  const segments = [];
  
  while ((match = linkRegex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      segments.push({ text: text.substring(lastIndex, match.index), type: 'text' });
    }
    segments.push({ text: match[1], url: match[2], type: 'link' });
    lastIndex = match.index + match[0].length;
  }
  
  if (lastIndex < text.length) {
    segments.push({ text: text.substring(lastIndex), type: 'text' });
  }
  
  if (segments.length === 0) {
    segments.push({ text: text, type: 'text' });
  }
  
  segments.forEach(segment => {
    if (segment.type === 'link') {
      const textElement = element.appendText(segment.text);
      textElement.setLinkUrl(segment.url);
      textElement.setForegroundColor('#0969da');
      textElement.setUnderline(true);
    } else {
      processInlineStyles(element, segment.text);
    }
  });
}

function processInlineStyles(element, text) {
  const tokens = [];
  let current = '';
  let i = 0;
  
  while (i < text.length) {
    if (text.substr(i, 3) === '***' || text.substr(i, 3) === '___') {
      if (current) tokens.push({ text: current, style: null });
      tokens.push({ type: 'toggle', style: 'bold-italic', marker: text.substr(i, 3) });
      current = '';
      i += 3;
    } else if (text.substr(i, 2) === '**' || text.substr(i, 2) === '__') {
      if (current) tokens.push({ text: current, style: null });
      tokens.push({ type: 'toggle', style: 'bold', marker: text.substr(i, 2) });
      current = '';
      i += 2;
    } else if (text.substr(i, 2) === '~~') {
      if (current) tokens.push({ text: current, style: null });
      tokens.push({ type: 'toggle', style: 'strike', marker: '~~' });
      current = '';
      i += 2;
    } else if ((text[i] === '*' || text[i] === '_') && text[i-1] !== '\\') {
      if (current) tokens.push({ text: current, style: null });
      tokens.push({ type: 'toggle', style: 'italic', marker: text[i] });
      current = '';
      i++;
    } else if (text[i] === '`') {
      if (current) tokens.push({ text: current, style: null });
      tokens.push({ type: 'toggle', style: 'code', marker: '`' });
      current = '';
      i++;
    } else {
      current += text[i];
      i++;
    }
  }
  
  if (current) tokens.push({ text: current, style: null });
  
  const activeStyles = new Set();
  tokens.forEach(token => {
    if (token.type === 'toggle') {
      if (activeStyles.has(token.style)) {
        activeStyles.delete(token.style);
      } else {
        activeStyles.add(token.style);
      }
    } else if (token.text) {
      const textElem = element.appendText(token.text);
      
      if (activeStyles.has('bold') || activeStyles.has('bold-italic')) {
        textElem.setBold(true);
      }
      if (activeStyles.has('italic') || activeStyles.has('bold-italic')) {
        textElem.setItalic(true);
      }
      if (activeStyles.has('strike')) {
        textElem.setStrikethrough(true);
      }
      if (activeStyles.has('code')) {
        textElem.setFontFamily('Courier New');
        textElem.setFontSize(10);
        textElem.setBackgroundColor('#eff1f3');
        textElem.setForegroundColor('#24292e');
      }
    }
  });
}