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

function decodeHtmlEntities(text) {
  if (!text) return text;
  
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
  
  if (decoded.includes('&amp;') || decoded.includes('&lt;') || decoded.includes('&gt;')) {
    decoded = decoded
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>');
  }
  
  return decoded;
}

function convertMarkdown(markdown) {
  try {
    if (!markdown || typeof markdown !== 'string') {
      return { success: false, message: 'No markdown content provided' };
    }
    
    markdown = decodeHtmlEntities(markdown);
    
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
    console.error('Conversion error:', error);
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function processContent(body, markdown, startIndex, docId) {
  const lines = markdown.split(/\r?\n/);
  let i = 0;
  let currentIndex = startIndex;
  const settings = getSettings();
  
  while (i < lines.length) {
    let line = lines[i] || '';
    
    // Code blocks & Mermaid
    if (line.match(/^```/)) {
      const langMatch = line.match(/^```(\w*)/);
      const lang = langMatch ? langMatch[1] : '';
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
    
    // Tables
    if (line.match(/^\s*\|.*\|\s*$/)) {
      const tableLines = [];
      let alignments = [];
      
      while (i < lines.length && lines[i].match(/^\s*\|.*\|\s*$/)) {
        const currentLine = lines[i].trim();
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
    
    // Blockquotes
    const bqMatch = line.match(/^(\s*)>\s?(.*)$/);
    if (bqMatch) {
      const quoteLines = [bqMatch[2]];
      i++;
      while (i < lines.length) {
        const nextLine = lines[i];
        const nextMatch = nextLine.match(/^(\s*)>\s?(.*)$/);
        if (nextMatch) {
          quoteLines.push(nextMatch[2]);
          i++;
        } else if (nextLine.trim() === '') {
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
        const match = lines[i].match(/^(\s*)([-*+]|\d+\.)\s+(.*)$/);
        if (!match) break;
        
        const indent = match[1].length;
        const level = Math.max(0, Math.floor((indent - baseIndent) / 2));
        items.push({ level, text: match[3] });
        i++;
        
        while (i < lines.length) {
          const nextLine = lines[i];
          
          if (nextLine.match(/^```/)) break;
          if (nextLine.match(/^\s*\|.*\|/)) break;
          if (nextLine.match(/^>\s?/)) break;
          if (nextLine.match(/^#{1,6}\s/)) break;
          if (nextLine.match(/^(-{3,}|\*{3,}|_{3,})$/)) break;
          if (nextLine.trim() === '') break;
          if (nextLine.match(/^(\s*)([-*+]|\d+\.)\s+/)) break;
          
          items[items.length - 1].text += ' ' + nextLine.trim();
          i++;
        }
      }
      
      currentIndex = insertNestedList(body, items, isOrdered, currentIndex);
      continue;
    }
    
    // Regular paragraph
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
  
  let errorLog = [];
  
  const providers = [
    { name: 'kroki', fn: () => fetchKrokiDiagram(code, 'mermaid', settings.theme) },
    { name: 'mermaid.ink', fn: () => fetchMermaidInk(code, settings.theme) }
  ];
  
  if (settings.mermaidProvider === 'mermaidchart' && settings.apiKey) {
    providers.unshift({ 
      name: 'mermaidchart', 
      fn: () => fetchMermaidChart(code, settings.apiKey, settings.theme) 
    });
  } else if (settings.mermaidProvider === 'custom' && settings.apiKey) {
    providers.unshift({ 
      name: 'custom', 
      fn: () => fetchCustomDiagram(code, settings.apiKey, settings.theme) 
    });
  }
  
  for (const provider of providers) {
    try {
      console.log(`Trying ${provider.name}...`);
      const imageBlob = provider.fn();
      
      if (imageBlob) {
        console.log(`${provider.name} succeeded`);
        
        // Insert image - ONLY ONCE
        const image = body.insertImage(index, imageBlob);
        
        // Google Docs sometimes auto-resizes to page width
        // We need to force it back to natural size or scaled proportionally
        
        // Lock aspect ratio first (in case it was unlocked and stretched)
        // Note: Apps Script doesn't expose lockAspectRatio, so we calculate both dimensions
        
        const docWidth = 580; // approximate usable width in points
        const imgWidth = image.getWidth();
        const imgHeight = image.getHeight();
        
        // If image is wider than page, scale it down proportionally
        if (imgWidth > docWidth) {
          const ratio = docWidth / imgWidth;
          image.setWidth(docWidth);
          image.setHeight(imgHeight * ratio);
        } else {
          // Keep natural size, but ensure both are set to prevent stretching
          image.setWidth(imgWidth);
          image.setHeight(imgHeight);
        }
        
        // Center it
        const para = image.getParent().asParagraph();
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        
        handleCodeRetention(body, image, code, index, docId, settings.codeRetention);
        return index + 1;
      }
    } catch (e) {
      console.error(`${provider.name} failed:`, e);
      errorLog.push(`${provider.name}: ${e.message}`);
    }
  }
  
  // All failed
  const errorMsg = errorLog.join('; ');
  const fallbackCode = `%% Rendering Error: ${errorMsg}\n%% Check syntax at https://mermaid.live\n\n${code}`;
  return insertCodeBlock(body, fallbackCode, 'mermaid', index);
}

function applyMermaidTheme(code, theme) {
  if (!theme || theme === 'neutral' || theme === 'default') {
    return code;
  }
  
  if (code.match(/^%%{init:/)) {
    return code.replace(/(theme\s*:\s*['"])\w*(['"])/, `$1${theme}$2`);
  }
  
  return `%%{init: {'theme': '${theme}'}}%%\n${code}`;
}

function fetchMermaidInk(diagramCode, theme) {
  const themedCode = applyMermaidTheme(diagramCode, theme);
  const encoded = Utilities.base64Encode(themedCode, Utilities.Charset.UTF_8);
  const url = `https://mermaid.ink/img/${encoded}?width=800`;
  
  console.log('Mermaid.ink URL length:', url.length, 'Theme:', theme);
  
  if (url.length > 8000) {
    throw new Error('Diagram too large for mermaid.ink');
  }
  
  const response = UrlFetchApp.fetch(url, { 
    method: 'GET',
    muteHttpExceptions: true 
  });
  
  const responseCode = response.getResponseCode();
  console.log('Mermaid.ink response:', responseCode);
  
  if (responseCode === 200) {
    return response.getBlob();
  } else if (responseCode === 414) {
    throw new Error('URI Too Long');
  } else {
    const errorText = response.getContentText().substring(0, 200);
    throw new Error(`HTTP ${responseCode}: ${errorText}`);
  }
}

function fetchKrokiDiagram(diagramCode, diagramType, theme) {
  const themedCode = applyMermaidTheme(diagramCode, theme);
  const url = `https://kroki.io/${diagramType}/png`;
  
  console.log('Kroki endpoint:', url, 'Theme:', theme);
  
  const response = UrlFetchApp.fetch(url, {
    method: 'POST',
    contentType: 'text/plain',
    payload: themedCode,
    muteHttpExceptions: true
  });
  
  const responseCode = response.getResponseCode();
  console.log('Kroki response:', responseCode);
  
  if (responseCode === 200) {
    return response.getBlob();
  } else if (responseCode === 400) {
    const errorText = response.getContentText().substring(0, 200);
    throw new Error(`Invalid syntax: ${errorText}`);
  } else {
    throw new Error(`HTTP ${responseCode}`);
  }
}

function fetchMermaidChart(diagramCode, apiKey, theme) {
  const url = 'https://api.mermaidchart.com/rest-api/diagrams/render';
  
  const payload = {
    code: diagramCode,
    format: 'png'
  };
  
  if (theme && theme !== 'neutral' && theme !== 'default') {
    payload.theme = theme;
  }
  
  const response = UrlFetchApp.fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const responseCode = response.getResponseCode();
  
  if (responseCode === 200) {
    return response.getBlob();
  } else if (responseCode === 401) {
    throw new Error('Invalid API key');
  } else {
    const errorText = response.getContentText().substring(0, 200);
    throw new Error(`HTTP ${responseCode}: ${errorText}`);
  }
}

function fetchCustomDiagram(diagramCode, apiUrl, theme) {
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      code: diagramCode,
      theme: theme
    }),
    muteHttpExceptions: true
  });
  
  const responseCode = response.getResponseCode();
  
  if (responseCode === 200) {
    return response.getBlob();
  } else {
    throw new Error(`HTTP ${responseCode}`);
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
      caption.appendText('Source: ' + preview.replace(/\n/g, ' '));
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

function addCommentToImage(docId, image, code) {
  try {
    if (typeof Drive !== 'undefined' && Drive.Comments) {
      const resource = {
        content: `Mermaid Source:\n\n${code}`,
        context: { type: 'text', value: 'Mermaid Diagram' }
      };
      Drive.Comments.insert(resource, docId);
    } else {
      image.setAltDescription(code);
    }
  } catch (e) {
    image.setAltDescription(code);
  }
}

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
  const cells = lines.map(line => {
    const cleaned = line.replace(/^\s*\|/, '').replace(/\|\s*$/, '');
    return cleaned.split('|').map(cell => cell.trim());
  });
  
  if (cells.length === 0 || cells[0].length === 0) return index;
  
  const table = body.insertTable(index, cells);
  table.setBorderWidth(1);
  table.setBorderColor('#d0d7de');
  
  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    const cell = headerRow.getCell(i);
    cell.setBackgroundColor('#f6f8fa');
    const para = cell.getChild(0).asParagraph();
    const cellText = cell.getText();
    para.clear();
    formatInlineText(para, cellText);
    para.setBold(true);
    para.setFontSize(11);
    
    if (alignments[i] === 'center') {
      para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } else if (alignments[i] === 'right') {
      para.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    }
  }
  
  for (let r = 1; r < table.getNumRows(); r++) {
    for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
      const cell = table.getRow(r).getCell(c);
      cell.setPaddingTop(8);
      cell.setPaddingBottom(8);
      
      const para = cell.getChild(0).asParagraph();
      const cellText = cell.getText();
      para.clear();
      formatInlineText(para, cellText);
      
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
  const cells = line.split('|').filter(c => c.trim() !== '' && c.trim().match(/^[\s:\-]+$/));
  return cells.map(cell => {
    cell = cell.trim();
    if (cell.startsWith(':') && cell.endsWith(':')) return 'center';
    if (cell.endsWith(':')) return 'right';
    return 'left';
  });
}

function insertBlockquote(body, text, index) {
  const para = body.insertParagraph(index, '');
  formatInlineText(para, text);
  para.setIndentStart(36);
  para.setIndentFirstLine(36);
  para.setItalic(true);
  para.setForegroundColor('#57606a');
  para.setSpacingBefore(6);
  para.setSpacingAfter(6);
  return index + 1;
}

function insertHorizontalRule(body, index) {
  const para = body.insertParagraph(index, '');
  para.setSpacingBefore(12);
  para.setSpacingAfter(12);
  const text = para.appendText('----------------------------------------');
  text.setForegroundColor('#d0d7de');
  text.setFontSize(11);
  return index + 1;
}

function insertNestedList(body, items, isOrdered, index) {
  items.forEach((item, i) => {
    const listItem = body.insertListItem(index + i, item.text);
    listItem.setGlyphType(isOrdered ? DocumentApp.GlyphType.NUMBER : DocumentApp.GlyphType.BULLET);
    
    if (item.level > 0) {
      listItem.setIndentStart(36 * item.level);
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
