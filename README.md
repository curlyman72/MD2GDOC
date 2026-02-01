# Markdown Converter for Google Docs

**Author:** Jerry Lee  
**Copyright:** 2026 Jerry Lee  
**License:** MIT License  
**Version:** 1.0

A Google Docs add-on that converts AI-generated markdown (ChatGPT, Claude, Gemini, etc.) into professionally formatted Google Docs with full support for tables, syntax-highlighted code blocks, and Mermaid diagrams.

Enjoying this tool? Buy me a coffee! Send spare change via PayPal: [paypal.me/curlyman72](https://paypal.me/curlyman72)

---

## Disclaimer

This software is provided "as is", without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose and noninfringement. In no event shall the authors or copyright holders be liable for any claim, damages or other liability, whether in an action of contract, tort or otherwise, arising from, out of or in connection with the software or the use or other dealings in the software.

Use at your own risk. Always backup your documents before running conversions.

---

## Features

- Complete typography support: Headers (H1-H6), bold, italic, strikethrough, inline code
- Lists: Ordered, unordered, and nested lists with proper indentation
- Tables: GitHub-flavored markdown with column alignment
- Code blocks: Syntax highlighting with monospace fonts and background colors
- Mermaid diagrams: Flowcharts, sequence diagrams, Gantt charts, class diagrams, state diagrams, and ER diagrams rendered as images
- Source code retention: Store Mermaid source in alt text, captions, or Google Doc comments
- Multiple providers: Support for mermaid.ink, MermaidChart, Kroki.io, and custom endpoints
- Source retrieval: Click any rendered diagram to view and copy its original code
- HTML entity decoding: Handles copied content from web sources

---

## Installation

### Step 1: Open Apps Script
1. Open any Google Doc (or create a new blank document)
2. Click **Extensions** -> **Apps Script**
3. A new browser tab opens with the Apps Script editor

### Step 2: Create the Main Script
1. In the Apps Script editor, you should see a file called `Code.gs` on the left
2. Delete any default content inside `Code.gs`
3. Copy and paste the contents of the `Code.gs` file from this repository
4. Press **Ctrl+S** (or Cmd+S on Mac) to save
5. Rename the project: Click "Untitled project" at the top, type "GenAI Markdown Converter", then click **Rename**

### Step 3: Create the User Interface
1. Click the **+** icon next to "Files" on the left sidebar
2. Select **HTML**
3. Name it exactly: `Sidebar` (case-sensitive)
4. Delete the default HTML content
5. Copy and paste the contents of the `Sidebar.html` file from this repository
6. Press **Ctrl+S** to save

### Step 4: Authorize the Script
1. Click on Code.gs to select that file and then click on **Run** in the menu bar
2. Google will ask for authorization:
   - Click **Review Permissions**
   - Choose your Google account
   - Click **Advanced** (if shown)
   - Click **Go to {Document title || Untitled Document} (unsafe)** - this is normal for custom scripts
   - Click **Allow** to grant document editing permissions
3. If prompted about "This app isn't verified", click **Advanced** then **Go to...**

### Step 5: Test the Installation
1. Return to your Google Doc and refresh the page (F5)
2. You should see **"GenAI Markdown"** in the top menu bar
3. Click **GenAI Markdown** -> **Open Converter**
4. A sidebar should appear on the right side of your document

---

## Usage

### Basic Conversion
1. Copy markdown output from your AI tool (ChatGPT, Claude, Gemini, etc.)
2. In your Google Doc, click **GenAI Markdown** -> **Open Converter**
3. Paste the markdown into the text area in the sidebar
4. Click **Convert to Doc**
5. Content will be inserted at your cursor position (or at the end of the document if no cursor is placed)

### Configuring Settings
1. In the sidebar, click the **Settings** tab
2. Configure the following options:
   - **Mermaid Provider:** Choose between mermaid.ink (free), MermaidChart (pro), Kroki.io (free), or Custom
   - **API Key:** Enter your API key if using MermaidChart or a custom endpoint
   - **Theme:** Select diagram theme (neutral, dark, forest, base)
   - **Code Retention:** Choose how to store Mermaid source code:
     - *Alt Text:* Hidden in image metadata (right-click image -> Alt text to view)
     - *Caption:* Shows truncated source below the image
     - *Comment:* Creates a Google Doc comment with full source (requires Drive API)
     - *None:* Image only, no source stored
3. Click **Save Settings**

### Retrieving Source Code
If you stored source code when converting:
1. Click on any rendered diagram in your document to select it
2. Open the **GenAI Markdown** sidebar
3. Click the **View Source** tab
4. Click **Get Source from Selection**
5. The original Mermaid code will appear for copying or editing

---

## Supported Markdown Syntax

### Headers
    # Heading 1
    ## Heading 2
    ### Heading 3
    #### Heading 4
    ##### Heading 5
    ###### Heading 6

### Typography
- `**bold**` or `__bold__` renders as bold text
- `*italic*` or `_italic_` renders as italic text
- `***bold italic***` renders as bold italic text
- `~~strikethrough~~` renders as strikethrough text
- `` `inline code` `` renders as inline code
- `[link text](https://example.com)` renders as clickable links

### Lists

Unordered lists:
    - Item one
    - Item two
      - Nested item A
      - Nested item B
    - Item three

Ordered lists:
    1. First step
    2. Second step
       1. Sub-step A
       2. Sub-step B
    3. Third step

### Code Blocks
Use triple backticks with optional language identifier:

    def hello_world():
        print("Hello, World!")
        return True

### Tables
    | Header 1 | Header 2 | Header 3 |
    |:---------|:--------:|---------:|
    | Left     | Center   | Right    |
    | Cell A   | Cell B   | Cell C   |

Use `:` in the separator line for alignment:
- `:---` for left align
- `:---:` for center align
- `---:` for right align

### Blockquotes
    > This is a blockquote.
    > It can span multiple lines.
    > **Bold** and *italic* work inside blockquotes too.

### Horizontal Rules
Use three or more dashes, asterisks, or underscores on their own line.

### Mermaid Diagrams
    graph TD
        A[Start] --> B{Decision}
        B -->|Yes| C[Action 1]
        B -->|No| D[Action 2]

Supported diagram types include flowcharts, sequence diagrams, Gantt charts, class diagrams, state diagrams, and ER diagrams.

---

## Troubleshooting

### "Script Error" or "Authorization Required"
- Return to the Apps Script editor (Extensions -> Apps Script)
- Click **Run** -> **onOpen** again
- Accept all permissions that appear

### Sidebar Won't Open
1. Refresh your Google Doc (F5)
2. Ensure both `Code.gs` and `Sidebar.html` are saved in Apps Script
3. Check that you ran `onOpen()` at least once to initialize the menu

### Mermaid Diagrams Not Rendering
- Check your internet connection (diagrams render via external API)
- Try switching providers in Settings (if mermaid.ink fails, try Kroki)
- If using MermaidChart, verify your API key is valid
- Check error codes:
  - 414: Diagram too large
  - 400: Syntax error in diagram
  - 429: Rate limited

### Tables or Blockquotes Formatting Incorrectly
- Ensure there are no spaces before table separator lines
- Blockquotes must start with `>` at the beginning of the line
- Lists cannot contain tables or blockquotes without proper spacing

### Code Blocks Look Wrong
- Use standard markdown fences: three backticks on their own lines
- Language specification should be immediately after opening backticks

### "Drive API Not Enabled" Error
- If using "Comment" retention mode, you must enable Drive API:
  1. In Apps Script editor, click **Services** (+) next to "Libraries"
  2. Select **Drive API** and add it
- Or switch Code Retention setting to "Alt Text" which does not require Drive API

---

## Security and Privacy

### API Key Storage
- API keys are stored in your personal Google account via `PropertiesService`
- Keys are user-specific and are **not** shared with other document collaborators
- Keys are **not** transmitted to the script author or any third party except your chosen Mermaid provider
- To remove stored keys: Apps Script editor -> **File** -> **Project Properties** -> **User Properties** -> Delete entries

### Data Processing
- Markdown content is processed locally within Google's infrastructure
- Mermaid diagrams are sent to your selected rendering service (mermaid.ink, Kroki.io, or MermaidChart)
- No document content is logged, stored, or transmitted to the script author
- Google Docs permissions are required solely to insert content into your document

---

## Support

If you find this tool useful, consider supporting development:

Buy me a coffee: **[paypal.me/curlyman72](https://paypal.me/curlyman72)**

Your support helps maintain and improve this project!

---

## License

MIT License

Copyright (c) 2026 Jerry Lee

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

