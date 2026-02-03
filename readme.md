# Google Doc Hyperlink Replacement Solution

A **Google Apps Script** utility to selectively replace hyperlinks in a Google Doc.  
It only replaces text if the matched text is already a hyperlink, preserving all formatting.

---

## Features
- Replace hyperlinks while keeping anchor text and formatting intact.
- Ignores plain text (won’t touch non-linked text).
- Handles partial links within a text element.
- Works in paragraphs, lists, tables, headers, and footers.
- Easy to customize old text, replacement text, and new URL.

---

## Usage

1. Open your Google Doc.
2. Go to **Extensions → Apps Script**.
3. Copy the script below into your project:

```javascript
const changeTextWithUrlIfLinked = () => {
  const searchText = 'Blue Widgets Inc.';
  const replaceText = 'Orange Inc.';
  const replaceUrl = 'https://digitalinspiration.com/';

  const document = DocumentApp.getActiveDocument();
  const body = document.getBody();
  let search = null;

  while ((search = body.findText(searchText, search))) {
    const element = search.getElement();
    const startIndex = search.getStartOffset();
    const endIndex = search.getEndOffsetInclusive();
    const textElement = element.asText();

    // Only replace if any part of the matched text is linked
    let isLinked = false;
    for (let i = startIndex; i <= endIndex; i++) {
      if (textElement.getLinkUrl(i)) {
        isLinked = true;
        break;
      }
    }
    if (!isLinked) continue;

    // Replace linked text with new text and URL
    textElement.deleteText(startIndex, endIndex);
    textElement.insertText(startIndex, replaceText);
    textElement.setLinkUrl(startIndex, startIndex + replaceText.length - 1, replaceUrl);
  }

  document.saveAndClose();
};
