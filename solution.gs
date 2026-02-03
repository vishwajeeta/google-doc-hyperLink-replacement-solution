const changeTextWithUrlIfLinked = () => {
  const searchText = 'NCHM JEE';
  const replaceText = 'NCHM JEE1';
  const replaceUrl = 'https://asvsi.com/';

  const document = DocumentApp.getActiveDocument();
  const body = document.getBody();
  let search = null;

  while ((search = body.findText(searchText, search))) {
    const element = search.getElement();
    const startIndex = search.getStartOffset();
    const endIndex = search.getEndOffsetInclusive();
    const textElement = element.asText();

    // ✅ Only replace if the matched text is already a hyperlink
    const currentLink = textElement.getLinkUrl(startIndex);
    if (!currentLink) {
      continue; // skip plain text
    }

    // Delete old text and insert new
    textElement.deleteText(startIndex, endIndex);
    textElement.insertText(startIndex, replaceText);
    textElement.setLinkUrl(startIndex, startIndex + replaceText.length - 1, replaceUrl);
  }

  document.saveAndClose();
};
