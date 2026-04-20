const addOnlyOneLinkToParagraphsOnly = () => {
  const searchText = 'KCET 3rd Round Seat Allotment 2026';
  const linkUrl = 'https://www.getmycollege.com/colleges/bangalore/article/kcet-3rd-round-seat-allotment';

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const total = body.getNumChildren();

  for (let i = 0; i < total; i++) {
    const element = body.getChild(i);

    // Only PARAGRAPH elements
    if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const paragraph = element.asParagraph();

    // Only NORMAL paragraphs (no headings)
    if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) continue;

    const text = paragraph.editAsText();

    // Find only the FIRST occurrence
    const match = text.findText(searchText);

    if (match) {
      const start = match.getStartOffset();
      const end = match.getEndOffsetInclusive();

      // Skip if already linked
      if (!text.getLinkUrl(start)) {
        text.setLinkUrl(start, end, linkUrl);
      }
    }
  }

  doc.saveAndClose();
};
