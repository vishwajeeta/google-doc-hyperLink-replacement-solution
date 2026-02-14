const addLinkToParagraphsOnly = () => {
  const searchText = 'Awards for Best Preschool in Bangalore';
  const linkUrl = 'https://www.google.com/maps/place/Podar+Juniors+Preschool+%26+Day+care+-Yelahanka+new+town/@13.0969154,77.5867624,17z/data=!3m1!4b1!4m6!3m5!1s0x3bae195977bbaf6d:0x163efdde1ad52672!8m2!3d13.0969154!4d77.5893427!16s%2Fg%2F11h2d9my1r?entry=ttu&g_ep=EgoyMDI2MDIwNC4wIKXMDSoASAFQAw%3D%3D';

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const total = body.getNumChildren();

  for (let i = 0; i < total; i++) {
    const element = body.getChild(i);

    // Only allow PARAGRAPH elements
    if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const paragraph = element.asParagraph();

    // Strictly allow only NORMAL paragraphs (exclude title & headings)
    if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) continue;

    const text = paragraph.editAsText();
    let match = null;

    while ((match = text.findText(searchText, match))) {
      const start = match.getStartOffset();
      const end = match.getEndOffsetInclusive();

      // Skip if already linked
      if (text.getLinkUrl(start)) continue;

      text.setLinkUrl(start, end, linkUrl);
    }
  }

  doc.saveAndClose();
};
