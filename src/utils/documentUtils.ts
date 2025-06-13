export const clearDocumentBody = async (context: Word.RequestContext) => {
  context.document.body.clear();
};

export const clearDocumentHeadersAndFooters = async (context: Word.RequestContext) => {
  const sections = context.document.sections;
  context.load(sections, "items");

  await context.sync();

  for (const section of sections.items) {
    const header = section.getHeader("Primary");
    const footer = section.getFooter("Primary");

    header.clear();
    footer.clear();

    const firstPageHeader = section.getHeader("FirstPage");
    const firstPageFooter = section.getFooter("FirstPage");

    firstPageHeader.clear();
    firstPageFooter.clear();

    const evenPageHeader = section.getHeader("EvenPages");
    const evenPageFooter = section.getFooter("EvenPages");

    evenPageHeader.clear();
    evenPageFooter.clear();
  }
};

export const insertDocument = async (context: Word.RequestContext, base64String: string) => {
  // clear document headers and footers
  const sections = context.document.sections;
  context.load(sections, "items");

  await context.sync();

  for (const section of sections.items) {
    const header = section.getHeader("Primary");
    const footer = section.getFooter("Primary");

    header.clear();
    footer.clear();

    const firstPageHeader = section.getHeader("FirstPage");
    const firstPageFooter = section.getFooter("FirstPage");

    firstPageHeader.clear();
    firstPageFooter.clear();

    const evenPageHeader = section.getHeader("EvenPages");
    const evenPageFooter = section.getFooter("EvenPages");

    evenPageHeader.clear();
    evenPageFooter.clear();
  }

  // clear document body
  context.document.body.clear();

  //insert new file, import new styles
  context.document.insertFileFromBase64(base64String, "Replace", {
    importTheme: true,
    importStyles: true,
    importParagraphSpacing: true,
    importPageColor: true,
    importDifferentOddEvenPages: true,
    importCustomProperties: true,
    importCustomXmlParts: true,
    importChangeTrackingMode: false,
  });

  await context.sync();
};

export const openInANewView = (context: Word.RequestContext, base64String: string) => {
  context.application.createDocument(base64String).open();
};
