## Office-JS insertFileFromBase64 bug demo

**The issue:** When a document is divided into multiple sections and each section has custom headers and footers (which are not "Linked to Previous"), the document is inserted incorrectly using insertFileFromBase64 method  

### Environment

- Platform: PC desktop
- Host: Word
- Office version number: Version 2505 (Build 18825.20150)
- Operating System: Windows 11

### Document properties

To reproduce the issue the document should have:

- multiple sections (added via Layout -> Breaks -> Section Breaks)
- custom headers and footers for each section (they should NOT be "Linked to Previous")

Document examples:
 [Document_1.docx](src/documents/Document_1.docx), 
[Document_2.docx](src/documents/Document_2.docx)

Base64 versions of the documents:
[document1.ts](src/base64/document1.ts),
[document2.ts](src/base64/document2.ts)

### Expected behavior

When the base64 inserted after some other one was inserted, headers and footers are inserted correctly

(as if the base64 was inserted into a new document)

### Current behavior 

When the base64 is inserted after some other one was inserted, headers and footers are inserted incorrectly.
The header content is missing. The footer content is missing. And sometimes the footer content is merged with previously inserted document 

### Steps to reproduce (case 1)

1. insert a base64 document file using insertFileFromBase64 
[Document_1.docx](src/documents/Document_1.docx)

```js
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
```

Result: The document is inserted correctly

2. insert another base64  document using the same method
[Document_2.docx](src/documents/Document_2.docx)

Result: The heading of the first section is missing: [Document_1_case_1_error.docx](src/documents/Document_2_case_1_error.docx)

3. insert the first base64 document again
[Document_1.docx](src/documents/Document_1.docx). 

Result: Header is missing and footer has extra content [Document_1_case_1_error.docx](src/documents/Document_1_case_1_error.docx)

### Steps to reproduce (case 2)

In the case 1 we only use `context.document.body.clear();` method before inserting a new document

If we try to clear headers and footers before inserting a new document, the merging of content doesn't happen, but document is still inserted incorrectly

1. insert a base64 document using insertFileFromBase64, clearing the headers, footers and body before that
[Document_2.docx](src/documents/Document_2.docx)

```js
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
```
Result: The header and footer for the first sections are missing [Document_2_case_2_error.docx](src/documents/Document_2_case_2_error.docx)


### Context

Solving this issue is essential for us because our application involves switching through many documents.


### Demo

The demo project lies withing this repository. You can install and run it locally

![demo-app.png](assets/demo-app.png)