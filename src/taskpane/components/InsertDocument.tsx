import React from "react";
import { Button } from "@fluentui/react-components";
import { DocumentDto } from "../../dto/documentDto";
import { insertDocument, openInANewView } from "../../utils/documentUtils";
import { document1 } from "../../base64/document1";
import { document2 } from "../../base64/document2";
import { DocumentTextRegular, OpenFilled } from "@fluentui/react-icons";

const InsertDocument = () => {
  const documents = [document1, document2];

  const onDocumentInsert = async (document: DocumentDto) => {
    await Word.run(async (context) => {
      const base64String = document.base64;
      await insertDocument(context, base64String);
    });
  };
  const onDocumentOpenInANewView = async (document: DocumentDto) => {
    await Word.run(async (context) => {
      const base64String = document.base64;
      await openInANewView(context, base64String);
    });
  };

  return (
    <div>
      <p>Insert file from base64:</p>
      <div className="stack">
        {documents.map((document) => (
          <div key={document.name} className="row">
            <Button
              icon={<DocumentTextRegular />}
              appearance="primary"
              onClick={() => onDocumentInsert(document)}
            >
              Insert {document.name}
            </Button>
            <Button onClick={() => onDocumentOpenInANewView(document)} icon={<OpenFilled />} />
          </div>
        ))}
      </div>
    </div>
  );
};

export default InsertDocument;
