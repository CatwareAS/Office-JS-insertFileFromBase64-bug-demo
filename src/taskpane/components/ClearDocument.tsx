import React from "react";
import { Button } from "@fluentui/react-components";
import { clearDocumentBody, clearDocumentHeadersAndFooters } from "../../utils/documentUtils";
import { DocumentBorderFilled, DocumentHeaderFooterRegular } from "@fluentui/react-icons";

const ClearDocument = () => {
  const onDocumentBodyClear = async () => {
    await Word.run(async (context) => {
      await clearDocumentBody(context);
      await context.sync();
    });
  };

  const onHeadersAndFootersClear = async () => {
    await Word.run(async (context) => {
      await clearDocumentHeadersAndFooters(context);
      await context.sync();
    });
  };

  return (
    <div>
      <p>Clear document:</p>
      <div className="row">
        <Button icon={<DocumentBorderFilled />} onClick={onDocumentBodyClear}>
          Clear document body
        </Button>
        <Button icon={<DocumentHeaderFooterRegular />} onClick={onHeadersAndFootersClear}>
          Clear headers and footers
        </Button>
      </div>
    </div>
  );
};

export default ClearDocument;
