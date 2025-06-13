import React from "react";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
} from "@fluentui/react-components";

const Description = () => {
  return (
    <Accordion style={{ marginTop: 20 }}>
      <AccordionItem value="1">
        <AccordionHeader>Description</AccordionHeader>
        <AccordionPanel>
          <ul>
            <li>"Insert Document" buttons insert base64 files into the current document</li>
            <li>Each document has helper test "Header/Footer should be"</li>
            <li>Before each document insertion document body is being cleared</li>
            <li>
              The OpenLink button near each document opens base64 file in a new view (can be used to
              see the original document appearance)
            </li>
          </ul>
        </AccordionPanel>
      </AccordionItem>
      <AccordionItem value="2">
        <AccordionHeader>Steps to reproduce (case 1)</AccordionHeader>
        <AccordionPanel>
          The header/footer content is being merged with previous
          <ol>
            <li>Click "Insert Document 1" button - the document is inserted correctly</li>
            <li>
              Click "Insert Document 2" button - the document header for -Section 1- is missing
            </li>
            <li>
              Click "Insert Document 1" button again - the document footer is merged with the
              previous document
            </li>
          </ol>
        </AccordionPanel>
      </AccordionItem>
      <AccordionItem value="3">
        <AccordionHeader>Steps to reproduce (case 2)</AccordionHeader>
        <AccordionPanel>
          The header/footer content is being not inserted
          <ol>
            <li>Click "Clear document body" - the document body is cleared</li>
            <li>Click "Clear headers and footers" - the document header and footer are cleared</li>
            <li>
              Click "Insert Document 2" button - the document is inserted without headers and
              footers for the first section
            </li>
            <li>Click "Clear document body" - the document body is cleared</li>
            <li>Click "Clear headers and footers" - the document header and footer are cleared</li>
            <li>
              Click "Insert Document 1" button - the document is inserted without headers and
              footers for the first section and footer has some unexpected content
            </li>
          </ol>
        </AccordionPanel>
      </AccordionItem>
    </Accordion>
  );
};

export default Description;
