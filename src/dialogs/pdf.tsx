/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* Everything used in here must be self contained */
import * as React from "react";
import * as ReactDOM from "react-dom";
import { Dropbox } from "dropbox";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { Spinner, SpinnerSize, DefaultButton } from "office-ui-fabric-react";
import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from "office-ui-fabric-react/lib/Stack";
import { Document, Page, pdfjs } from "react-pdf";

initializeIcons();

pdfjs.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjs.version}/pdf.worker.js`;

const fileName = decodeURI(new URLSearchParams(window.location.search).get("cell"));
const dbx = new Dropbox({
  accessToken: process.env.TOKEN
});

// Styles definition
const stackStyles: IStackStyles = {
  root: {
    height: 250
  }
};
const stackItemStyles: IStackItemStyles = {
  root: {
    alignItems: "center",
    display: "flex",
    justifyContent: "center"
  }
};

// Tokens definition
const outerStackTokens: IStackTokens = { childrenGap: 5 };
const innerStackTokens: IStackTokens = {
  childrenGap: 5,
  padding: 10
};

const PDF = () => {
  const [message, setMessage] = React.useState("");
  const [src, setSrc] = React.useState("");
  const [numPages, setNumPages] = React.useState(null);
  const [pageNumber, setPageNumber] = React.useState(1);

  const goToPrevPage = () => setPageNumber(pageNumber === 1 ? 1 : pageNumber - 1);
  const goToNextPage = () => setPageNumber(pageNumber === numPages ? numPages : pageNumber + 1);

  React.useEffect(() => {
    const download = async () => {
      const meta: any = await dbx.filesDownload({
        path: `/Ideal Image/${fileName}.pdf`
      });

      if (meta.status !== 200) {
        setMessage(meta.result);
      }

      const blob = URL.createObjectURL(new Blob([meta.result.fileBlob], { type: "application/pdf" }));
      setSrc(blob);
    };
    download();
  }, []);

  if (!src) {
    return <Spinner style={{ padding: 50 }} size={SpinnerSize.large} label="Working on it .." />;
  }

  return (
    <Stack tokens={outerStackTokens}>
      <Stack styles={stackStyles} tokens={innerStackTokens}>
        <Stack.Item grow={2} styles={stackItemStyles}>
          {message}
        </Stack.Item>
        <Stack.Item grow={4} styles={stackItemStyles}>
          <Document onLoadError={console.error} onLoadSuccess={({ numPages }) => setNumPages(numPages)} file={src}>
            <Page pageNumber={pageNumber} />
          </Document>
        </Stack.Item>
        <Stack.Item grow styles={stackItemStyles}>
          <p>
            Page {pageNumber} of {numPages}
          </p>
        </Stack.Item>
        <Stack.Item grow styles={stackItemStyles}>
          <nav>
            <DefaultButton iconProps={{ iconName: "ChevronLeft" }} onClick={goToPrevPage} />
            <DefaultButton iconProps={{ iconName: "ChevronRight" }} onClick={goToNextPage} />
          </nav>
        </Stack.Item>
      </Stack>
    </Stack>
  );
};

ReactDOM.render(<PDF />, document.getElementById("container"));
