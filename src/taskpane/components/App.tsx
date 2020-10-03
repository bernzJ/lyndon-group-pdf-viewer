import * as React from "react";
import { DefaultButton } from "office-ui-fabric-react";
import axios from "axios";
import Header from "./Header";
import Progress from "./Progress";
import Download from "./Download";
import DownloadDesktop from "./DownloadDesktop";

import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
import "../../../assets/logo-filled.png";

interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App = ({ title, isOfficeInitialized }: AppProps) => {
  const [fileName, setFileName] = React.useState(null);
  const [downloading, setDownloading] = React.useState(false);
  const source = axios.CancelToken.source();

  const renderDownload = () => {
    const isOfficeDesktop = () => Office && Office.context && Office.context.platform.toString() === "OfficeOnline";
    if (isOfficeDesktop()) {
      return (
        <Download
          loading={downloading}
          fileName={fileName}
          cts={source}
          onDownloadFinish={() => setDownloading(false)}
        />
      );
    }
    return <DownloadDesktop fileName={fileName} loading={downloading} onDownloadFinish={() => setDownloading(false)} />;
  };

  const onDownloadClick = () => {
    if (downloading) {
      source.cancel("Aborted by user.");
      setDownloading(false);
      return;
    }
    const excelTask = async () => {
      try {
        await Excel.run(async context => {
          /**
           * Insert your Excel code here
           */
          const range = context.workbook.getSelectedRange();
          range.load("values");
          await context.sync();
          setFileName(range.values[0][0]);
          setDownloading(true);
        });
      } catch (error) {
        console.error(error);
      }
    };
    excelTask();
  };

  if (!isOfficeInitialized) {
    return (
      <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
    );
  }
  return (
    <div className="ms-welcome">
      <Header logo="assets/logo-filled.png" title={title} message="PDF Viewer" />
      <main className="ms-welcome__main">
        <p className="ms-font-l">
          Select cell and press <b>Download</b>
        </p>
        <DefaultButton
          className="ms-welcome__action"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={onDownloadClick}
        >
          Download
        </DefaultButton>
        {renderDownload()}
      </main>
    </div>
  );
};

export default App;
