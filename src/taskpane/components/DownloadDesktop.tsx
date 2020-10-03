import * as React from "react";

interface DownloadProps {
  fileName?: string;
  loading: boolean;
  onDownloadFinish(): void;
}

const openAPI = (url: string) => Office.context.ui.openBrowserWindow(url);
const wrapUrl = (url: string) =>
  location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + url;

export default function DownloadDesktop({ fileName, loading, onDownloadFinish }: DownloadProps): JSX.Element {
  React.useEffect(() => {
    if (loading) {
      openAPI(wrapUrl(`/pdf.html?cell=${encodeURI(fileName)}`));
      onDownloadFinish();
    }
  }, [loading]);
  return <></>;
}
