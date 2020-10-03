import * as React from "react";
import axios, { CancelTokenSource } from "axios";
import { Dropbox } from "dropbox";
import { ProgressIndicator } from "office-ui-fabric-react";

interface DownloadProps {
  fileName?: string;
  loading: boolean;
  cts: CancelTokenSource;
  onDownloadFinish(): void;
}

const dbx = new Dropbox({
  accessToken: process.env.TOKEN
});

export default function Download({ fileName, loading, onDownloadFinish, cts }: DownloadProps): JSX.Element {
  const [percentage, setPercentage] = React.useState(0);
  const [message, setMessage] = React.useState(<></>);

  React.useEffect(() => {
    let didCancel = false;
    const download = async () => {
      try {
        const meta: any = await dbx.filesDownload({
          path: `/Ideal Image/${fileName}.pdf`
        });
        if (meta.status !== 200) {
          throw new Error(meta.result);
        }
        const blob = URL.createObjectURL(meta.result.fileBlob);
        const result = await axios({
          url: blob,
          responseType: "blob",
          onDownloadProgress(progressEvent) {
            setPercentage(Math.round((progressEvent.loaded / progressEvent.total) * 100));
            if (didCancel) {
              cts.cancel();
            }
          },
          cancelToken: cts.token
        });
        URL.revokeObjectURL(blob);
        //@TODO: check if this leaks.
        setMessage(
          <a
            href={URL.createObjectURL(new Blob([result.data], { type: "application/pdf" }))}
            rel="noopener noreferrer"
            target="_blank"
          >
            {meta.result.name}
          </a>
        );
        URL.revokeObjectURL(blob);
      } catch (error) {
        switch (true) {
          case error.error !== undefined: {
            const msg = JSON.parse(error.error).error_summary;
            setMessage(<span>{msg}</span>);
            break;
          }
          default: {
            setMessage(<span>{error.message}</span>);
          }
        }
      } finally {
        onDownloadFinish();
      }
    };

    if (loading) {
      setMessage(<></>);
      setPercentage(0);
      download();
    }
    return () => {
      didCancel = true;
    };
  }, [fileName, loading, cts]);

  return (
    <ProgressIndicator
      className="ms__download"
      label={`${percentage} %`}
      description={message}
      percentComplete={percentage}
    />
  );
}
