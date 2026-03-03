import { useState, useRef, useCallback, useEffect } from 'react';

interface FileLoaderProps {
  onLoad: (buffer: ArrayBuffer, fileName: string) => void;
  isLoading: boolean;
  error: string | null;
  sourceUrl?: string;
}

function guessFileNameFromUrl(url: string) {
  try {
    const u = new URL(url);
    const fileParam = u.searchParams.get('file');
    if (fileParam) return fileParam;
  } catch {
    // ignore
  }
  return 'https://mckessoncorp.sharepoint.com/:x:/r/sites/GRPProductCommercialLeadershipTeam/_layouts/15/Doc.aspx?sourcedoc=%7B940AC279-49B3-4839-B3A4-52847757824D%7D&file=FY27%20PLT%20Priorities.xlsx&action=default&mobileredirect=true&DefaultItemOpen=1';
}

export const FileLoader: React.FC<FileLoaderProps> = ({ onLoad, isLoading, error, sourceUrl }) => {
  const [dragOver, setDragOver] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [urlLoadError, setUrlLoadError] = useState<string | null>(null);
  const [hasAutoLoaded, setHasAutoLoaded] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const processFile = useCallback(
    (file: File) => {
      if (!file.name.match(/\.xlsx?$/i)) return;

      setFileName(file.name);

      const reader = new FileReader();
      reader.onload = () => {
        if (reader.result instanceof ArrayBuffer) {
          onLoad(reader.result, file.name);
        }
      };
      reader.readAsArrayBuffer(file);
    },
    [onLoad],
  );

  const loadFromUrl = useCallback(
    async (url: string) => {
      setUrlLoadError(null);

      // NOTE: SharePoint often requires auth cookies and may still block CORS.
      // credentials: 'include' helps *if* the browser has a valid SP session cookie.
      const res = await fetch(url, {
        method: 'GET',
        credentials: 'include',
        redirect: 'follow',
      });

      if (!res.ok) {
        throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      }

      // Defensive check: sometimes SharePoint returns HTML (login page) instead of XLSX.
      const contentType = res.headers.get('content-type') || '';
      const isProbablyHtml = contentType.includes('text/html');
      if (isProbablyHtml) {
        throw new Error(
          'Got an HTML response instead of an Excel file (likely SharePoint login/CORS). Try downloading manually and uploading.',
        );
      }

      const buffer = await res.arrayBuffer();
      const name = guessFileNameFromUrl(url);

      setFileName(name);
      onLoad(buffer, name);
    },
    [onLoad],
  );

  // Auto-load once on mount if sourceUrl is provided
  useEffect(() => {
    if (!sourceUrl) return;
    if (hasAutoLoaded) return;

    setHasAutoLoaded(true);

    // Only attempt auto-load if we aren't already loading from another source
    if (!isLoading && !fileName) {
      loadFromUrl(sourceUrl).catch((e) => {
        console.error('Auto-load from URL failed:', e);
        setUrlLoadError(e instanceof Error ? e.message : 'Failed to load file from URL');
      });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sourceUrl, hasAutoLoaded]);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);

      const file = e.dataTransfer.files[0];
      if (file) processFile(file);
    },
    [processFile],
  );

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = () => {
    setDragOver(false);
  };

  const handleBrowse = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const handleLoadLatestClick = async () => {
    if (!sourceUrl) return;
    try {
      await loadFromUrl(sourceUrl);
    } catch (e) {
      console.error('Manual URL load failed:', e);
      setUrlLoadError(e instanceof Error ? e.message : 'Failed to load file from URL');
    }
  };

  return (
    <div className="file-loader">
      {/* If auto-load fails, explain + offer retry */}
      {sourceUrl && urlLoadError && (
        <div className="error-banner">
          <p>
            <strong>Couldn’t load the spreadsheet automatically.</strong> {urlLoadError}
          </p>
          <p className="hint-text">
            This is usually caused by SharePoint requiring sign-in or blocking cross-site downloads.
            You can still download it manually and drop it below.
          </p>
          <button
            type="button"
            className="view-toggle-btn"
            onClick={handleLoadLatestClick}
            disabled={isLoading}
          >
            Try “Load from SharePoint” again
          </button>
        </div>
      )}

      <div
        className={`file-drop-zone${dragOver ? ' file-drop-zone--active' : ''}${
          isLoading ? ' file-drop-zone--disabled' : ''
        }`}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onClick={handleBrowse}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileChange}
          className="file-input-hidden"
          disabled={isLoading}
        />
        {isLoading ? (
          <div className="file-drop-content">
            <span className="spinner spinner--large" />
            <span>Loading&hellip;</span>
          </div>
        ) : (
          <div className="file-drop-content">
            <span className="file-drop-icon">&#128196;</span>
            <span>
              {fileName ? (
                <>Loaded <strong>{fileName}</strong> &mdash; drop another to reload</>
              ) : (
                <>Drag &amp; drop an Excel file here, or <strong>click to browse</strong></>
              )}
            </span>
          </div>
        )}
      </div>

      {sourceUrl && (
        <p className="file-source-hint">
          Latest file:{' '}
          <a href={sourceUrl} target="_blank" rel="noopener noreferrer">
            SharePoint
          </a>
          {' '}
          <button
            type="button"
            className="view-toggle-btn"
            onClick={handleLoadLatestClick}
            disabled={isLoading}
            style={{ marginLeft: 8 }}
          >
            Load from SharePoint
          </button>
        </p>
      )}

      {error && (
        <div className="error-banner">
          <p>
            <strong>Error loading file:</strong> {error}
          </p>
          <p className="hint-text">
            Make sure the file is a valid .xlsx Excel file with a &ldquo;Product Initiatives&rdquo; sheet.
          </p>
        </div>
      )}
    </div>
  );
};