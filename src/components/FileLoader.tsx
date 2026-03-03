import { useState, useRef, useCallback } from 'react';

interface FileLoaderProps {
  onLoad: (buffer: ArrayBuffer, fileName: string) => void;
  isLoading: boolean;
  error: string | null;
  sourceUrl?: string;
}

export const FileLoader: React.FC<FileLoaderProps> = ({ onLoad, isLoading, error, sourceUrl }) => {
  const [dragOver, setDragOver] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const processFile = useCallback(
    (file: File) => {
      if (!file.name.match(/\.xlsx?$/i)) {
        return;
      }
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

  return (
    <div className="file-loader">
      <div
        className={`file-drop-zone${dragOver ? ' file-drop-zone--active' : ''}${isLoading ? ' file-drop-zone--disabled' : ''}`}
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
          Download the latest file from{' '}
          <a href={sourceUrl} target="_blank" rel="noopener noreferrer">
            SharePoint
          </a>
          , then drop it above.
        </p>
      )}

      {error && (
        <div className="error-banner">
          <p><strong>Error loading file:</strong> {error}</p>
          <p className="hint-text">
            Make sure the file is a valid .xlsx Excel file with a &ldquo;Product Initiatives&rdquo; sheet.
          </p>
        </div>
      )}
    </div>
  );
};
