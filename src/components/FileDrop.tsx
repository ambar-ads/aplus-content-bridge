import { useRef, useState } from 'react';

interface Props {
  accept: string;
  label: string;
  fileName?: string;
  onFile: (file: File) => void;
}

/** A plain upload target: click it, or drop a file onto it. */
export function FileDrop({ accept, label, fileName, onFile }: Props) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);

  return (
    <div
      className={`drop${over ? ' over' : ''}`}
      role="button"
      tabIndex={0}
      onClick={() => inputRef.current?.click()}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          inputRef.current?.click();
        }
      }}
      onDragOver={(e) => {
        e.preventDefault();
        setOver(true);
      }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setOver(false);
        const f = e.dataTransfer.files?.[0];
        if (f) onFile(f);
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onFile(f);
          // Reset it so choosing the same file twice still fires onChange.
          e.target.value = '';
        }}
      />
      {fileName ? (
        <>
          <div className="file">{fileName}</div>
          <div className="sub">Click to replace</div>
        </>
      ) : (
        <div className="sub">{label}</div>
      )}
    </div>
  );
}
