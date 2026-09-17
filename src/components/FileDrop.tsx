import { useRef, useState } from 'react';

interface Props {
  accept: string;
  label: string;
  fileName?: string;
  /** Set while this particular file is being read, so the spinner sits where the file was dropped. */
  busy?: string | null;
  onFile: (file: File) => void;
}

/** A plain upload target: click it, or drop a file onto it. */
export function FileDrop({ accept, label, fileName, busy, onFile }: Props) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  const [size, setSize] = useState<number | null>(null);

  const choose = (f: File) => {
    setSize(f.size);
    onFile(f);
  };

  if (busy) {
    // Reading a 20 MB price list blocks the main thread for a few seconds. The spinner is a CSS
    // transform, which Chrome keeps running on the compositor while the main thread is busy, so
    // it stays moving rather than freezing mid-turn and looking like a hang.
    return (
      <div className="drop busy" aria-live="polite" aria-busy="true">
        <div className="file">
          <span className="spinner dark" />
          {busy}
        </div>
        <div className="sub">
          {fileName || 'Reading the file'}
          {size !== null && ` · ${(size / 1024 / 1024).toFixed(1)} MB`}
        </div>
      </div>
    );
  }

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
        if (f) choose(f);
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) choose(f);
          // Reset it so choosing the same file twice still fires onChange.
          e.target.value = '';
        }}
      />
      {fileName ? (
        <>
          <div className="file">
            <span className="tick" aria-hidden="true">
              ✓
            </span>
            {fileName}
          </div>
          <div className="sub">
            {size !== null && `${(size / 1024 / 1024).toFixed(1)} MB · `}
            Click to replace
          </div>
        </>
      ) : (
        <div className="sub">{label}</div>
      )}
    </div>
  );
}
