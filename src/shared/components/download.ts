/** Offers a Blob as a file download in the browser. */
export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Give the browser time to start the download before the URL is released.
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

/** "Projekt sablon" + date → "Projekt-sablon-2026-09-24.json". */
export function templateFileName(name: string, date: Date, extension: string): string {
  const base = name.trim().replace(/[\\/:*?"<>|#%]+/g, '').replace(/\s+/g, '-') || 'copyjet';
  const d = date.toISOString().slice(0, 10);
  return `${base}-${d}.${extension}`;
}
