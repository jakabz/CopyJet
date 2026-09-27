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

/** Local date and time for file names, without characters Windows forbids: "2026-09-25_09-31-05". */
export function fileTimestamp(date: Date): string {
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `${date.getFullYear()}-${p(date.getMonth() + 1)}-${p(date.getDate())}_${p(date.getHours())}-${p(date.getMinutes())}-${p(date.getSeconds())}`;
}

/** "Projekt sablon" + local time → "Projekt-sablon-2026-09-25_09-31-05.json". */
export function templateFileName(name: string, date: Date, extension: string): string {
  const base = name.trim().replace(/[\\/:*?"<>|#%]+/g, '').replace(/\s+/g, '-') || 'copyjet';
  return `${base}-${fileTimestamp(date)}.${extension}`;
}
