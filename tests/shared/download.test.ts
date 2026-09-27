import { fileTimestamp, templateFileName } from '../../src/shared/components/download';

describe('templateFileName', () => {
  const at = new Date(2026, 8, 5, 9, 3, 7); // local time

  it('adds the local date and time', () => {
    expect(fileTimestamp(at)).toBe('2026-09-05_09-03-07');
    expect(templateFileName('Projekt sablon', at, 'zip')).toBe('Projekt-sablon-2026-09-05_09-03-07.zip');
  });

  it('drops characters file systems forbid and falls back to copyjet', () => {
    expect(templateFileName('A/B: "C"?', at, 'json')).toBe('AB-C-2026-09-05_09-03-07.json');
    expect(templateFileName('  ', at, 'json')).toBe('copyjet-2026-09-05_09-03-07.json');
  });
});
