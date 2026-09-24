import { Logger, type ILogEntry } from '../../src/core/logger';

const fixedNow = (): Date => new Date('2026-09-24T10:00:00.000Z');

describe('Logger', () => {
  it('collects entries, counts levels and notifies subscribers until unsubscribed', () => {
    const log = new Logger(fixedNow);
    const seen: ILogEntry[] = [];
    const unsubscribe = log.subscribe((e) => seen.push(e));

    log.info('List created', { artifact: { kind: 'list', key: 'list:Projektek' }, step: 'list' });
    log.warn('User not found', { code: 'USER_NOT_FOUND' });
    unsubscribe();
    log.error('Failed');

    expect(seen.map((e) => e.message)).toEqual(['List created', 'User not found']);
    expect(log.entries).toHaveLength(3);
    expect(log.counts).toEqual({ info: 1, warn: 1, error: 1 });
    expect(log.entries[0].time).toBe('2026-09-24T10:00:00.000Z');
  });

  it('isolates a throwing listener', () => {
    const log = new Logger(fixedNow);
    const other = jest.fn();
    log.subscribe(() => {
      throw new Error('ui bug');
    });
    log.subscribe(other);
    expect(() => log.info('x')).not.toThrow();
    expect(other).toHaveBeenCalledTimes(1);
  });

  it('exports CSV with quoting and JSON', () => {
    const log = new Logger(fixedNow);
    log.error('Bad "value", retry', { artifact: { kind: 'items', key: 'items:Projektek' }, detail: { id: 7 } });
    const lines = log.toCsv().split('\r\n');
    expect(lines[0]).toBe('time,level,artifactKind,artifactKey,step,code,message,detail');
    expect(lines[1]).toBe('2026-09-24T10:00:00.000Z,error,items,items:Projektek,,,"Bad ""value"", retry","{""id"":7}"');
    expect(JSON.parse(log.toJson())).toHaveLength(1);
  });
});
