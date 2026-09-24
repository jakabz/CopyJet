/**
 * Base class for errors raised by the core. `code` is stable and is what the UI localizes;
 * `message` is an English developer-facing description.
 */
export class CopyJetError extends Error {
  public readonly code: string;
  public readonly detail?: unknown;

  constructor(code: string, message: string, detail?: unknown) {
    super(message);
    // Restore the prototype chain: the build targets ES5, where extending Error breaks instanceof.
    Object.setPrototypeOf(this, new.target.prototype);
    this.name = 'CopyJetError';
    this.code = code;
    this.detail = detail;
  }
}

/** Thrown when an operation is cancelled through its AbortSignal. */
export class AbortError extends CopyJetError {
  constructor() {
    super('ABORTED', 'The operation was aborted.');
    this.name = 'AbortError';
  }
}

export function throwIfAborted(signal?: AbortSignal): void {
  if (signal && signal.aborted) {
    throw new AbortError();
  }
}
