/** Typed collaboration errors — safe to surface without leaking document content */

export class CollaborationError extends Error {
  readonly code: string;
  constructor(code: string, message: string) {
    super(message);
    this.name = 'CollaborationError';
    this.code = code;
  }
}

export class AuthenticationError extends CollaborationError {
  constructor(message = 'Authentication failed') {
    super('AUTHENTICATION', message);
    this.name = 'AuthenticationError';
  }
}

export class AuthorizationError extends CollaborationError {
  constructor(message = 'Not authorized for this document') {
    super('AUTHORIZATION', message);
    this.name = 'AuthorizationError';
  }
}

export class ConfigurationError extends CollaborationError {
  constructor(message: string) {
    super('CONFIGURATION', message);
    this.name = 'ConfigurationError';
  }
}

export class ConnectionError extends CollaborationError {
  constructor(message: string) {
    super('CONNECTION', message);
    this.name = 'ConnectionError';
  }
}

export class StorageError extends CollaborationError {
  constructor(message: string) {
    super('STORAGE', message);
    this.name = 'StorageError';
  }
}

export class SyncError extends CollaborationError {
  constructor(message: string) {
    super('SYNC', message);
    this.name = 'SyncError';
  }
}

export class ExportError extends CollaborationError {
  constructor(message: string) {
    super('EXPORT', message);
    this.name = 'ExportError';
  }
}
