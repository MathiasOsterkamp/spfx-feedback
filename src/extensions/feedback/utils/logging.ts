import { ConsoleListener, Logger, LogLevel } from '@pnp/logging';

/** Configurable in the extension properties. */
export type LogLevelName = 'verbose' | 'info' | 'warning' | 'error' | 'off';

const LOG_LEVEL_NAMES: LogLevelName[] = ['verbose', 'info', 'warning', 'error', 'off'];
const DEFAULT_LOG_LEVEL: LogLevelName = 'warning';

const LEVELS: { [K in LogLevelName]: LogLevel } = {
  verbose: LogLevel.Verbose,
  info: LogLevel.Info,
  warning: LogLevel.Warning,
  error: LogLevel.Error,
  off: LogLevel.Off
};

const PREFIX = '[Feedback] ';
let subscribed = false;

/** Stored properties are untrusted: unknown values fall back to the default level. */
export function sanitizeLogLevel(value: unknown): LogLevelName {
  const index = LOG_LEVEL_NAMES.indexOf(value as LogLevelName);
  return index >= 0 ? LOG_LEVEL_NAMES[index] : DEFAULT_LOG_LEVEL;
}

/** The PnP logger is shared by all components on a page: the level set last applies. */
export function configureLogging(level: LogLevelName): void {
  if (!subscribed) {
    Logger.subscribe(new ConsoleListener());
    subscribed = true;
  }
  Logger.activeLogLevel = LEVELS[level];
}

function write(level: LogLevel, message: string, data?: unknown): void {
  Logger.log({ level, message: PREFIX + message, data });
}

export function logVerbose(message: string, data?: unknown): void {
  write(LogLevel.Verbose, message, data);
}

export function logInfo(message: string, data?: unknown): void {
  write(LogLevel.Info, message, data);
}

export function logWarning(message: string, data?: unknown): void {
  write(LogLevel.Warning, message, data);
}

export function logError(message: string, data?: unknown): void {
  write(LogLevel.Error, message, data);
}
