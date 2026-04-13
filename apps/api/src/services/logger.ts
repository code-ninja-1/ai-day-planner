import fs from "node:fs";
import path from "node:path";
import { randomUUID } from "node:crypto";
import { env } from "../env.js";
import { recordAuditEvent } from "../db.js";
import type { AuditEventStatus, AuditLogLevel } from "../types.js";

export interface LogEventInput {
  level?: AuditLogLevel;
  eventType: string;
  requestId?: string | null;
  runId?: string | null;
  entityType?: string | null;
  entityId?: string | null;
  provider?: string | null;
  status?: AuditEventStatus;
  source?: string | null;
  message: string;
  metadata?: unknown;
}

const LOG_LEVEL_ORDER: Record<AuditLogLevel, number> = {
  debug: 10,
  info: 20,
  warn: 30,
  error: 40
};

fs.mkdirSync(env.logDirectory, { recursive: true });

function shouldWrite(level: AuditLogLevel) {
  const configured = (env.logLevel as AuditLogLevel) in LOG_LEVEL_ORDER ? (env.logLevel as AuditLogLevel) : "info";
  return LOG_LEVEL_ORDER[level] >= LOG_LEVEL_ORDER[configured];
}

function redactValue(value: unknown): unknown {
  if (env.enableSensitiveDebugLogs) {
    return value;
  }

  if (typeof value === "string") {
    if (value.length > 240) {
      return `${value.slice(0, 240)}…[redacted]`;
    }
    return value;
  }

  if (Array.isArray(value)) {
    return value.map((item) => redactValue(item));
  }

  if (value && typeof value === "object") {
    const entries = Object.entries(value as Record<string, unknown>).map(([key, entry]) => {
      if (/(body|content|token|secret|authorization|accessToken|refreshToken|apiToken|description)/i.test(key)) {
        return [key, "[redacted]"];
      }
      return [key, redactValue(entry)];
    });
    return Object.fromEntries(entries);
  }

  return value;
}

function writeToFile(serialized: string) {
  const dayKey = new Date().toISOString().slice(0, 10);
  const filePath = path.join(env.logDirectory, `${dayKey}.ndjson`);
  fs.appendFileSync(filePath, `${serialized}\n`);
}

export function createCorrelationId() {
  return randomUUID();
}

export function logEvent(input: LogEventInput) {
  const timestamp = new Date().toISOString();
  const level = input.level ?? "info";
  if (!shouldWrite(level)) {
    return;
  }

  const payload = {
    timestamp,
    level,
    eventType: input.eventType,
    requestId: input.requestId ?? null,
    runId: input.runId ?? null,
    entityType: input.entityType ?? null,
    entityId: input.entityId ?? null,
    provider: input.provider ?? null,
    status: input.status ?? "info",
    source: input.source ?? "server",
    message: input.message,
    metadata: redactValue(input.metadata ?? null)
  };

  const serialized = JSON.stringify(payload);

  if (env.logLevel === "debug" || level !== "debug") {
    const writer = level === "error" ? console.error : level === "warn" ? console.warn : console.log;
    writer(serialized);
  }

  writeToFile(serialized);
  recordAuditEvent({
    timestamp,
    level,
    eventType: input.eventType,
    requestId: input.requestId ?? null,
    runId: input.runId ?? null,
    entityType: input.entityType ?? null,
    entityId: input.entityId ?? null,
    provider: input.provider ?? null,
    status: input.status ?? "info",
    source: input.source ?? "server",
    message: input.message,
    metadataJson: payload.metadata ? JSON.stringify(payload.metadata) : null
  });
}
