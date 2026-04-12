import fs from "node:fs";
import path from "node:path";
import Database from "better-sqlite3";
import { env } from "./env.js";
import type {
  AutomationSettings,
  IntegrationConnection,
  Meeting,
  Reminder,
  ReminderKind,
  ReminderStatus,
  Task,
  TaskEffortBucket,
  TaskPriority,
  TaskSource,
  TaskStatus
} from "./types.js";

const databaseDir = path.dirname(env.databasePath);
fs.mkdirSync(databaseDir, { recursive: true });

export const db = new Database(env.databasePath);

db.pragma("journal_mode = WAL");

db.exec(`
  CREATE TABLE IF NOT EXISTS tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    source TEXT NOT NULL,
    priority TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Not Started',
    source_link TEXT,
    source_ref TEXT,
    source_thread_ref TEXT,
    jira_status TEXT,
    ignored INTEGER NOT NULL DEFAULT 0,
    deferred_until TEXT,
    reminder_state TEXT,
    last_reminded_at TEXT,
    estimated_effort_bucket TEXT,
    priority_score REAL,
    priority_explanation TEXT,
    task_age_days INTEGER NOT NULL DEFAULT 0,
    carry_forward_count INTEGER NOT NULL DEFAULT 0,
    completed_at TEXT,
    last_activity_at TEXT,
    manual_override_flags TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
  );

  CREATE UNIQUE INDEX IF NOT EXISTS idx_tasks_source_ref
  ON tasks(source, source_ref)
  WHERE source_ref IS NOT NULL;

  CREATE TABLE IF NOT EXISTS meetings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    external_id TEXT,
    title TEXT NOT NULL,
    start_time TEXT NOT NULL,
    end_time TEXT NOT NULL,
    time_zone TEXT,
    duration_minutes INTEGER NOT NULL,
    meeting_link TEXT,
    meeting_link_type TEXT,
    is_cancelled INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS integration_connections (
    provider TEXT PRIMARY KEY,
    status TEXT NOT NULL,
    account_label TEXT,
    config_json TEXT,
    access_token TEXT,
    refresh_token TEXT,
    expires_at TEXT,
    error_message TEXT,
    updated_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS sync_state (
    provider TEXT PRIMARY KEY,
    last_sync_at TEXT
  );

  CREATE TABLE IF NOT EXISTS reminders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    reminder_key TEXT NOT NULL UNIQUE,
    task_id INTEGER,
    kind TEXT NOT NULL,
    title TEXT NOT NULL,
    reason TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'active',
    source_link TEXT,
    source_label TEXT,
    scheduled_for TEXT,
    dismissed_at TEXT,
    throttle_until TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS automation_settings (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    schedule_enabled INTEGER NOT NULL DEFAULT 0,
    schedule_time_local TEXT NOT NULL DEFAULT '08:30',
    schedule_timezone TEXT NOT NULL DEFAULT 'UTC',
    reminders_enabled INTEGER NOT NULL DEFAULT 1,
    reminder_cadence_hours INTEGER NOT NULL DEFAULT 6,
    desktop_notifications_enabled INTEGER NOT NULL DEFAULT 0,
    last_auto_generated_at TEXT,
    scheduler_last_run_at TEXT,
    scheduler_last_status TEXT NOT NULL DEFAULT 'idle',
    scheduler_last_error TEXT
  );

  CREATE TABLE IF NOT EXISTS generation_runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    trigger_type TEXT NOT NULL,
    generated_at TEXT NOT NULL,
    warnings_json TEXT
  );
`);

function ensureColumn(table: string, column: string, definition: string) {
  const columns = db.prepare(`PRAGMA table_info(${table})`).all() as Array<{ name: string }>;
  if (!columns.some((item) => item.name === column)) {
    db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`);
  }
}

ensureColumn("meetings", "time_zone", "TEXT");
ensureColumn("meetings", "meeting_link_type", "TEXT");
ensureColumn("meetings", "is_cancelled", "INTEGER NOT NULL DEFAULT 0");

ensureColumn("tasks", "deferred_until", "TEXT");
ensureColumn("tasks", "reminder_state", "TEXT");
ensureColumn("tasks", "last_reminded_at", "TEXT");
ensureColumn("tasks", "estimated_effort_bucket", "TEXT");
ensureColumn("tasks", "priority_score", "REAL");
ensureColumn("tasks", "priority_explanation", "TEXT");
ensureColumn("tasks", "task_age_days", "INTEGER NOT NULL DEFAULT 0");
ensureColumn("tasks", "carry_forward_count", "INTEGER NOT NULL DEFAULT 0");
ensureColumn("tasks", "completed_at", "TEXT");
ensureColumn("tasks", "last_activity_at", "TEXT");
ensureColumn("tasks", "manual_override_flags", "TEXT NOT NULL DEFAULT '[]'");

db.prepare(
  `
  INSERT INTO automation_settings (id)
  VALUES (1)
  ON CONFLICT(id) DO NOTHING
  `
).run();

function parseJsonArray(value: unknown) {
  if (typeof value !== "string" || !value.trim()) return [];
  try {
    const parsed = JSON.parse(value) as unknown;
    return Array.isArray(parsed) ? parsed.filter((entry): entry is string => typeof entry === "string") : [];
  } catch {
    return [];
  }
}

const taskRowToTask = (row: Record<string, unknown>): Task => ({
  id: Number(row.id),
  title: String(row.title),
  source: row.source as TaskSource,
  priority: row.priority as TaskPriority,
  status: row.status as TaskStatus,
  sourceLink: (row.source_link as string | null) ?? null,
  sourceRef: (row.source_ref as string | null) ?? null,
  sourceThreadRef: (row.source_thread_ref as string | null) ?? null,
  jiraStatus: (row.jira_status as string | null) ?? null,
  ignored: Number(row.ignored),
  deferredUntil: (row.deferred_until as string | null) ?? null,
  reminderState: (row.reminder_state as ReminderStatus | null) ?? null,
  lastRemindedAt: (row.last_reminded_at as string | null) ?? null,
  estimatedEffortBucket: (row.estimated_effort_bucket as TaskEffortBucket | null) ?? null,
  priorityScore: row.priority_score === null || row.priority_score === undefined ? null : Number(row.priority_score),
  priorityExplanation: (row.priority_explanation as string | null) ?? null,
  taskAgeDays: Number(row.task_age_days ?? 0),
  carryForwardCount: Number(row.carry_forward_count ?? 0),
  completedAt: (row.completed_at as string | null) ?? null,
  lastActivityAt: (row.last_activity_at as string | null) ?? null,
  manualOverrideFlags: parseJsonArray(row.manual_override_flags),
  createdAt: String(row.created_at),
  updatedAt: String(row.updated_at)
});

const meetingRowToMeeting = (row: Record<string, unknown>): Meeting => ({
  id: Number(row.id),
  externalId: (row.external_id as string | null) ?? null,
  title: String(row.title),
  startTime: String(row.start_time),
  endTime: String(row.end_time),
  timeZone: (row.time_zone as string | null) ?? null,
  durationMinutes: Number(row.duration_minutes),
  meetingLink: (row.meeting_link as string | null) ?? null,
  meetingLinkType: (row.meeting_link_type as "join" | "calendar" | null) ?? null,
  isCancelled: Number(row.is_cancelled) === 1,
  createdAt: String(row.created_at)
});

const reminderRowToReminder = (row: Record<string, unknown>): Reminder => ({
  id: Number(row.id),
  reminderKey: String(row.reminder_key),
  taskId: row.task_id === null || row.task_id === undefined ? null : Number(row.task_id),
  kind: row.kind as ReminderKind,
  title: String(row.title),
  reason: String(row.reason),
  status: row.status as ReminderStatus,
  sourceLink: (row.source_link as string | null) ?? null,
  sourceLabel: (row.source_label as string | null) ?? null,
  scheduledFor: (row.scheduled_for as string | null) ?? null,
  createdAt: String(row.created_at),
  updatedAt: String(row.updated_at),
  dismissedAt: (row.dismissed_at as string | null) ?? null,
  throttleUntil: (row.throttle_until as string | null) ?? null
});

const automationRowToSettings = (row: Record<string, unknown>): AutomationSettings => ({
  scheduleEnabled: Number(row.schedule_enabled) === 1,
  scheduleTimeLocal: String(row.schedule_time_local),
  scheduleTimezone: String(row.schedule_timezone),
  remindersEnabled: Number(row.reminders_enabled) === 1,
  reminderCadenceHours: Number(row.reminder_cadence_hours),
  desktopNotificationsEnabled: Number(row.desktop_notifications_enabled) === 1,
  lastAutoGeneratedAt: (row.last_auto_generated_at as string | null) ?? null,
  schedulerLastRunAt: (row.scheduler_last_run_at as string | null) ?? null,
  schedulerLastStatus: row.scheduler_last_status as AutomationSettings["schedulerLastStatus"],
  schedulerLastError: (row.scheduler_last_error as string | null) ?? null
});

function orderByPrioritySql() {
  return "CASE priority WHEN 'High' THEN 1 WHEN 'Medium' THEN 2 ELSE 3 END, priority_score DESC, updated_at DESC";
}

export function listTasks(
  status?: TaskStatus,
  options?: {
    includeDeferred?: boolean;
    onlyDeferred?: boolean;
  }
) {
  const filters = ["ignored = 0"];
  const values: Array<string> = [];

  if (status) {
    filters.push("status = ?");
    values.push(status);
  }

  if (options?.onlyDeferred) {
    filters.push("deferred_until IS NOT NULL");
  } else if (!options?.includeDeferred) {
    filters.push("(deferred_until IS NULL OR deferred_until <= datetime('now'))");
  }

  const sql = `SELECT * FROM tasks WHERE ${filters.join(" AND ")} ORDER BY ${orderByPrioritySql()}`;
  return db.prepare(sql).all(...values).map((row: unknown) => taskRowToTask(row as Record<string, unknown>));
}

export function listDeferredTasks() {
  return listTasks(undefined, { onlyDeferred: true });
}

export function listReminderItems(status: ReminderStatus[] = ["active"]) {
  const placeholders = status.map(() => "?").join(", ");
  return db
    .prepare(`SELECT * FROM reminders WHERE status IN (${placeholders}) ORDER BY created_at DESC`)
    .all(...status)
    .map((row: unknown) => reminderRowToReminder(row as Record<string, unknown>));
}

export function getReminderById(id: number) {
  const row = db.prepare("SELECT * FROM reminders WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  return row ? reminderRowToReminder(row) : null;
}

export function upsertReminder(input: {
  reminderKey: string;
  taskId?: number | null;
  kind: ReminderKind;
  title: string;
  reason: string;
  status?: ReminderStatus;
  sourceLink?: string | null;
  sourceLabel?: string | null;
  scheduledFor?: string | null;
  dismissedAt?: string | null;
  throttleUntil?: string | null;
}) {
  const now = new Date().toISOString();
  db.prepare(
    `
    INSERT INTO reminders (
      reminder_key, task_id, kind, title, reason, status, source_link, source_label, scheduled_for,
      dismissed_at, throttle_until, created_at, updated_at
    ) VALUES (
      @reminderKey, @taskId, @kind, @title, @reason, @status, @sourceLink, @sourceLabel, @scheduledFor,
      @dismissedAt, @throttleUntil, @createdAt, @updatedAt
    )
    ON CONFLICT(reminder_key) DO UPDATE SET
      task_id = excluded.task_id,
      kind = excluded.kind,
      title = excluded.title,
      reason = excluded.reason,
      status = excluded.status,
      source_link = excluded.source_link,
      source_label = excluded.source_label,
      scheduled_for = excluded.scheduled_for,
      dismissed_at = excluded.dismissed_at,
      throttle_until = excluded.throttle_until,
      updated_at = excluded.updated_at
    `
  ).run({
    reminderKey: input.reminderKey,
    taskId: input.taskId ?? null,
    kind: input.kind,
    title: input.title,
    reason: input.reason,
    status: input.status ?? "active",
    sourceLink: input.sourceLink ?? null,
    sourceLabel: input.sourceLabel ?? null,
    scheduledFor: input.scheduledFor ?? null,
    dismissedAt: input.dismissedAt ?? null,
    throttleUntil: input.throttleUntil ?? null,
    createdAt: now,
    updatedAt: now
  });
}

export function updateReminder(
  id: number,
  patch: Partial<Pick<Reminder, "status" | "reason" | "scheduledFor" | "throttleUntil" | "dismissedAt">>
) {
  const existing = db.prepare("SELECT * FROM reminders WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  if (!existing) return null;
  db.prepare(
    `
    UPDATE reminders
    SET status = @status,
        reason = @reason,
        scheduled_for = @scheduledFor,
        throttle_until = @throttleUntil,
        dismissed_at = @dismissedAt,
        updated_at = @updatedAt
    WHERE id = @id
    `
  ).run({
    id,
    status: patch.status ?? existing.status,
    reason: patch.reason ?? existing.reason,
    scheduledFor: patch.scheduledFor ?? existing.scheduled_for,
    throttleUntil: patch.throttleUntil ?? existing.throttle_until,
    dismissedAt: patch.dismissedAt ?? existing.dismissed_at,
    updatedAt: new Date().toISOString()
  });
  return getReminderById(id);
}

export function resolveStaleReminders(activeKeys: string[]) {
  const rows = db
    .prepare("SELECT reminder_key FROM reminders WHERE status = 'active'")
    .all() as Array<{ reminder_key: string }>;
  const keys = new Set(activeKeys);
  const now = new Date().toISOString();
  for (const row of rows) {
    if (keys.has(row.reminder_key)) continue;
    db.prepare(
      `
      UPDATE reminders
      SET status = 'resolved',
          updated_at = ?
      WHERE reminder_key = ?
      `
    ).run(now, row.reminder_key);
  }
}

export function getTaskById(id: number) {
  const row = db.prepare("SELECT * FROM tasks WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  return row ? taskRowToTask(row) : null;
}

export function listMeetings() {
  return db
    .prepare("SELECT * FROM meetings ORDER BY start_time ASC")
    .all()
    .map((row: unknown) => meetingRowToMeeting(row as Record<string, unknown>));
}

function mergeOverrideFlags(existing: string[], next: string[]) {
  return [...new Set([...existing, ...next])];
}

export function upsertTask(input: {
  title: string;
  source: TaskSource;
  priority: TaskPriority;
  status?: TaskStatus;
  sourceLink?: string | null;
  sourceRef?: string | null;
  sourceThreadRef?: string | null;
  jiraStatus?: string | null;
}) {
  const now = new Date().toISOString();
  const existing = input.sourceRef
    ? (db
        .prepare("SELECT * FROM tasks WHERE source = ? AND source_ref = ?")
        .get(input.source, input.sourceRef) as Record<string, unknown> | undefined)
    : undefined;

  if (existing) {
    const overrideFlags = parseJsonArray(existing.manual_override_flags);
    db.prepare(
      `
      UPDATE tasks
      SET title = @title,
          priority = @priority,
          status = @status,
          source_link = @sourceLink,
          source_thread_ref = @sourceThreadRef,
          jira_status = @jiraStatus,
          updated_at = @updatedAt
      WHERE source = @source AND source_ref = @sourceRef
      `
    ).run({
      title: input.title,
      priority: overrideFlags.includes("priority") ? existing.priority : input.priority,
      status: overrideFlags.includes("status") ? existing.status : (input.status ?? existing.status),
      sourceLink: input.sourceLink ?? null,
      sourceThreadRef: input.sourceThreadRef ?? null,
      jiraStatus: input.jiraStatus ?? null,
      updatedAt: now,
      source: input.source,
      sourceRef: input.sourceRef
    });
  } else {
    db.prepare(
      `
      INSERT INTO tasks (
        title, source, priority, status, source_link, source_ref, source_thread_ref, jira_status,
        ignored, deferred_until, reminder_state, last_reminded_at, estimated_effort_bucket,
        priority_score, priority_explanation, task_age_days, carry_forward_count, completed_at,
        last_activity_at, manual_override_flags, created_at, updated_at
      ) VALUES (
        @title, @source, @priority, @status, @sourceLink, @sourceRef, @sourceThreadRef, @jiraStatus,
        0, NULL, NULL, NULL, NULL, NULL, NULL, 0, 0, NULL, @createdAt, '[]', @createdAt, @updatedAt
      )
      `
    ).run({
      title: input.title,
      source: input.source,
      priority: input.priority,
      status: input.status ?? "Not Started",
      sourceLink: input.sourceLink ?? null,
      sourceRef: input.sourceRef ?? null,
      sourceThreadRef: input.sourceThreadRef ?? null,
      jiraStatus: input.jiraStatus ?? null,
      createdAt: now,
      updatedAt: now
    });
  }
}

export function createManualTask(input: {
  title: string;
  priority?: TaskPriority;
  status?: TaskStatus;
}) {
  const now = new Date().toISOString();
  const status = input.status ?? "Not Started";
  const result = db
    .prepare(
      `
      INSERT INTO tasks (
        title, source, priority, status, ignored, deferred_until, reminder_state, last_reminded_at,
        estimated_effort_bucket, priority_score, priority_explanation, task_age_days, carry_forward_count,
        completed_at, last_activity_at, manual_override_flags, created_at, updated_at
      ) VALUES (
        @title, 'Manual', @priority, @status, 0, NULL, NULL, NULL,
        NULL, NULL, NULL, 0, 0, @completedAt, @lastActivityAt, '["priority","status"]', @createdAt, @updatedAt
      )
      `
    )
    .run({
      title: input.title,
      priority: input.priority ?? "Medium",
      status,
      completedAt: status === "Completed" ? now : null,
      lastActivityAt: now,
      createdAt: now,
      updatedAt: now
    });

  return db
    .prepare("SELECT * FROM tasks WHERE id = ?")
    .get(result.lastInsertRowid) as Record<string, unknown>;
}

export function updateTask(
  id: number,
  patch: Partial<Pick<Task, "title" | "priority" | "status" | "deferredUntil">> & {
    manualOverrideFlags?: string[];
    priorityScore?: number | null;
    priorityExplanation?: string | null;
    estimatedEffortBucket?: TaskEffortBucket | null;
    taskAgeDays?: number;
    carryForwardCount?: number;
    reminderState?: ReminderStatus | null;
    lastRemindedAt?: string | null;
    lastActivityAt?: string | null;
  }
) {
  const existing = db.prepare("SELECT * FROM tasks WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  if (!existing) return null;

  const nextStatus = patch.status ?? (existing.status as TaskStatus);
  const overrideFlags = patch.manualOverrideFlags
    ? mergeOverrideFlags(parseJsonArray(existing.manual_override_flags), patch.manualOverrideFlags)
    : parseJsonArray(existing.manual_override_flags);
  const updatedAt = new Date().toISOString();

  db.prepare(
    `
    UPDATE tasks
    SET title = @title,
        priority = @priority,
        status = @status,
        deferred_until = @deferredUntil,
        estimated_effort_bucket = @estimatedEffortBucket,
        priority_score = @priorityScore,
        priority_explanation = @priorityExplanation,
        task_age_days = @taskAgeDays,
        carry_forward_count = @carryForwardCount,
        reminder_state = @reminderState,
        last_reminded_at = @lastRemindedAt,
        completed_at = @completedAt,
        last_activity_at = @lastActivityAt,
        manual_override_flags = @manualOverrideFlags,
        updated_at = @updatedAt
    WHERE id = @id
    `
  ).run({
    id,
    title: patch.title ?? existing.title,
    priority: patch.priority ?? existing.priority,
    status: nextStatus,
    deferredUntil: patch.deferredUntil ?? existing.deferred_until,
    estimatedEffortBucket: patch.estimatedEffortBucket ?? existing.estimated_effort_bucket ?? null,
    priorityScore: patch.priorityScore ?? existing.priority_score ?? null,
    priorityExplanation: patch.priorityExplanation ?? existing.priority_explanation ?? null,
    taskAgeDays: patch.taskAgeDays ?? existing.task_age_days ?? 0,
    carryForwardCount: patch.carryForwardCount ?? existing.carry_forward_count ?? 0,
    reminderState: patch.reminderState ?? existing.reminder_state ?? null,
    lastRemindedAt: patch.lastRemindedAt ?? existing.last_reminded_at ?? null,
    completedAt: nextStatus === "Completed" ? (existing.completed_at ?? updatedAt) : null,
    lastActivityAt: patch.lastActivityAt ?? updatedAt,
    manualOverrideFlags: JSON.stringify(overrideFlags),
    updatedAt
  });

  return db.prepare("SELECT * FROM tasks WHERE id = ?").get(id);
}

export function deleteTask(id: number) {
  const existing = db.prepare("SELECT * FROM tasks WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  if (!existing) return false;

  if (existing.source !== "Manual") {
    db.prepare(
      `
      UPDATE tasks
      SET ignored = 1,
          updated_at = ?
      WHERE id = ?
      `
    ).run(new Date().toISOString(), id);
  } else {
    db.prepare("DELETE FROM tasks WHERE id = ?").run(id);
  }

  return true;
}

export function replaceMeetings(meetings: Omit<Meeting, "id" | "createdAt">[]) {
  const now = new Date().toISOString();
  const transaction = db.transaction(() => {
    db.prepare("DELETE FROM meetings").run();
    const insert = db.prepare(
      `
      INSERT INTO meetings (
        external_id, title, start_time, end_time, time_zone, duration_minutes, meeting_link,
        meeting_link_type, is_cancelled, created_at
      ) VALUES (
        @externalId, @title, @startTime, @endTime, @timeZone, @durationMinutes, @meetingLink,
        @meetingLinkType, @isCancelled, @createdAt
      )
      `
    );
    for (const meeting of meetings) {
      insert.run({
        ...meeting,
        isCancelled: meeting.isCancelled ? 1 : 0,
        createdAt: now
      });
    }
  });
  transaction();
}

export function saveIntegrationConnection(input: Omit<IntegrationConnection, "updatedAt">) {
  db.prepare(
    `
    INSERT INTO integration_connections (
      provider, status, account_label, config_json, access_token, refresh_token, expires_at, error_message, updated_at
    ) VALUES (
      @provider, @status, @accountLabel, @configJson, @accessToken, @refreshToken, @expiresAt, @errorMessage, @updatedAt
    )
    ON CONFLICT(provider) DO UPDATE SET
      status = excluded.status,
      account_label = excluded.account_label,
      config_json = excluded.config_json,
      access_token = excluded.access_token,
      refresh_token = excluded.refresh_token,
      expires_at = excluded.expires_at,
      error_message = excluded.error_message,
      updated_at = excluded.updated_at
    `
  ).run({
    ...input,
    updatedAt: new Date().toISOString()
  });
}

export function deleteIntegrationConnection(provider: "microsoft" | "jira") {
  db.prepare("DELETE FROM integration_connections WHERE provider = ?").run(provider);
}

export function getIntegrationConnection(provider: "microsoft" | "jira") {
  const row = db
    .prepare("SELECT * FROM integration_connections WHERE provider = ?")
    .get(provider) as Record<string, unknown> | undefined;
  if (!row) return null;
  return {
    provider: row.provider,
    status: row.status,
    accountLabel: row.account_label,
    configJson: row.config_json,
    accessToken: row.access_token,
    refreshToken: row.refresh_token,
    expiresAt: row.expires_at,
    errorMessage: row.error_message,
    updatedAt: row.updated_at
  } as IntegrationConnection;
}

export function listIntegrationConnections() {
  return db.prepare("SELECT * FROM integration_connections ORDER BY provider ASC").all() as Record<string, unknown>[];
}

export function setSyncState(provider: string, timestamp: string) {
  db.prepare(
    `
    INSERT INTO sync_state (provider, last_sync_at)
    VALUES (?, ?)
    ON CONFLICT(provider) DO UPDATE SET last_sync_at = excluded.last_sync_at
    `
  ).run(provider, timestamp);
}

export function getSyncState(provider: string) {
  const row = db
    .prepare("SELECT last_sync_at FROM sync_state WHERE provider = ?")
    .get(provider) as { last_sync_at: string } | undefined;
  return row?.last_sync_at ?? null;
}

export function getAutomationSettings() {
  const row = db.prepare("SELECT * FROM automation_settings WHERE id = 1").get() as Record<string, unknown>;
  return automationRowToSettings(row);
}

export function saveAutomationSettings(
  patch: Partial<
    Pick<
      AutomationSettings,
      | "scheduleEnabled"
      | "scheduleTimeLocal"
      | "scheduleTimezone"
      | "remindersEnabled"
      | "reminderCadenceHours"
      | "desktopNotificationsEnabled"
      | "lastAutoGeneratedAt"
      | "schedulerLastRunAt"
      | "schedulerLastStatus"
      | "schedulerLastError"
    >
  >
) {
  const existing = getAutomationSettings();
  db.prepare(
    `
    UPDATE automation_settings
    SET schedule_enabled = @scheduleEnabled,
        schedule_time_local = @scheduleTimeLocal,
        schedule_timezone = @scheduleTimezone,
        reminders_enabled = @remindersEnabled,
        reminder_cadence_hours = @reminderCadenceHours,
        desktop_notifications_enabled = @desktopNotificationsEnabled,
        last_auto_generated_at = @lastAutoGeneratedAt,
        scheduler_last_run_at = @schedulerLastRunAt,
        scheduler_last_status = @schedulerLastStatus,
        scheduler_last_error = @schedulerLastError
    WHERE id = 1
    `
  ).run({
    scheduleEnabled: Number(patch.scheduleEnabled ?? existing.scheduleEnabled),
    scheduleTimeLocal: patch.scheduleTimeLocal ?? existing.scheduleTimeLocal,
    scheduleTimezone: patch.scheduleTimezone ?? existing.scheduleTimezone,
    remindersEnabled: Number(patch.remindersEnabled ?? existing.remindersEnabled),
    reminderCadenceHours: patch.reminderCadenceHours ?? existing.reminderCadenceHours,
    desktopNotificationsEnabled: Number(
      patch.desktopNotificationsEnabled ?? existing.desktopNotificationsEnabled
    ),
    lastAutoGeneratedAt: patch.lastAutoGeneratedAt ?? existing.lastAutoGeneratedAt,
    schedulerLastRunAt: patch.schedulerLastRunAt ?? existing.schedulerLastRunAt,
    schedulerLastStatus: patch.schedulerLastStatus ?? existing.schedulerLastStatus,
    schedulerLastError: patch.schedulerLastError ?? existing.schedulerLastError
  });
  return getAutomationSettings();
}

export function recordGenerationRun(triggerType: "manual" | "scheduled", warnings: string[]) {
  db.prepare(
    `
    INSERT INTO generation_runs (trigger_type, generated_at, warnings_json)
    VALUES (?, ?, ?)
    `
  ).run(triggerType, new Date().toISOString(), JSON.stringify(warnings));
}

export function groupTasksByPriority(tasks: Task[]) {
  return {
    High: tasks.filter((task) => task.priority === "High"),
    Medium: tasks.filter((task) => task.priority === "Medium"),
    Low: tasks.filter((task) => task.priority === "Low")
  };
}

export function normalizeTask(row: Record<string, unknown>) {
  return taskRowToTask(row);
}
