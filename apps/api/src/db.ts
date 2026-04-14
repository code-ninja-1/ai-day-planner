import fs from "node:fs";
import path from "node:path";
import Database from "better-sqlite3";
import { env } from "./env.js";
import type {
  AuditEvent,
  AuditEventStatus,
  AuditLogLevel,
  AutomationSettings,
  BehaviorFeedbackEvent,
  FeedbackAction,
  FeedbackPolarity,
  IntegrationConnection,
  InsightsOverview,
  Meeting,
  HomeScheduleEntry,
  PlannerRunDetail,
  PersonalizationInsight,
  Reminder,
  ReminderKind,
  ReminderStatus,
  RejectedTask,
  ReasonTag,
  ScoreBreakdownItem,
  Task,
  TaskInsightsPayload,
  TaskDecisionState,
  TaskEffortBucket,
  TaskPriority,
  TaskStage,
  TaskStateEvent,
  TaskSource,
  TaskStatus,
  UserPriorityProfile
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
    stage TEXT NOT NULL DEFAULT 'Later',
    stage_order INTEGER NOT NULL DEFAULT 0,
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
    attendance_status TEXT NOT NULL DEFAULT 'attending',
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
    workday_start_local TEXT NOT NULL DEFAULT '09:30',
    workday_end_local TEXT NOT NULL DEFAULT '18:00',
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

  CREATE TABLE IF NOT EXISTS daily_plan_snapshots (
    day_key TEXT PRIMARY KEY,
    weekday INTEGER NOT NULL,
    base_workday_minutes INTEGER NOT NULL,
    adapted_task_capacity_minutes INTEGER NOT NULL,
    remaining_task_capacity_minutes INTEGER NOT NULL,
    meeting_minutes INTEGER NOT NULL,
    planned_task_minutes INTEGER NOT NULL,
    completed_task_minutes INTEGER NOT NULL,
    remaining_task_minutes INTEGER NOT NULL,
    spillover_task_count INTEGER NOT NULL,
    free_minutes INTEGER NOT NULL,
    focus_factor REAL NOT NULL DEFAULT 1,
    completion_rate REAL NOT NULL DEFAULT 0,
    planned_task_ids_json TEXT NOT NULL DEFAULT '[]',
    summary_json TEXT NOT NULL DEFAULT '{}',
    blocks_json TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS user_priority_profile (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    personalization_enabled INTEGER NOT NULL DEFAULT 1,
    role_focus TEXT,
    prioritization_prompt TEXT,
    important_work_json TEXT NOT NULL DEFAULT '[]',
    noise_work_json TEXT NOT NULL DEFAULT '[]',
    must_not_miss_json TEXT NOT NULL DEFAULT '[]',
    important_sources_json TEXT NOT NULL DEFAULT '[]',
    important_people_json TEXT NOT NULL DEFAULT '[]',
    important_projects_json TEXT NOT NULL DEFAULT '[]',
    positive_reason_tags_json TEXT NOT NULL DEFAULT '[]',
    negative_reason_tags_json TEXT NOT NULL DEFAULT '[]',
    filtering_style TEXT NOT NULL DEFAULT 'conservative',
    priority_bias TEXT NOT NULL DEFAULT 'balanced',
    questionnaire_json TEXT,
    example_rankings_json TEXT,
    last_profile_refresh_at TEXT,
    updated_at TEXT
  );

  CREATE TABLE IF NOT EXISTS rejected_tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    source TEXT NOT NULL,
    source_link TEXT,
    source_ref TEXT,
    source_thread_ref TEXT,
    jira_status TEXT,
    proposed_priority TEXT NOT NULL,
    decision_state TEXT NOT NULL DEFAULT 'rejected',
    decision_confidence REAL,
    decision_reason TEXT,
    decision_reason_tags TEXT NOT NULL DEFAULT '[]',
    personalization_version INTEGER,
    candidate_payload_json TEXT,
    rejected_at TEXT NOT NULL,
    restored_at TEXT,
    updated_at TEXT NOT NULL
  );

  CREATE UNIQUE INDEX IF NOT EXISTS idx_rejected_tasks_source_ref
  ON rejected_tasks(source, source_ref)
  WHERE source_ref IS NOT NULL;

  CREATE TABLE IF NOT EXISTS task_decision_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id INTEGER,
    source TEXT NOT NULL,
    source_ref TEXT,
    source_thread_ref TEXT,
    action TEXT NOT NULL,
    before_priority TEXT,
    after_priority TEXT,
    system_decision_state TEXT,
    decision_confidence REAL,
    decision_reason TEXT,
    decision_reason_tags TEXT NOT NULL DEFAULT '[]',
    features_json TEXT,
    feedback_payload_json TEXT,
    inferred_reason TEXT,
    inferred_reason_tag TEXT,
    preference_polarity TEXT NOT NULL DEFAULT 'neutral',
    created_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS preference_memory_snapshots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    snapshot_json TEXT NOT NULL,
    insights_json TEXT NOT NULL DEFAULT '[]',
    source_event_count INTEGER NOT NULL DEFAULT 0,
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS audit_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    level TEXT NOT NULL,
    event_type TEXT NOT NULL,
    request_id TEXT,
    run_id TEXT,
    entity_type TEXT,
    entity_id TEXT,
    provider TEXT,
    status TEXT NOT NULL,
    source TEXT,
    message TEXT NOT NULL,
    metadata_json TEXT
  );

  CREATE TABLE IF NOT EXISTS planner_run_details (
    run_id TEXT PRIMARY KEY,
    trigger_type TEXT NOT NULL,
    preferred_time_zone TEXT,
    warnings_json TEXT NOT NULL DEFAULT '[]',
    meeting_count INTEGER NOT NULL DEFAULT 0,
    active_task_count INTEGER NOT NULL DEFAULT 0,
    rejected_task_count INTEGER NOT NULL DEFAULT 0,
    deferred_task_count INTEGER NOT NULL DEFAULT 0,
    workload_state TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS task_state_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id INTEGER,
    source TEXT NOT NULL,
    source_ref TEXT,
    source_thread_ref TEXT,
    event_type TEXT NOT NULL,
    actor TEXT NOT NULL,
    reason TEXT,
    before_json TEXT,
    after_json TEXT,
    created_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS home_schedule_entries (
    entry_id TEXT PRIMARY KEY,
    day_key TEXT NOT NULL,
    task_id INTEGER NOT NULL,
    start_minutes INTEGER NOT NULL,
    duration_minutes INTEGER NOT NULL,
    source TEXT NOT NULL DEFAULT 'planner',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
  );

  CREATE INDEX IF NOT EXISTS idx_home_schedule_entries_day_key
  ON home_schedule_entries(day_key, start_minutes ASC);

  CREATE TABLE IF NOT EXISTS home_meeting_overrides (
    day_key TEXT NOT NULL,
    meeting_id INTEGER NOT NULL,
    visibility TEXT NOT NULL DEFAULT 'active',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    PRIMARY KEY(day_key, meeting_id)
  );

  CREATE TABLE IF NOT EXISTS home_schedule_days (
    day_key TEXT PRIMARY KEY,
    updated_at TEXT NOT NULL
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
ensureColumn("meetings", "attendance_status", "TEXT NOT NULL DEFAULT 'attending'");

ensureColumn("tasks", "deferred_until", "TEXT");
ensureColumn("tasks", "stage", "TEXT NOT NULL DEFAULT 'Later'");
ensureColumn("tasks", "stage_order", "INTEGER NOT NULL DEFAULT 0");
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
ensureColumn("tasks", "decision_state", "TEXT");
ensureColumn("tasks", "decision_confidence", "REAL");
ensureColumn("tasks", "decision_reason", "TEXT");
ensureColumn("tasks", "decision_reason_tags", "TEXT NOT NULL DEFAULT '[]'");
ensureColumn("tasks", "personalization_version", "INTEGER");
ensureColumn("tasks", "was_user_overridden", "INTEGER NOT NULL DEFAULT 0");
ensureColumn("tasks", "restored_at", "TEXT");
ensureColumn("tasks", "rejected_at", "TEXT");
ensureColumn("tasks", "jira_estimate_seconds", "INTEGER");
ensureColumn("tasks", "jira_subtask_estimate_seconds", "INTEGER");
ensureColumn("tasks", "jira_planning_subtasks_json", "TEXT NOT NULL DEFAULT '[]'");
ensureColumn("user_priority_profile", "prioritization_prompt", "TEXT");
ensureColumn("tasks", "selection_reason", "TEXT");
ensureColumn("tasks", "priority_reason", "TEXT");
ensureColumn("tasks", "score_breakdown_json", "TEXT NOT NULL DEFAULT '[]'");
ensureColumn("tasks", "history_signals_json", "TEXT NOT NULL DEFAULT '[]'");
ensureColumn("tasks", "last_changed_by", "TEXT");
ensureColumn("tasks", "last_changed_at", "TEXT");
ensureColumn("automation_settings", "workday_start_local", "TEXT NOT NULL DEFAULT '09:30'");
ensureColumn("automation_settings", "workday_end_local", "TEXT NOT NULL DEFAULT '18:00'");

db.prepare(
  `
  INSERT INTO automation_settings (id)
  VALUES (1)
  ON CONFLICT(id) DO NOTHING
  `
).run();

db.prepare(
  `
  INSERT INTO user_priority_profile (id, updated_at)
  VALUES (1, ?)
  ON CONFLICT(id) DO NOTHING
  `
).run(new Date().toISOString());

function parseJsonArray(value: unknown) {
  if (typeof value !== "string" || !value.trim()) return [];
  try {
    const parsed = JSON.parse(value) as unknown;
    return Array.isArray(parsed) ? parsed.filter((entry): entry is string => typeof entry === "string") : [];
  } catch {
    return [];
  }
}

function parseJsonValue<T>(value: unknown, fallback: T): T {
  if (typeof value !== "string" || !value.trim()) return fallback;
  try {
    return JSON.parse(value) as T;
  } catch {
    return fallback;
  }
}

function normalizeOptionalString(value: unknown) {
  if (typeof value !== "string") return null;
  const normalized = value.trim();
  return normalized ? normalized : null;
}

function buildOutlookCalendarItemLink(eventId: string) {
  return `https://outlook.office365.com/owa/?itemid=${encodeURIComponent(eventId)}&exvsurl=1&path=/calendar/item`;
}

const taskRowToTask = (row: Record<string, unknown>): Task => ({
  id: Number(row.id),
  title: String(row.title),
  source: row.source as TaskSource,
  stage: ((row.stage as TaskStage | null) ?? "Later") as TaskStage,
  stageOrder: Number(row.stage_order ?? 0),
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
  jiraEstimateSeconds:
    row.jira_estimate_seconds === null || row.jira_estimate_seconds === undefined
      ? null
      : Number(row.jira_estimate_seconds),
  jiraSubtaskEstimateSeconds:
    row.jira_subtask_estimate_seconds === null || row.jira_subtask_estimate_seconds === undefined
      ? null
      : Number(row.jira_subtask_estimate_seconds),
  jiraPlanningSubtasks: (() => {
    const parsed = parseJsonValue<unknown>(row.jira_planning_subtasks_json, []);
    return Array.isArray(parsed)
      ? parsed
          .filter((entry): entry is Record<string, unknown> => Boolean(entry) && typeof entry === "object")
          .map((entry) => ({
            key: String(entry.key ?? ""),
            title: String(entry.title ?? entry.key ?? ""),
            status: typeof entry.status === "string" ? entry.status : null,
            estimateSeconds:
              typeof entry.estimateSeconds === "number"
                ? entry.estimateSeconds
                : typeof entry.estimate_seconds === "number"
                  ? entry.estimate_seconds
                  : null
          }))
          .filter((entry) => entry.key)
      : [];
  })(),
  priorityScore: row.priority_score === null || row.priority_score === undefined ? null : Number(row.priority_score),
  priorityExplanation: (row.priority_explanation as string | null) ?? null,
  selectionReason: (row.selection_reason as string | null) ?? null,
  priorityReason: (row.priority_reason as string | null) ?? null,
  scoreBreakdown: parseJsonValue<ScoreBreakdownItem[]>(row.score_breakdown_json, []),
  historySignals: parseJsonArray(row.history_signals_json),
  taskAgeDays: Number(row.task_age_days ?? 0),
  carryForwardCount: Number(row.carry_forward_count ?? 0),
  completedAt: (row.completed_at as string | null) ?? null,
  lastActivityAt: (row.last_activity_at as string | null) ?? null,
  lastChangedBy: (row.last_changed_by as string | null) ?? null,
  lastChangedAt: (row.last_changed_at as string | null) ?? null,
  manualOverrideFlags: parseJsonArray(row.manual_override_flags),
  decisionState: (row.decision_state as TaskDecisionState | null) ?? null,
  decisionConfidence:
    row.decision_confidence === null || row.decision_confidence === undefined
      ? null
      : Number(row.decision_confidence),
  decisionReason: (row.decision_reason as string | null) ?? null,
  decisionReasonTags: parseJsonArray(row.decision_reason_tags) as ReasonTag[],
  personalizationVersion:
    row.personalization_version === null || row.personalization_version === undefined
      ? null
      : Number(row.personalization_version),
  wasUserOverridden: Number(row.was_user_overridden ?? 0) === 1,
  restoredAt: (row.restored_at as string | null) ?? null,
  rejectedAt: (row.rejected_at as string | null) ?? null,
  createdAt: String(row.created_at),
  updatedAt: String(row.updated_at)
});

const meetingRowToMeeting = (row: Record<string, unknown>): Meeting => {
  const externalId = normalizeOptionalString(row.external_id);
  const isCancelled = Number(row.is_cancelled) === 1;
  const storedLink = normalizeOptionalString(row.meeting_link);
  const storedLinkType = (row.meeting_link_type as "join" | "calendar" | null) ?? null;
  const fallbackCalendarLink = !isCancelled && externalId ? buildOutlookCalendarItemLink(externalId) : null;
  const meetingLink = storedLink ?? fallbackCalendarLink;
  const meetingLinkType = meetingLink ? (storedLinkType === "join" && storedLink ? "join" : "calendar") : null;

  return {
    id: Number(row.id),
    externalId,
    title: String(row.title),
    startTime: String(row.start_time),
    endTime: String(row.end_time),
    timeZone: (row.time_zone as string | null) ?? null,
    durationMinutes: Number(row.duration_minutes),
    meetingLink,
    meetingLinkType,
    isCancelled,
    attendanceStatus: ((row.attendance_status as string | null) ?? "attending") === "unattending" ? "unattending" : "attending",
    createdAt: String(row.created_at)
  };
};

const homeScheduleEntryRowToEntry = (row: Record<string, unknown>): HomeScheduleEntry => ({
  entryId: String(row.entry_id),
  dayKey: String(row.day_key),
  taskId: Number(row.task_id),
  startMinutes: Number(row.start_minutes),
  durationMinutes: Number(row.duration_minutes),
  source: row.source === "user" ? "user" : "planner",
  createdAt: String(row.created_at),
  updatedAt: String(row.updated_at)
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
  workdayStartLocal: String(row.workday_start_local ?? "09:30"),
  workdayEndLocal: String(row.workday_end_local ?? "18:00"),
  remindersEnabled: Number(row.reminders_enabled) === 1,
  reminderCadenceHours: Number(row.reminder_cadence_hours),
  desktopNotificationsEnabled: Number(row.desktop_notifications_enabled) === 1,
  lastAutoGeneratedAt: (row.last_auto_generated_at as string | null) ?? null,
  schedulerLastRunAt: (row.scheduler_last_run_at as string | null) ?? null,
  schedulerLastStatus: row.scheduler_last_status as AutomationSettings["schedulerLastStatus"],
  schedulerLastError: (row.scheduler_last_error as string | null) ?? null
});

const rejectedTaskRowToRejectedTask = (row: Record<string, unknown>): RejectedTask => ({
  id: Number(row.id),
  title: String(row.title),
  source: row.source as TaskSource,
  sourceLink: (row.source_link as string | null) ?? null,
  sourceRef: (row.source_ref as string | null) ?? null,
  sourceThreadRef: (row.source_thread_ref as string | null) ?? null,
  jiraStatus: (row.jira_status as string | null) ?? null,
  proposedPriority: row.proposed_priority as TaskPriority,
  decisionState: row.decision_state as TaskDecisionState,
  decisionConfidence:
    row.decision_confidence === null || row.decision_confidence === undefined
      ? null
      : Number(row.decision_confidence),
  decisionReason: (row.decision_reason as string | null) ?? null,
  decisionReasonTags: parseJsonArray(row.decision_reason_tags) as ReasonTag[],
  personalizationVersion:
    row.personalization_version === null || row.personalization_version === undefined
      ? null
      : Number(row.personalization_version),
  candidatePayloadJson: (row.candidate_payload_json as string | null) ?? null,
  rejectedAt: String(row.rejected_at),
  restoredAt: (row.restored_at as string | null) ?? null,
  updatedAt: String(row.updated_at)
});

const priorityProfileRowToProfile = (row: Record<string, unknown>): UserPriorityProfile => ({
  personalizationEnabled: Number(row.personalization_enabled) === 1,
  roleFocus: (row.role_focus as string | null) ?? null,
  prioritizationPrompt: (row.prioritization_prompt as string | null) ?? null,
  importantWork: parseJsonArray(row.important_work_json),
  noiseWork: parseJsonArray(row.noise_work_json),
  mustNotMiss: parseJsonArray(row.must_not_miss_json),
  importantSources: parseJsonArray(row.important_sources_json),
  importantPeople: parseJsonArray(row.important_people_json),
  importantProjects: parseJsonArray(row.important_projects_json),
  positiveReasonTags: parseJsonArray(row.positive_reason_tags_json) as ReasonTag[],
  negativeReasonTags: parseJsonArray(row.negative_reason_tags_json) as ReasonTag[],
  filteringStyle: row.filtering_style as UserPriorityProfile["filteringStyle"],
  priorityBias: row.priority_bias as UserPriorityProfile["priorityBias"],
  questionnaireJson: (row.questionnaire_json as string | null) ?? null,
  exampleRankingsJson: (row.example_rankings_json as string | null) ?? null,
  lastProfileRefreshAt: (row.last_profile_refresh_at as string | null) ?? null,
  updatedAt: (row.updated_at as string | null) ?? null
});

function orderByPrioritySql() {
  return [
    "CASE stage WHEN 'Now' THEN 1 WHEN 'Next' THEN 2 WHEN 'Later' THEN 3 WHEN 'Review' THEN 4 ELSE 5 END",
    "CASE WHEN status = 'In Progress' THEN 0 WHEN status = 'Not Started' THEN 1 ELSE 2 END",
    "stage_order ASC",
    "CASE priority WHEN 'High' THEN 1 WHEN 'Medium' THEN 2 ELSE 3 END",
    "priority_score DESC",
    "updated_at DESC"
  ].join(", ");
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
  return listTasks().filter((task) => task.stage === "Next");
}

export function listTasksByStage(stage: TaskStage) {
  return listTasks(undefined, { includeDeferred: true }).filter((task) => task.stage === stage && task.status !== "Completed");
}

export function getNextStageOrder(stage: TaskStage) {
  const row = db
    .prepare("SELECT COALESCE(MAX(stage_order), -1) as max_order FROM tasks WHERE stage = ? AND ignored = 0")
    .get(stage) as { max_order: number | null };
  return Number(row.max_order ?? -1) + 1;
}

export function reorderTasksWithinStage(stage: TaskStage, orderedTaskIds: number[]) {
  const transaction = db.transaction(() => {
    const update = db.prepare("UPDATE tasks SET stage = ?, stage_order = ?, updated_at = ? WHERE id = ?");
    const now = new Date().toISOString();
    orderedTaskIds.forEach((taskId, index) => {
      update.run(stage, index, now, taskId);
    });
  });
  transaction();
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

export function getUserPriorityProfile() {
  const row = db.prepare("SELECT * FROM user_priority_profile WHERE id = 1").get() as Record<string, unknown>;
  return priorityProfileRowToProfile(row);
}

export function saveUserPriorityProfile(
  patch: Partial<
    Pick<
      UserPriorityProfile,
      | "personalizationEnabled"
      | "roleFocus"
      | "prioritizationPrompt"
      | "importantWork"
      | "noiseWork"
      | "mustNotMiss"
      | "importantSources"
      | "importantPeople"
      | "importantProjects"
      | "positiveReasonTags"
      | "negativeReasonTags"
      | "filteringStyle"
      | "priorityBias"
      | "questionnaireJson"
      | "exampleRankingsJson"
      | "lastProfileRefreshAt"
    >
  >
) {
  const existing = getUserPriorityProfile();
  const updatedAt = new Date().toISOString();
  db.prepare(
    `
    UPDATE user_priority_profile
    SET personalization_enabled = @personalizationEnabled,
        role_focus = @roleFocus,
        prioritization_prompt = @prioritizationPrompt,
        important_work_json = @importantWork,
        noise_work_json = @noiseWork,
        must_not_miss_json = @mustNotMiss,
        important_sources_json = @importantSources,
        important_people_json = @importantPeople,
        important_projects_json = @importantProjects,
        positive_reason_tags_json = @positiveReasonTags,
        negative_reason_tags_json = @negativeReasonTags,
        filtering_style = @filteringStyle,
        priority_bias = @priorityBias,
        questionnaire_json = @questionnaireJson,
        example_rankings_json = @exampleRankingsJson,
        last_profile_refresh_at = @lastProfileRefreshAt,
        updated_at = @updatedAt
    WHERE id = 1
    `
  ).run({
    personalizationEnabled: Number(patch.personalizationEnabled ?? existing.personalizationEnabled),
    roleFocus: patch.roleFocus ?? existing.roleFocus,
    prioritizationPrompt: patch.prioritizationPrompt ?? existing.prioritizationPrompt,
    importantWork: JSON.stringify(patch.importantWork ?? existing.importantWork),
    noiseWork: JSON.stringify(patch.noiseWork ?? existing.noiseWork),
    mustNotMiss: JSON.stringify(patch.mustNotMiss ?? existing.mustNotMiss),
    importantSources: JSON.stringify(patch.importantSources ?? existing.importantSources),
    importantPeople: JSON.stringify(patch.importantPeople ?? existing.importantPeople),
    importantProjects: JSON.stringify(patch.importantProjects ?? existing.importantProjects),
    positiveReasonTags: JSON.stringify(patch.positiveReasonTags ?? existing.positiveReasonTags),
    negativeReasonTags: JSON.stringify(patch.negativeReasonTags ?? existing.negativeReasonTags),
    filteringStyle: patch.filteringStyle ?? existing.filteringStyle,
    priorityBias: patch.priorityBias ?? existing.priorityBias,
    questionnaireJson: patch.questionnaireJson ?? existing.questionnaireJson,
    exampleRankingsJson: patch.exampleRankingsJson ?? existing.exampleRankingsJson,
    lastProfileRefreshAt: patch.lastProfileRefreshAt ?? existing.lastProfileRefreshAt,
    updatedAt
  });
  return getUserPriorityProfile();
}

export function listRejectedTasks() {
  const tasks = db
    .prepare("SELECT * FROM rejected_tasks WHERE decision_state NOT IN ('restored', 'ignored') ORDER BY updated_at DESC, rejected_at DESC")
    .all()
    .map((row: unknown) => rejectedTaskRowToRejectedTask(row as Record<string, unknown>));

  const seenExact = new Set<string>();
  const seenEmailThreads = new Set<string>();

  return tasks.filter((task) => {
    if (task.source === "Email" && task.sourceThreadRef) {
      const threadKey = `${task.source}:${task.sourceThreadRef}`;
      if (seenEmailThreads.has(threadKey)) {
        return false;
      }
      seenEmailThreads.add(threadKey);
      return true;
    }

    if (task.sourceRef) {
      const exactKey = `${task.source}:${task.sourceRef}`;
      if (seenExact.has(exactKey)) {
        return false;
      }
      seenExact.add(exactKey);
    }

    return true;
  });
}

export function listIgnoredRejectedTasks() {
  const tasks = db
    .prepare("SELECT * FROM rejected_tasks WHERE decision_state = 'ignored' ORDER BY updated_at DESC, rejected_at DESC")
    .all()
    .map((row: unknown) => rejectedTaskRowToRejectedTask(row as Record<string, unknown>));

  const seenExact = new Set<string>();
  const seenEmailThreads = new Set<string>();

  return tasks.filter((task) => {
    if (task.source === "Email" && task.sourceThreadRef) {
      const threadKey = `${task.source}:${task.sourceThreadRef}`;
      if (seenEmailThreads.has(threadKey)) {
        return false;
      }
      seenEmailThreads.add(threadKey);
      return true;
    }

    if (task.sourceRef) {
      const exactKey = `${task.source}:${task.sourceRef}`;
      if (seenExact.has(exactKey)) {
        return false;
      }
      seenExact.add(exactKey);
    }

    return true;
  });
}

export function getRejectedTaskById(id: number) {
  const row = db.prepare("SELECT * FROM rejected_tasks WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  return row ? rejectedTaskRowToRejectedTask(row) : null;
}

export function getRejectedTaskBySource(source: TaskSource, sourceRef: string | null) {
  if (!sourceRef) return null;
  const row = db
    .prepare("SELECT * FROM rejected_tasks WHERE source = ? AND source_ref = ?")
    .get(source, sourceRef) as Record<string, unknown> | undefined;
  return row ? rejectedTaskRowToRejectedTask(row) : null;
}

export function getRejectedTaskBySourceThread(source: TaskSource, sourceThreadRef: string | null) {
  if (!sourceThreadRef) return null;
  const row = db
    .prepare(
      "SELECT * FROM rejected_tasks WHERE source = ? AND source_thread_ref = ? AND decision_state NOT IN ('restored', 'ignored') ORDER BY updated_at DESC LIMIT 1"
    )
    .get(source, sourceThreadRef) as Record<string, unknown> | undefined;
  return row ? rejectedTaskRowToRejectedTask(row) : null;
}

export function upsertRejectedTask(input: {
  title: string;
  source: TaskSource;
  sourceLink?: string | null;
  sourceRef?: string | null;
  sourceThreadRef?: string | null;
  jiraStatus?: string | null;
  proposedPriority: TaskPriority;
  decisionState?: TaskDecisionState;
  decisionConfidence?: number | null;
  decisionReason?: string | null;
  decisionReasonTags?: ReasonTag[];
  personalizationVersion?: number | null;
  candidatePayloadJson?: string | null;
  rejectedAt?: string;
  restoredAt?: string | null;
}) {
  const now = new Date().toISOString();
  const rejectedAt = input.rejectedAt ?? now;
  const existingBySourceRef =
    input.sourceRef
      ? (db
          .prepare("SELECT id, decision_state FROM rejected_tasks WHERE source = ? AND source_ref = ?")
          .get(input.source, input.sourceRef) as { id: number; decision_state: TaskDecisionState } | undefined)
      : undefined;
  const threadExisting =
    input.source === "Email" && input.sourceThreadRef
      ? (db
          .prepare(
            "SELECT id FROM rejected_tasks WHERE source = 'Email' AND source_thread_ref = ? AND decision_state NOT IN ('restored', 'ignored') ORDER BY updated_at DESC LIMIT 1"
          )
          .get(input.sourceThreadRef) as { id: number } | undefined)
      : undefined;
  const ignoredThreadExisting =
    input.source === "Email" && input.sourceThreadRef
      ? (db
          .prepare(
            "SELECT id, decision_state FROM rejected_tasks WHERE source = 'Email' AND source_thread_ref = ? AND decision_state = 'ignored' ORDER BY updated_at DESC LIMIT 1"
          )
          .get(input.sourceThreadRef) as { id: number; decision_state: TaskDecisionState } | undefined)
      : undefined;

  const ignoredExistingId =
    existingBySourceRef?.decision_state === "ignored"
      ? existingBySourceRef.id
      : ignoredThreadExisting?.decision_state === "ignored"
        ? ignoredThreadExisting.id
        : null;

  if (ignoredExistingId !== null) {
    return getRejectedTaskById(ignoredExistingId);
  }

  if (input.sourceRef) {
    if (existingBySourceRef || threadExisting) {
      db.prepare(
        `
        UPDATE rejected_tasks
        SET title = @title,
            source_link = @sourceLink,
            source_ref = @sourceRef,
            source_thread_ref = @sourceThreadRef,
            jira_status = @jiraStatus,
            proposed_priority = @proposedPriority,
            decision_state = @decisionState,
            decision_confidence = @decisionConfidence,
            decision_reason = @decisionReason,
            decision_reason_tags = @decisionReasonTags,
            personalization_version = @personalizationVersion,
            candidate_payload_json = @candidatePayloadJson,
            rejected_at = @rejectedAt,
            restored_at = @restoredAt,
            updated_at = @updatedAt
        WHERE id = @id
        `
      ).run({
        id: existingBySourceRef?.id ?? threadExisting?.id,
        title: input.title,
        sourceLink: input.sourceLink ?? null,
        sourceRef: input.sourceRef ?? null,
        sourceThreadRef: input.sourceThreadRef ?? null,
        jiraStatus: input.jiraStatus ?? null,
        proposedPriority: input.proposedPriority,
        decisionState: input.decisionState ?? "rejected",
        decisionConfidence: input.decisionConfidence ?? null,
        decisionReason: input.decisionReason ?? null,
        decisionReasonTags: JSON.stringify(input.decisionReasonTags ?? []),
        personalizationVersion: input.personalizationVersion ?? null,
        candidatePayloadJson: input.candidatePayloadJson ?? null,
        rejectedAt,
        restoredAt: input.restoredAt ?? null,
        updatedAt: now
      });
    } else {
      db.prepare(
        `
        INSERT INTO rejected_tasks (
          title, source, source_link, source_ref, source_thread_ref, jira_status, proposed_priority,
          decision_state, decision_confidence, decision_reason, decision_reason_tags, personalization_version,
          candidate_payload_json, rejected_at, restored_at, updated_at
        ) VALUES (
          @title, @source, @sourceLink, @sourceRef, @sourceThreadRef, @jiraStatus, @proposedPriority,
          @decisionState, @decisionConfidence, @decisionReason, @decisionReasonTags, @personalizationVersion,
          @candidatePayloadJson, @rejectedAt, @restoredAt, @updatedAt
        )
        `
      ).run({
        title: input.title,
        source: input.source,
        sourceLink: input.sourceLink ?? null,
        sourceRef: input.sourceRef,
        sourceThreadRef: input.sourceThreadRef ?? null,
        jiraStatus: input.jiraStatus ?? null,
        proposedPriority: input.proposedPriority,
        decisionState: input.decisionState ?? "rejected",
        decisionConfidence: input.decisionConfidence ?? null,
        decisionReason: input.decisionReason ?? null,
        decisionReasonTags: JSON.stringify(input.decisionReasonTags ?? []),
        personalizationVersion: input.personalizationVersion ?? null,
        candidatePayloadJson: input.candidatePayloadJson ?? null,
        rejectedAt,
        restoredAt: input.restoredAt ?? null,
        updatedAt: now
      });
    }
  } else {
    db.prepare(
      `
      INSERT INTO rejected_tasks (
        title, source, source_link, source_ref, source_thread_ref, jira_status, proposed_priority,
        decision_state, decision_confidence, decision_reason, decision_reason_tags, personalization_version,
        candidate_payload_json, rejected_at, restored_at, updated_at
      ) VALUES (
        @title, @source, @sourceLink, NULL, @sourceThreadRef, @jiraStatus, @proposedPriority,
        @decisionState, @decisionConfidence, @decisionReason, @decisionReasonTags, @personalizationVersion,
        @candidatePayloadJson, @rejectedAt, @restoredAt, @updatedAt
      )
      `
    ).run({
      title: input.title,
      source: input.source,
      sourceLink: input.sourceLink ?? null,
      sourceThreadRef: input.sourceThreadRef ?? null,
      jiraStatus: input.jiraStatus ?? null,
      proposedPriority: input.proposedPriority,
      decisionState: input.decisionState ?? "rejected",
      decisionConfidence: input.decisionConfidence ?? null,
      decisionReason: input.decisionReason ?? null,
      decisionReasonTags: JSON.stringify(input.decisionReasonTags ?? []),
      personalizationVersion: input.personalizationVersion ?? null,
      candidatePayloadJson: input.candidatePayloadJson ?? null,
      rejectedAt,
      restoredAt: input.restoredAt ?? null,
      updatedAt: now
    });
  }

  if (input.sourceRef) {
    return getRejectedTaskBySource(input.source, input.sourceRef);
  }

  if (input.source === "Email" && input.sourceThreadRef) {
    return getRejectedTaskBySourceThread(input.source, input.sourceThreadRef);
  }

  return null;
}

export function updateRejectedTask(
  id: number,
  patch: Partial<
    Pick<
      RejectedTask,
      "title" | "sourceLink" | "jiraStatus" | "proposedPriority" | "decisionState" | "decisionConfidence" | "decisionReason" | "personalizationVersion" | "restoredAt"
    > & { decisionReasonTags?: ReasonTag[]; candidatePayloadJson?: string | null }
  >
) {
  const existing = db.prepare("SELECT * FROM rejected_tasks WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  if (!existing) return null;
  db.prepare(
    `
    UPDATE rejected_tasks
    SET title = @title,
        source_link = @sourceLink,
        jira_status = @jiraStatus,
        proposed_priority = @proposedPriority,
        decision_state = @decisionState,
        decision_confidence = @decisionConfidence,
        decision_reason = @decisionReason,
        decision_reason_tags = @decisionReasonTags,
        personalization_version = @personalizationVersion,
        candidate_payload_json = @candidatePayloadJson,
        restored_at = @restoredAt,
        updated_at = @updatedAt
    WHERE id = @id
    `
  ).run({
    id,
    title: patch.title ?? existing.title,
    sourceLink: patch.sourceLink ?? existing.source_link,
    jiraStatus: patch.jiraStatus ?? existing.jira_status,
    proposedPriority: patch.proposedPriority ?? existing.proposed_priority,
    decisionState: patch.decisionState ?? existing.decision_state,
    decisionConfidence: patch.decisionConfidence ?? existing.decision_confidence,
    decisionReason: patch.decisionReason ?? existing.decision_reason,
    decisionReasonTags: JSON.stringify(patch.decisionReasonTags ?? parseJsonArray(existing.decision_reason_tags)),
    personalizationVersion: patch.personalizationVersion ?? existing.personalization_version,
    candidatePayloadJson: patch.candidatePayloadJson ?? existing.candidate_payload_json ?? null,
    restoredAt: patch.restoredAt ?? existing.restored_at ?? null,
    updatedAt: new Date().toISOString()
  });
  return getRejectedTaskById(id);
}

export function clearRejectedTasksBySourceThread(source: TaskSource, sourceThreadRef: string | null) {
  if (!sourceThreadRef) return 0;
  const now = new Date().toISOString();
  const result = db
    .prepare(
      `
      UPDATE rejected_tasks
      SET decision_state = 'restored',
          restored_at = COALESCE(restored_at, @restoredAt),
          updated_at = @updatedAt
      WHERE source = @source
        AND source_thread_ref = @sourceThreadRef
        AND decision_state != 'restored'
      `
    )
    .run({
      source,
      sourceThreadRef,
      restoredAt: now,
      updatedAt: now
    });
  return Number(result.changes ?? 0);
}

export function logTaskDecisionEvent(input: {
  taskId?: number | null;
  source: TaskSource | "Calibration";
  sourceRef?: string | null;
  sourceThreadRef?: string | null;
  action: FeedbackAction;
  beforePriority?: TaskPriority | null;
  afterPriority?: TaskPriority | null;
  systemDecisionState?: TaskDecisionState | null;
  decisionConfidence?: number | null;
  decisionReason?: string | null;
  decisionReasonTags?: ReasonTag[];
  featuresJson?: string | null;
  feedbackPayloadJson?: string | null;
  inferredReason?: string | null;
  inferredReasonTag?: ReasonTag | null;
  preferencePolarity?: FeedbackPolarity;
}) {
  db.prepare(
    `
    INSERT INTO task_decision_log (
      task_id, source, source_ref, source_thread_ref, action, before_priority, after_priority, system_decision_state,
      decision_confidence, decision_reason, decision_reason_tags, features_json, feedback_payload_json,
      inferred_reason, inferred_reason_tag, preference_polarity, created_at
    ) VALUES (
      @taskId, @source, @sourceRef, @sourceThreadRef, @action, @beforePriority, @afterPriority, @systemDecisionState,
      @decisionConfidence, @decisionReason, @decisionReasonTags, @featuresJson, @feedbackPayloadJson,
      @inferredReason, @inferredReasonTag, @preferencePolarity, @createdAt
    )
    `
  ).run({
    taskId: input.taskId ?? null,
    source: input.source,
    sourceRef: input.sourceRef ?? null,
    sourceThreadRef: input.sourceThreadRef ?? null,
    action: input.action,
    beforePriority: input.beforePriority ?? null,
    afterPriority: input.afterPriority ?? null,
    systemDecisionState: input.systemDecisionState ?? null,
    decisionConfidence: input.decisionConfidence ?? null,
    decisionReason: input.decisionReason ?? null,
    decisionReasonTags: JSON.stringify(input.decisionReasonTags ?? []),
    featuresJson: input.featuresJson ?? null,
    feedbackPayloadJson: input.feedbackPayloadJson ?? null,
    inferredReason: input.inferredReason ?? null,
    inferredReasonTag: input.inferredReasonTag ?? null,
    preferencePolarity: input.preferencePolarity ?? "neutral",
    createdAt: new Date().toISOString()
  });
}

export function listRecentDecisionLogs(limit = 60) {
  return db
    .prepare("SELECT * FROM task_decision_log ORDER BY created_at DESC LIMIT ?")
    .all(limit) as Array<Record<string, unknown>>;
}

export function getDecisionEventCount() {
  const row = db.prepare("SELECT COUNT(*) as count FROM task_decision_log").get() as { count: number };
  return Number(row.count ?? 0);
}

const auditRowToEvent = (row: Record<string, unknown>): AuditEvent => ({
  id: Number(row.id),
  timestamp: String(row.timestamp),
  level: row.level as AuditLogLevel,
  eventType: String(row.event_type),
  requestId: (row.request_id as string | null) ?? null,
  runId: (row.run_id as string | null) ?? null,
  entityType: (row.entity_type as string | null) ?? null,
  entityId: (row.entity_id as string | null) ?? null,
  provider: (row.provider as string | null) ?? null,
  status: row.status as AuditEventStatus,
  source: (row.source as string | null) ?? null,
  message: String(row.message),
  metadataJson: (row.metadata_json as string | null) ?? null
});

const plannerRunRowToDetail = (row: Record<string, unknown>): PlannerRunDetail => ({
  runId: String(row.run_id),
  triggerType: row.trigger_type as PlannerRunDetail["triggerType"],
  preferredTimeZone: (row.preferred_time_zone as string | null) ?? null,
  warnings: parseJsonValue<string[]>(row.warnings_json, []),
  meetingCount: Number(row.meeting_count ?? 0),
  activeTaskCount: Number(row.active_task_count ?? 0),
  rejectedTaskCount: Number(row.rejected_task_count ?? 0),
  deferredTaskCount: Number(row.deferred_task_count ?? 0),
  workloadState: (row.workload_state as PlannerRunDetail["workloadState"]) ?? null,
  createdAt: String(row.created_at),
  updatedAt: String(row.updated_at)
});

const taskStateEventRowToEvent = (row: Record<string, unknown>): TaskStateEvent => ({
  id: Number(row.id),
  taskId: row.task_id === null || row.task_id === undefined ? null : Number(row.task_id),
  source: row.source as TaskStateEvent["source"],
  sourceRef: (row.source_ref as string | null) ?? null,
  sourceThreadRef: (row.source_thread_ref as string | null) ?? null,
  eventType: String(row.event_type),
  actor: row.actor as TaskStateEvent["actor"],
  reason: (row.reason as string | null) ?? null,
  beforeJson: (row.before_json as string | null) ?? null,
  afterJson: (row.after_json as string | null) ?? null,
  createdAt: String(row.created_at)
});

export function recordAuditEvent(input: Omit<AuditEvent, "id" | "timestamp"> & { timestamp?: string }) {
  db.prepare(
    `
    INSERT INTO audit_events (
      timestamp, level, event_type, request_id, run_id, entity_type, entity_id, provider, status, source, message, metadata_json
    ) VALUES (
      @timestamp, @level, @eventType, @requestId, @runId, @entityType, @entityId, @provider, @status, @source, @message, @metadataJson
    )
    `
  ).run({
    timestamp: input.timestamp ?? new Date().toISOString(),
    level: input.level,
    eventType: input.eventType,
    requestId: input.requestId ?? null,
    runId: input.runId ?? null,
    entityType: input.entityType ?? null,
    entityId: input.entityId ?? null,
    provider: input.provider ?? null,
    status: input.status,
    source: input.source ?? null,
    message: input.message,
    metadataJson: input.metadataJson ?? null
  });
}

export function listAuditEvents(limit = 200) {
  return db
    .prepare("SELECT * FROM audit_events ORDER BY timestamp DESC, id DESC LIMIT ?")
    .all(limit)
    .map((row: unknown) => auditRowToEvent(row as Record<string, unknown>));
}

export function upsertPlannerRunDetail(input: Omit<PlannerRunDetail, "createdAt" | "updatedAt">) {
  const now = new Date().toISOString();
  db.prepare(
    `
    INSERT INTO planner_run_details (
      run_id, trigger_type, preferred_time_zone, warnings_json, meeting_count, active_task_count,
      rejected_task_count, deferred_task_count, workload_state, created_at, updated_at
    ) VALUES (
      @runId, @triggerType, @preferredTimeZone, @warningsJson, @meetingCount, @activeTaskCount,
      @rejectedTaskCount, @deferredTaskCount, @workloadState, @createdAt, @updatedAt
    )
    ON CONFLICT(run_id) DO UPDATE SET
      trigger_type = excluded.trigger_type,
      preferred_time_zone = excluded.preferred_time_zone,
      warnings_json = excluded.warnings_json,
      meeting_count = excluded.meeting_count,
      active_task_count = excluded.active_task_count,
      rejected_task_count = excluded.rejected_task_count,
      deferred_task_count = excluded.deferred_task_count,
      workload_state = excluded.workload_state,
      updated_at = excluded.updated_at
    `
  ).run({
    runId: input.runId,
    triggerType: input.triggerType,
    preferredTimeZone: input.preferredTimeZone ?? null,
    warningsJson: JSON.stringify(input.warnings ?? []),
    meetingCount: input.meetingCount,
    activeTaskCount: input.activeTaskCount,
    rejectedTaskCount: input.rejectedTaskCount,
    deferredTaskCount: input.deferredTaskCount,
    workloadState: input.workloadState ?? null,
    createdAt: now,
    updatedAt: now
  });
}

export function listPlannerRunDetails(limit = 30) {
  return db
    .prepare("SELECT * FROM planner_run_details ORDER BY created_at DESC LIMIT ?")
    .all(limit)
    .map((row: unknown) => plannerRunRowToDetail(row as Record<string, unknown>));
}

export function recordTaskStateEvent(input: Omit<TaskStateEvent, "id" | "createdAt"> & { createdAt?: string }) {
  db.prepare(
    `
    INSERT INTO task_state_events (
      task_id, source, source_ref, source_thread_ref, event_type, actor, reason, before_json, after_json, created_at
    ) VALUES (
      @taskId, @source, @sourceRef, @sourceThreadRef, @eventType, @actor, @reason, @beforeJson, @afterJson, @createdAt
    )
    `
  ).run({
    taskId: input.taskId ?? null,
    source: input.source,
    sourceRef: input.sourceRef ?? null,
    sourceThreadRef: input.sourceThreadRef ?? null,
    eventType: input.eventType,
    actor: input.actor,
    reason: input.reason ?? null,
    beforeJson: input.beforeJson ?? null,
    afterJson: input.afterJson ?? null,
    createdAt: input.createdAt ?? new Date().toISOString()
  });
}

export function listTaskStateEvents(options?: { taskId?: number; dayKey?: string; startDayKey?: string; endDayKey?: string; limit?: number }) {
  const filters: string[] = [];
  const values: Array<string | number> = [];

  if (options?.taskId !== undefined) {
    filters.push("task_id = ?");
    values.push(options.taskId);
  }

  if (options?.dayKey) {
    filters.push("substr(created_at, 1, 10) = ?");
    values.push(options.dayKey);
  }

  if (options?.startDayKey) {
    filters.push("substr(created_at, 1, 10) >= ?");
    values.push(options.startDayKey);
  }

  if (options?.endDayKey) {
    filters.push("substr(created_at, 1, 10) <= ?");
    values.push(options.endDayKey);
  }

  const whereClause = filters.length ? `WHERE ${filters.join(" AND ")}` : "";
  const limitClause = options?.limit ? `LIMIT ${Number(options.limit)}` : "";

  return db
    .prepare(`SELECT * FROM task_state_events ${whereClause} ORDER BY created_at DESC, id DESC ${limitClause}`)
    .all(...values)
    .map((row: unknown) => taskStateEventRowToEvent(row as Record<string, unknown>));
}

export function savePreferenceMemorySnapshot(input: {
  snapshotJson: string;
  insights: PersonalizationInsight[];
  sourceEventCount: number;
}) {
  db.prepare("UPDATE preference_memory_snapshots SET active = 0 WHERE active = 1").run();
  db.prepare(
    `
    INSERT INTO preference_memory_snapshots (snapshot_json, insights_json, source_event_count, active, created_at)
    VALUES (?, ?, ?, 1, ?)
    `
  ).run(input.snapshotJson, JSON.stringify(input.insights), input.sourceEventCount, new Date().toISOString());
}

export function getLatestPreferenceMemorySnapshot() {
  const row = db
    .prepare("SELECT * FROM preference_memory_snapshots WHERE active = 1 ORDER BY created_at DESC LIMIT 1")
    .get() as Record<string, unknown> | undefined;
  if (!row) {
    return {
      snapshotJson: "{}",
      insights: [] as PersonalizationInsight[],
      sourceEventCount: 0,
      createdAt: null as string | null
    };
  }
  return {
    snapshotJson: String(row.snapshot_json),
    insights: (() => {
      try {
        return JSON.parse(String(row.insights_json)) as PersonalizationInsight[];
      } catch {
        return [] as PersonalizationInsight[];
      }
    })(),
    sourceEventCount: Number(row.source_event_count ?? 0),
    createdAt: String(row.created_at)
  };
}

export function getTaskById(id: number) {
  const row = db.prepare("SELECT * FROM tasks WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  return row ? taskRowToTask(row) : null;
}

export function getTaskBySource(source: TaskSource, sourceRef: string | null, options?: { includeIgnored?: boolean }) {
  if (!sourceRef) return null;
  const row = db
    .prepare(
      `SELECT * FROM tasks WHERE source = ? AND source_ref = ?${options?.includeIgnored ? "" : " AND ignored = 0"} LIMIT 1`
    )
    .get(source, sourceRef) as Record<string, unknown> | undefined;
  return row ? taskRowToTask(row) : null;
}

export function getTaskBySourceThread(
  source: TaskSource,
  sourceThreadRef: string | null,
  options?: { includeIgnored?: boolean }
) {
  if (!sourceThreadRef) return null;
  const row = db
    .prepare(
      `SELECT * FROM tasks WHERE source = ? AND source_thread_ref = ?${options?.includeIgnored ? "" : " AND ignored = 0"} ORDER BY updated_at DESC LIMIT 1`
    )
    .get(source, sourceThreadRef) as Record<string, unknown> | undefined;
  return row ? taskRowToTask(row) : null;
}

export function listMeetings() {
  return db
    .prepare("SELECT * FROM meetings ORDER BY start_time ASC")
    .all()
    .map((row: unknown) => meetingRowToMeeting(row as Record<string, unknown>));
}

export function getMeetingById(id: number) {
  const row = db.prepare("SELECT * FROM meetings WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  return row ? meetingRowToMeeting(row) : null;
}

export function updateMeetingAttendanceStatus(id: number, attendanceStatus: "attending" | "unattending") {
  const existing = db.prepare("SELECT * FROM meetings WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  if (!existing) return null;
  db.prepare(
    `
    UPDATE meetings
    SET attendance_status = ?
    WHERE id = ?
    `
  ).run(attendanceStatus, id);
  const row = db.prepare("SELECT * FROM meetings WHERE id = ?").get(id) as Record<string, unknown> | undefined;
  return row ? meetingRowToMeeting(row) : null;
}

export function listHomeScheduleEntries(dayKey: string) {
  return db
    .prepare("SELECT * FROM home_schedule_entries WHERE day_key = ? ORDER BY start_minutes ASC, entry_id ASC")
    .all(dayKey)
    .map((row: unknown) => homeScheduleEntryRowToEntry(row as Record<string, unknown>));
}

export function hasHomeScheduleOverrideDay(dayKey: string) {
  const row = db.prepare("SELECT day_key FROM home_schedule_days WHERE day_key = ?").get(dayKey) as { day_key?: string } | undefined;
  return Boolean(row?.day_key);
}

export function markHomeScheduleOverrideDay(dayKey: string) {
  db.prepare(
    `
    INSERT INTO home_schedule_days (day_key, updated_at)
    VALUES (?, ?)
    ON CONFLICT(day_key) DO UPDATE SET
      updated_at = excluded.updated_at
    `
  ).run(dayKey, new Date().toISOString());
}

export function replaceHomeScheduleEntries(
  dayKey: string,
  entries: Array<Pick<HomeScheduleEntry, "entryId" | "taskId" | "startMinutes" | "durationMinutes" | "source">>
) {
  const now = new Date().toISOString();
  const transaction = db.transaction(() => {
    db.prepare(
      `
      INSERT INTO home_schedule_days (day_key, updated_at)
      VALUES (?, ?)
      ON CONFLICT(day_key) DO UPDATE SET
        updated_at = excluded.updated_at
      `
    ).run(dayKey, now);
    db.prepare("DELETE FROM home_schedule_entries WHERE day_key = ?").run(dayKey);
    const insert = db.prepare(
      `
      INSERT INTO home_schedule_entries (
        entry_id, day_key, task_id, start_minutes, duration_minutes, source, created_at, updated_at
      ) VALUES (
        @entryId, @dayKey, @taskId, @startMinutes, @durationMinutes, @source, @createdAt, @updatedAt
      )
      `
    );
    entries.forEach((entry) => {
      insert.run({
        entryId: entry.entryId,
        dayKey,
        taskId: entry.taskId,
        startMinutes: entry.startMinutes,
        durationMinutes: entry.durationMinutes,
        source: entry.source,
        createdAt: now,
        updatedAt: now
      });
    });
  });
  transaction();
  return listHomeScheduleEntries(dayKey);
}

export function deleteHomeScheduleEntry(entryId: string) {
  const row = db
    .prepare("SELECT * FROM home_schedule_entries WHERE entry_id = ?")
    .get(entryId) as Record<string, unknown> | undefined;
  if (!row) return null;
  markHomeScheduleOverrideDay(String(row.day_key));
  db.prepare("DELETE FROM home_schedule_entries WHERE entry_id = ?").run(entryId);
  return homeScheduleEntryRowToEntry(row);
}

export function listHiddenHomeMeetingIds(dayKey: string) {
  return db
    .prepare("SELECT meeting_id FROM home_meeting_overrides WHERE day_key = ? AND visibility = 'removed' ORDER BY meeting_id ASC")
    .all(dayKey)
    .map((row: unknown) => Number((row as { meeting_id: number }).meeting_id));
}

export function setHomeMeetingVisibility(dayKey: string, meetingId: number, visibility: "active" | "removed") {
  const now = new Date().toISOString();
  markHomeScheduleOverrideDay(dayKey);
  if (visibility === "active") {
    db.prepare("DELETE FROM home_meeting_overrides WHERE day_key = ? AND meeting_id = ?").run(dayKey, meetingId);
    return;
  }
  db.prepare(
    `
    INSERT INTO home_meeting_overrides (day_key, meeting_id, visibility, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?)
    ON CONFLICT(day_key, meeting_id) DO UPDATE SET
      visibility = excluded.visibility,
      updated_at = excluded.updated_at
    `
  ).run(dayKey, meetingId, visibility, now, now);
}

function mergeOverrideFlags(existing: string[], next: string[]) {
  return [...new Set([...existing, ...next])];
}

export function upsertTask(input: {
  title: string;
  source: TaskSource;
  stage?: TaskStage;
  stageOrder?: number | null;
  priority: TaskPriority;
  status?: TaskStatus;
  sourceLink?: string | null;
  sourceRef?: string | null;
  sourceThreadRef?: string | null;
  jiraStatus?: string | null;
  estimatedEffortBucket?: TaskEffortBucket | null;
  priorityExplanation?: string | null;
  decisionState?: TaskDecisionState | null;
  decisionConfidence?: number | null;
  decisionReason?: string | null;
  decisionReasonTags?: ReasonTag[];
  personalizationVersion?: number | null;
  restoredAt?: string | null;
  rejectedAt?: string | null;
  lastActivityAt?: string | null;
  selectionReason?: string | null;
  priorityReason?: string | null;
  scoreBreakdown?: ScoreBreakdownItem[];
  historySignals?: string[];
  lastChangedBy?: string | null;
  lastChangedAt?: string | null;
  jiraEstimateSeconds?: number | null;
  jiraSubtaskEstimateSeconds?: number | null;
  jiraPlanningSubtasks?: Array<{
    key: string;
    title: string;
    status: string | null;
    estimateSeconds: number | null;
  }>;
  reviveIgnored?: boolean;
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
          stage = @stage,
          stage_order = @stageOrder,
          priority = @priority,
          status = @status,
          source_link = @sourceLink,
          source_thread_ref = @sourceThreadRef,
          jira_status = @jiraStatus,
          estimated_effort_bucket = @estimatedEffortBucket,
          jira_estimate_seconds = @jiraEstimateSeconds,
          jira_subtask_estimate_seconds = @jiraSubtaskEstimateSeconds,
          jira_planning_subtasks_json = @jiraPlanningSubtasksJson,
          ignored = @ignored,
          decision_state = @decisionState,
          decision_confidence = @decisionConfidence,
          decision_reason = @decisionReason,
          decision_reason_tags = @decisionReasonTags,
          personalization_version = @personalizationVersion,
          priority_explanation = @priorityExplanation,
          selection_reason = @selectionReason,
          priority_reason = @priorityReason,
          score_breakdown_json = @scoreBreakdownJson,
          history_signals_json = @historySignalsJson,
          last_changed_by = @lastChangedBy,
          last_changed_at = @lastChangedAt,
          restored_at = @restoredAt,
          rejected_at = @rejectedAt,
          last_activity_at = @lastActivityAt,
          updated_at = @updatedAt
      WHERE source = @source AND source_ref = @sourceRef
      `
    ).run({
      title: input.title,
      stage: overrideFlags.includes("stage") ? existing.stage : (input.stage ?? existing.stage ?? "Later"),
      stageOrder:
        overrideFlags.includes("stageOrder") || overrideFlags.includes("stage")
          ? Number(existing.stage_order ?? 0)
          : (input.stageOrder ?? Number(existing.stage_order ?? 0)),
      priority: overrideFlags.includes("priority") ? existing.priority : input.priority,
      status: overrideFlags.includes("status") ? existing.status : (input.status ?? existing.status),
      sourceLink: input.sourceLink ?? null,
      sourceThreadRef: input.sourceThreadRef ?? null,
      jiraStatus: input.jiraStatus ?? null,
      estimatedEffortBucket: input.estimatedEffortBucket ?? existing.estimated_effort_bucket ?? null,
      jiraEstimateSeconds: input.jiraEstimateSeconds ?? existing.jira_estimate_seconds ?? null,
      jiraSubtaskEstimateSeconds: input.jiraSubtaskEstimateSeconds ?? existing.jira_subtask_estimate_seconds ?? null,
      jiraPlanningSubtasksJson: JSON.stringify(input.jiraPlanningSubtasks ?? (() => {
        try {
          return JSON.parse(String(existing.jira_planning_subtasks_json ?? "[]"));
        } catch {
          return [];
        }
      })()),
      ignored: input.reviveIgnored ? 0 : existing.ignored,
      decisionState: input.decisionState ?? existing.decision_state ?? "accepted",
      decisionConfidence: input.decisionConfidence ?? existing.decision_confidence ?? null,
      decisionReason: input.decisionReason ?? existing.decision_reason ?? null,
      decisionReasonTags: JSON.stringify(input.decisionReasonTags ?? parseJsonArray(existing.decision_reason_tags)),
      personalizationVersion: input.personalizationVersion ?? existing.personalization_version ?? null,
      priorityExplanation: input.priorityExplanation ?? existing.priority_explanation ?? null,
      selectionReason: input.selectionReason ?? existing.selection_reason ?? null,
      priorityReason: input.priorityReason ?? existing.priority_reason ?? null,
      scoreBreakdownJson: JSON.stringify(input.scoreBreakdown ?? parseJsonValue(existing.score_breakdown_json, [])),
      historySignalsJson: JSON.stringify(input.historySignals ?? parseJsonArray(existing.history_signals_json)),
      lastChangedBy: input.lastChangedBy ?? existing.last_changed_by ?? null,
      lastChangedAt: input.lastChangedAt ?? existing.last_changed_at ?? now,
      restoredAt: input.restoredAt ?? existing.restored_at ?? null,
      rejectedAt: input.rejectedAt ?? existing.rejected_at ?? null,
      lastActivityAt: input.lastActivityAt ?? existing.last_activity_at ?? now,
      updatedAt: now,
      source: input.source,
      sourceRef: input.sourceRef
    });
  } else {
    db.prepare(
      `
      INSERT INTO tasks (
        title, source, stage, stage_order, priority, status, source_link, source_ref, source_thread_ref, jira_status,
        jira_estimate_seconds,
        jira_subtask_estimate_seconds, jira_planning_subtasks_json,
        ignored, deferred_until, reminder_state, last_reminded_at, estimated_effort_bucket,
        priority_score, priority_explanation, task_age_days, carry_forward_count, completed_at,
        last_activity_at, manual_override_flags, decision_state, decision_confidence, decision_reason,
        decision_reason_tags, personalization_version, selection_reason, priority_reason, score_breakdown_json,
        history_signals_json, last_changed_by, last_changed_at, was_user_overridden, restored_at, rejected_at,
        created_at, updated_at
      ) VALUES (
        @title, @source, @stage, @stageOrder, @priority, @status, @sourceLink, @sourceRef, @sourceThreadRef, @jiraStatus,
        @jiraEstimateSeconds,
        @jiraSubtaskEstimateSeconds, @jiraPlanningSubtasksJson,
        0, NULL, NULL, NULL, @estimatedEffortBucket, NULL, @priorityExplanation, 0, 0, NULL, @lastActivityAt, '[]', @decisionState, @decisionConfidence,
        @decisionReason, @decisionReasonTags, @personalizationVersion, @selectionReason, @priorityReason, @scoreBreakdownJson,
        @historySignalsJson, @lastChangedBy, @lastChangedAt, 0, @restoredAt, @rejectedAt, @createdAt, @updatedAt
      )
      `
    ).run({
      title: input.title,
      source: input.source,
      stage: input.stage ?? "Later",
      stageOrder: input.stageOrder ?? 0,
      priority: input.priority,
      status: input.status ?? "Not Started",
      sourceLink: input.sourceLink ?? null,
      sourceRef: input.sourceRef ?? null,
      sourceThreadRef: input.sourceThreadRef ?? null,
      jiraStatus: input.jiraStatus ?? null,
      estimatedEffortBucket: input.estimatedEffortBucket ?? null,
      priorityExplanation: input.priorityExplanation ?? null,
      jiraEstimateSeconds: input.jiraEstimateSeconds ?? null,
      jiraSubtaskEstimateSeconds: input.jiraSubtaskEstimateSeconds ?? null,
      jiraPlanningSubtasksJson: JSON.stringify(input.jiraPlanningSubtasks ?? []),
      decisionState: input.decisionState ?? "accepted",
      decisionConfidence: input.decisionConfidence ?? null,
      decisionReason: input.decisionReason ?? null,
      decisionReasonTags: JSON.stringify(input.decisionReasonTags ?? []),
      personalizationVersion: input.personalizationVersion ?? null,
      selectionReason: input.selectionReason ?? null,
      priorityReason: input.priorityReason ?? null,
      scoreBreakdownJson: JSON.stringify(input.scoreBreakdown ?? []),
      historySignalsJson: JSON.stringify(input.historySignals ?? []),
      lastChangedBy: input.lastChangedBy ?? "system",
      lastChangedAt: input.lastChangedAt ?? now,
      restoredAt: input.restoredAt ?? null,
      rejectedAt: input.rejectedAt ?? null,
      lastActivityAt: input.lastActivityAt ?? now,
      createdAt: now,
      updatedAt: now
    });
  }
}

export function createManualTask(input: {
  title: string;
  stage?: TaskStage;
  stageOrder?: number;
  priority?: TaskPriority;
  status?: TaskStatus;
}) {
  const now = new Date().toISOString();
  const status = input.status ?? "Not Started";
  const result = db
    .prepare(
      `
      INSERT INTO tasks (
        title, source, stage, stage_order, priority, status, ignored, deferred_until, reminder_state, last_reminded_at,
        estimated_effort_bucket, priority_score, priority_explanation, task_age_days, carry_forward_count,
        completed_at, last_activity_at, manual_override_flags, selection_reason, priority_reason,
        score_breakdown_json, history_signals_json, last_changed_by, last_changed_at, created_at, updated_at
      ) VALUES (
        @title, 'Manual', @stage, @stageOrder, @priority, @status, 0, NULL, NULL, NULL,
        NULL, NULL, NULL, 0, 0, @completedAt, @lastActivityAt, '["priority","status"]', @selectionReason, @priorityReason,
        '[]', '[]', 'user', @lastChangedAt, @createdAt, @updatedAt
      )
      `
    )
    .run({
      title: input.title,
      stage: input.stage ?? "Later",
      stageOrder: input.stageOrder ?? 0,
      priority: input.priority ?? "Medium",
      status,
      completedAt: status === "Completed" ? now : null,
      lastActivityAt: now,
      selectionReason: "Created manually by you.",
      priorityReason: input.priority ? `You set this as ${input.priority.toLowerCase()} priority.` : "Manual task defaults to medium priority.",
      lastChangedAt: now,
      createdAt: now,
      updatedAt: now
    });

  return db
    .prepare("SELECT * FROM tasks WHERE id = ?")
    .get(result.lastInsertRowid) as Record<string, unknown>;
}

export function updateTask(
  id: number,
  patch: Partial<Pick<Task, "title" | "stage" | "stageOrder" | "priority" | "status" | "deferredUntil" | "sourceLink" | "jiraStatus">> & {
    manualOverrideFlags?: string[];
    priorityScore?: number | null;
    priorityExplanation?: string | null;
    estimatedEffortBucket?: TaskEffortBucket | null;
    jiraEstimateSeconds?: number | null;
    jiraSubtaskEstimateSeconds?: number | null;
    jiraPlanningSubtasks?: Array<{
      key: string;
      title: string;
      status: string | null;
      estimateSeconds: number | null;
    }>;
    taskAgeDays?: number;
    carryForwardCount?: number;
    reminderState?: ReminderStatus | null;
    lastRemindedAt?: string | null;
    lastActivityAt?: string | null;
    decisionState?: TaskDecisionState | null;
    decisionConfidence?: number | null;
    decisionReason?: string | null;
    decisionReasonTags?: ReasonTag[];
    personalizationVersion?: number | null;
    selectionReason?: string | null;
    priorityReason?: string | null;
    scoreBreakdown?: ScoreBreakdownItem[];
    historySignals?: string[];
    lastChangedBy?: string | null;
    lastChangedAt?: string | null;
    wasUserOverridden?: boolean;
    restoredAt?: string | null;
    rejectedAt?: string | null;
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
        stage = @stage,
        stage_order = @stageOrder,
        priority = @priority,
        status = @status,
        source_link = @sourceLink,
        jira_status = @jiraStatus,
        deferred_until = @deferredUntil,
        estimated_effort_bucket = @estimatedEffortBucket,
        jira_estimate_seconds = @jiraEstimateSeconds,
        jira_subtask_estimate_seconds = @jiraSubtaskEstimateSeconds,
        jira_planning_subtasks_json = @jiraPlanningSubtasksJson,
        priority_score = @priorityScore,
        priority_explanation = @priorityExplanation,
        task_age_days = @taskAgeDays,
        carry_forward_count = @carryForwardCount,
        reminder_state = @reminderState,
        last_reminded_at = @lastRemindedAt,
        completed_at = @completedAt,
        last_activity_at = @lastActivityAt,
        manual_override_flags = @manualOverrideFlags,
        decision_state = @decisionState,
        decision_confidence = @decisionConfidence,
        decision_reason = @decisionReason,
        decision_reason_tags = @decisionReasonTags,
        personalization_version = @personalizationVersion,
        selection_reason = @selectionReason,
        priority_reason = @priorityReason,
        score_breakdown_json = @scoreBreakdownJson,
        history_signals_json = @historySignalsJson,
        last_changed_by = @lastChangedBy,
        last_changed_at = @lastChangedAt,
        was_user_overridden = @wasUserOverridden,
        restored_at = @restoredAt,
        rejected_at = @rejectedAt,
        updated_at = @updatedAt
    WHERE id = @id
    `
  ).run({
    id,
    title: patch.title ?? existing.title,
    stage: patch.stage ?? existing.stage ?? "Later",
    stageOrder: patch.stageOrder ?? Number(existing.stage_order ?? 0),
    priority: patch.priority ?? existing.priority,
    status: nextStatus,
    sourceLink: patch.sourceLink ?? existing.source_link ?? null,
    jiraStatus: patch.jiraStatus ?? existing.jira_status ?? null,
    deferredUntil: patch.deferredUntil ?? existing.deferred_until,
    estimatedEffortBucket: patch.estimatedEffortBucket ?? existing.estimated_effort_bucket ?? null,
    jiraEstimateSeconds: patch.jiraEstimateSeconds ?? existing.jira_estimate_seconds ?? null,
    jiraSubtaskEstimateSeconds: patch.jiraSubtaskEstimateSeconds ?? existing.jira_subtask_estimate_seconds ?? null,
    jiraPlanningSubtasksJson: JSON.stringify(
      patch.jiraPlanningSubtasks ??
        (() => {
          try {
            return JSON.parse(String(existing.jira_planning_subtasks_json ?? "[]"));
          } catch {
            return [];
          }
        })()
    ),
    priorityScore: patch.priorityScore ?? existing.priority_score ?? null,
    priorityExplanation: patch.priorityExplanation ?? existing.priority_explanation ?? null,
    taskAgeDays: patch.taskAgeDays ?? existing.task_age_days ?? 0,
    carryForwardCount: patch.carryForwardCount ?? existing.carry_forward_count ?? 0,
    reminderState: patch.reminderState ?? existing.reminder_state ?? null,
    lastRemindedAt: patch.lastRemindedAt ?? existing.last_reminded_at ?? null,
    completedAt: nextStatus === "Completed" ? (existing.completed_at ?? updatedAt) : null,
    lastActivityAt: patch.lastActivityAt ?? updatedAt,
    manualOverrideFlags: JSON.stringify(overrideFlags),
    decisionState: patch.decisionState ?? existing.decision_state ?? null,
    decisionConfidence: patch.decisionConfidence ?? existing.decision_confidence ?? null,
    decisionReason: patch.decisionReason ?? existing.decision_reason ?? null,
    decisionReasonTags: JSON.stringify(patch.decisionReasonTags ?? parseJsonArray(existing.decision_reason_tags)),
    personalizationVersion: patch.personalizationVersion ?? existing.personalization_version ?? null,
    selectionReason: patch.selectionReason ?? existing.selection_reason ?? null,
    priorityReason: patch.priorityReason ?? existing.priority_reason ?? null,
    scoreBreakdownJson: JSON.stringify(patch.scoreBreakdown ?? parseJsonValue(existing.score_breakdown_json, [])),
    historySignalsJson: JSON.stringify(patch.historySignals ?? parseJsonArray(existing.history_signals_json)),
    lastChangedBy: patch.lastChangedBy ?? existing.last_changed_by ?? null,
    lastChangedAt: patch.lastChangedAt ?? existing.last_changed_at ?? updatedAt,
    wasUserOverridden: Number(
      patch.wasUserOverridden ?? (Number(existing.was_user_overridden ?? 0) === 1)
    ),
    restoredAt: patch.restoredAt ?? existing.restored_at ?? null,
    rejectedAt: patch.rejectedAt ?? existing.rejected_at ?? null,
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
    const existingAttendance = new Map<string, "attending" | "unattending">(
      (db.prepare("SELECT external_id, attendance_status FROM meetings WHERE external_id IS NOT NULL").all() as Array<{
        external_id: string;
        attendance_status: string | null;
      }>).map((row) => [row.external_id, row.attendance_status === "unattending" ? "unattending" : "attending"])
    );
    db.prepare("DELETE FROM meetings").run();
    const insert = db.prepare(
      `
      INSERT INTO meetings (
        external_id, title, start_time, end_time, time_zone, duration_minutes, meeting_link,
        meeting_link_type, is_cancelled, attendance_status, created_at
      ) VALUES (
        @externalId, @title, @startTime, @endTime, @timeZone, @durationMinutes, @meetingLink,
        @meetingLinkType, @isCancelled, @attendanceStatus, @createdAt
      )
      `
    );
    for (const meeting of meetings) {
      insert.run({
        ...meeting,
        isCancelled: meeting.isCancelled ? 1 : 0,
        attendanceStatus: meeting.externalId ? existingAttendance.get(meeting.externalId) ?? meeting.attendanceStatus ?? "attending" : meeting.attendanceStatus ?? "attending",
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
      | "workdayStartLocal"
      | "workdayEndLocal"
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
        workday_start_local = @workdayStartLocal,
        workday_end_local = @workdayEndLocal,
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
    workdayStartLocal: patch.workdayStartLocal ?? existing.workdayStartLocal,
    workdayEndLocal: patch.workdayEndLocal ?? existing.workdayEndLocal,
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

export function upsertDailyPlanSnapshot(input: {
  dayKey: string;
  weekday: number;
  baseWorkdayMinutes: number;
  adaptedTaskCapacityMinutes: number;
  remainingTaskCapacityMinutes: number;
  meetingMinutes: number;
  plannedTaskMinutes: number;
  completedTaskMinutes: number;
  remainingTaskMinutes: number;
  spilloverTaskCount: number;
  freeMinutes: number;
  focusFactor: number;
  completionRate: number;
  plannedTaskIds: number[];
  summaryJson: string;
  blocksJson: string;
}) {
  const now = new Date().toISOString();
  db.prepare(
    `
    INSERT INTO daily_plan_snapshots (
      day_key, weekday, base_workday_minutes, adapted_task_capacity_minutes, remaining_task_capacity_minutes,
      meeting_minutes, planned_task_minutes, completed_task_minutes, remaining_task_minutes, spillover_task_count,
      free_minutes, focus_factor, completion_rate, planned_task_ids_json, summary_json, blocks_json, created_at, updated_at
    ) VALUES (
      @dayKey, @weekday, @baseWorkdayMinutes, @adaptedTaskCapacityMinutes, @remainingTaskCapacityMinutes,
      @meetingMinutes, @plannedTaskMinutes, @completedTaskMinutes, @remainingTaskMinutes, @spilloverTaskCount,
      @freeMinutes, @focusFactor, @completionRate, @plannedTaskIdsJson, @summaryJson, @blocksJson, @createdAt, @updatedAt
    )
    ON CONFLICT(day_key) DO UPDATE SET
      weekday = excluded.weekday,
      base_workday_minutes = excluded.base_workday_minutes,
      adapted_task_capacity_minutes = excluded.adapted_task_capacity_minutes,
      remaining_task_capacity_minutes = excluded.remaining_task_capacity_minutes,
      meeting_minutes = excluded.meeting_minutes,
      planned_task_minutes = excluded.planned_task_minutes,
      completed_task_minutes = excluded.completed_task_minutes,
      remaining_task_minutes = excluded.remaining_task_minutes,
      spillover_task_count = excluded.spillover_task_count,
      free_minutes = excluded.free_minutes,
      focus_factor = excluded.focus_factor,
      completion_rate = excluded.completion_rate,
      planned_task_ids_json = excluded.planned_task_ids_json,
      summary_json = excluded.summary_json,
      blocks_json = excluded.blocks_json,
      updated_at = excluded.updated_at
    `
  ).run({
    dayKey: input.dayKey,
    weekday: input.weekday,
    baseWorkdayMinutes: input.baseWorkdayMinutes,
    adaptedTaskCapacityMinutes: input.adaptedTaskCapacityMinutes,
    remainingTaskCapacityMinutes: input.remainingTaskCapacityMinutes,
    meetingMinutes: input.meetingMinutes,
    plannedTaskMinutes: input.plannedTaskMinutes,
    completedTaskMinutes: input.completedTaskMinutes,
    remainingTaskMinutes: input.remainingTaskMinutes,
    spilloverTaskCount: input.spilloverTaskCount,
    freeMinutes: input.freeMinutes,
    focusFactor: input.focusFactor,
    completionRate: input.completionRate,
    plannedTaskIdsJson: JSON.stringify(input.plannedTaskIds),
    summaryJson: input.summaryJson,
    blocksJson: input.blocksJson,
    createdAt: now,
    updatedAt: now
  });
}

export function listRecentDailyPlanSnapshots(limit = 28) {
  return db
    .prepare("SELECT * FROM daily_plan_snapshots ORDER BY day_key DESC LIMIT ?")
    .all(limit) as Array<Record<string, unknown>>;
}

export function getDailyPlanSnapshot(dayKey: string) {
  return db
    .prepare("SELECT * FROM daily_plan_snapshots WHERE day_key = ? LIMIT 1")
    .get(dayKey) as Record<string, unknown> | undefined;
}

export function getRejectedTaskCount() {
  const row = db
    .prepare("SELECT COUNT(*) as count FROM rejected_tasks WHERE decision_state IN ('rejected', 'uncertain')")
    .get() as { count: number };
  return Number(row.count ?? 0);
}

export function getIgnoredRejectedTaskCount() {
  const row = db
    .prepare("SELECT COUNT(*) as count FROM rejected_tasks WHERE decision_state = 'ignored'")
    .get() as { count: number };
  return Number(row.count ?? 0);
}

export function clearSoftDevelopmentData() {
  const transaction = db.transaction(() => {
    db.prepare("DELETE FROM tasks").run();
    db.prepare("DELETE FROM meetings").run();
    db.prepare("DELETE FROM reminders").run();
    db.prepare("DELETE FROM rejected_tasks").run();
    db.prepare("DELETE FROM task_decision_log").run();
    db.prepare("DELETE FROM preference_memory_snapshots").run();
    db.prepare("DELETE FROM generation_runs").run();
    db.prepare("DELETE FROM daily_plan_snapshots").run();
    db.prepare("DELETE FROM audit_events").run();
    db.prepare("DELETE FROM planner_run_details").run();
    db.prepare("DELETE FROM task_state_events").run();
    db.prepare("DELETE FROM home_schedule_entries").run();
    db.prepare("DELETE FROM home_meeting_overrides").run();
    db.prepare("DELETE FROM home_schedule_days").run();
  });

  transaction();
}

export function clearHardDevelopmentData() {
  db.exec(`
    DELETE FROM tasks;
    DELETE FROM meetings;
    DELETE FROM integration_connections;
    DELETE FROM sync_state;
    DELETE FROM reminders;
    DELETE FROM automation_settings;
    DELETE FROM generation_runs;
    DELETE FROM daily_plan_snapshots;
    DELETE FROM user_priority_profile;
    DELETE FROM rejected_tasks;
    DELETE FROM task_decision_log;
    DELETE FROM preference_memory_snapshots;
    DELETE FROM audit_events;
    DELETE FROM planner_run_details;
    DELETE FROM task_state_events;
    DELETE FROM home_schedule_entries;
    DELETE FROM home_meeting_overrides;
    DELETE FROM home_schedule_days;
  `);

  db.prepare(
    `
    INSERT INTO automation_settings (id)
    VALUES (1)
    ON CONFLICT(id) DO NOTHING
    `
  ).run();

  db.prepare(
    `
    INSERT INTO user_priority_profile (id, updated_at)
    VALUES (1, ?)
    ON CONFLICT(id) DO NOTHING
    `
  ).run(new Date().toISOString());
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
