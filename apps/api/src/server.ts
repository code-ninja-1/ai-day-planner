import cors from "cors";
import express from "express";
import { z } from "zod";
import { acquireGraphTokenOnBehalfOf, getOptionalMicrosoftSession } from "./auth/microsoftAuth.js";
import {
  clearRejectedTasksBySourceThread,
  recordTaskStateEvent,
  createManualTask,
  deleteIntegrationConnection,
  deleteTask,
  getMeetingById,
  listAuditEvents,
  getAutomationSettings,
  getIgnoredRejectedTaskCount,
  getLatestPreferenceMemorySnapshot,
  getReminderById,
  getRejectedTaskById,
  getRejectedTaskBySource,
  getRejectedTaskBySourceThread,
  getTaskById,
  getTaskBySource,
  getTaskBySourceThread,
  getSyncState,
  getUserPriorityProfile,
  listIntegrationConnections,
  listIgnoredRejectedTasks,
  listMeetings,
  listPlannerRunDetails,
  listReminderItems,
  listRejectedTasks,
  listTasks,
  listDeferredTasks,
  logTaskDecisionEvent,
  normalizeTask,
  saveAutomationSettings,
  saveIntegrationConnection,
  saveUserPriorityProfile,
  updateTask,
  updateRejectedTask,
  updateReminder,
  upsertRejectedTask,
  upsertTask,
  groupTasksByPriority
} from "./db.js";
import { env } from "./env.js";
import {
  buildJiraIssueBrowseUrl,
  fetchJiraIssueDetail,
  fetchJiraIssueForSync,
  fetchJiraIssuePlanningContext,
  mapJiraWorkflowStatus,
  normalizeJiraBaseUrl,
  transitionJiraIssue,
  transitionJiraIssueToTaskStatus,
  validateJiraCredentials
} from "./providers/jira.js";
import {
  exchangeMicrosoftCode,
  fetchMeetingDetailWithAccessToken,
  fetchEmailDetailWithAccessToken,
  fetchMicrosoftProfileWithAccessToken,
  getMicrosoftAuthUrl
} from "./providers/microsoft.js";
import { sendMailWithAccessToken } from "./providers/microsoft.js";
import {
  generatePlan,
  getDeferredTasksPayload,
  getReminderCenterPayload,
  getTodaySnapshot,
  syncMeetingsOnly,
  syncTasksOnly
} from "./services/planService.js";
import {
  getDiagnosticsPayload,
  getInsightsHistoryDayPayload,
  getInsightsHistoryPayload,
  getInsightsOverviewPayload,
  getInsightsTodayPayload,
  getTaskInsightsPayload
} from "./services/insights.js";
import { createCorrelationId, logEvent } from "./services/logger.js";
import { analyzeFeedbackReason, defaultPriorityProfile, synthesizePriorityProfile } from "./services/personalization.js";
import { generateEmailReplyDraft, generateMeetingPrep } from "./services/assistant.js";
import { scheduleAutomation } from "./services/scheduler.js";

const app = express();
app.use(cors({ origin: env.appOrigin, credentials: true }));
app.use(express.json());
app.use((req, res, next) => {
  const requestId = createCorrelationId();
  const startedAt = Date.now();
  (res.locals as Record<string, unknown>).requestId = requestId;
  logEvent({
    eventType: "http.request",
    requestId,
    status: "started",
    source: "server",
    message: `${req.method} ${req.path} started.`,
    metadata: {
      method: req.method,
      path: req.path,
      query: req.query
    }
  });

  res.on("finish", () => {
    logEvent({
      eventType: "http.request",
      requestId,
      status: res.statusCode >= 500 ? "failure" : "success",
      source: "server",
      message: `${req.method} ${req.path} completed.`,
      metadata: {
        method: req.method,
        path: req.path,
        statusCode: res.statusCode,
        durationMs: Date.now() - startedAt
      }
    });
  });

  next();
});

const taskCreateSchema = z.object({
  title: z.string().min(1),
  priority: z.enum(["High", "Medium", "Low"]).optional(),
  status: z.enum(["Not Started", "In Progress", "Completed"]).optional()
});

const taskUpdateSchema = taskCreateSchema
  .partial()
  .extend({ deferredUntil: z.string().datetime().nullable().optional() });

const reminderUpdateSchema = z.object({
  status: z.enum(["active", "dismissed", "resolved"]).optional(),
  reason: z.string().min(1).optional(),
  scheduledFor: z.string().datetime().nullable().optional(),
  throttleUntil: z.string().datetime().nullable().optional()
});

const automationSettingsSchema = z.object({
  scheduleEnabled: z.boolean().optional(),
  scheduleTimeLocal: z.string().regex(/^\d{2}:\d{2}$/).optional(),
  scheduleTimezone: z.string().min(1).optional(),
  remindersEnabled: z.boolean().optional(),
  reminderCadenceHours: z.number().int().min(1).max(72).optional(),
  desktopNotificationsEnabled: z.boolean().optional()
});

const personalizationProfileSchema = z.object({
  personalizationEnabled: z.boolean().optional(),
  roleFocus: z.string().nullable().optional(),
  prioritizationPrompt: z.string().nullable().optional(),
  importantWork: z.array(z.string()).optional(),
  noiseWork: z.array(z.string()).optional(),
  mustNotMiss: z.array(z.string()).optional(),
  importantSources: z.array(z.string()).optional(),
  importantPeople: z.array(z.string()).optional(),
  importantProjects: z.array(z.string()).optional(),
  positiveReasonTags: z.array(z.string()).optional(),
  negativeReasonTags: z.array(z.string()).optional(),
  filteringStyle: z.enum(["conservative", "balanced", "aggressive"]).optional(),
  priorityBias: z.enum(["focus", "balanced", "coverage"]).optional()
});

const calibrationSchema = z.object({
  roleFocus: z.string().default(""),
  prioritizationPrompt: z.string().default(""),
  importantWork: z.array(z.string()).default([]),
  noiseWork: z.array(z.string()).default([]),
  mustNotMiss: z.array(z.string()).default([]),
  importantPeople: z.array(z.string()).default([]),
  importantProjects: z.array(z.string()).default([]),
  filteringStyle: z.enum(["conservative", "balanced", "aggressive"]),
  priorityBias: z.enum(["focus", "balanced", "coverage"]),
  exampleRankings: z.array(
    z.object({
      title: z.string(),
      source: z.enum(["Email", "Jira", "Manual"]),
      decision: z.enum(["show_today", "keep_low", "reject_noise"])
    })
  )
});

const taskFeedbackSchema = z.object({
  action: z.enum([
    "reject",
    "restore",
    "priority_changed",
    "status_changed",
    "deferred",
    "completed",
    "always_ignore_similar",
    "should_have_been_included"
  ]),
  beforePriority: z.enum(["High", "Medium", "Low"]).nullable().optional(),
  afterPriority: z.enum(["High", "Medium", "Low"]).nullable().optional(),
  context: z.string().nullable().optional()
});

const rejectedTaskPatchSchema = z.object({
  action: z.enum(["always_ignore_exact", "always_ignore_similar", "should_have_been_included", "keep_rejected"])
});

const clientEventSchema = z.object({
  eventType: z.string().min(1),
  level: z.enum(["debug", "info", "warn", "error"]).optional(),
  message: z.string().min(1),
  entityType: z.string().nullable().optional(),
  entityId: z.string().nullable().optional(),
  status: z.enum(["started", "success", "failure", "updated", "skipped", "info"]).optional(),
  metadata: z.unknown().optional()
});

const emailReplyDraftSchema = z.object({
  userIntent: z.string().nullable().optional()
});

const emailReplySendSchema = z.object({
  to: z.array(z.string().email()).min(1),
  cc: z.array(z.string().email()).default([]),
  subject: z.string().min(1),
  body: z.string().min(1)
});

const meetingPrepSchema = z.object({
  userNotes: z.string().nullable().optional()
});

const jiraTransitionSchema = z.object({
  issueKey: z.string().min(1).optional(),
  transitionId: z.string().min(1),
  parentTaskId: z.number().int().positive().optional()
});

async function captureTaskFeedback(input: {
  taskId?: number | null;
  source: "Email" | "Jira" | "Manual" | "Calibration";
  sourceRef?: string | null;
  sourceThreadRef?: string | null;
  action:
    | "system_evaluated"
    | "reject"
    | "restore"
    | "priority_changed"
    | "status_changed"
    | "deferred"
    | "completed"
    | "always_ignore_similar"
    | "should_have_been_included";
  title: string;
  beforePriority?: "High" | "Medium" | "Low" | null;
  afterPriority?: "High" | "Medium" | "Low" | null;
  decisionReason?: string | null;
  decisionReasonTags?: string[];
  context?: string | null;
}) {
  const analysis = await analyzeFeedbackReason({
    action: input.action,
    taskTitle: input.title,
    source: input.source,
    beforePriority: input.beforePriority,
    afterPriority: input.afterPriority,
    decisionReason: input.decisionReason,
    decisionTags: (input.decisionReasonTags ?? []) as never,
    context: input.context
  });

  logTaskDecisionEvent({
    taskId: input.taskId ?? null,
    source: input.source,
    sourceRef: input.sourceRef ?? null,
    sourceThreadRef: input.sourceThreadRef ?? null,
    action: input.action,
    beforePriority: input.beforePriority ?? null,
    afterPriority: input.afterPriority ?? null,
    decisionReason: input.decisionReason ?? null,
    decisionReasonTags: (input.decisionReasonTags ?? []) as never,
    feedbackPayloadJson: JSON.stringify({ title: input.title, context: input.context }),
    inferredReason: analysis.likelyReason,
    inferredReasonTag: analysis.reasonTag,
    preferencePolarity: analysis.positiveOrNegativePreference
  });
}

function getRequestId(res: express.Response) {
  return ((res.locals as Record<string, unknown>).requestId as string | undefined) ?? null;
}

function recordUserTaskMutation(input: {
  taskId?: number | null;
  source: "Email" | "Jira" | "Manual" | "Calibration";
  sourceRef?: string | null;
  sourceThreadRef?: string | null;
  eventType: string;
  reason?: string | null;
  before?: unknown;
  after?: unknown;
}) {
  recordTaskStateEvent({
    taskId: input.taskId ?? null,
    source: input.source,
    sourceRef: input.sourceRef ?? null,
    sourceThreadRef: input.sourceThreadRef ?? null,
    eventType: input.eventType,
    actor: "user",
    reason: input.reason ?? null,
    beforeJson: input.before ? JSON.stringify(input.before) : null,
    afterJson: input.after ? JSON.stringify(input.after) : null
  });
}

function extractIssueKeys(value: string) {
  return [...new Set((value.toUpperCase().match(/\b[A-Z][A-Z0-9]+-\d+\b/g) ?? []).map((key) => key.trim()))];
}

async function refreshJiraTaskFromSource(taskId: number, issueKey: string) {
  const issue = await fetchJiraIssueForSync(issueKey);
  let planningContext: Awaited<ReturnType<typeof fetchJiraIssuePlanningContext>> | null = null;
  try {
    planningContext = await fetchJiraIssuePlanningContext(issueKey);
  } catch {
    planningContext = null;
  }

  const jiraConnection = listIntegrationConnections().find((entry) => entry.provider === "jira");
  const jiraConfig =
    typeof jiraConnection?.configJson === "string"
      ? (JSON.parse(jiraConnection.configJson) as { baseUrl?: string })
      : null;
  const row = updateTask(taskId, {
    title: `${issue.key} ${issue.fields.summary}`,
    status: mapJiraWorkflowStatus(issue.fields.status?.name),
    sourceLink: jiraConfig?.baseUrl ? buildJiraIssueBrowseUrl(jiraConfig.baseUrl, issue.key) : issue.self,
    jiraStatus: issue.fields.status?.name ?? null,
    jiraEstimateSeconds:
      issue.fields.timeoriginalestimate ?? issue.fields.timetracking?.originalEstimateSeconds ?? null,
    jiraSubtaskEstimateSeconds: planningContext?.openSubtaskEstimateSeconds ?? null,
    jiraPlanningSubtasks: planningContext?.subtasks ?? [],
    lastActivityAt: issue.fields.updated ?? new Date().toISOString(),
    lastChangedBy: "user",
    lastChangedAt: new Date().toISOString(),
    wasUserOverridden: true
  });

  if (!row) {
    throw new Error("Failed to refresh Jira task after transition");
  }
  return normalizeTask(row as Record<string, unknown>);
}

app.get("/api/health", (_req, res) => {
  res.json({ ok: true });
});

app.get("/api/today", (_req, res) => {
  res.json(getTodaySnapshot());
});

app.get("/api/insights/overview", (_req, res) => {
  res.json(getInsightsOverviewPayload());
});

app.get("/api/insights/today", (_req, res) => {
  res.json(getInsightsTodayPayload());
});

app.get("/api/insights/history", (req, res) => {
  const limit = Number(req.query.limit ?? 30);
  res.json(getInsightsHistoryPayload(Number.isFinite(limit) ? limit : 30));
});

app.get("/api/insights/history/:dayKey", (req, res) => {
  const payload = getInsightsHistoryDayPayload(req.params.dayKey);
  if (!payload) {
    return res.status(404).json({ message: "History not found for this day" });
  }
  return res.json(payload);
});

app.get("/api/insights/tasks/:taskId", (req, res) => {
  const payload = getTaskInsightsPayload(Number(req.params.taskId));
  if (!payload) {
    return res.status(404).json({ message: "Task insights not found" });
  }
  return res.json(payload);
});

app.get("/api/debug/runs", (_req, res) => {
  res.json({ runs: listPlannerRunDetails(30), diagnostics: getDiagnosticsPayload() });
});

app.get("/api/debug/logs", (req, res) => {
  const limit = Number(req.query.limit ?? 200);
  res.json({ logs: listAuditEvents(Number.isFinite(limit) ? limit : 200) });
});

app.post("/api/debug/client-events", (req, res) => {
  const parsed = clientEventSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid client event payload" });
  }
  logEvent({
    level: parsed.data.level ?? "info",
    eventType: parsed.data.eventType,
    requestId: getRequestId(res),
    entityType: parsed.data.entityType ?? null,
    entityId: parsed.data.entityId ?? null,
    status: parsed.data.status ?? "info",
    source: "client",
    message: parsed.data.message,
    metadata: parsed.data.metadata
  });
  return res.status(201).json({ ok: true });
});

app.post("/api/plan/generate", async (req, res) => {
  try {
    const runId = createCorrelationId();
    let microsoftGraphAccessToken: string | null = null;
    let microsoftWarning: string | null = null;
    const preferredTimeZone =
      typeof req.body?.timeZone === "string" && req.body.timeZone.trim() ? req.body.timeZone.trim() : null;

    try {
      const session = await getOptionalMicrosoftSession(req);
      if (session) {
        microsoftGraphAccessToken = await acquireGraphTokenOnBehalfOf(session);
        saveIntegrationConnection({
          provider: "microsoft",
          status: "connected",
          accountLabel: session.accountLabel ?? session.displayName,
          configJson: JSON.stringify({ mode: "msal-obo", oid: session.oid }),
          accessToken: null,
          refreshToken: null,
          expiresAt: null,
          errorMessage: null
        });
      } else {
        microsoftWarning = "Microsoft is not connected for this browser session.";
      }
    } catch (error) {
      microsoftWarning =
        error instanceof Error
          ? `Microsoft session is unavailable: ${error.message}`
          : "Microsoft session is unavailable.";
    }

    const payload = await generatePlan({
      microsoftGraphAccessToken,
      microsoftWarning,
      preferredTimeZone,
      runId
    }, "manual");
    res.json(payload);
  } catch (error) {
    res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to generate plan"
    });
  }
});

app.post("/api/plan/generate-now", async (req, res) => {
  try {
    const runId = createCorrelationId();
    let microsoftGraphAccessToken: string | null = null;
    let microsoftWarning: string | null = null;
    const preferredTimeZone =
      typeof req.body?.timeZone === "string" && req.body.timeZone.trim() ? req.body.timeZone.trim() : null;

    try {
      const session = await getOptionalMicrosoftSession(req);
      if (session) {
        microsoftGraphAccessToken = await acquireGraphTokenOnBehalfOf(session);
      } else {
        microsoftWarning = "Microsoft is not connected for this browser session.";
      }
    } catch (error) {
      microsoftWarning =
        error instanceof Error
          ? `Microsoft session is unavailable: ${error.message}`
          : "Microsoft session is unavailable.";
    }

    const payload = await generatePlan(
      {
        microsoftGraphAccessToken,
        microsoftWarning,
        preferredTimeZone,
        runId
      },
      "manual"
    );
    res.json(payload);
  } catch (error) {
    res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to generate plan"
    });
  }
});

app.post("/api/sync/meetings", async (req, res) => {
  try {
    const runId = createCorrelationId();
    let microsoftGraphAccessToken: string | null = null;
    let microsoftWarning: string | null = null;
    const preferredTimeZone =
      typeof req.body?.timeZone === "string" && req.body.timeZone.trim() ? req.body.timeZone.trim() : null;

    try {
      const session = await getOptionalMicrosoftSession(req);
      if (session) {
        microsoftGraphAccessToken = await acquireGraphTokenOnBehalfOf(session);
      } else {
        microsoftWarning = "Microsoft is not connected for this browser session.";
      }
    } catch (error) {
      microsoftWarning =
        error instanceof Error
          ? `Microsoft session is unavailable: ${error.message}`
          : "Microsoft session is unavailable.";
    }

    res.json(
      await syncMeetingsOnly({
        microsoftGraphAccessToken,
        microsoftWarning,
        preferredTimeZone,
        runId
      })
    );
  } catch (error) {
    res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to sync meetings"
    });
  }
});

app.post("/api/sync/tasks", async (req, res) => {
  try {
    const runId = createCorrelationId();
    let microsoftGraphAccessToken: string | null = null;
    let microsoftWarning: string | null = null;
    const preferredTimeZone =
      typeof req.body?.timeZone === "string" && req.body.timeZone.trim() ? req.body.timeZone.trim() : null;

    try {
      const session = await getOptionalMicrosoftSession(req);
      if (session) {
        microsoftGraphAccessToken = await acquireGraphTokenOnBehalfOf(session);
      } else {
        microsoftWarning = "Microsoft is not connected for this browser session.";
      }
    } catch (error) {
      microsoftWarning =
        error instanceof Error
          ? `Microsoft session is unavailable: ${error.message}`
          : "Microsoft session is unavailable.";
    }

    res.json(
      await syncTasksOnly({
        microsoftGraphAccessToken,
        microsoftWarning,
        preferredTimeZone,
        runId
      })
    );
  } catch (error) {
    res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to sync tasks"
    });
  }
});

app.get("/api/tasks", (req, res) => {
  const status = req.query.status as "Not Started" | "In Progress" | "Completed" | undefined;
  res.json({ tasks: listTasks(status) });
});

app.get("/api/tasks/deferred", (_req, res) => {
  res.json(getDeferredTasksPayload());
});

app.get("/api/tasks/rejected", (_req, res) => {
  const jiraKeys = new Set(
    listTasks(undefined, { includeDeferred: true })
      .filter((task) => task.source === "Jira")
      .flatMap((task) => [
        ...(task.sourceRef ? [String(task.sourceRef)] : []),
        ...task.jiraPlanningSubtasks.map((subtask) => subtask.key)
      ])
  );

  const allowVisibleRejectedTask = (task: { source: string; decisionState: string; title: string; candidatePayloadJson: string | null }) => {
    if (task.decisionState === "restored" || task.decisionState === "ignored") {
      return false;
    }

    if (task.source !== "Email") {
      return true;
    }

    const keys = extractIssueKeys(`${task.title} ${task.candidatePayloadJson ?? ""}`);
    return !keys.some((key) => jiraKeys.has(key));
  };

  const tasks = listRejectedTasks().filter(allowVisibleRejectedTask);
  const ignoredTasks = listIgnoredRejectedTasks().filter((task) => {
    if (task.source !== "Email") {
      return true;
    }

    const keys = extractIssueKeys(`${task.title} ${task.candidatePayloadJson ?? ""}`);
    return !keys.some((key) => jiraKeys.has(key));
  });

  res.json({ tasks, ignoredTasks });
});

app.get("/api/tasks/:id/details", async (req, res) => {
  const task = getTaskById(Number(req.params.id));
  if (!task) {
    return res.status(404).json({ message: "Task not found" });
  }

  if (task.source === "Manual") {
    return res.status(400).json({ message: "Manual tasks do not have source details" });
  }

  try {
    if (task.source === "Jira") {
      if (!task.sourceRef) {
        return res.status(400).json({ message: "Jira task is missing source reference" });
      }
      return res.json({ detail: await fetchJiraIssueDetail(task.sourceRef) });
    }

    if (!task.sourceRef) {
      return res.status(400).json({ message: "Email task is missing source reference" });
    }

    const session = await getOptionalMicrosoftSession(req);
    if (!session) {
      return res.status(401).json({ message: "Microsoft is not connected for this browser session." });
    }

    const graphToken = await acquireGraphTokenOnBehalfOf(session);
    return res.json({
      detail: await fetchEmailDetailWithAccessToken(task.sourceRef, task.sourceThreadRef, graphToken)
    });
  } catch (error) {
    return res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to fetch task details"
    });
  }
});

app.post("/api/tasks/:id/email-reply/draft", async (req, res) => {
  const parsed = emailReplyDraftSchema.safeParse(req.body ?? {});
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid email draft payload" });
  }
  const task = getTaskById(Number(req.params.id));
  if (!task || task.source !== "Email" || !task.sourceRef) {
    return res.status(404).json({ message: "Email task not found" });
  }

  try {
    const session = await getOptionalMicrosoftSession(req);
    if (!session) {
      return res.status(401).json({ message: "Microsoft is not connected for this browser session." });
    }
    const graphToken = await acquireGraphTokenOnBehalfOf(session);
    const detail = await fetchEmailDetailWithAccessToken(task.sourceRef, task.sourceThreadRef, graphToken);
    const draft = await generateEmailReplyDraft(detail, parsed.data.userIntent ?? null);
    logEvent({
      eventType: "email.reply_draft",
      requestId: getRequestId(res),
      entityType: "task",
      entityId: String(task.id),
      provider: "microsoft",
      status: "success",
      message: "Generated email reply draft.",
      metadata: { sourceRef: task.sourceRef }
    });
    return res.json({ draft });
  } catch (error) {
    return res.status(500).json({ message: error instanceof Error ? error.message : "Failed to generate email draft" });
  }
});

app.post("/api/tasks/:id/email-reply/send", async (req, res) => {
  const parsed = emailReplySendSchema.safeParse(req.body ?? {});
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid email send payload" });
  }
  const task = getTaskById(Number(req.params.id));
  if (!task || task.source !== "Email") {
    return res.status(404).json({ message: "Email task not found" });
  }

  try {
    const session = await getOptionalMicrosoftSession(req);
    if (!session) {
      return res.status(401).json({ message: "Microsoft is not connected for this browser session." });
    }
    const graphToken = await acquireGraphTokenOnBehalfOf(session);
    await sendMailWithAccessToken(graphToken, parsed.data);
    logEvent({
      eventType: "email.reply_send",
      requestId: getRequestId(res),
      entityType: "task",
      entityId: String(task.id),
      provider: "microsoft",
      status: "success",
      message: "Sent email reply from planner.",
      metadata: { sourceRef: task.sourceRef, toCount: parsed.data.to.length, ccCount: parsed.data.cc.length }
    });
    return res.json({ ok: true });
  } catch (error) {
    return res.status(500).json({ message: error instanceof Error ? error.message : "Failed to send email reply" });
  }
});

app.post("/api/meetings/:id/prepare", async (req, res) => {
  const parsed = meetingPrepSchema.safeParse(req.body ?? {});
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid meeting prep payload" });
  }
  const meeting = getMeetingById(Number(req.params.id));
  if (!meeting?.externalId) {
    return res.status(404).json({ message: "Meeting not found" });
  }

  try {
    const session = await getOptionalMicrosoftSession(req);
    if (!session) {
      return res.status(401).json({ message: "Microsoft is not connected for this browser session." });
    }
    const graphToken = await acquireGraphTokenOnBehalfOf(session);
    const detail = await fetchMeetingDetailWithAccessToken(meeting.externalId, graphToken);
    const prep = await generateMeetingPrep({
      title: detail.subject ?? meeting.title,
      startTime: detail.start?.dateTime ?? meeting.startTime,
      endTime: detail.end?.dateTime ?? meeting.endTime,
      timeZone: detail.start?.timeZone ?? meeting.timeZone,
      description: detail.body?.content ?? detail.bodyPreview ?? "",
      organizer: detail.organizer?.emailAddress?.name ?? detail.organizer?.emailAddress?.address ?? null,
      attendees: (detail.attendees ?? [])
        .map((entry) => entry.emailAddress?.name ?? entry.emailAddress?.address ?? "")
        .filter(Boolean),
      userNotes: parsed.data.userNotes ?? null
    });
    logEvent({
      eventType: "meeting.prep_generate",
      requestId: getRequestId(res),
      entityType: "meeting",
      entityId: String(meeting.id),
      provider: "microsoft",
      status: "success",
      message: "Generated meeting preparation notes.",
      metadata: { externalId: meeting.externalId }
    });
    return res.json({ prep });
  } catch (error) {
    return res.status(500).json({ message: error instanceof Error ? error.message : "Failed to generate meeting prep" });
  }
});

app.post("/api/tasks", (req, res) => {
  const parsed = taskCreateSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid task payload" });
  }
  const row = createManualTask(parsed.data);
  const task = normalizeTask(row as Record<string, unknown>);
  recordUserTaskMutation({
    taskId: task.id,
    source: task.source,
    eventType: "create_manual",
    reason: "Created manual task.",
    after: task
  });
  logEvent({
    eventType: "task.create",
    requestId: getRequestId(res),
    entityType: "task",
    entityId: String(task.id),
    status: "success",
    message: "Manual task created.",
    metadata: { taskId: task.id, title: task.title }
  });
  return res.status(201).json({ task });
});

app.patch("/api/tasks/:id", (req, res) => {
  const parsed = taskUpdateSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid task payload" });
  }
  const existingTask = getTaskById(Number(req.params.id));
  if (existingTask?.source === "Jira" && parsed.data.status && parsed.data.status !== existingTask.status) {
    if (!existingTask.sourceRef) {
      return res.status(400).json({ message: "Jira task is missing its issue key" });
    }
    transitionJiraIssueToTaskStatus(existingTask.sourceRef ?? "", parsed.data.status, existingTask.jiraStatus)
      .then(async () => {
        const task = await refreshJiraTaskFromSource(existingTask.id, existingTask.sourceRef ?? "");
        recordUserTaskMutation({
          taskId: task.id,
          source: task.source,
          sourceRef: task.sourceRef ?? null,
          sourceThreadRef: task.sourceThreadRef ?? null,
          eventType: "task_updated",
          reason: "Jira task transitioned from planner.",
          before: existingTask,
          after: task
        });
        if (parsed.data.status !== existingTask.status) {
          await captureTaskFeedback({
            taskId: task.id,
            source: task.source,
            sourceRef: task.sourceRef ?? null,
            sourceThreadRef: task.sourceThreadRef ?? null,
            action: parsed.data.status === "Completed" ? "completed" : "status_changed",
            title: task.title,
            beforePriority: task.priority,
            afterPriority: task.priority,
            decisionReason: task.decisionReason,
            decisionReasonTags: task.decisionReasonTags
          });
        }
        logEvent({
          eventType: "task.update",
          requestId: getRequestId(res),
          entityType: "task",
          entityId: String(task.id),
          provider: "jira",
          status: "updated",
          message: "Jira task transitioned and refreshed.",
          metadata: {
            jiraIssueKey: task.sourceRef,
            targetStatus: parsed.data.status,
            jiraStatus: task.jiraStatus
          }
        });
        return res.json({ task });
      })
      .catch((error) => {
        res.status(400).json({ message: error instanceof Error ? error.message : "Failed to transition Jira task" });
      });
    return;
  }
  const row = updateTask(Number(req.params.id), {
    ...parsed.data,
    lastChangedBy: "user",
    lastChangedAt: new Date().toISOString(),
    wasUserOverridden: true
  });
  if (!row) {
    return res.status(404).json({ message: "Task not found" });
  }
  const task = normalizeTask(row as Record<string, unknown>);
  recordUserTaskMutation({
    taskId: task.id,
    source: task.source,
    sourceRef: task.sourceRef ?? null,
    sourceThreadRef: task.sourceThreadRef ?? null,
    eventType: "task_updated",
    reason: "Task edited by user.",
    before: existingTask,
    after: task
  });
  if (existingTask) {
    if (parsed.data.priority && parsed.data.priority !== existingTask.priority) {
      void captureTaskFeedback({
        taskId: task.id,
        source: task.source,
        sourceRef: task.sourceRef ?? null,
        sourceThreadRef: task.sourceThreadRef ?? null,
        action: "priority_changed",
        title: task.title,
        beforePriority: existingTask.priority,
        afterPriority: task.priority,
        decisionReason: task.decisionReason,
        decisionReasonTags: task.decisionReasonTags
      });
    }
    if (parsed.data.status && parsed.data.status !== existingTask.status) {
      void captureTaskFeedback({
        taskId: task.id,
        source: task.source,
        sourceRef: task.sourceRef ?? null,
        sourceThreadRef: task.sourceThreadRef ?? null,
        action: parsed.data.status === "Completed" ? "completed" : "status_changed",
        title: task.title,
        beforePriority: task.priority,
        afterPriority: task.priority,
        decisionReason: task.decisionReason,
        decisionReasonTags: task.decisionReasonTags
      });
    }
  }
  logEvent({
    eventType: "task.update",
    requestId: getRequestId(res),
    entityType: "task",
    entityId: String(task.id),
    status: "updated",
    message: "Task updated.",
    metadata: {
      taskId: task.id,
      changedPriority: parsed.data.priority ?? null,
      changedStatus: parsed.data.status ?? null
    }
  });
  return res.json({ task });
});

app.post("/api/jira/issues/:issueKey/transition", async (req, res) => {
  const parsed = jiraTransitionSchema.safeParse({
    ...req.body,
    issueKey: req.params.issueKey
  });
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid Jira transition payload" });
  }
  const issueKey = parsed.data.issueKey;
  if (!issueKey) {
    return res.status(400).json({ message: "Missing Jira issue key" });
  }

  try {
    await transitionJiraIssue(issueKey, parsed.data.transitionId);
    let linkedTask = parsed.data.parentTaskId ? getTaskById(parsed.data.parentTaskId) : null;

    if (linkedTask?.source === "Jira" && linkedTask.sourceRef) {
      linkedTask = await refreshJiraTaskFromSource(linkedTask.id, linkedTask.sourceRef);
    }

    const detail = await fetchJiraIssueDetail(issueKey);
    logEvent({
      eventType: "jira.transition",
      requestId: getRequestId(res),
      entityType: "jira_issue",
      entityId: issueKey,
      provider: "jira",
      status: "updated",
      message: "Jira issue transitioned from planner UI.",
      metadata: {
        issueKey,
        transitionId: parsed.data.transitionId,
        parentTaskId: parsed.data.parentTaskId ?? null
      }
    });
    return res.json({
      detail,
      task: linkedTask
    });
  } catch (error) {
    return res.status(400).json({
      message: error instanceof Error ? error.message : "Failed to update Jira status"
    });
  }
});

app.patch("/api/tasks/:id/defer", (req, res) => {
  const schema = z.object({
    deferredUntil: z.string().datetime().nullable()
  });
  const parsed = schema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid defer payload" });
  }
  const row = updateTask(Number(req.params.id), {
    deferredUntil: parsed.data.deferredUntil,
    manualOverrideFlags: ["deferredUntil"],
    lastChangedBy: "user",
    lastChangedAt: new Date().toISOString(),
    wasUserOverridden: true
  });
  if (!row) {
    return res.status(404).json({ message: "Task not found" });
  }
  const task = normalizeTask(row as Record<string, unknown>);
  recordUserTaskMutation({
    taskId: task.id,
    source: task.source,
    sourceRef: task.sourceRef ?? null,
    sourceThreadRef: task.sourceThreadRef ?? null,
    eventType: "deferred",
    reason: parsed.data.deferredUntil ? "Task deferred." : "Task restored from defer.",
    after: task
  });
  void captureTaskFeedback({
    taskId: task.id,
    source: task.source,
    sourceRef: task.sourceRef ?? null,
    sourceThreadRef: task.sourceThreadRef ?? null,
    action: "deferred",
    title: task.title,
    beforePriority: task.priority,
    afterPriority: task.priority,
    decisionReason: task.decisionReason,
    decisionReasonTags: task.decisionReasonTags,
    context: parsed.data.deferredUntil
  });
  logEvent({
    eventType: "task.defer",
    requestId: getRequestId(res),
    entityType: "task",
    entityId: String(task.id),
    status: "updated",
    message: "Task defer state updated.",
    metadata: { taskId: task.id, deferredUntil: parsed.data.deferredUntil }
  });
  return res.json({ task });
});

app.post("/api/tasks/:id/feedback", async (req, res) => {
  const parsed = taskFeedbackSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid feedback payload" });
  }
  const task = getTaskById(Number(req.params.id));
  if (!task) {
    return res.status(404).json({ message: "Task not found" });
  }
  await captureTaskFeedback({
    taskId: task.id,
    source: task.source,
    sourceRef: task.sourceRef ?? null,
    sourceThreadRef: task.sourceThreadRef ?? null,
    action: parsed.data.action,
    title: task.title,
    beforePriority: parsed.data.beforePriority ?? task.priority,
    afterPriority: parsed.data.afterPriority ?? task.priority,
    decisionReason: task.decisionReason,
    decisionReasonTags: task.decisionReasonTags,
    context: parsed.data.context
  });
  return res.json({ ok: true });
});

app.delete("/api/tasks/:id", (req, res) => {
  const task = getTaskById(Number(req.params.id));
  if (task && task.source !== "Manual") {
    upsertRejectedTask({
      title: task.title,
      source: task.source,
      sourceLink: task.sourceLink,
      sourceRef: task.sourceRef ?? null,
      sourceThreadRef: task.sourceThreadRef ?? null,
      jiraStatus: task.jiraStatus,
      proposedPriority: task.priority,
      decisionState: "rejected",
      decisionConfidence: task.decisionConfidence,
      decisionReason: task.decisionReason ?? "User rejected this task from the active plan.",
      decisionReasonTags: task.decisionReasonTags,
      personalizationVersion: task.personalizationVersion,
      candidatePayloadJson: JSON.stringify({
        title: task.title,
        source: task.source,
        sourceLink: task.sourceLink,
        sourceRef: task.sourceRef,
        sourceThreadRef: task.sourceThreadRef,
        jiraStatus: task.jiraStatus
      }),
      rejectedAt: new Date().toISOString()
    });
    void captureTaskFeedback({
      taskId: task.id,
      source: task.source,
      sourceRef: task.sourceRef ?? null,
      sourceThreadRef: task.sourceThreadRef ?? null,
      action: "reject",
      title: task.title,
      beforePriority: task.priority,
      afterPriority: task.priority,
      decisionReason: task.decisionReason,
      decisionReasonTags: task.decisionReasonTags
    });
  }
  const deleted = deleteTask(Number(req.params.id));
  if (!deleted) {
    return res.status(404).json({ message: "Task not found" });
  }
  if (task) {
    recordUserTaskMutation({
      taskId: task.id,
      source: task.source,
      sourceRef: task.sourceRef ?? null,
      sourceThreadRef: task.sourceThreadRef ?? null,
      eventType: task.source === "Manual" ? "remove_manual" : "reject",
      reason: task.source === "Manual" ? "Manual task removed." : "Task rejected from active plan.",
      before: task
    });
  }
  logEvent({
    eventType: "task.delete",
    requestId: getRequestId(res),
    entityType: "task",
    entityId: task ? String(task.id) : req.params.id,
    status: "updated",
    message: task?.source === "Manual" ? "Manual task removed." : "Task rejected from active plan."
  });
  return res.status(204).send();
});

app.patch("/api/tasks/rejected/:id", async (req, res) => {
  const parsed = rejectedTaskPatchSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid rejected-task payload" });
  }
  const rejected = getRejectedTaskById(Number(req.params.id));
  if (!rejected) {
    return res.status(404).json({ message: "Rejected task not found" });
  }

  if (parsed.data.action === "always_ignore_exact") {
    const updated = updateRejectedTask(rejected.id, {
      decisionState: "ignored",
      decisionReason: `${rejected.decisionReason ?? "Rejected"} Marked to always ignore this exact item.`,
      restoredAt: null
    });
    recordUserTaskMutation({
      source: rejected.source,
      sourceRef: rejected.sourceRef,
      sourceThreadRef: rejected.sourceThreadRef,
      eventType: "ignore_exact",
      reason: "Ignored this exact rejected item.",
      before: rejected,
      after: updated
    });
    return res.json({ task: updated });
  }

  if (parsed.data.action === "always_ignore_similar") {
    const profile = getUserPriorityProfile();
    const nextTags = [...new Set([...profile.negativeReasonTags, ...rejected.decisionReasonTags])];
    saveUserPriorityProfile({
      negativeReasonTags: nextTags,
      lastProfileRefreshAt: new Date().toISOString()
    });
    await captureTaskFeedback({
      source: rejected.source,
      sourceRef: rejected.sourceRef,
      sourceThreadRef: rejected.sourceThreadRef,
      action: "always_ignore_similar",
      title: rejected.title,
      beforePriority: rejected.proposedPriority,
      afterPriority: rejected.proposedPriority,
      decisionReason: rejected.decisionReason,
      decisionReasonTags: rejected.decisionReasonTags
    });
  }

  if (parsed.data.action === "should_have_been_included") {
    const profile = getUserPriorityProfile();
    const nextTags = [...new Set([...profile.positiveReasonTags, ...rejected.decisionReasonTags])];
    saveUserPriorityProfile({
      positiveReasonTags: nextTags,
      lastProfileRefreshAt: new Date().toISOString()
    });
    await captureTaskFeedback({
      source: rejected.source,
      sourceRef: rejected.sourceRef,
      sourceThreadRef: rejected.sourceThreadRef,
      action: "should_have_been_included",
      title: rejected.title,
      beforePriority: rejected.proposedPriority,
      afterPriority: rejected.proposedPriority,
      decisionReason: rejected.decisionReason,
      decisionReasonTags: rejected.decisionReasonTags
    });
  }

  const updated = updateRejectedTask(rejected.id, {
    decisionState:
      parsed.data.action === "always_ignore_similar"
        ? "ignored"
        : parsed.data.action === "should_have_been_included"
          ? "uncertain"
          : rejected.decisionState,
    decisionConfidence:
      parsed.data.action === "should_have_been_included"
        ? Math.max(0.35, Math.min(0.95, (rejected.decisionConfidence ?? 0.5) - 0.2))
        : rejected.decisionConfidence,
    decisionReason:
      parsed.data.action === "always_ignore_similar"
        ? `${rejected.decisionReason ?? "Rejected"} Marked to ignore similar items in the future.`
        : parsed.data.action === "should_have_been_included"
          ? `${rejected.decisionReason ?? "Rejected"} You marked this as something that should have been included, so similar items will be ranked higher.`
          : rejected.decisionReason
  });
  recordUserTaskMutation({
    source: rejected.source,
    sourceRef: rejected.sourceRef,
    sourceThreadRef: rejected.sourceThreadRef,
    eventType: parsed.data.action,
    reason: parsed.data.action === "always_ignore_similar" ? "Ignore similar items saved." : "Rejected item marked as under-included.",
    before: rejected,
    after: updated
  });
  return res.json({ task: updated });
});

app.post("/api/tasks/rejected/:id/restore", async (req, res) => {
  const rejected = getRejectedTaskById(Number(req.params.id));
  if (!rejected) {
    return res.status(404).json({ message: "Rejected task not found" });
  }
  const profile = getUserPriorityProfile();
  const nextTags = [...new Set([...profile.positiveReasonTags, ...rejected.decisionReasonTags])];
  saveUserPriorityProfile({
    positiveReasonTags: nextTags,
    lastProfileRefreshAt: new Date().toISOString()
  });
  const restoredAt = new Date().toISOString();
  const exactTask =
    rejected.sourceRef !== null ? getTaskBySource(rejected.source, rejected.sourceRef, { includeIgnored: true }) : null;
  const threadTask =
    rejected.source === "Email" && !exactTask
      ? getTaskBySourceThread(rejected.source, rejected.sourceThreadRef, { includeIgnored: false })
      : null;

  if (threadTask && threadTask.sourceRef !== rejected.sourceRef) {
    updateTask(threadTask.id, {
      title: rejected.title,
      priority: threadTask.manualOverrideFlags.includes("priority") ? undefined : rejected.proposedPriority,
      lastActivityAt: restoredAt,
      decisionState: "restored",
      decisionConfidence: rejected.decisionConfidence,
      decisionReason: rejected.decisionReason ?? "User restored this task from the rejected queue.",
      decisionReasonTags: rejected.decisionReasonTags,
      personalizationVersion: rejected.personalizationVersion,
      restoredAt,
      rejectedAt: null
    });
  } else {
    upsertTask({
      title: rejected.title,
      source: rejected.source,
      priority: rejected.proposedPriority,
      sourceLink: rejected.sourceLink,
      sourceRef: rejected.sourceRef,
      sourceThreadRef: rejected.sourceThreadRef,
      jiraStatus: rejected.jiraStatus,
      decisionState: "restored",
      decisionConfidence: rejected.decisionConfidence,
      decisionReason: rejected.decisionReason ?? "User restored this task from the rejected queue.",
      decisionReasonTags: rejected.decisionReasonTags,
      personalizationVersion: rejected.personalizationVersion,
      restoredAt,
      reviveIgnored: true
    });
  }

  if (rejected.source === "Email") {
    clearRejectedTasksBySourceThread(rejected.source, rejected.sourceThreadRef);
  }

  const updated = updateRejectedTask(rejected.id, {
    decisionState: "restored",
    restoredAt
  });
  recordUserTaskMutation({
    source: rejected.source,
    sourceRef: rejected.sourceRef,
    sourceThreadRef: rejected.sourceThreadRef,
    eventType: "restore",
    reason: "Rejected item restored to plan.",
    before: rejected,
    after: updated
  });
  await captureTaskFeedback({
    source: rejected.source,
    sourceRef: rejected.sourceRef,
    sourceThreadRef: rejected.sourceThreadRef,
    action: "restore",
    title: rejected.title,
    beforePriority: rejected.proposedPriority,
    afterPriority: rejected.proposedPriority,
    decisionReason: rejected.decisionReason,
    decisionReasonTags: rejected.decisionReasonTags
  });
  const restoredTask = listTasks(undefined, { includeDeferred: true }).find((task) => {
    if (task.source !== rejected.source) return false;
    if (rejected.sourceRef && task.sourceRef === rejected.sourceRef) return true;
    return Boolean(rejected.sourceThreadRef && task.sourceThreadRef === rejected.sourceThreadRef);
  });
  return res.json({ task: restoredTask, rejectedTask: updated });
});

app.get("/api/personalization/profile", (_req, res) => {
  res.json({ profile: getUserPriorityProfile() ?? defaultPriorityProfile });
});

app.patch("/api/personalization/profile", (req, res) => {
  const parsed = personalizationProfileSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid personalization profile" });
  }
  const profile = saveUserPriorityProfile({
    ...parsed.data,
    positiveReasonTags: (parsed.data.positiveReasonTags ?? []) as never,
    negativeReasonTags: (parsed.data.negativeReasonTags ?? []) as never,
    lastProfileRefreshAt: new Date().toISOString()
  });
  return res.json({ profile });
});

app.post("/api/personalization/calibrate", async (req, res) => {
  const parsed = calibrationSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid calibration payload" });
  }
  const synthesized = await synthesizePriorityProfile(parsed.data);
  const profile = saveUserPriorityProfile({
    ...synthesized,
    positiveReasonTags: (synthesized.positiveReasonTags ?? []) as never,
    negativeReasonTags: (synthesized.negativeReasonTags ?? []) as never,
    prioritizationPrompt: parsed.data.prioritizationPrompt || synthesized.prioritizationPrompt || defaultPriorityProfile.prioritizationPrompt,
    questionnaireJson: JSON.stringify({
      roleFocus: parsed.data.roleFocus,
      prioritizationPrompt: parsed.data.prioritizationPrompt,
      importantWork: parsed.data.importantWork,
      noiseWork: parsed.data.noiseWork,
      mustNotMiss: parsed.data.mustNotMiss,
      importantPeople: parsed.data.importantPeople,
      importantProjects: parsed.data.importantProjects,
      filteringStyle: parsed.data.filteringStyle,
      priorityBias: parsed.data.priorityBias
    }),
    exampleRankingsJson: JSON.stringify(parsed.data.exampleRankings),
    personalizationEnabled: true,
    lastProfileRefreshAt: new Date().toISOString()
  });
  await captureTaskFeedback({
    source: "Calibration",
    action: "system_evaluated",
    title: "Priority calibration",
    decisionReason: "User completed priority calibration.",
    context: JSON.stringify(parsed.data)
  });
  return res.status(201).json({ profile });
});

app.get("/api/personalization/insights", (_req, res) => {
  const latest = getLatestPreferenceMemorySnapshot();
  res.json({
    insights: latest.insights,
    sourceEventCount: latest.sourceEventCount,
    createdAt: latest.createdAt
  });
});

app.get("/api/settings/integrations", async (req, res) => {
  const rows = listIntegrationConnections().map((row) => ({
    provider: row.provider,
    status: row.status,
    accountLabel: row.account_label,
    errorMessage: row.error_message,
    updatedAt: row.updated_at,
    lastSyncAt: getSyncState(String(row.provider)),
    config:
      row.provider === "jira" && row.config_json
        ? (() => {
            try {
              const parsed = JSON.parse(String(row.config_json)) as {
                baseUrl?: string;
                email?: string;
                apiToken?: string;
              };
              return {
                baseUrl: parsed.baseUrl ?? "",
                email: parsed.email ?? "",
                apiToken: parsed.apiToken ?? ""
              };
            } catch {
              return null;
            }
          })()
        : null
  }));

  const microsoft = rows.find((row) => row.provider === "microsoft") ?? {
    provider: "microsoft",
    status: "disconnected",
    accountLabel: null,
    errorMessage: null,
    updatedAt: null,
    lastSyncAt: null,
    config: null
  };
  const jira = rows.find((row) => row.provider === "jira") ?? {
    provider: "jira",
    status: "disconnected",
    accountLabel: null,
    errorMessage: null,
    updatedAt: null,
    lastSyncAt: null,
    config: null
  };

  try {
    const session = await getOptionalMicrosoftSession(req);
    if (session) {
      const graphToken = await acquireGraphTokenOnBehalfOf(session);
      const profile = await fetchMicrosoftProfileWithAccessToken(graphToken);
      microsoft.status = "connected";
      microsoft.accountLabel =
        profile.userPrincipalName ?? profile.displayName ?? session.accountLabel ?? session.displayName;
      microsoft.errorMessage = null;
    }
  } catch (error) {
    microsoft.status = "error";
    microsoft.errorMessage =
      error instanceof Error
        ? `Microsoft session is unavailable: ${error.message}`
        : "Microsoft session is unavailable";
  }

  res.json({ integrations: { microsoft, jira } });
});

app.get("/api/settings/automation", (_req, res) => {
  res.json({
    automation: getAutomationSettings(),
    reminders: listReminderItems(["active", "dismissed", "resolved"])
  });
});

app.patch("/api/settings/schedule", (req, res) => {
  const parsed = automationSettingsSchema.pick({
    scheduleEnabled: true,
    scheduleTimeLocal: true,
    scheduleTimezone: true
  }).safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid schedule settings" });
  }
  const automation = saveAutomationSettings(parsed.data);
  scheduleAutomation();
  return res.json({ automation });
});

app.patch("/api/settings/reminders", (req, res) => {
  const parsed = automationSettingsSchema.pick({
    remindersEnabled: true,
    reminderCadenceHours: true,
    desktopNotificationsEnabled: true
  }).safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid reminder settings" });
  }
  const automation = saveAutomationSettings(parsed.data);
  return res.json({ automation });
});

app.get("/api/reminders", (_req, res) => {
  res.json(getReminderCenterPayload());
});

app.patch("/api/reminders/:id", (req, res) => {
  const parsed = reminderUpdateSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid reminder payload" });
  }
  const reminder = getReminderById(Number(req.params.id));
  if (!reminder) {
    return res.status(404).json({ message: "Reminder not found" });
  }
  const nextStatus = parsed.data.status ?? reminder.status;
  const updated = updateReminder(reminder.id, {
    ...parsed.data,
    status: nextStatus,
    dismissedAt: nextStatus === "dismissed" ? new Date().toISOString() : null
  });
  logEvent({
    eventType: "reminder.update",
    requestId: getRequestId(res),
    entityType: "reminder",
    entityId: String(reminder.id),
    status: "updated",
    message: "Reminder updated.",
    metadata: { status: nextStatus }
  });
  return res.json({ reminder: updated });
});

app.post("/api/settings/integrations/jira", async (req, res) => {
  const schema = z.object({
    baseUrl: z.string().url(),
    email: z.string().email(),
    apiToken: z.string().min(1)
  });
  const parsed = schema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid Jira settings" });
  }

  try {
    const normalizedInput = {
      ...parsed.data,
      baseUrl: normalizeJiraBaseUrl(parsed.data.baseUrl),
      email: parsed.data.email.trim(),
      apiToken: parsed.data.apiToken.trim()
    };
    const validation = await validateJiraCredentials(normalizedInput);
    saveIntegrationConnection({
      provider: "jira",
      status: "connected",
      accountLabel:
        validation.profile.emailAddress ?? validation.profile.displayName ?? normalizedInput.email,
      configJson: JSON.stringify({
        ...normalizedInput,
        authType: validation.authType
      }),
      accessToken: null,
      refreshToken: null,
      expiresAt: null,
      errorMessage: null
    });
    logEvent({
      eventType: "integration.jira.save",
      requestId: getRequestId(res),
      entityType: "integration",
      entityId: "jira",
      provider: "jira",
      status: "success",
      message: "Jira integration saved.",
      metadata: { accountLabel: validation.profile.emailAddress ?? validation.profile.displayName ?? normalizedInput.email }
    });
    return res.status(201).json({ ok: true });
  } catch (error) {
    saveIntegrationConnection({
      provider: "jira",
      status: "error",
      accountLabel: parsed.data.email.trim(),
      configJson: JSON.stringify({
        ...parsed.data,
        baseUrl: (() => {
          try {
            return normalizeJiraBaseUrl(parsed.data.baseUrl);
          } catch {
            return parsed.data.baseUrl.trim();
          }
        })()
      }),
      accessToken: null,
      refreshToken: null,
      expiresAt: null,
      errorMessage: error instanceof Error ? error.message : "Jira validation failed"
    });
    logEvent({
      level: "error",
      eventType: "integration.jira.save",
      requestId: getRequestId(res),
      entityType: "integration",
      entityId: "jira",
      provider: "jira",
      status: "failure",
      message: error instanceof Error ? error.message : "Jira validation failed"
    });
    return res.status(400).json({
      message: error instanceof Error ? error.message : "Jira validation failed"
    });
  }
});

app.delete("/api/settings/integrations/:provider", (req, res) => {
  const provider = req.params.provider;
  if (provider !== "microsoft" && provider !== "jira") {
    return res.status(400).json({ message: "Unsupported integration provider" });
  }

  deleteIntegrationConnection(provider);
  logEvent({
    eventType: "integration.revoke",
    requestId: getRequestId(res),
    entityType: "integration",
    entityId: provider,
    provider,
    status: "updated",
    message: `${provider} integration revoked.`
  });
  return res.status(204).send();
});

app.get("/api/auth/microsoft/start", (_req, res) => {
  if (!env.microsoftClientId || !env.microsoftClientSecret) {
    return res.status(400).json({ message: "Microsoft OAuth is not configured in apps/api/.env" });
  }
  logEvent({
    eventType: "integration.microsoft.start",
    requestId: getRequestId(res),
    entityType: "integration",
    entityId: "microsoft",
    provider: "microsoft",
    status: "started",
    message: "Microsoft auth flow requested."
  });
  res.json({ url: getMicrosoftAuthUrl() });
});

app.get("/api/auth/microsoft/callback", async (req, res) => {
  const code = String(req.query.code ?? "");
  if (!code) {
    return res.status(400).send("Missing code");
  }
  try {
    await exchangeMicrosoftCode(code);
    logEvent({
      eventType: "integration.microsoft.callback",
      requestId: getRequestId(res),
      entityType: "integration",
      entityId: "microsoft",
      provider: "microsoft",
      status: "success",
      message: "Microsoft auth callback succeeded."
    });
    return res.redirect(`${env.appOrigin}/settings?connected=microsoft`);
  } catch (error) {
    logEvent({
      level: "error",
      eventType: "integration.microsoft.callback",
      requestId: getRequestId(res),
      entityType: "integration",
      entityId: "microsoft",
      provider: "microsoft",
      status: "failure",
      message: error instanceof Error ? error.message : "Microsoft connection failed"
    });
    return res
      .status(400)
      .send(error instanceof Error ? error.message : "Microsoft connection failed");
  }
});

app.use((error: unknown, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
  logEvent({
    level: "error",
    eventType: "http.unhandled_error",
    requestId: getRequestId(res),
    status: "failure",
    source: "server",
    message: error instanceof Error ? error.message : "Unhandled server error"
  });
  res.status(500).json({
    message: error instanceof Error ? error.message : "Internal server error"
  });
});

app.listen(env.port, () => {
  logEvent({
    eventType: "server.start",
    entityType: "server",
    entityId: String(env.port),
    status: "success",
    source: "server",
    message: `API listening on http://localhost:${env.port}`
  });
});

scheduleAutomation();
