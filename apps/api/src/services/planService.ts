import {
  getAutomationSettings,
  getIntegrationConnection,
  getSyncState,
  groupTasksByPriority,
  listDeferredTasks,
  listMeetings,
  listReminderItems,
  listTasks,
  recordGenerationRun,
  replaceMeetings,
  resolveStaleReminders,
  saveAutomationSettings,
  setSyncState,
  updateTask,
  upsertReminder,
  upsertTask
} from "../db.js";
import {
  buildJiraIssueBrowseUrl,
  fetchRecentAssignedIssues,
  getMappedJiraPriority
} from "../providers/jira.js";
import {
  fetchRecentEmails,
  fetchRecentEmailsWithAccessToken,
  fetchTodaysMeetings,
  fetchTodaysMeetingsWithAccessToken
} from "../providers/microsoft.js";
import { classifyEmail } from "./emailClassifier.js";
import type { GraphEvent, GraphMail } from "../providers/microsoft.js";
import type {
  ReminderKind,
  Task,
  TaskEffortBucket,
  TaskPriority,
  TaskStatus,
  TodayPayload,
  WorkloadSummary
} from "../types.js";

function startAndEndOfAgendaWindow() {
  const start = new Date();
  start.setDate(start.getDate() - 2);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + 7);
  return { startIso: start.toISOString(), endIso: end.toISOString() };
}

function mapJiraWorkflowStatus(status?: string | null): TaskStatus {
  const value = (status ?? "").toLowerCase();
  if (/(done|closed|resolved|complete|completed)/.test(value)) {
    return "Completed";
  }
  if (/(progress|coding|review|testing|qa|blocked|in dev|development)/.test(value)) {
    return "In Progress";
  }
  return "Not Started";
}

function parseGraphCalendarDateTime(value?: string | null) {
  if (!value) return null;
  if (/[zZ]|[+-]\d{2}:\d{2}$/.test(value)) {
    return new Date(value);
  }
  return new Date(`${value}Z`);
}

async function applyMicrosoftResults(
  emails: GraphMail[],
  meetings: GraphEvent[],
  calendarTimeZone: string | null
) {
  for (const email of emails) {
    const classification = await classifyEmail(email);
    if (!classification.actionable) continue;

    upsertTask({
      title: classification.title,
      source: "Email",
      priority: classification.priority,
      sourceLink: email.webLink ?? null,
      sourceRef: email.id,
      sourceThreadRef: email.conversationId ?? null
    });
  }

  if (meetings.length) {
    replaceMeetings(
      meetings.map((meeting) => {
        const start = parseGraphCalendarDateTime(meeting.start?.dateTime);
        const end = parseGraphCalendarDateTime(meeting.end?.dateTime);
        const isCancelled =
          meeting.isCancelled === true ||
          /^canceled:/i.test(meeting.subject ?? "") ||
          /^cancelled:/i.test(meeting.subject ?? "");
        const meetingLink = isCancelled ? null : meeting.onlineMeetingUrl ?? meeting.webLink ?? null;
        const meetingLinkType = isCancelled ? null : meeting.onlineMeetingUrl ? "join" : meeting.webLink ? "calendar" : null;
        return {
          externalId: meeting.id,
          title: meeting.subject || "Untitled meeting",
          startTime: meeting.start?.dateTime ?? "",
          endTime: meeting.end?.dateTime ?? "",
          timeZone: meeting.start?.timeZone ?? meeting.end?.timeZone ?? calendarTimeZone ?? null,
          durationMinutes:
            start && end ? Math.max(0, Math.round((end.getTime() - start.getTime()) / 60000)) : 0,
          meetingLink,
          meetingLinkType,
          isCancelled
        };
      })
    );
  }
}

async function syncMicrosoftTasks(options?: {
  microsoftGraphAccessToken?: string | null;
  microsoftWarning?: string | null;
}) {
  const warnings: string[] = [];
  const sinceIso = new Date(Date.now() - 48 * 60 * 60 * 1000).toISOString();
  const now = new Date().toISOString();

  if (options?.microsoftGraphAccessToken) {
    try {
      const emails = await fetchRecentEmailsWithAccessToken(sinceIso, options.microsoftGraphAccessToken);
      await applyMicrosoftResults(emails, [], null);
      setSyncState("microsoft", now);
    } catch (error) {
      warnings.push(
        error instanceof Error ? `Microsoft task sync failed: ${error.message}` : "Microsoft task sync failed"
      );
    }
  } else {
    const microsoftConnection = getIntegrationConnection("microsoft");
    if (microsoftConnection?.status === "connected" && microsoftConnection.accessToken) {
      try {
        const emails = await fetchRecentEmails(sinceIso);
        await applyMicrosoftResults(emails, [], null);
        setSyncState("microsoft", now);
      } catch (error) {
        warnings.push(
          error instanceof Error ? `Microsoft task sync failed: ${error.message}` : "Microsoft task sync failed"
        );
      }
    } else if (options?.microsoftWarning) {
      warnings.push(options.microsoftWarning);
    }
  }

  return warnings;
}

async function syncMicrosoftMeetings(options?: {
  microsoftGraphAccessToken?: string | null;
  microsoftWarning?: string | null;
  preferredTimeZone?: string | null;
}) {
  const warnings: string[] = [];
  const now = new Date().toISOString();
  const { startIso, endIso } = startAndEndOfAgendaWindow();

  if (options?.microsoftGraphAccessToken) {
    try {
      const meetingsResult = await fetchTodaysMeetingsWithAccessToken(
        startIso,
        endIso,
        options.microsoftGraphAccessToken,
        options.preferredTimeZone
      );
      await applyMicrosoftResults([], meetingsResult.events, meetingsResult.timeZone);
      setSyncState("microsoft", now);
    } catch (error) {
      warnings.push(
        error instanceof Error ? `Microsoft meeting sync failed: ${error.message}` : "Microsoft meeting sync failed"
      );
    }
  } else {
    const microsoftConnection = getIntegrationConnection("microsoft");
    if (microsoftConnection?.status === "connected" && microsoftConnection.accessToken) {
      try {
        const meetingsResult = await fetchTodaysMeetings(startIso, endIso, options?.preferredTimeZone);
        await applyMicrosoftResults([], meetingsResult.events, meetingsResult.timeZone);
        setSyncState("microsoft", now);
      } catch (error) {
        warnings.push(
          error instanceof Error ? `Microsoft meeting sync failed: ${error.message}` : "Microsoft meeting sync failed"
        );
      }
    } else if (options?.microsoftWarning) {
      warnings.push(options.microsoftWarning);
    }
  }

  return warnings;
}

async function syncJiraTasks() {
  const warnings: string[] = [];
  const now = new Date().toISOString();
  const sinceIso = new Date(Date.now() - 48 * 60 * 60 * 1000).toISOString();

  const jiraConnection = getIntegrationConnection("jira");
  if (jiraConnection?.status === "connected") {
    try {
      const issues = await fetchRecentAssignedIssues(sinceIso);
      const jiraConfig = jiraConnection.configJson
        ? (JSON.parse(jiraConnection.configJson) as { baseUrl?: string })
        : null;
      for (const issue of issues) {
        upsertTask({
          title: `${issue.key} ${issue.fields.summary}`,
          source: "Jira",
          priority: getMappedJiraPriority(issue.fields.priority?.name),
          status: mapJiraWorkflowStatus(issue.fields.status?.name),
          sourceLink: jiraConfig?.baseUrl ? buildJiraIssueBrowseUrl(jiraConfig.baseUrl, issue.key) : issue.self,
          sourceRef: issue.key,
          jiraStatus: issue.fields.status?.name ?? null
        });
      }
      setSyncState("jira", now);
    } catch (error) {
      warnings.push(error instanceof Error ? `Jira sync failed: ${error.message}` : "Jira sync failed");
    }
  }

  return warnings;
}

function daysBetween(fromIso: string, toIso: string) {
  return Math.max(0, Math.floor((new Date(toIso).getTime() - new Date(fromIso).getTime()) / 86_400_000));
}

function inferEffortBucket(task: Task): TaskEffortBucket {
  const text = `${task.title} ${task.jiraStatus ?? ""}`.toLowerCase();
  if (/(migration|epic|refactor|design|investigation|spike|replace|build)/.test(text)) return "2+ hours";
  if (/(review|testing|qa|implement|feature|story)/.test(text)) return "1 hour";
  if (/(follow up|comment|reply|triage|check|standup|prep)/.test(text)) return "30 min";
  return "15 min";
}

function minutesForEffort(bucket: TaskEffortBucket | null) {
  switch (bucket) {
    case "30 min":
      return 30;
    case "1 hour":
      return 60;
    case "2+ hours":
      return 150;
    case "15 min":
    default:
      return 15;
  }
}

function computeTaskSignals(task: Task, meetings: ReturnType<typeof listMeetings>) {
  let score = 0;
  const reasons: string[] = [];
  const ageDays = daysBetween(task.createdAt, new Date().toISOString());
  const carryForwardCount = task.status === "Completed" ? 0 : Math.max(0, ageDays - 1);
  const effort = inferEffortBucket(task);

  if (task.priority === "High") score += 34;
  else if (task.priority === "Medium") score += 18;
  else score += 8;

  if (task.source === "Jira") score += 10;
  if (task.source === "Email") score += 6;
  if (task.status === "In Progress") {
    score += 12;
    reasons.push("Already in progress");
  }
  if (task.deferredUntil) {
    const deferredTime = new Date(task.deferredUntil).getTime();
    const hoursUntil = (deferredTime - Date.now()) / 3_600_000;
    if (hoursUntil <= 0) {
      score += 20;
      reasons.push("Deferred task is due again");
    } else {
      score -= 18;
    }
  }
  if (ageDays >= 3 && task.status !== "Completed") {
    score += 10;
    reasons.push(`Unfinished for ${ageDays} days`);
  }
  if (carryForwardCount > 0) {
    score += Math.min(16, carryForwardCount * 4);
    reasons.push(`Carried forward ${carryForwardCount} day${carryForwardCount === 1 ? "" : "s"}`);
  }
  if (task.source === "Email" && /(action required|follow up|approval|urgent)/i.test(task.title)) {
    score += 12;
    reasons.push("Urgency signal detected");
  }
  if (task.source === "Jira" && task.jiraStatus && /blocked|review|qa/i.test(task.jiraStatus)) {
    score += 8;
    reasons.push(task.jiraStatus);
  }
  const nextMeeting = meetings.find(
    (meeting) => !meeting.isCancelled && new Date(meeting.startTime).getTime() > Date.now()
  );
  if (nextMeeting && /prep|agenda|review/i.test(task.title)) {
    score += 10;
    reasons.push("Relevant to an upcoming meeting");
  }

  const explanation =
    reasons[0] ??
    (task.source === "Jira" ? "Assigned Jira work updated recently" : task.source === "Email" ? "Recent email needs attention" : "Manual task still active");

  return {
    priorityScore: Math.round(score),
    priorityExplanation: explanation,
    estimatedEffortBucket: effort,
    taskAgeDays: ageDays,
    carryForwardCount
  };
}

function derivePriorityFromScore(score: number): TaskPriority {
  if (score >= 48) return "High";
  if (score >= 24) return "Medium";
  return "Low";
}

function applyTaskIntelligence() {
  const meetings = listMeetings();
  const tasks = listTasks(undefined, { includeDeferred: true });
  for (const task of tasks) {
    const intelligence = computeTaskSignals(task, meetings);
    const overridePriority = task.manualOverrideFlags.includes("priority");
    updateTask(task.id, {
      priority: overridePriority ? task.priority : derivePriorityFromScore(intelligence.priorityScore),
      priorityScore: intelligence.priorityScore,
      priorityExplanation: intelligence.priorityExplanation,
      estimatedEffortBucket: intelligence.estimatedEffortBucket,
      taskAgeDays: intelligence.taskAgeDays,
      carryForwardCount: intelligence.carryForwardCount
    });
  }
}

function syncReminders() {
  const settings = getAutomationSettings();
  if (!settings.remindersEnabled) {
    resolveStaleReminders([]);
    return;
  }

  const activeTasks = listTasks(undefined, { includeDeferred: true });
  const meetings = listMeetings();
  const activeKeys: string[] = [];
  const cadenceMs = settings.reminderCadenceHours * 3_600_000;

  for (const task of activeTasks) {
    if (task.status === "Completed") continue;
    let kind: ReminderKind | null = null;
    let reason = "";

    if (task.deferredUntil && new Date(task.deferredUntil).getTime() <= Date.now()) {
      kind = "deferred_due";
      reason = "Deferred task is active again.";
    } else if (task.source === "Email" && task.taskAgeDays >= 1 && task.status === "Not Started") {
      kind = "email_follow_up";
      reason = "Important email has not been answered yet.";
    } else if (task.source === "Jira" && task.taskAgeDays >= 2 && task.status !== "In Progress") {
      kind = "jira_stale";
      reason = "Assigned Jira work shows no recent progress.";
    }

    if (!kind) continue;

    const reminderKey = `${kind}:${task.id}`;
    activeKeys.push(reminderKey);
    const lastRemindedAt = task.lastRemindedAt ? new Date(task.lastRemindedAt).getTime() : 0;
    if (lastRemindedAt && Date.now() - lastRemindedAt < cadenceMs) continue;

    upsertReminder({
      reminderKey,
      taskId: task.id,
      kind,
      title: task.title,
      reason,
      status: "active",
      sourceLink: task.sourceLink,
      sourceLabel: task.source,
      scheduledFor: task.deferredUntil ?? null,
      throttleUntil: new Date(Date.now() + cadenceMs).toISOString()
    });
    updateTask(task.id, { reminderState: "active", lastRemindedAt: new Date().toISOString() });
  }

  for (const meeting of meetings) {
    if (meeting.isCancelled) continue;
    const start = new Date(meeting.startTime).getTime();
    const hoursUntil = (start - Date.now()) / 3_600_000;
    if (hoursUntil > 0 && hoursUntil <= 2) {
      const key = `meeting_prep:${meeting.externalId ?? meeting.id}`;
      activeKeys.push(key);
      upsertReminder({
        reminderKey: key,
        kind: "meeting_prep",
        title: meeting.title,
        reason: "Upcoming meeting starts soon. Prep while the context is fresh.",
        status: "active",
        sourceLink: meeting.meetingLink,
        sourceLabel: "Calendar",
        scheduledFor: meeting.startTime,
        throttleUntil: new Date(start).toISOString()
      });
    }
  }

  resolveStaleReminders(activeKeys);
}

function buildWorkloadSummary(tasks: Task[]): WorkloadSummary {
  const meetingsTodayMinutes = listMeetings()
    .filter((meeting) => {
      const start = new Date(meeting.startTime);
      const now = new Date();
      return (
        start.getFullYear() === now.getFullYear() &&
        start.getMonth() === now.getMonth() &&
        start.getDate() === now.getDate() &&
        !meeting.isCancelled
      );
    })
    .reduce((sum, meeting) => sum + meeting.durationMinutes, 0);

  const taskMinutes = tasks
    .filter((task) => task.status !== "Completed")
    .reduce((sum, task) => sum + minutesForEffort(task.estimatedEffortBucket), 0);

  const totalPlannedMinutes = meetingsTodayMinutes + taskMinutes;
  const state =
    totalPlannedMinutes < 240 ? "Underloaded" : totalPlannedMinutes <= 480 ? "Balanced" : "Overloaded";

  return {
    totalMeetingMinutes: meetingsTodayMinutes,
    totalTaskMinutes: taskMinutes,
    totalPlannedMinutes,
    state
  };
}

function buildPayload(warnings: string[] = []): TodayPayload {
  const tasks = listTasks();
  return {
    meetings: listMeetings(),
    tasks: groupTasksByPriority(tasks),
    reminders: listReminderItems(["active", "dismissed"]),
    workload: buildWorkloadSummary(tasks),
    deferredTaskCount: listDeferredTasks().length,
    automation: getAutomationSettings(),
    sync: {
      microsoft: getSyncState("microsoft"),
      jira: getSyncState("jira"),
      lastGeneratedAt: getSyncState("plan")
    },
    warnings
  };
}

export function getTodaySnapshot() {
  return buildPayload([]);
}

async function postSyncMaintenance(triggerType: "manual" | "scheduled", warnings: string[]) {
  applyTaskIntelligence();
  syncReminders();
  const now = new Date().toISOString();
  setSyncState("plan", now);
  recordGenerationRun(triggerType, warnings);
  if (triggerType === "scheduled") {
    saveAutomationSettings({
      lastAutoGeneratedAt: now,
      schedulerLastRunAt: now,
      schedulerLastStatus: "ok",
      schedulerLastError: null
    });
  }
}

export async function generatePlan(
  options?: {
    microsoftGraphAccessToken?: string | null;
    microsoftWarning?: string | null;
    preferredTimeZone?: string | null;
  },
  triggerType: "manual" | "scheduled" = "manual"
): Promise<TodayPayload> {
  const [microsoftTaskWarnings, microsoftMeetingWarnings, jiraWarnings] = await Promise.all([
    syncMicrosoftTasks(options),
    syncMicrosoftMeetings(options),
    syncJiraTasks()
  ]);
  const warnings = [...microsoftTaskWarnings, ...microsoftMeetingWarnings, ...jiraWarnings];
  await postSyncMaintenance(triggerType, warnings);
  return buildPayload(warnings);
}

export async function syncMeetingsOnly(options?: {
  microsoftGraphAccessToken?: string | null;
  microsoftWarning?: string | null;
  preferredTimeZone?: string | null;
}) {
  const warnings = await syncMicrosoftMeetings(options);
  syncReminders();
  return buildPayload(warnings);
}

export async function syncTasksOnly(options?: {
  microsoftGraphAccessToken?: string | null;
  microsoftWarning?: string | null;
  preferredTimeZone?: string | null;
}) {
  const [microsoftWarnings, jiraWarnings] = await Promise.all([syncMicrosoftTasks(options), syncJiraTasks()]);
  applyTaskIntelligence();
  syncReminders();
  return buildPayload([...microsoftWarnings, ...jiraWarnings]);
}

export function getDeferredTasksPayload() {
  return { tasks: listDeferredTasks() };
}

export function getReminderCenterPayload() {
  return { reminders: listReminderItems(["active", "dismissed", "resolved"]) };
}
