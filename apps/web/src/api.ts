import type {
  AuditEvent,
  AutomationSettings,
  DayHistoryDetail,
  DayHistorySummary,
  DiagnosticsPayload,
  EmailReplyDraft,
  InsightsOverview,
  InsightsTodayPayload,
  IntegrationStatus,
  MeetingPrep,
  PersonalizationInsight,
  PlannerRunDetail,
  RejectedTask,
  Reminder,
  Task,
  TaskDetail,
  TaskInsightsPayload,
  TaskPriority,
  TaskStatus,
  TodayResponse,
  UserPriorityProfile
} from "./types";

const API_ROOT = import.meta.env.VITE_API_ROOT ?? "http://localhost:4000/api";

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${API_ROOT}${path}`, {
    headers: {
      "Content-Type": "application/json",
      ...(init?.headers ?? {})
    },
    ...init
  });

  if (!response.ok) {
    const json = (await response.json().catch(() => ({}))) as { message?: string };
    throw new Error(json.message ?? `Request failed: ${response.status}`);
  }

  if (response.status === 204) {
    return undefined as T;
  }

  return (await response.json()) as T;
}

async function authedRequest<T>(path: string, accessToken: string, init?: RequestInit): Promise<T> {
  return request<T>(path, {
    ...init,
    headers: {
      ...(init?.headers ?? {}),
      Authorization: `Bearer ${accessToken}`
    }
  });
}

export const api = {
  getToday: () => request<TodayResponse>("/today"),
  getAutomationSettings: () =>
    request<{
      automation: AutomationSettings;
      reminders: Reminder[];
    }>("/settings/automation"),
  generatePlan: (timeZone?: string) =>
    request<TodayResponse>("/plan/generate", {
      method: "POST",
      body: JSON.stringify({ timeZone })
    }),
  generatePlanWithMicrosoftSession: (accessToken: string, timeZone?: string) =>
    authedRequest<TodayResponse>("/plan/generate", accessToken, {
      method: "POST",
      body: JSON.stringify({ timeZone })
    }),
  syncMeetings: (timeZone?: string) =>
    request<TodayResponse>("/sync/meetings", {
      method: "POST",
      body: JSON.stringify({ timeZone })
    }),
  syncMeetingsWithMicrosoftSession: (accessToken: string, timeZone?: string) =>
    authedRequest<TodayResponse>("/sync/meetings", accessToken, {
      method: "POST",
      body: JSON.stringify({ timeZone })
    }),
  syncTasks: (timeZone?: string) =>
    request<TodayResponse>("/sync/tasks", {
      method: "POST",
      body: JSON.stringify({ timeZone })
    }),
  syncTasksWithMicrosoftSession: (accessToken: string, timeZone?: string) =>
    authedRequest<TodayResponse>("/sync/tasks", accessToken, {
      method: "POST",
      body: JSON.stringify({ timeZone })
    }),
  getDeferredTasks: () => request<{ tasks: Task[] }>("/tasks/deferred"),
  getRejectedTasks: () => request<{ tasks: RejectedTask[]; ignoredTasks: RejectedTask[] }>("/tasks/rejected"),
  getInsightsOverview: () => request<InsightsOverview>("/insights/overview"),
  getInsightsToday: () => request<InsightsTodayPayload>("/insights/today"),
  getInsightsHistory: (limit = 30) => request<{ days: DayHistorySummary[] }>(`/insights/history?limit=${limit}`),
  getInsightsHistoryDay: (dayKey: string) => request<DayHistoryDetail>(`/insights/history/${encodeURIComponent(dayKey)}`),
  getTaskInsights: (taskId: number) => request<TaskInsightsPayload>(`/insights/tasks/${taskId}`),
  getDebugRuns: () => request<{ runs: PlannerRunDetail[]; diagnostics: DiagnosticsPayload }>("/debug/runs"),
  getDebugLogs: (limit = 200) => request<{ logs: AuditEvent[] }>(`/debug/logs?limit=${limit}`),
  logClientEvent: (input: {
    eventType: string;
    level?: "debug" | "info" | "warn" | "error";
    message: string;
    entityType?: string | null;
    entityId?: string | null;
    status?: "started" | "success" | "failure" | "updated" | "skipped" | "info";
    metadata?: unknown;
  }) =>
    request<{ ok: boolean }>("/debug/client-events", {
      method: "POST",
      body: JSON.stringify(input)
    }),
  restoreRejectedTask: (id: number) =>
    request<{ task: Task | null; rejectedTask: RejectedTask | null }>(`/tasks/rejected/${id}/restore`, {
      method: "POST"
    }),
  updateRejectedTask: (
    id: number,
    input: { action: "always_ignore_exact" | "always_ignore_similar" | "should_have_been_included" | "keep_rejected" }
  ) =>
    request<{ task: RejectedTask | null }>(`/tasks/rejected/${id}`, {
      method: "PATCH",
      body: JSON.stringify(input)
    }),
  deferTask: (id: number, deferredUntil: string | null) =>
    request<{ task: Task }>(`/tasks/${id}/defer`, {
      method: "PATCH",
      body: JSON.stringify({ deferredUntil })
    }),
  sendTaskFeedback: (
    id: number,
    input: {
      action:
        | "reject"
        | "restore"
        | "priority_changed"
        | "status_changed"
        | "deferred"
        | "completed"
        | "always_ignore_similar"
        | "should_have_been_included";
      beforePriority?: TaskPriority | null;
      afterPriority?: TaskPriority | null;
      context?: string | null;
    }
  ) =>
    request<{ ok: boolean }>(`/tasks/${id}/feedback`, {
      method: "POST",
      body: JSON.stringify(input)
    }),
  getReminders: () => request<{ reminders: Reminder[] }>("/reminders"),
  updateReminder: (
    id: number,
    input: Partial<{
      status: "active" | "dismissed" | "resolved";
      reason: string;
      scheduledFor: string | null;
      throttleUntil: string | null;
    }>
  ) =>
    request<{ reminder: Reminder }>(`/reminders/${id}`, {
      method: "PATCH",
      body: JSON.stringify(input)
    }),
  getTasks: (status?: TaskStatus) =>
    request<{ tasks: Task[] }>(`/tasks${status ? `?status=${encodeURIComponent(status)}` : ""}`),
  createTask: (input: { title: string; priority?: TaskPriority; status?: TaskStatus }) =>
    request<{ task: Task }>("/tasks", {
      method: "POST",
      body: JSON.stringify(input)
    }),
  updateTask: (id: number, input: Partial<{ title: string; priority: TaskPriority; status: TaskStatus }>) =>
    request<{ task: Task }>(`/tasks/${id}`, {
      method: "PATCH",
      body: JSON.stringify(input)
    }),
  getTaskDetails: (id: number) =>
    request<{ detail: TaskDetail }>(`/tasks/${id}/details`),
  getTaskDetailsWithMicrosoftSession: (id: number, accessToken: string) =>
    authedRequest<{ detail: TaskDetail }>(`/tasks/${id}/details`, accessToken),
  generateEmailReplyDraft: (id: number, input: { userIntent?: string | null }) =>
    request<{ draft: EmailReplyDraft }>(`/tasks/${id}/email-reply/draft`, {
      method: "POST",
      body: JSON.stringify(input)
    }),
  generateEmailReplyDraftWithMicrosoftSession: (id: number, accessToken: string, input: { userIntent?: string | null }) =>
    authedRequest<{ draft: EmailReplyDraft }>(`/tasks/${id}/email-reply/draft`, accessToken, {
      method: "POST",
      body: JSON.stringify(input)
    }),
  sendEmailReply: (id: number, input: { to: string[]; cc: string[]; subject: string; body: string }) =>
    request<{ ok: boolean }>(`/tasks/${id}/email-reply/send`, {
      method: "POST",
      body: JSON.stringify(input)
    }),
  sendEmailReplyWithMicrosoftSession: (
    id: number,
    accessToken: string,
    input: { to: string[]; cc: string[]; subject: string; body: string }
  ) =>
    authedRequest<{ ok: boolean }>(`/tasks/${id}/email-reply/send`, accessToken, {
      method: "POST",
      body: JSON.stringify(input)
    }),
  generateMeetingPrep: (id: number, input: { userNotes?: string | null }) =>
    request<{ prep: MeetingPrep }>(`/meetings/${id}/prepare`, {
      method: "POST",
      body: JSON.stringify(input)
    }),
  generateMeetingPrepWithMicrosoftSession: (id: number, accessToken: string, input: { userNotes?: string | null }) =>
    authedRequest<{ prep: MeetingPrep }>(`/meetings/${id}/prepare`, accessToken, {
      method: "POST",
      body: JSON.stringify(input)
    }),
  transitionJiraIssue: (issueKey: string, input: { transitionId: string; parentTaskId?: number }) =>
    request<{ detail: Extract<TaskDetail, { type: "jira" }>; task: Task | null }>(
      `/jira/issues/${encodeURIComponent(issueKey)}/transition`,
      {
        method: "POST",
        body: JSON.stringify(input)
      }
    ),
  deleteTask: (id: number) =>
    request<void>(`/tasks/${id}`, {
      method: "DELETE"
    }),
  getIntegrations: () =>
    request<{
      integrations: {
        microsoft: IntegrationStatus;
        jira: IntegrationStatus;
      };
    }>("/settings/integrations"),
  getIntegrationsWithMicrosoftSession: (accessToken: string) =>
    authedRequest<{
      integrations: {
        microsoft: IntegrationStatus;
        jira: IntegrationStatus;
      };
    }>("/settings/integrations", accessToken),
  saveJira: (input: { baseUrl: string; email: string; apiToken: string }) =>
    request<{ ok: boolean }>("/settings/integrations/jira", {
      method: "POST",
      body: JSON.stringify(input)
    }),
  revokeIntegration: (provider: "microsoft" | "jira") =>
    request<void>(`/settings/integrations/${provider}`, {
      method: "DELETE"
    }),
  updateScheduleSettings: (input: Partial<Pick<AutomationSettings, "scheduleEnabled" | "scheduleTimeLocal" | "scheduleTimezone">>) =>
    request<{ automation: AutomationSettings }>("/settings/schedule", {
      method: "PATCH",
      body: JSON.stringify(input)
    }),
  updateReminderSettings: (
    input: Partial<Pick<AutomationSettings, "remindersEnabled" | "reminderCadenceHours" | "desktopNotificationsEnabled">>
  ) =>
    request<{ automation: AutomationSettings }>("/settings/reminders", {
      method: "PATCH",
      body: JSON.stringify(input)
    }),
  getPersonalizationProfile: () => request<{ profile: UserPriorityProfile }>("/personalization/profile"),
  updatePersonalizationProfile: (input: Partial<UserPriorityProfile>) =>
    request<{ profile: UserPriorityProfile }>("/personalization/profile", {
      method: "PATCH",
      body: JSON.stringify(input)
    }),
  calibratePersonalization: (input: {
    roleFocus: string;
    prioritizationPrompt: string;
    importantWork: string[];
    noiseWork: string[];
    mustNotMiss: string[];
    importantPeople: string[];
    importantProjects: string[];
    filteringStyle: UserPriorityProfile["filteringStyle"];
    priorityBias: UserPriorityProfile["priorityBias"];
    exampleRankings: Array<{ title: string; source: "Email" | "Jira" | "Manual"; decision: "show_today" | "keep_low" | "reject_noise" }>;
  }) =>
    request<{ profile: UserPriorityProfile }>("/personalization/calibrate", {
      method: "POST",
      body: JSON.stringify(input)
    }),
  getPersonalizationInsights: () =>
    request<{ insights: PersonalizationInsight[]; sourceEventCount: number; createdAt: string | null }>(
      "/personalization/insights"
    ),
  getMicrosoftAuthUrl: () => request<{ url: string }>("/auth/microsoft/start")
};
