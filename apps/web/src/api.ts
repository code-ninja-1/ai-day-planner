import type {
  AutomationSettings,
  IntegrationStatus,
  Reminder,
  Task,
  TaskDetail,
  TaskPriority,
  TaskStatus,
  TodayResponse
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
  deferTask: (id: number, deferredUntil: string | null) =>
    request<{ task: Task }>(`/tasks/${id}/defer`, {
      method: "PATCH",
      body: JSON.stringify({ deferredUntil })
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
  getMicrosoftAuthUrl: () => request<{ url: string }>("/auth/microsoft/start")
};
