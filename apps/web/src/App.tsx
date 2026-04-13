import { useEffect, useMemo, useRef, useState } from "react";
import { api } from "./api";
import {
  acquireMicrosoftApiToken,
  getMicrosoftAccount,
  loginWithMicrosoft,
  logoutFromMicrosoft,
  type MicrosoftAccount
} from "./auth";
import type {
  AutomationSettings,
  EmailTaskDetail,
  IntegrationStatus,
  JiraTaskDetail,
  PersonalizationInsight,
  RejectedTask,
  Reminder,
  Task,
  TaskDetail,
  TaskPriority,
  TaskStatus,
  TodayResponse,
  UserPriorityProfile
} from "./types";

type View = "today" | "tasks" | "deferred" | "rejected" | "reminders" | "settings";
type TaskFilter = TaskStatus | "All";

const priorityOrder: TaskPriority[] = ["High", "Medium", "Low"];
const statusOptions: TaskStatus[] = ["Not Started", "In Progress", "Completed"];

function formatDateTime(value: string | null) {
  if (!value) return "Never";
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short"
  }).format(new Date(value));
}

function formatDate(value: string | null) {
  if (!value) return "Not set";
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium"
  }).format(new Date(value));
}

function formatMinutesAsHours(minutes: number) {
  const hours = minutes / 60;
  return `${hours % 1 === 0 ? hours.toFixed(0) : hours.toFixed(1)}h`;
}

function formatPercent(value: number) {
  return `${Math.round(value * 100)}%`;
}

function formatPreferenceLines(values: string[]) {
  return values.join("\n");
}

function parsePreferenceLines(value: string) {
  return [...new Set(value.split(/\n|,/).map((item) => item.trim()).filter(Boolean))];
}

function getBrowserTimeZone() {
  return Intl.DateTimeFormat().resolvedOptions().timeZone || undefined;
}

function parseMeetingDateValue(value: string, sourceTimeZone?: string | null) {
  const hasZone = /[zZ]|[+-]\d{2}:\d{2}$/.test(value);
  const normalizedSourceTimeZone = sourceTimeZone?.trim().toUpperCase();
  return {
    date: new Date(
      hasZone || normalizedSourceTimeZone !== "UTC" ? value : `${value}Z`
    )
  };
}

function meetingInstant(meeting: TodayResponse["meetings"][number], edge: "start" | "end" = "start") {
  return parseMeetingDateValue(edge === "start" ? meeting.startTime : meeting.endTime, meeting.timeZone).date;
}

function meetingActionLabel(meeting: TodayResponse["meetings"][number]) {
  if (meeting.meetingLinkType === "join") return "Join meeting";
  if (meeting.meetingLinkType === "calendar") return "Open in Calendar";
  return null;
}

function getUpcomingMeetingId(meetings: TodayResponse["meetings"]) {
  const now = Date.now();
  const next = meetings.find((meeting) => !meeting.isCancelled && meetingInstant(meeting, "end").getTime() >= now);
  return next?.id ?? meetings.find((meeting) => !meeting.isCancelled)?.id ?? meetings[0]?.id ?? null;
}

function getUpcomingJoinableMeetingId(meetings: TodayResponse["meetings"]) {
  const now = Date.now();
  const nextJoinable = meetings.find(
    (meeting) =>
      !meeting.isCancelled &&
      meeting.meetingLinkType === "join" &&
      Boolean(meeting.meetingLink) &&
      meetingInstant(meeting, "end").getTime() >= now
  );
  return nextJoinable?.id ?? null;
}

function getMeetingDayKey(meeting: TodayResponse["meetings"][number]) {
  const parsed = parseMeetingDateValue(meeting.startTime, meeting.timeZone);
  return new Intl.DateTimeFormat("en-CA", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(parsed.date);
}

function formatMeetingTime(value: string, sourceTimeZone?: string | null) {
  const parsed = parseMeetingDateValue(value, sourceTimeZone);
  return new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit"
  }).format(parsed.date);
}

function formatMeetingDayLabel(value: string, sourceTimeZone?: string | null) {
  const parsed = parseMeetingDateValue(value, sourceTimeZone);
  const meetingDate = parsed.date;
  const today = new Date();
  const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const startOfMeeting = new Date(meetingDate.getFullYear(), meetingDate.getMonth(), meetingDate.getDate());
  const diffDays = Math.round((startOfMeeting.getTime() - startOfToday.getTime()) / 86_400_000);

  if (diffDays === 0) return "Today";
  if (diffDays === -1) return "Yesterday";
  if (diffDays === 1) return "Tomorrow";

  return new Intl.DateTimeFormat(undefined, {
    weekday: "long",
    month: "short",
    day: "numeric"
  }).format(meetingDate);
}

function formatMeetingDateStamp(value: string, sourceTimeZone?: string | null) {
  const parsed = parseMeetingDateValue(value, sourceTimeZone);
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric"
  }).format(parsed.date);
}

function groupMeetingsByDay(meetings: TodayResponse["meetings"]) {
  const groups: Array<{ key: string; label: string; stamp: string; meetings: TodayResponse["meetings"] }> = [];

  for (const meeting of meetings) {
    const dateKey = getMeetingDayKey(meeting);
    const existing = groups.find((group) => group.key === dateKey);
    if (existing) {
      existing.meetings.push(meeting);
      continue;
    }

    groups.push({
      key: dateKey,
      label: formatMeetingDayLabel(meeting.startTime, meeting.timeZone),
      stamp: formatMeetingDateStamp(meeting.startTime, meeting.timeZone),
      meetings: [meeting]
    });
  }

  return groups;
}

function dayPlanBlockStatusLabel(block: TodayResponse["dayPlan"]["blocks"][number]) {
  switch (block.status) {
    case "in_progress":
      return block.kind === "meeting" ? "Live" : "In progress";
    case "up_next":
      return "Up next";
    case "completed":
      return "Completed";
    case "ended":
      return "Ended";
    default:
      return block.kind === "meeting" ? "Planned" : "Planned";
  }
}

function dayPlanBlockActionLabel(block: TodayResponse["dayPlan"]["blocks"][number]) {
  if (!block.link) return null;
  if (block.kind === "meeting") {
    return block.note?.toLowerCase().includes("join") ? "Join now" : "Open in Calendar";
  }
  return "Open source";
}

function sourceLabel(task: Task) {
  if (task.source === "Jira" && task.jiraStatus) {
    return task.jiraStatus;
  }
  return task.status;
}

function nextJiraSubtask(task: Task) {
  return (
    task.jiraPlanningSubtasks.find((subtask) =>
      /(progress|coding|review|testing|qa|blocked|in dev|development)/i.test(subtask.status ?? "")
    ) ??
    task.jiraPlanningSubtasks[0] ??
    null
  );
}

function jiraStorySummary(task: Task) {
  if (task.source !== "Jira") return null;
  const nextSubtask = nextJiraSubtask(task);
  const openCount = task.jiraPlanningSubtasks.length;
  const openHours =
    task.jiraSubtaskEstimateSeconds && task.jiraSubtaskEstimateSeconds > 0
      ? `${(task.jiraSubtaskEstimateSeconds / 3600).toFixed(task.jiraSubtaskEstimateSeconds % 3600 === 0 ? 0 : 1)}h`
      : null;
  const parts = [`Story ${task.sourceRef ?? "Jira item"}`];
  if (openCount > 0) {
    parts.push(`${openCount} open subtask${openCount === 1 ? "" : "s"}`);
  }
  if (nextSubtask) {
    parts.push(`Next ${nextSubtask.key}`);
  }
  if (openHours) {
    parts.push(`${openHours} remaining`);
  }
  return parts.join(" • ");
}

function compareTasks(left: Task, right: Task) {
  return new Date(right.updatedAt ?? 0).getTime() - new Date(left.updatedAt ?? 0).getTime();
}

function simplifyEmailTaskTitle(title: string) {
  return title
    .replace(/\s*\[[^\]]*account:\s*\d+[^\]]*\]/gi, "")
    .replace(/\b\d{10,}\b/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function emailGroupKey(task: Task) {
  if (task.source !== "Email") return null;
  const simplified = simplifyEmailTaskTitle(task.title);
  return simplified ? `email:${simplified.toLowerCase()}` : null;
}

type TaskPresentationItem =
  | { kind: "single"; key: string; task: Task }
  | { kind: "cluster"; key: string; title: string; tasks: Task[] };

function buildTaskPresentationItems(tasks: Task[]) {
  const ordered = [...tasks].sort(compareTasks);
  const emailGroups = new Map<string, Task[]>();

  for (const task of ordered) {
    const key = emailGroupKey(task);
    if (!key) continue;
    const current = emailGroups.get(key) ?? [];
    current.push(task);
    emailGroups.set(key, current);
  }

  const seen = new Set<string>();
  const items: TaskPresentationItem[] = [];

  for (const task of ordered) {
    const key = emailGroupKey(task);
    const grouped = key ? emailGroups.get(key) ?? [] : [];
    if (key && grouped.length > 1) {
      if (seen.has(key)) continue;
      seen.add(key);
      items.push({
        kind: "cluster",
        key,
        title: simplifyEmailTaskTitle(grouped[0]?.title ?? task.title),
        tasks: grouped
      });
      continue;
    }

    items.push({ kind: "single", key: `task:${task.id}`, task });
  }

  return items;
}

function groupTasksByPriority(tasks: Task[]) {
  return {
    High: tasks.filter((task) => task.priority === "High").sort(compareTasks),
    Medium: tasks.filter((task) => task.priority === "Medium").sort(compareTasks),
    Low: tasks.filter((task) => task.priority === "Low").sort(compareTasks)
  } satisfies Record<TaskPriority, Task[]>;
}

function flattenTaskGroups(groups: TodayResponse["tasks"]) {
  return [...groups.High, ...groups.Medium, ...groups.Low];
}

function joinList(values: string[]) {
  return values.length ? values.join(", ") : "None";
}

function hasText(value: string | null | undefined) {
  return Boolean(value?.trim());
}

function hasItems(values: string[] | null | undefined) {
  return Boolean(values?.length);
}

function EmailMetaField(props: { label: string; value?: string | null; values?: string[]; compact?: boolean }) {
  const items = props.values?.filter((entry) => entry.trim()) ?? [];
  const text = props.value?.trim() ?? "";
  const content = text || (items.length ? items.join(", ") : "");

  if (!content) {
    return null;
  }

  return (
    <div className={props.compact ? "email-meta-field compact" : "email-meta-field"}>
      <span>{props.label}</span>
      <div className="email-meta-value">{content}</div>
    </div>
  );
}

function AppHeader(props: { active: View; onChange: (view: View) => void }) {
  return (
    <aside className="sidebar">
      <div className="sidebar-top">
        <p className="eyebrow">AI Day Planner</p>
        <h1>Focus your workday with calm structure and sharp priorities.</h1>
        <p className="sidebar-copy">
          A refined control center for meetings, Jira work, and email follow-ups.
        </p>
      </div>
      <nav className="nav">
        {[
          ["today", "Today"],
          ["tasks", "Tasks"],
          ["deferred", "Deferred"],
          ["rejected", "Rejected"],
          ["reminders", "Reminders"],
          ["settings", "Settings"]
        ].map(([id, label]) => (
          <button
            key={id}
            className={props.active === id ? "nav-link active" : "nav-link"}
            onClick={() => props.onChange(id as View)}
          >
            <span className="nav-link-label">{label}</span>
          </button>
        ))}
      </nav>
    </aside>
  );
}

function IconSyncButton(props: { label: string; loading?: boolean; onClick: () => Promise<void> }) {
  return (
    <button className="icon-button" onClick={() => void props.onClick()} disabled={props.loading} title={props.label}>
      <span className={props.loading ? "spin" : ""}>↻</span>
      <span>{props.loading ? "Syncing..." : props.label}</span>
    </button>
  );
}

function TodaySkeleton() {
  return (
    <section className="panel-stack">
      <div className="hero-card skeleton-block hero-skeleton" />
      <div className="panel skeleton-panel">
        <div className="skeleton-line wide" />
        <div className="skeleton-line medium" />
        <div className="skeleton-stack">
          <div className="skeleton-card" />
          <div className="skeleton-card" />
        </div>
      </div>
      <div className="dashboard-grid">
        <div className="panel tall-panel skeleton-panel">
          <div className="skeleton-line wide" />
          <div className="skeleton-line medium" />
          <div className="skeleton-stack">
            <div className="skeleton-card tall" />
            <div className="skeleton-card tall" />
            <div className="skeleton-card tall" />
          </div>
        </div>
        <div className="panel tall-panel skeleton-panel">
          <div className="skeleton-line wide" />
          <div className="skeleton-line short" />
          <div className="skeleton-stack">
            <div className="skeleton-card" />
            <div className="skeleton-card" />
            <div className="skeleton-card" />
          </div>
        </div>
      </div>
    </section>
  );
}

function DayPlanPanel(props: {
  data: TodayResponse;
  onOpenDetails: (task: Task) => Promise<void>;
}) {
  const taskLookup = new Map(flattenTaskGroups(props.data.tasks).map((task) => [task.id, task]));
  const summary = props.data.dayPlan.summary;

  return (
    <div className="panel day-plan-panel">
      <div className="panel-header">
        <div>
          <h3>Daily Plan</h3>
          <p className="timeline-header-note">{summary.guidance}</p>
        </div>
        <div className="panel-header-actions">
          <span>{summary.dayKey}</span>
          <span className="day-plan-factor">Pattern fit {formatPercent(summary.focusFactor)}</span>
        </div>
      </div>

      <div className="day-plan-summary-grid">
        <div className="overview-card day-plan-summary-card">
          <span>Capacity</span>
          <strong>{formatMinutesAsHours(summary.adaptedTaskCapacityMinutes)}</strong>
          <p>{formatMinutesAsHours(summary.remainingTaskCapacityMinutes)} still available for task work</p>
        </div>
        <div className="overview-card day-plan-summary-card">
          <span>Completed today</span>
          <strong>{formatMinutesAsHours(summary.completedTaskMinutes)}</strong>
          <p>{formatPercent(summary.completionRate)} of today’s planned task load already closed</p>
        </div>
        <div className="overview-card day-plan-summary-card">
          <span>Planned next</span>
          <strong>{formatMinutesAsHours(summary.plannedTaskMinutes)}</strong>
          <p>{props.data.dayPlan.blocks.filter((block) => block.kind === "task").length} work blocks arranged around meetings</p>
        </div>
        <div className="overview-card day-plan-summary-card">
          <span>Spillover</span>
          <strong>{summary.spilloverTaskCount}</strong>
          <p>{formatMinutesAsHours(summary.freeMinutes)} flexible time remains after the scheduled blocks</p>
        </div>
      </div>

      <div className="day-plan-list">
        {props.data.dayPlan.blocks.length ? (
          props.data.dayPlan.blocks.map((block) => {
            const task = block.taskId !== null ? taskLookup.get(block.taskId) ?? null : null;
            const actionLabel = dayPlanBlockActionLabel(block);
            return (
              <div
                key={block.id}
                className={`day-plan-item day-plan-item-${block.kind} day-plan-item-${block.status}`}
              >
                <div className="day-plan-time">
                  <strong>{formatMeetingTime(block.startTime, block.timeZone)}</strong>
                  <span>{block.durationMinutes} min</span>
                </div>
                <div className="day-plan-body">
                  <div className="day-plan-title-row">
                    <strong>{block.title}</strong>
                    <span className={`timeline-state ${block.status === "planned" ? "queued" : block.status}`}>
                      {dayPlanBlockStatusLabel(block)}
                    </span>
                    {block.priority ? (
                      <span className={`priority-pill priority-pill-${block.priority.toLowerCase()}`}>{block.priority}</span>
                    ) : null}
                  </div>
                  <p>
                    {formatMeetingTime(block.startTime, block.timeZone)} to{" "}
                    {formatMeetingTime(block.endTime, block.timeZone)}
                  </p>
                  {block.note ? <p className="day-plan-note">{block.note}</p> : null}
                  <div className="day-plan-actions">
                    {task ? (
                      <button className="ghost-button subtle-action" onClick={() => void props.onOpenDetails(task)}>
                        View task
                      </button>
                    ) : null}
                    {block.link && actionLabel ? (
                      <a className="source-link" href={block.link} target="_blank" rel="noreferrer">
                        {actionLabel}
                      </a>
                    ) : null}
                  </div>
                </div>
              </div>
            );
          })
        ) : (
          <p className="empty-state">No remaining work blocks are needed today.</p>
        )}
      </div>

      {props.data.dayPlan.spilloverTasks.length ? (
        <div className="day-plan-spillover">
          <div className="task-group-header">
            <h4>Spillover candidates</h4>
            <span>{props.data.dayPlan.spilloverTasks.length}</span>
          </div>
          <div className="day-plan-spillover-list">
            {props.data.dayPlan.spilloverTasks.slice(0, 4).map((task) => (
              <button key={task.id} className="day-plan-spillover-item" onClick={() => void props.onOpenDetails(task)}>
                <strong>{task.title}</strong>
                {task.source === "Jira" && jiraStorySummary(task) ? (
                  <span>{jiraStorySummary(task)}</span>
                ) : null}
                <span>
                  {task.priority} • {task.estimatedEffortBucket ?? "15 min"}
                </span>
              </button>
            ))}
          </div>
        </div>
      ) : null}
    </div>
  );
}

function TasksSkeleton() {
  return (
    <section className="panel-stack">
      <div className="panel skeleton-panel">
        <div className="skeleton-line wide" />
        <div className="skeleton-line medium" />
        <div className="skeleton-card" />
        <div className="skeleton-stack">
          <div className="skeleton-card" />
          <div className="skeleton-card" />
          <div className="skeleton-card" />
        </div>
      </div>
    </section>
  );
}

function SettingsSkeleton() {
  return (
    <section className="panel-stack">
      <div className="settings-grid">
        <div className="panel integration-card skeleton-panel">
          <div className="skeleton-line wide" />
          <div className="skeleton-line medium" />
          <div className="skeleton-stack">
            <div className="skeleton-line short" />
            <div className="skeleton-line short" />
            <div className="skeleton-card" />
          </div>
        </div>
        <div className="panel integration-card skeleton-panel">
          <div className="skeleton-line wide" />
          <div className="skeleton-line medium" />
          <div className="skeleton-stack">
            <div className="skeleton-card" />
            <div className="skeleton-card" />
            <div className="skeleton-card" />
          </div>
        </div>
      </div>
    </section>
  );
}

function StatusSelect(props: {
  value: TaskStatus;
  onChange: (status: TaskStatus) => Promise<void>;
  compact?: boolean;
  disabled?: boolean;
}) {
  return (
    <label className={props.compact ? "status-select compact" : "status-select"}>
      <span>Status</span>
      <select
        value={props.value}
        disabled={props.disabled}
        onChange={(event) => void props.onChange(event.target.value as TaskStatus)}
      >
        {statusOptions.map((status) => (
          <option key={status} value={status}>
            {status}
          </option>
        ))}
      </select>
    </label>
  );
}

function PrioritySelect(props: {
  value: TaskPriority;
  onChange: (priority: TaskPriority) => Promise<void>;
  compact?: boolean;
  disabled?: boolean;
}) {
  return (
    <label className={props.compact ? "status-select compact" : "status-select"}>
      <span>Priority</span>
      <select
        value={props.value}
        disabled={props.disabled}
        onChange={(event) => void props.onChange(event.target.value as TaskPriority)}
      >
        {priorityOrder.map((priority) => (
          <option key={priority} value={priority}>
            {priority}
          </option>
        ))}
      </select>
    </label>
  );
}

function JiraDetailView(props: { detail: JiraTaskDetail }) {
  return (
    <div className="detail-stack">
      <div className="detail-grid">
        <div className="detail-stat">
          <span>Status</span>
          <strong>{props.detail.status ?? "Unknown"}</strong>
        </div>
        <div className="detail-stat">
          <span>Priority</span>
          <strong>{props.detail.priority ?? "Unknown"}</strong>
        </div>
        <div className="detail-stat">
          <span>Story points</span>
          <strong>{props.detail.storyPoints ?? "Not set"}</strong>
        </div>
        <div className="detail-stat">
          <span>Assignee</span>
          <strong>{props.detail.assignee ?? "Unassigned"}</strong>
        </div>
      </div>

      <section className="detail-section">
        <h4>Description</h4>
        <pre className="detail-content">{props.detail.description ?? "No description available."}</pre>
      </section>

      <section className="detail-section">
        <h4>Context</h4>
        <div className="detail-grid compact">
          <div className="detail-stat">
            <span>Reporter</span>
            <strong>{props.detail.reporter ?? "Unknown"}</strong>
          </div>
          <div className="detail-stat">
            <span>Labels</span>
            <strong>{props.detail.labels.length ? props.detail.labels.join(", ") : "None"}</strong>
          </div>
        </div>
      </section>

      <section className="detail-section">
        <h4>Subtasks</h4>
        {props.detail.subtasks.length ? (
          <div className="detail-list">
            {props.detail.subtasks.map((subtask) => (
              <article className="detail-row" key={subtask.key}>
                <strong>{subtask.key}</strong>
                <p>{subtask.title}</p>
                <span className="subtle-pill">{subtask.status ?? "Unknown"}</span>
              </article>
            ))}
          </div>
        ) : (
          <p className="empty-state">No subtasks.</p>
        )}
      </section>

      <section className="detail-section">
        <h4>Comments</h4>
        {props.detail.comments.length ? (
          <div className="detail-list">
            {props.detail.comments.map((comment, index) => (
              <article className="detail-row stacked" key={`${comment.author}-${index}`}>
                <div className="detail-row-header">
                  <strong>{comment.author}</strong>
                  <span>{formatDateTime(comment.createdAt)}</span>
                </div>
                <pre className="detail-content">{comment.body || "No comment body."}</pre>
              </article>
            ))}
          </div>
        ) : (
          <p className="empty-state">No comments.</p>
        )}
      </section>

      <section className="detail-section">
        <h4>Worklog</h4>
        {props.detail.worklogs.length ? (
          <div className="detail-list">
            {props.detail.worklogs.map((worklog, index) => (
              <article className="detail-row stacked" key={`${worklog.author}-${index}`}>
                <div className="detail-row-header">
                  <strong>{worklog.author}</strong>
                  <span>
                    {worklog.timeSpent ?? "No time"} • {formatDateTime(worklog.startedAt)}
                  </span>
                </div>
                <pre className="detail-content">{worklog.comment ?? "No worklog comment."}</pre>
              </article>
            ))}
          </div>
        ) : (
          <p className="empty-state">No worklog entries.</p>
        )}
      </section>
    </div>
  );
}

function EmailDetailView(props: { detail: EmailTaskDetail }) {
  const showThread = props.detail.thread.length > 0;
  const hasFrom = hasText(props.detail.from);
  const hasTo = hasItems(props.detail.to);
  const hasCc = hasItems(props.detail.cc);
  const hasReceived = hasText(props.detail.receivedAt);

  return (
    <div className="detail-stack">
      {hasFrom || hasTo || hasCc || hasReceived ? (
        <section className="email-meta-panel">
          {hasFrom ? <EmailMetaField label="From" value={props.detail.from} /> : null}
          {hasTo ? <EmailMetaField label="To" values={props.detail.to} /> : null}
          {hasCc || hasReceived ? (
            <div className="email-meta-secondary">
              {hasCc ? <EmailMetaField label="CC" values={props.detail.cc} compact /> : null}
              {hasReceived ? (
                <EmailMetaField label="Received" value={formatDateTime(props.detail.receivedAt)} compact />
              ) : null}
            </div>
          ) : null}
        </section>
      ) : null}

      {hasText(props.detail.subject) ? (
        <section className="detail-section">
          <h4>Subject</h4>
          <pre className="detail-content">{props.detail.subject}</pre>
        </section>
      ) : null}

      <section className="detail-section">
        <h4>Email content</h4>
        <pre className="detail-content">{props.detail.body || "No message content available."}</pre>
      </section>

      {showThread ? (
        <section className="detail-section">
          <h4>Previous thread emails</h4>
          <div className="detail-list">
            {props.detail.thread.map((message) => (
              <article className="detail-row stacked" key={message.id}>
                <div className="detail-row-header">
                  <strong>{message.subject ?? "No subject"}</strong>
                  <span>{formatDateTime(message.receivedAt)}</span>
                </div>
                {hasText(message.from) || hasItems(message.to) || hasItems(message.cc) ? (
                  <div className="detail-thread-meta">
                    {hasText(message.from) ? (
                      <div className="thread-meta-row">
                        <span className="thread-meta-label">From</span>
                        <span>{message.from}</span>
                      </div>
                    ) : null}
                    {hasItems(message.to) ? (
                      <div className="thread-meta-row">
                        <span className="thread-meta-label">To</span>
                        <span>{joinList(message.to)}</span>
                      </div>
                    ) : null}
                    {hasItems(message.cc) ? (
                      <div className="thread-meta-row">
                        <span className="thread-meta-label">CC</span>
                        <span>{joinList(message.cc)}</span>
                      </div>
                    ) : null}
                  </div>
                ) : null}
                <pre className="detail-content">{message.body || "No message content available."}</pre>
              </article>
            ))}
          </div>
        </section>
      ) : null}
    </div>
  );
}

function TaskDetailsDialog(props: {
  task: Task | null;
  detail: TaskDetail | null;
  loading: boolean;
  error: string | null;
  onClose: () => void;
}) {
  if (!props.task) return null;

  return (
    <div className="detail-overlay" onClick={props.onClose}>
      <div className="detail-modal" onClick={(event) => event.stopPropagation()}>
        <div className="detail-modal-header">
          <div>
            <p className="eyebrow">Task Details</p>
            <h3>{props.task.title}</h3>
          </div>
          <button className="ghost-button subtle-action" onClick={props.onClose}>
            Close
          </button>
        </div>

        {props.loading ? <p className="muted">Loading source details…</p> : null}
        {props.error ? <p className="error-text">{props.error}</p> : null}
        {!props.loading && !props.error && props.detail?.type === "jira" ? <JiraDetailView detail={props.detail} /> : null}
        {!props.loading && !props.error && props.detail?.type === "email" ? <EmailDetailView detail={props.detail} /> : null}
      </div>
    </div>
  );
}

function TaskCard(props: {
  task: Task;
  onStatusChange: (task: Task, status: TaskStatus) => Promise<void>;
  onPriorityChange: (task: Task, priority: TaskPriority) => Promise<void>;
  onOpenDetails?: (task: Task) => Promise<void>;
  dense?: boolean;
  showPriority?: boolean;
  onDelete?: (task: Task) => Promise<void>;
  draggable?: boolean;
  onDragStart?: (task: Task) => void;
  onDragEnd?: () => void;
  onDeferUntilTomorrow?: (task: Task) => Promise<void>;
  onBringBackNow?: (task: Task) => Promise<void>;
  disableControls?: boolean;
}) {
  const canOpenDetails = Boolean(props.onOpenDetails && props.task.source !== "Manual");

  return (
    <article
      className={props.dense ? "task-card dense" : "task-card"}
      draggable={props.draggable}
      onDragStart={() => props.onDragStart?.(props.task)}
      onDragEnd={() => props.onDragEnd?.()}
    >
      <div
        className={canOpenDetails ? "task-main task-main-clickable" : "task-main"}
        onClick={canOpenDetails ? () => void props.onOpenDetails?.(props.task) : undefined}
        onKeyDown={
          canOpenDetails
            ? (event) => {
                if (event.key === "Enter" || event.key === " ") {
                  event.preventDefault();
                  void props.onOpenDetails?.(props.task);
                }
              }
            : undefined
        }
        role={canOpenDetails ? "button" : undefined}
        tabIndex={canOpenDetails ? 0 : undefined}
      >
        <div className="task-meta">
          <span className={`pill pill-${props.task.source.toLowerCase()}`}>{props.task.source}</span>
          <span className="subtle-pill">{sourceLabel(props.task)}</span>
          {props.showPriority !== false ? (
            <span className={`priority-dot ${props.task.priority.toLowerCase()}`}>{props.task.priority}</span>
          ) : null}
        </div>
        <strong>{props.task.title}</strong>
        {props.task.source === "Jira" && jiraStorySummary(props.task) ? (
          <p className="task-story-context">{jiraStorySummary(props.task)}</p>
        ) : null}
        {props.task.priorityExplanation ? (
          <p className="task-why">Why now: {props.task.priorityExplanation}</p>
        ) : null}
        <div className="task-link-row">
          {props.task.sourceLink ? (
            <a
              className="source-link"
              href={props.task.sourceLink}
              target="_blank"
              rel="noreferrer"
              onClick={(event) => event.stopPropagation()}
            >
              Open source
            </a>
          ) : (
            <span className="muted small-text">No source link</span>
          )}
          {props.task.estimatedEffortBucket ? <span className="subtle-pill">{props.task.estimatedEffortBucket}</span> : null}
          {props.task.carryForwardCount > 0 ? (
            <span className="subtle-pill">Carry-forward {props.task.carryForwardCount}</span>
          ) : null}
          {props.task.deferredUntil ? <span className="subtle-pill">Deferred until {formatDate(props.task.deferredUntil)}</span> : null}
        </div>
      </div>
      <div className="task-actions">
        {props.showPriority !== false ? (
          <PrioritySelect
            value={props.task.priority}
            compact={props.dense}
            disabled={props.disableControls}
            onChange={(priority) => props.onPriorityChange(props.task, priority)}
          />
        ) : null}
        <StatusSelect
          value={props.task.status}
          compact={props.dense}
          disabled={props.disableControls}
          onChange={(status) => props.onStatusChange(props.task, status)}
        />
        {props.onDelete ? (
          <button className="ghost-button subtle-action" onClick={() => void props.onDelete?.(props.task)}>
            {props.task.source === "Manual" ? "Remove" : "Reject"}
          </button>
        ) : null}
        {props.onDeferUntilTomorrow && !props.task.deferredUntil ? (
          <button className="ghost-button subtle-action" onClick={() => void props.onDeferUntilTomorrow?.(props.task)}>
            Defer to tomorrow
          </button>
        ) : null}
        {props.onBringBackNow && props.task.deferredUntil ? (
          <button className="ghost-button subtle-action" onClick={() => void props.onBringBackNow?.(props.task)}>
            Bring back now
          </button>
        ) : null}
      </div>
    </article>
  );
}

function TaskClusterCard(props: {
  title: string;
  tasks: Task[];
  onStatusChange: (task: Task, status: TaskStatus) => Promise<void>;
  onPriorityChange: (task: Task, priority: TaskPriority) => Promise<void>;
  onDelete?: (task: Task) => Promise<void>;
  onOpenDetails?: (task: Task) => Promise<void>;
  onDeferUntilTomorrow?: (task: Task) => Promise<void>;
  onBringBackNow?: (task: Task) => Promise<void>;
  disableControls?: boolean;
}) {
  return (
    <section className="task-cluster">
      <div className="task-cluster-header">
        <div>
          <div className="task-meta">
            <span className="pill pill-email">Email</span>
            <span className="subtle-pill">{props.tasks.length} related messages</span>
          </div>
          <strong>{props.title}</strong>
          <p className="muted">Grouped together as one stream of work while each email keeps its own controls.</p>
        </div>
      </div>
      <div className="task-cluster-list">
        {props.tasks.map((task) => (
          <TaskCard
            key={task.id}
            task={task}
            dense
            onStatusChange={(currentTask, status) => props.onStatusChange(currentTask, status)}
            onPriorityChange={(currentTask, priority) => props.onPriorityChange(currentTask, priority)}
            onDelete={props.onDelete ? (currentTask) => props.onDelete!(currentTask) : undefined}
            onOpenDetails={props.onOpenDetails}
            onDeferUntilTomorrow={props.onDeferUntilTomorrow}
            onBringBackNow={props.onBringBackNow}
            disableControls={props.disableControls}
          />
        ))}
      </div>
    </section>
  );
}

function TodayView(props: {
  data: TodayResponse | null;
  loading: boolean;
  onGenerate: () => Promise<void>;
  onSyncMeetings: () => Promise<void>;
  onSyncTasks: () => Promise<void>;
  syncMeetingsLoading: boolean;
  syncTasksLoading: boolean;
  onTaskStatusChange: (task: Task, status: TaskStatus) => Promise<void>;
  onTaskPriorityChange: (task: Task, priority: TaskPriority) => Promise<void>;
  onOpenDetails: (task: Task) => Promise<void>;
}) {
  const [draggedTask, setDraggedTask] = useState<Task | null>(null);
  const [dropPriority, setDropPriority] = useState<TaskPriority | null>(null);
  const timelineRef = useRef<HTMLDivElement | null>(null);
  const meetingGroups = props.data ? groupMeetingsByDay(props.data.meetings) : [];
  const meetingTimeZone = props.data?.meetings.find((meeting) => meeting.timeZone)?.timeZone ?? null;
  const orderedMeetings = props.data?.meetings ?? [];
  const focusedMeetingId = getUpcomingMeetingId(orderedMeetings);
  const joinableMeetingId = getUpcomingJoinableMeetingId(orderedMeetings);
  const focusedMeeting = focusedMeetingId
    ? orderedMeetings.find((meeting) => meeting.id === focusedMeetingId) ?? null
    : null;
  const upcomingDayKey = focusedMeeting ? getMeetingDayKey(focusedMeeting) : null;

  useEffect(() => {
    if (!upcomingDayKey || !timelineRef.current) return;
    const section = timelineRef.current.querySelector<HTMLElement>(`[data-day-key="${upcomingDayKey}"]`);
    if (!section) return;
    timelineRef.current.scrollTo({
      top: Math.max(0, section.offsetTop - 10),
      behavior: "smooth"
    });
  }, [upcomingDayKey]);

  return (
    <section className="panel-stack">
      <div className="hero-card">
        <div className="hero-copy">
          <p className="eyebrow">Today Dashboard</p>
          <h2>Your day, resolved into what matters now.</h2>
          <p className="muted">
            Generate a fresh plan from Outlook, Calendar, and Jira whenever you need a reset.
          </p>
        </div>
        <button className="primary-button hero-button" onClick={() => props.onGenerate()} disabled={props.loading}>
          {props.loading ? "Refreshing plan..." : "Generate Today's Plan"}
        </button>
      </div>

      {props.data?.warnings.length ? (
        <div className="warning-list">
          {props.data.warnings.map((warning) => (
            <div key={warning} className="warning-item">
              {warning}
            </div>
          ))}
        </div>
      ) : null}

      {props.data ? (
        <div className="overview-strip">
          <div className="overview-card">
            <span>Workload</span>
            <strong>{props.data.workload.state}</strong>
            <p>
              {formatMinutesAsHours(props.data.workload.totalTaskMinutes)} tasks • {formatMinutesAsHours(props.data.workload.totalMeetingMinutes)} meetings
            </p>
          </div>
          <div className="overview-card">
            <span>Reminder center</span>
            <strong>{props.data.reminders.filter((item) => item.status === "active").length} active</strong>
            <p>Follow-ups, deferred work, and meeting prep in one place.</p>
          </div>
          <div className="overview-card">
            <span>Deferred queue</span>
            <strong>{props.data.deferredTaskCount} hidden</strong>
            <p>Deferred tasks stay out of the active plan until they are due.</p>
          </div>
          <div className="overview-card">
            <span>Rejected queue</span>
            <strong>{props.data.rejectedTaskCount} hidden</strong>
            <p>Rejected tasks remain recoverable with explanations and feedback controls.</p>
          </div>
        </div>
      ) : null}

      {props.data ? <DayPlanPanel data={props.data} onOpenDetails={props.onOpenDetails} /> : null}

      <div className="dashboard-grid">
        <div className="panel tall-panel">
          <div className="panel-header">
            <div>
              <h3>Meetings Timeline</h3>
              <p className="timeline-header-note">
                {meetingTimeZone ? `All times shown in ${meetingTimeZone}` : "All times shown in your calendar timezone"}
              </p>
            </div>
            <div className="panel-header-actions">
              <span>{props.data?.meetings.length ?? 0} across 7 days</span>
              <IconSyncButton
                label="Sync meetings"
                loading={props.syncMeetingsLoading}
                onClick={props.onSyncMeetings}
              />
            </div>
          </div>
          <div className="timeline" ref={timelineRef}>
            {meetingGroups.length ? (
              meetingGroups.map((group) => (
                <section
                  className={group.key === upcomingDayKey ? "timeline-day timeline-day-upcoming" : "timeline-day"}
                  key={group.key}
                  data-day-key={group.key}
                >
                  <div className="timeline-day-header">
                    <div>
                      <h4>{group.label}</h4>
                      <p>{group.meetings.length} meeting{group.meetings.length === 1 ? "" : "s"}</p>
                    </div>
                    <span>{group.stamp}</span>
                  </div>
                  <div className="timeline-day-list">
                    {group.meetings.map((meeting) => (
                      <div
                        className={
                          meeting.isCancelled
                            ? meeting.id === focusedMeetingId
                              ? "timeline-item cancelled focused"
                              : "timeline-item cancelled"
                            : meetingInstant(meeting, "end").getTime() < Date.now()
                              ? "timeline-item ended"
                            : meeting.id === focusedMeetingId
                              ? "timeline-item focused"
                              : "timeline-item"
                        }
                        key={meeting.id}
                        id={`meeting-${meeting.id}`}
                      >
                        <div className="timeline-rail">
                          <div className="timeline-dot" />
                          <div className="timeline-line" />
                        </div>
                        <div className="timeline-copy">
                          <div className="timeline-time-block">
                            <strong>{formatMeetingTime(meeting.startTime, meeting.timeZone)}</strong>
                            <span>{meeting.durationMinutes} min</span>
                          </div>
                          <div className="timeline-body">
                            <div className="timeline-title-row">
                              <strong>{meeting.title}</strong>
                              {meeting.isCancelled ? <span className="timeline-state cancelled">Cancelled</span> : null}
                              {!meeting.isCancelled && meetingInstant(meeting, "end").getTime() < Date.now() ? (
                                <span className="timeline-state ended">Ended</span>
                              ) : null}
                              {!meeting.isCancelled && meeting.id === focusedMeetingId ? (
                                <span className="timeline-state live">Up next</span>
                              ) : null}
                            </div>
                            <p>
                              {formatMeetingTime(meeting.startTime, meeting.timeZone)} to{" "}
                              {formatMeetingTime(meeting.endTime, meeting.timeZone)}
                            </p>
                            {meeting.meetingLink && !meeting.isCancelled && meetingActionLabel(meeting) ? (
                              <a className="source-link" href={meeting.meetingLink} target="_blank" rel="noreferrer">
                                {meeting.id === joinableMeetingId && meeting.meetingLinkType === "join"
                                  ? "Join now"
                                  : meetingActionLabel(meeting)}
                              </a>
                            ) : null}
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </section>
              ))
            ) : (
              <p className="empty-state">No meetings scheduled in this testing window.</p>
            )}
          </div>
        </div>

        <div className="panel tall-panel">
          <div className="panel-header">
            <h3>Priority Tasks</h3>
            <div className="panel-header-actions">
              <span>Last plan: {formatDateTime(props.data?.sync.lastGeneratedAt ?? null)}</span>
              <IconSyncButton label="Sync tasks" loading={props.syncTasksLoading} onClick={props.onSyncTasks} />
            </div>
          </div>
          <div className="task-groups">
            {priorityOrder.map((priority) => (
              <div
                key={priority}
                className={dropPriority === priority ? "task-group drop-target" : "task-group"}
                onDragOver={(event) => {
                  event.preventDefault();
                  if (draggedTask) {
                    setDropPriority(priority);
                  }
                }}
                onDragLeave={() => {
                  setDropPriority((current) => (current === priority ? null : current));
                }}
                onDrop={(event) => {
                  event.preventDefault();
                  if (draggedTask && draggedTask.priority !== priority) {
                    void props.onTaskPriorityChange(draggedTask, priority);
                  }
                  setDraggedTask(null);
                  setDropPriority(null);
                }}
              >
                <div className="task-group-header">
                  <h4 className={`priority-heading priority-heading-${priority.toLowerCase()}`}>{priority}</h4>
                  <span className="priority-count">{props.data?.tasks[priority]?.length ?? 0}</span>
                </div>
                {props.data?.tasks[priority]?.length ? (
                  props.data.tasks[priority].map((task) => (
                    <TaskCard
                      key={task.id}
                      task={task}
                      onStatusChange={(currentTask, status) => props.onTaskStatusChange(currentTask, status)}
                      onPriorityChange={(currentTask, nextPriority) =>
                        props.onTaskPriorityChange(currentTask, nextPriority)
                      }
                      onOpenDetails={props.onOpenDetails}
                      showPriority={false}
                      draggable
                      onDragStart={(currentTask) => {
                        setDraggedTask(currentTask);
                        setDropPriority(currentTask.priority);
                      }}
                      onDragEnd={() => {
                        setDraggedTask(null);
                        setDropPriority(null);
                      }}
                    />
                  ))
                ) : (
                  <p className="empty-state">Nothing in {priority.toLowerCase()} priority right now.</p>
                )}
              </div>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}

function TasksView(props: {
  tasks: Task[];
  loading: boolean;
  filter: TaskFilter;
  onFilterChange: (filter: TaskFilter) => void;
  onCreate: (title: string) => Promise<void>;
  onUpdateStatus: (task: Task, status: TaskStatus) => Promise<void>;
  onUpdatePriority: (task: Task, priority: TaskPriority) => Promise<void>;
  onDelete: (task: Task) => Promise<void>;
  onOpenDetails: (task: Task) => Promise<void>;
  onDeferUntilTomorrow: (task: Task) => Promise<void>;
}) {
  const [title, setTitle] = useState("");
  const items = useMemo(() => buildTaskPresentationItems(props.tasks), [props.tasks]);

  return (
    <section className="panel-stack">
      <div className="panel">
        <div className="tasks-header">
          <div>
            <h3>Task List</h3>
            <p className="muted">Manage manual items and source-backed tasks without leaving the planner.</p>
          </div>
          <div className="filter-shell">
            <label className="status-select compact">
              <span>Filter</span>
              <select value={props.filter} onChange={(e) => props.onFilterChange(e.target.value as TaskFilter)}>
                <option value="All">All statuses</option>
                {statusOptions.map((status) => (
                  <option key={status} value={status}>
                    {status}
                  </option>
                ))}
              </select>
            </label>
          </div>
        </div>

        <form
          className="create-task-bar"
          onSubmit={async (event) => {
            event.preventDefault();
            if (!title.trim()) return;
            await props.onCreate(title.trim());
            setTitle("");
          }}
        >
          <input
            value={title}
            onChange={(event) => setTitle(event.target.value)}
            placeholder="Capture a manual task quickly"
          />
          <button className="primary-button" type="submit">
            Add Task
          </button>
        </form>

        <div className="task-table">
          {props.loading ? <p className="muted">Loading tasks…</p> : null}
          {items.map((item) =>
            item.kind === "cluster" ? (
              <TaskClusterCard
                key={item.key}
                title={item.title}
                tasks={item.tasks}
                onStatusChange={(currentTask, status) => props.onUpdateStatus(currentTask, status)}
                onPriorityChange={(currentTask, priority) => props.onUpdatePriority(currentTask, priority)}
                onDelete={(currentTask) => props.onDelete(currentTask)}
                onOpenDetails={props.onOpenDetails}
                onDeferUntilTomorrow={props.onDeferUntilTomorrow}
              />
            ) : (
              <TaskCard
                key={item.key}
                task={item.task}
                dense
                onStatusChange={(currentTask, status) => props.onUpdateStatus(currentTask, status)}
                onPriorityChange={(currentTask, priority) => props.onUpdatePriority(currentTask, priority)}
                onDelete={(currentTask) => props.onDelete(currentTask)}
                onOpenDetails={props.onOpenDetails}
                onDeferUntilTomorrow={props.onDeferUntilTomorrow}
              />
            )
          )}
          {!props.tasks.length ? <p className="empty-state">No tasks match this filter yet.</p> : null}
        </div>
      </div>
    </section>
  );
}

function DeferredView(props: {
  tasks: Task[];
  loading: boolean;
  onBringBackNow: (task: Task) => Promise<void>;
  onOpenDetails: (task: Task) => Promise<void>;
}) {
  const items = useMemo(() => buildTaskPresentationItems(props.tasks), [props.tasks]);

  return (
    <section className="panel-stack">
      <div className="panel">
        <div className="tasks-header">
          <div>
            <h3>Deferred Tasks</h3>
            <p className="muted">Tasks you intentionally moved out of the active plan until later.</p>
          </div>
        </div>
        <div className="task-table">
          {props.loading ? <p className="muted">Loading deferred tasks…</p> : null}
          {items.map((item) =>
            item.kind === "cluster" ? (
              <TaskClusterCard
                key={item.key}
                title={item.title}
                tasks={item.tasks}
                onStatusChange={async () => undefined}
                onPriorityChange={async () => undefined}
                onOpenDetails={props.onOpenDetails}
                onBringBackNow={props.onBringBackNow}
                disableControls
              />
            ) : (
              <TaskCard
                key={item.key}
                task={item.task}
                dense
                onStatusChange={async () => undefined}
                onPriorityChange={async () => undefined}
                onOpenDetails={item.task.source === "Manual" ? undefined : props.onOpenDetails}
                onBringBackNow={props.onBringBackNow}
                disableControls
              />
            )
          )}
          {!props.tasks.length ? <p className="empty-state">No deferred tasks right now.</p> : null}
        </div>
      </div>
    </section>
  );
}

function ReminderCenterView(props: {
  reminders: Reminder[];
  loading: boolean;
  onDismiss: (reminder: Reminder) => Promise<void>;
  onReactivate: (reminder: Reminder) => Promise<void>;
}) {
  return (
    <section className="panel-stack">
      <div className="panel">
        <div className="tasks-header">
          <div>
            <h3>Reminder Center</h3>
            <p className="muted">Passive follow-ups that help work stay visible without becoming noisy.</p>
          </div>
        </div>
        <div className="reminder-list">
          {props.loading ? <p className="muted">Loading reminders…</p> : null}
          {props.reminders.map((reminder) => (
            <article className={`reminder-card ${reminder.status}`} key={reminder.id}>
              <div className="task-meta">
                <span className="subtle-pill">{reminder.kind.replace(/_/g, " ")}</span>
                <span className={`status-badge ${reminder.status}`}>{reminder.status}</span>
              </div>
              <strong>{reminder.title}</strong>
              <p className="muted">{reminder.reason}</p>
              <div className="reminder-footer">
                <span className="small-text">
                  {reminder.scheduledFor ? `Scheduled ${formatDateTime(reminder.scheduledFor)}` : formatDateTime(reminder.updatedAt)}
                </span>
                <div className="integration-actions">
                  {reminder.sourceLink ? (
                    <a className="ghost-button subtle-action" href={reminder.sourceLink} target="_blank" rel="noreferrer">
                      Open source
                    </a>
                  ) : null}
                  {reminder.status === "active" ? (
                    <button className="ghost-button subtle-action" onClick={() => void props.onDismiss(reminder)}>
                      Dismiss
                    </button>
                  ) : (
                    <button className="ghost-button subtle-action" onClick={() => void props.onReactivate(reminder)}>
                      Reactivate
                    </button>
                  )}
                </div>
              </div>
            </article>
          ))}
          {!props.reminders.length ? <p className="empty-state">No reminders yet. That is a good sign.</p> : null}
        </div>
      </div>
    </section>
  );
}

function RejectedView(props: {
  tasks: RejectedTask[];
  loading: boolean;
  onRestore: (task: RejectedTask) => Promise<void>;
  onAlwaysIgnore: (task: RejectedTask) => Promise<void>;
  onShouldInclude: (task: RejectedTask) => Promise<void>;
}) {
  const [sourceFilter, setSourceFilter] = useState<"All" | "Email" | "Jira">("All");
  const filtered = props.tasks.filter((task) => (sourceFilter === "All" ? true : task.source === sourceFilter));

  function hasIncludeLearning(task: RejectedTask) {
    return task.decisionState === "uncertain" || (task.decisionReason ?? "").includes("should have been included");
  }

  function hasIgnoreLearning(task: RejectedTask) {
    return (task.decisionReason ?? "").includes("ignore similar");
  }

  return (
    <section className="panel-stack">
      <div className="panel">
        <div className="tasks-header">
          <div>
            <h3>Rejected Queue</h3>
            <p className="muted">Items the planner hid from the active plan, with reasons and recovery actions.</p>
          </div>
          <div className="filter-shell">
            <label className="status-select compact">
              <span>Source</span>
              <select value={sourceFilter} onChange={(event) => setSourceFilter(event.target.value as "All" | "Email" | "Jira")}>
                <option value="All">All sources</option>
                <option value="Email">Email</option>
                <option value="Jira">Jira</option>
              </select>
            </label>
          </div>
        </div>
        <div className="reminder-list">
          {props.loading ? <p className="muted">Loading rejected tasks…</p> : null}
          {filtered.map((task) => (
            <article className="reminder-card rejected-card" key={task.id}>
              <div className="task-meta">
                <span className={`pill pill-${task.source.toLowerCase()}`}>{task.source}</span>
                <span className="subtle-pill">
                  Hidden • {Math.round((task.decisionConfidence ?? 0) * 100)}% confidence
                </span>
              </div>
              <strong>{task.title}</strong>
              <p className="muted">{task.decisionReason ?? "This item looked less relevant to your current preferences."}</p>
              <div className="task-link-row">
                {task.decisionReasonTags.map((tag) => (
                  <span className="subtle-pill" key={tag}>
                    {tag.replace(/_/g, " ")}
                  </span>
                ))}
                {hasIncludeLearning(task) ? <span className="subtle-pill learning-pill">Learning applied</span> : null}
                {hasIgnoreLearning(task) ? <span className="subtle-pill learning-pill">Ignore similar saved</span> : null}
              </div>
              <div className="reminder-footer">
                <span className="small-text">Rejected {formatDateTime(task.rejectedAt)}</span>
                <div className="integration-actions">
                  {task.sourceLink ? (
                    <a className="ghost-button subtle-action" href={task.sourceLink} target="_blank" rel="noreferrer">
                      Open source
                    </a>
                  ) : null}
                  <button
                    className={hasIgnoreLearning(task) ? "ghost-button subtle-action selected-action" : "ghost-button subtle-action"}
                    onClick={() => void props.onAlwaysIgnore(task)}
                  >
                    Always ignore similar
                  </button>
                  <button
                    className={hasIncludeLearning(task) ? "ghost-button subtle-action selected-action" : "ghost-button subtle-action"}
                    onClick={() => void props.onShouldInclude(task)}
                  >
                    This should have been included
                  </button>
                  <button className="primary-button" onClick={() => void props.onRestore(task)}>
                    Restore to plan
                  </button>
                </div>
              </div>
            </article>
          ))}
          {!filtered.length ? <p className="empty-state">Nothing is hidden right now.</p> : null}
        </div>
      </div>
    </section>
  );
}

function SettingsView(props: {
  integrations: {
    microsoft: IntegrationStatus;
    jira: IntegrationStatus;
  } | null;
  loading: boolean;
  automation: AutomationSettings | null;
  profile: UserPriorityProfile | null;
  insights: PersonalizationInsight[];
  microsoftAccount: MicrosoftAccount | null;
  microsoftStatusText: string | null;
  jiraStatusText: string | null;
  savingMicrosoft: boolean;
  savingJira: boolean;
  onUpdateSchedule: (input: Partial<Pick<AutomationSettings, "scheduleEnabled" | "scheduleTimeLocal" | "scheduleTimezone">>) => Promise<void>;
  onUpdateReminderSettings: (input: Partial<Pick<AutomationSettings, "remindersEnabled" | "reminderCadenceHours" | "desktopNotificationsEnabled">>) => Promise<void>;
  onUpdateProfile: (input: Partial<UserPriorityProfile>) => Promise<void>;
  onRunCalibration: (input: {
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
  }) => Promise<void>;
  onConnectMicrosoft: () => Promise<void>;
  onDisconnectMicrosoft: () => Promise<void>;
  onSaveJira: (input: { baseUrl: string; email: string; apiToken: string }) => Promise<void>;
  onDisconnectJira: () => Promise<void>;
}) {
  const [form, setForm] = useState({ baseUrl: "", email: "", apiToken: "" });
  const [scheduleForm, setScheduleForm] = useState({
    scheduleEnabled: false,
    scheduleTimeLocal: "08:30",
    scheduleTimezone: getBrowserTimeZone() ?? "UTC",
    remindersEnabled: true,
    reminderCadenceHours: 6,
    desktopNotificationsEnabled: false
  });
  const [calibrationForm, setCalibrationForm] = useState({
    roleFocus: "",
    prioritizationPrompt: "",
    importantWork: "",
    noiseWork: "",
    mustNotMiss: "",
    importantPeople: "",
    importantProjects: "",
    filteringStyle: "conservative" as UserPriorityProfile["filteringStyle"],
    priorityBias: "balanced" as UserPriorityProfile["priorityBias"]
  });
  const [exampleRankings, setExampleRankings] = useState<
    Array<{ title: string; source: "Email" | "Jira" | "Manual"; decision: "show_today" | "keep_low" | "reject_noise" }>
  >([
    { title: "Manager asks for design review by EOD", source: "Email", decision: "show_today" },
    { title: "Jira issue assigned to you and blocked in QA", source: "Jira", decision: "show_today" },
    { title: "Automated comment notification on Jira thread", source: "Email", decision: "reject_noise" },
    { title: "Weekly org newsletter", source: "Email", decision: "reject_noise" },
    { title: "Manual note for future cleanup", source: "Manual", decision: "keep_low" },
    { title: "Jira story updated with review request", source: "Jira", decision: "show_today" }
  ]);

  useEffect(() => {
    const jiraConfig = props.integrations?.jira.config;
    if (!jiraConfig) {
      return;
    }
    setForm({
      baseUrl: jiraConfig.baseUrl || "",
      email: jiraConfig.email || "",
      apiToken: jiraConfig.apiToken || ""
    });
  }, [props.integrations?.jira.config?.baseUrl, props.integrations?.jira.config?.email, props.integrations?.jira.config?.apiToken]);

  useEffect(() => {
    if (props.integrations?.jira.status !== "connected") {
      setForm({ baseUrl: "", email: "", apiToken: "" });
    }
  }, [props.integrations?.jira.status]);

  useEffect(() => {
    if (!props.automation) return;
    setScheduleForm({
      scheduleEnabled: props.automation.scheduleEnabled,
      scheduleTimeLocal: props.automation.scheduleTimeLocal,
      scheduleTimezone: props.automation.scheduleTimezone,
      remindersEnabled: props.automation.remindersEnabled,
      reminderCadenceHours: props.automation.reminderCadenceHours,
      desktopNotificationsEnabled: props.automation.desktopNotificationsEnabled
    });
  }, [props.automation]);

  useEffect(() => {
    if (!props.profile) return;
    setCalibrationForm({
      roleFocus: props.profile.roleFocus ?? "",
      prioritizationPrompt: props.profile.prioritizationPrompt ?? "",
      importantWork: formatPreferenceLines(props.profile.importantWork),
      noiseWork: formatPreferenceLines(props.profile.noiseWork),
      mustNotMiss: formatPreferenceLines(props.profile.mustNotMiss),
      importantPeople: formatPreferenceLines(props.profile.importantPeople),
      importantProjects: formatPreferenceLines(props.profile.importantProjects),
      filteringStyle: props.profile.filteringStyle,
      priorityBias: props.profile.priorityBias
    });
  }, [props.profile]);

  return (
    <section className="panel-stack">
      <div className="hero-card settings-hero">
        <div className="hero-copy">
          <p className="eyebrow">Settings</p>
          <h2>Connections, automation, and prioritization controls.</h2>
          <p className="muted">
            Tune integrations, scheduling, reminders, and the personalization instructions that shape your plan.
          </p>
        </div>
      </div>
      <div className="settings-grid">
        <div className="panel integration-card">
          <div className="panel-header">
            <h3>Microsoft Outlook + Calendar</h3>
            <span className={`status-badge ${props.integrations?.microsoft.status ?? "disconnected"}`}>
              {props.integrations?.microsoft.status ?? "disconnected"}
            </span>
          </div>
          <p className="muted">
            Connect a Microsoft account to read recent email and today&apos;s meetings through Graph API.
          </p>
          <div className="integration-facts">
            <p>
              <span>Account</span>
              <strong>{props.integrations?.microsoft.accountLabel ?? "Not connected"}</strong>
            </p>
            <p>
              <span>Last sync</span>
              <strong>{formatDateTime(props.integrations?.microsoft.lastSyncAt ?? null)}</strong>
            </p>
          </div>
          {props.integrations?.microsoft.errorMessage ? (
            <p className="error-text">{props.integrations.microsoft.errorMessage}</p>
          ) : null}
          {props.microsoftStatusText ? <p className="muted">{props.microsoftStatusText}</p> : null}
          <div className="integration-actions">
            <button
              className="primary-button"
              onClick={() => props.onConnectMicrosoft()}
              disabled={props.savingMicrosoft || props.loading}
            >
              {props.savingMicrosoft ? "Working..." : props.microsoftAccount ? "Reconnect Microsoft" : "Connect Microsoft"}
            </button>
            {props.microsoftAccount ? (
              <button
                className="ghost-button subtle-action"
                onClick={() => props.onDisconnectMicrosoft()}
                disabled={props.savingMicrosoft || props.loading}
              >
                Disconnect
              </button>
            ) : null}
          </div>
        </div>

        <div className="panel integration-card">
          <div className="panel-header">
            <h3>Jira</h3>
            <span className={`status-badge ${props.integrations?.jira.status ?? "disconnected"}`}>
              {props.integrations?.jira.status ?? "disconnected"}
            </span>
          </div>
          <p className="muted">Save your Jira site URL, email, and API token for read-only issue sync.</p>
          <div className="integration-facts">
            <p>
              <span>Account</span>
              <strong>{props.integrations?.jira.accountLabel ?? "Not connected"}</strong>
            </p>
            <p>
              <span>Last sync</span>
              <strong>{formatDateTime(props.integrations?.jira.lastSyncAt ?? null)}</strong>
            </p>
          </div>
          {props.integrations?.jira.errorMessage ? (
            <p className="error-text">{props.integrations.jira.errorMessage}</p>
          ) : null}
          {props.jiraStatusText ? <p className="muted">{props.jiraStatusText}</p> : null}
          <form
            className="settings-form refined"
            onSubmit={async (event) => {
              event.preventDefault();
              await props.onSaveJira(form);
            }}
          >
            <input
              value={form.baseUrl}
              onChange={(event) => setForm((current) => ({ ...current, baseUrl: event.target.value }))}
              placeholder="https://your-company.atlassian.net"
            />
            <input
              value={form.email}
              onChange={(event) => setForm((current) => ({ ...current, email: event.target.value }))}
              placeholder="Email"
            />
            <input
              value={form.apiToken}
              onChange={(event) => {
                setForm((current) => ({ ...current, apiToken: event.target.value }));
              }}
              placeholder="API token"
              type="password"
              autoComplete="current-password"
              onCopy={(event) => event.preventDefault()}
              onCut={(event) => event.preventDefault()}
              onPaste={(event) => event.preventDefault()}
            />
            <div className="integration-actions">
              <button className="primary-button" type="submit" disabled={props.savingJira || props.loading}>
                {props.savingJira ? "Saving..." : "Save Jira Connection"}
              </button>
              {props.integrations?.jira.status === "connected" ? (
                <button
                  className="ghost-button subtle-action"
                  type="button"
                  onClick={() => props.onDisconnectJira()}
                  disabled={props.savingJira || props.loading}
                >
                  Revoke Jira
                </button>
              ) : null}
            </div>
          </form>
        </div>
      </div>
      <div className="panel automation-panel">
        <div className="panel-header">
          <h3>Automation & Trust</h3>
          <span className={`status-badge ${props.automation?.schedulerLastStatus ?? "idle"}`}>
            {props.automation?.schedulerLastStatus ?? "idle"}
          </span>
        </div>
        <p className="muted">Automation stays opt-in, transparent, and easy to override.</p>
        <div className="automation-grid">
          <label className="toggle-card">
            <span>Auto-generate daily plan</span>
            <input
              type="checkbox"
              checked={scheduleForm.scheduleEnabled}
              onChange={async (event) => {
                const next = event.target.checked;
                setScheduleForm((current) => ({ ...current, scheduleEnabled: next }));
                await props.onUpdateSchedule({ scheduleEnabled: next });
              }}
            />
          </label>
          <label className="field-card">
            <span>Run time</span>
            <input
              type="time"
              value={scheduleForm.scheduleTimeLocal}
              onChange={(event) => setScheduleForm((current) => ({ ...current, scheduleTimeLocal: event.target.value }))}
              onBlur={async () => {
                await props.onUpdateSchedule({ scheduleTimeLocal: scheduleForm.scheduleTimeLocal });
              }}
            />
          </label>
          <label className="field-card">
            <span>Timezone</span>
            <input
              value={scheduleForm.scheduleTimezone}
              onChange={(event) => setScheduleForm((current) => ({ ...current, scheduleTimezone: event.target.value }))}
              onBlur={async () => {
                await props.onUpdateSchedule({ scheduleTimezone: scheduleForm.scheduleTimezone });
              }}
            />
          </label>
          <label className="toggle-card">
            <span>Reminder center enabled</span>
            <input
              type="checkbox"
              checked={scheduleForm.remindersEnabled}
              onChange={async (event) => {
                const next = event.target.checked;
                setScheduleForm((current) => ({ ...current, remindersEnabled: next }));
                await props.onUpdateReminderSettings({ remindersEnabled: next });
              }}
            />
          </label>
          <label className="field-card">
            <span>Reminder cadence</span>
            <select
              value={scheduleForm.reminderCadenceHours}
              onChange={async (event) => {
                const next = Number(event.target.value);
                setScheduleForm((current) => ({ ...current, reminderCadenceHours: next }));
                await props.onUpdateReminderSettings({ reminderCadenceHours: next });
              }}
            >
              {[2, 4, 6, 12, 24].map((hours) => (
                <option key={hours} value={hours}>
                  Every {hours} hours
                </option>
              ))}
            </select>
          </label>
          <label className="toggle-card">
            <span>Desktop notifications</span>
            <input
              type="checkbox"
              checked={scheduleForm.desktopNotificationsEnabled}
              onChange={async (event) => {
                const next = event.target.checked;
                setScheduleForm((current) => ({ ...current, desktopNotificationsEnabled: next }));
                await props.onUpdateReminderSettings({ desktopNotificationsEnabled: next });
              }}
            />
          </label>
        </div>
        <div className="integration-facts">
          <p>
            <span>Last scheduled run</span>
            <strong>{formatDateTime(props.automation?.schedulerLastRunAt ?? null)}</strong>
          </p>
          <p>
            <span>Last auto-generated plan</span>
            <strong>{formatDateTime(props.automation?.lastAutoGeneratedAt ?? null)}</strong>
          </p>
        </div>
        {props.automation?.schedulerLastError ? <p className="error-text">{props.automation.schedulerLastError}</p> : null}
      </div>
      <div className="panel automation-panel">
        <div className="panel-header">
          <h3>Planner Preferences</h3>
          <span className={`status-badge ${props.profile?.personalizationEnabled ? "connected" : "disconnected"}`}>
            {props.profile?.personalizationEnabled ? "enabled" : "disabled"}
          </span>
        </div>
        <p className="muted">
          Keep this simple: tell the planner what to surface early, what is safe to mute, and what should never disappear.
          The planner still analyzes the full mail or Jira content before deciding.
        </p>
        <div className="automation-grid">
          <label className="toggle-card">
            <span>Personalized filtering</span>
            <input
              type="checkbox"
              checked={props.profile?.personalizationEnabled ?? true}
              onChange={async (event) => {
                await props.onUpdateProfile({ personalizationEnabled: event.target.checked });
              }}
            />
          </label>
          <label className="field-card">
            <span>How selective should the planner be?</span>
            <select
              value={calibrationForm.filteringStyle}
              onChange={(event) =>
                setCalibrationForm((current) => ({
                  ...current,
                  filteringStyle: event.target.value as UserPriorityProfile["filteringStyle"]
                }))
              }
            >
              <option value="conservative">Tight focus</option>
              <option value="balanced">Balanced</option>
              <option value="aggressive">Show me more</option>
            </select>
          </label>
          <label className="field-card">
            <span>When in doubt, optimize for</span>
            <select
              value={calibrationForm.priorityBias}
              onChange={(event) =>
                setCalibrationForm((current) => ({
                  ...current,
                  priorityBias: event.target.value as UserPriorityProfile["priorityBias"]
                }))
              }
            >
              <option value="focus">Finishing what matters</option>
              <option value="balanced">Balanced</option>
              <option value="coverage">Keeping options visible</option>
            </select>
          </label>
        </div>
        <form
          className="settings-form refined"
          onSubmit={async (event) => {
            event.preventDefault();
            await props.onRunCalibration({
              roleFocus: calibrationForm.roleFocus,
              prioritizationPrompt: calibrationForm.prioritizationPrompt,
              importantWork: parsePreferenceLines(calibrationForm.importantWork),
              noiseWork: parsePreferenceLines(calibrationForm.noiseWork),
              mustNotMiss: parsePreferenceLines(calibrationForm.mustNotMiss),
              importantPeople: parsePreferenceLines(calibrationForm.importantPeople),
              importantProjects: parsePreferenceLines(calibrationForm.importantProjects),
              filteringStyle: calibrationForm.filteringStyle,
              priorityBias: calibrationForm.priorityBias,
              exampleRankings
            });
          }}
        >
          <div className="preference-grid">
            <label className="field-card preference-card">
              <span>Your role or current charter</span>
              <input
                value={calibrationForm.roleFocus}
                onChange={(event) => setCalibrationForm((current) => ({ ...current, roleFocus: event.target.value }))}
                placeholder="For example: Staff engineer focused on platform reliability"
              />
            </label>
            <label className="field-card preference-card">
              <span>Always prioritize</span>
              <textarea
                className="preference-editor"
                value={calibrationForm.mustNotMiss}
                onChange={(event) => setCalibrationForm((current) => ({ ...current, mustNotMiss: event.target.value }))}
                placeholder={"One item per line\nProduction issues\nManager requests\nCustomer escalations"}
              />
            </label>
            <label className="field-card preference-card">
              <span>Focus areas</span>
              <textarea
                className="preference-editor"
                value={calibrationForm.importantWork}
                onChange={(event) => setCalibrationForm((current) => ({ ...current, importantWork: event.target.value }))}
                placeholder={"One item per line\nCode reviews\nRelease work\nPlatform migrations"}
              />
            </label>
            <label className="field-card preference-card">
              <span>Priority people or senders</span>
              <textarea
                className="preference-editor"
                value={calibrationForm.importantPeople}
                onChange={(event) => setCalibrationForm((current) => ({ ...current, importantPeople: event.target.value }))}
                placeholder={"One item per line\nmanager@company.com\nDirector of Engineering\nProduct lead"}
              />
            </label>
            <label className="field-card preference-card">
              <span>Priority projects</span>
              <textarea
                className="preference-editor"
                value={calibrationForm.importantProjects}
                onChange={(event) => setCalibrationForm((current) => ({ ...current, importantProjects: event.target.value }))}
                placeholder={"One item per line\nSECCCAT\nPLATFORM\nAUTH"}
              />
            </label>
            <label className="field-card preference-card">
              <span>Usually safe to ignore</span>
              <textarea
                className="preference-editor"
                value={calibrationForm.noiseWork}
                onChange={(event) => setCalibrationForm((current) => ({ ...current, noiseWork: event.target.value }))}
                placeholder={"One item per line\nGeneric digests\nFYI only notices\nLow-signal comment notifications"}
              />
            </label>
          </div>
          <label className="field-card">
            <span>Notes to the planner</span>
            <textarea
              className="prompt-editor"
              value={calibrationForm.prioritizationPrompt}
              onChange={(event) =>
                setCalibrationForm((current) => ({ ...current, prioritizationPrompt: event.target.value }))
              }
              placeholder="Anything subtle the planner should know, like how you balance coding, reviews, meetings, or follow-ups."
            />
          </label>
          <div className="field-card">
            <span>Teach with examples</span>
            <p className="muted small-text">
              These examples help the planner tune filtering and ranking without hiding work permanently.
            </p>
          </div>
          <div className="calibration-list">
            {exampleRankings.map((item, index) => (
              <div className="field-card" key={`${item.title}-${index}`}>
                <span>
                  {item.source}: {item.title}
                </span>
                <select
                  value={item.decision}
                  onChange={(event) =>
                    setExampleRankings((current) =>
                      current.map((entry, currentIndex) =>
                        currentIndex === index
                          ? {
                              ...entry,
                              decision: event.target.value as "show_today" | "keep_low" | "reject_noise"
                            }
                          : entry
                      )
                    )
                  }
                >
                  <option value="show_today">Bring forward</option>
                  <option value="keep_low">Keep visible, lower</option>
                  <option value="reject_noise">Hide as low value</option>
                </select>
              </div>
            ))}
          </div>
          <div className="integration-actions">
            <button className="primary-button" type="submit">
              Refresh preferences
            </button>
            <button
              className="ghost-button subtle-action"
              type="button"
              onClick={() =>
                void props.onUpdateProfile({
                  filteringStyle: calibrationForm.filteringStyle,
                  priorityBias: calibrationForm.priorityBias,
                  roleFocus: calibrationForm.roleFocus,
                  prioritizationPrompt: calibrationForm.prioritizationPrompt,
                  importantWork: parsePreferenceLines(calibrationForm.importantWork),
                  noiseWork: parsePreferenceLines(calibrationForm.noiseWork),
                  mustNotMiss: parsePreferenceLines(calibrationForm.mustNotMiss),
                  importantPeople: parsePreferenceLines(calibrationForm.importantPeople),
                  importantProjects: parsePreferenceLines(calibrationForm.importantProjects)
                })
              }
            >
              Save preferences
            </button>
          </div>
        </form>
        <div className="field-card">
          <span>What the planner is learning</span>
          <p className="muted small-text">
            Rejections, restores, reprioritization, and completions continue to refine future filtering and ranking.
          </p>
        </div>
        <div className="detail-list">
          {props.insights.length ? (
            props.insights.map((insight, index) => (
              <article className="detail-row" key={`${insight.statement}-${index}`}>
                <strong>{insight.statement}</strong>
                <span className="subtle-pill">{Math.round(insight.confidence * 100)}%</span>
              </article>
            ))
          ) : (
            <p className="empty-state">Insights will appear after a few task decisions and calibrations.</p>
          )}
        </div>
      </div>
    </section>
  );
}

function replaceTaskInGroups(data: TodayResponse | null, taskId: number, updater: (task: Task) => Task) {
  if (!data) return data;
  const nextTasks = Object.fromEntries(
    priorityOrder.map((priority) => [
      priority,
      data.tasks[priority].map((task) => (task.id === taskId ? updater(task) : task))
    ])
  ) as Record<TaskPriority, Task[]>;
  return { ...data, tasks: nextTasks };
}

function replaceTaskInList(tasks: Task[], taskId: number, updater: (task: Task) => Task) {
  return tasks.map((task) => (task.id === taskId ? updater(task) : task));
}

function mergeTaskLists(existing: Task[], incoming: Task[]) {
  const byId = new Map<number, Task>();
  for (const task of existing) {
    byId.set(task.id, task);
  }
  for (const task of incoming) {
    byId.set(task.id, task);
  }
  return [...byId.values()].sort(compareTasks);
}

function applyTodayResponseState(
  payload: TodayResponse,
  setters: {
    setToday: (value: TodayResponse) => void;
    setAllTasks: React.Dispatch<React.SetStateAction<Task[]>>;
    setReminders: (value: Reminder[]) => void;
    setAutomation: (value: AutomationSettings) => void;
  }
) {
  setters.setToday(payload);
  setters.setAllTasks((current) => mergeTaskLists(current, flattenTaskGroups(payload.tasks)));
  setters.setReminders(payload.reminders);
  setters.setAutomation(payload.automation);
}

export function App() {
  const [view, setView] = useState<View>("today");
  const [today, setToday] = useState<TodayResponse | null>(null);
  const [allTasks, setAllTasks] = useState<Task[]>([]);
  const [deferredTasks, setDeferredTasks] = useState<Task[]>([]);
  const [rejectedTasks, setRejectedTasks] = useState<RejectedTask[]>([]);
  const [reminders, setReminders] = useState<Reminder[]>([]);
  const [automation, setAutomation] = useState<AutomationSettings | null>(null);
  const [profile, setProfile] = useState<UserPriorityProfile | null>(null);
  const [insights, setInsights] = useState<PersonalizationInsight[]>([]);
  const [taskFilter, setTaskFilter] = useState<TaskFilter>("All");
  const [integrations, setIntegrations] = useState<{ microsoft: IntegrationStatus; jira: IntegrationStatus } | null>(
    null
  );
  const [pageLoading, setPageLoading] = useState<Record<View, boolean>>({
    today: true,
    tasks: false,
    deferred: false,
    rejected: false,
    reminders: false,
    settings: false
  });
  const [loadedViews, setLoadedViews] = useState<Record<View, boolean>>({
    today: false,
    tasks: false,
    deferred: false,
    rejected: false,
    reminders: false,
    settings: false
  });
  const [loading, setLoading] = useState(false);
  const [syncMeetingsLoading, setSyncMeetingsLoading] = useState(false);
  const [syncTasksLoading, setSyncTasksLoading] = useState(false);
  const [savingMicrosoft, setSavingMicrosoft] = useState(false);
  const [savingJira, setSavingJira] = useState(false);
  const [microsoftStatusText, setMicrosoftStatusText] = useState<string | null>(null);
  const [jiraStatusText, setJiraStatusText] = useState<string | null>(null);
  const [microsoftAccount, setMicrosoftAccount] = useState<MicrosoftAccount | null>(() => getMicrosoftAccount());
  const [detailTask, setDetailTask] = useState<Task | null>(null);
  const [detailData, setDetailData] = useState<TaskDetail | null>(null);
  const [detailLoading, setDetailLoading] = useState(false);
  const [detailError, setDetailError] = useState<string | null>(null);

  async function getMicrosoftSessionToken() {
    const account = getMicrosoftAccount();
    setMicrosoftAccount(account);
    if (!account) {
      return null;
    }
    try {
      return await acquireMicrosoftApiToken();
    } catch (error) {
      console.error(error);
      return null;
    }
  }

  async function loadTodayPage() {
    setPageLoading((current) => ({ ...current, today: true }));
    try {
      const microsoftAccessToken = await getMicrosoftSessionToken();
      const [todayData, integrationsData, automationData, profileData, insightsData, rejectedData] = await Promise.all([
        api.getToday(),
        microsoftAccessToken
          ? api.getIntegrationsWithMicrosoftSession(microsoftAccessToken)
          : api.getIntegrations(),
        api.getAutomationSettings(),
        api.getPersonalizationProfile(),
        api.getPersonalizationInsights(),
        api.getRejectedTasks()
      ]);
      setToday(todayData);
      setIntegrations(integrationsData.integrations);
      setAutomation(automationData.automation);
      setReminders(automationData.reminders);
      setProfile(profileData.profile);
      setInsights(insightsData.insights);
      setRejectedTasks(rejectedData.tasks);
      setLoadedViews((current) => ({ ...current, today: true }));
    } finally {
      setPageLoading((current) => ({ ...current, today: false }));
    }
  }

  async function loadTasksPage() {
    setPageLoading((current) => ({ ...current, tasks: true }));
    try {
      const taskData = await api.getTasks();
      setAllTasks(taskData.tasks);
      setLoadedViews((current) => ({ ...current, tasks: true }));
    } finally {
      setPageLoading((current) => ({ ...current, tasks: false }));
    }
  }

  async function loadDeferredPage() {
    setPageLoading((current) => ({ ...current, deferred: true }));
    try {
      const response = await api.getDeferredTasks();
      setDeferredTasks(response.tasks);
      setLoadedViews((current) => ({ ...current, deferred: true }));
    } finally {
      setPageLoading((current) => ({ ...current, deferred: false }));
    }
  }

  async function loadRejectedPage() {
    setPageLoading((current) => ({ ...current, rejected: true }));
    try {
      const response = await api.getRejectedTasks();
      setRejectedTasks(response.tasks);
      setLoadedViews((current) => ({ ...current, rejected: true }));
    } finally {
      setPageLoading((current) => ({ ...current, rejected: false }));
    }
  }

  async function loadRemindersPage() {
    setPageLoading((current) => ({ ...current, reminders: true }));
    try {
      const response = await api.getReminders();
      setReminders(response.reminders);
      setLoadedViews((current) => ({ ...current, reminders: true }));
    } finally {
      setPageLoading((current) => ({ ...current, reminders: false }));
    }
  }

  async function loadSettingsPage() {
    setPageLoading((current) => ({ ...current, settings: true }));
    try {
      const microsoftAccessToken = await getMicrosoftSessionToken();
      const [integrationsData, automationData, profileData, insightsData] = await Promise.all([
        microsoftAccessToken
          ? api.getIntegrationsWithMicrosoftSession(microsoftAccessToken)
          : api.getIntegrations(),
        api.getAutomationSettings(),
        api.getPersonalizationProfile(),
        api.getPersonalizationInsights()
      ]);
      setIntegrations(integrationsData.integrations);
      setAutomation(automationData.automation);
      setReminders(automationData.reminders);
      setProfile(profileData.profile);
      setInsights(insightsData.insights);
      setLoadedViews((current) => ({ ...current, settings: true }));
    } finally {
      setPageLoading((current) => ({ ...current, settings: false }));
    }
  }

  useEffect(() => {
    void loadTodayPage();
  }, []);

  useEffect(() => {
    if (view === "today" && !loadedViews.today && !pageLoading.today) {
      void loadTodayPage();
    }
    if (view === "tasks" && !loadedViews.tasks && !pageLoading.tasks) {
      void loadTasksPage();
    }
    if (view === "settings" && !loadedViews.settings && !pageLoading.settings) {
      void loadSettingsPage();
    }
    if (view === "deferred" && !loadedViews.deferred && !pageLoading.deferred) {
      void loadDeferredPage();
    }
    if (view === "reminders" && !loadedViews.reminders && !pageLoading.reminders) {
      void loadRemindersPage();
    }
    if (view === "rejected" && !loadedViews.rejected && !pageLoading.rejected) {
      void loadRejectedPage();
    }
  }, [view, loadedViews, pageLoading]);

  const visibleTasks = useMemo(
    () => (taskFilter === "All" ? allTasks : allTasks.filter((task) => task.status === taskFilter)),
    [allTasks, taskFilter]
  );

  async function refreshTodayAndIntegrations() {
    const microsoftAccessToken = await getMicrosoftSessionToken();
    const [todayData, integrationsData, deferredData, reminderData, automationData, profileData, insightsData, rejectedData] = await Promise.all([
      api.getToday(),
      microsoftAccessToken
        ? api.getIntegrationsWithMicrosoftSession(microsoftAccessToken)
        : api.getIntegrations(),
      api.getDeferredTasks(),
      api.getReminders(),
      api.getAutomationSettings(),
      api.getPersonalizationProfile(),
      api.getPersonalizationInsights(),
      api.getRejectedTasks()
    ]);
    setToday(todayData);
    setIntegrations(integrationsData.integrations);
    setDeferredTasks(deferredData.tasks);
    setReminders(reminderData.reminders);
    setAutomation(automationData.automation);
    setProfile(profileData.profile);
    setInsights(insightsData.insights);
    setRejectedTasks(rejectedData.tasks);
  }

  async function refreshTasksPage() {
    const taskData = await api.getTasks();
    setAllTasks(taskData.tasks);
  }

  async function handleTaskStatusChange(task: Task, status: TaskStatus) {
    const originalStatus = task.status;
    setAllTasks((current) => current.map((item) => (item.id === task.id ? { ...item, status } : item)));
    setToday((current) => replaceTaskInGroups(current, task.id, (item) => ({ ...item, status })));

    try {
      const response = await api.updateTask(task.id, { status });
      setAllTasks((current) => current.map((item) => (item.id === task.id ? response.task : item)));
      setToday((current) => replaceTaskInGroups(current, task.id, () => response.task));
    } catch (error) {
      console.error(error);
      setAllTasks((current) =>
        current.map((item) => (item.id === task.id ? { ...item, status: originalStatus } : item))
      );
      setToday((current) => replaceTaskInGroups(current, task.id, (item) => ({ ...item, status: originalStatus })));
    }
  }

  async function handleTaskPriorityChange(task: Task, priority: TaskPriority) {
    if (task.priority === priority) return;

    const originalPriority = task.priority;
    const optimisticTask = { ...task, priority, updatedAt: new Date().toISOString() };

    setAllTasks((current) => replaceTaskInList(current, task.id, () => optimisticTask));
    setToday((current) =>
      current
        ? {
            ...current,
            tasks: groupTasksByPriority(
              replaceTaskInList(
                [...current.tasks.High, ...current.tasks.Medium, ...current.tasks.Low],
                task.id,
                () => optimisticTask
              )
            )
          }
        : current
    );

    try {
      const response = await api.updateTask(task.id, { priority });
      setAllTasks((current) => replaceTaskInList(current, task.id, () => response.task));
      setToday((current) =>
        current
          ? {
              ...current,
              tasks: groupTasksByPriority(
                replaceTaskInList(
                  [...current.tasks.High, ...current.tasks.Medium, ...current.tasks.Low],
                  task.id,
                  () => response.task
                )
              )
            }
          : current
      );
    } catch (error) {
      console.error(error);
      setAllTasks((current) =>
        replaceTaskInList(current, task.id, (item) => ({ ...item, priority: originalPriority }))
      );
      setToday((current) =>
        current
          ? {
              ...current,
              tasks: groupTasksByPriority(
                replaceTaskInList(
                  [...current.tasks.High, ...current.tasks.Medium, ...current.tasks.Low],
                  task.id,
                  (item) => ({ ...item, priority: originalPriority })
                )
              )
            }
          : current
      );
    }
  }

  async function openTaskDetails(task: Task) {
    setDetailTask(task);
    setDetailLoading(true);
    setDetailError(null);
    setDetailData(null);

    try {
      const microsoftAccessToken = task.source === "Email" ? await getMicrosoftSessionToken() : null;
      const response =
        task.source === "Email" && microsoftAccessToken
          ? await api.getTaskDetailsWithMicrosoftSession(task.id, microsoftAccessToken)
          : await api.getTaskDetails(task.id);
      setDetailData(response.detail);
    } catch (error) {
      setDetailError(error instanceof Error ? error.message : "Failed to load task details");
    } finally {
      setDetailLoading(false);
    }
  }

  async function handleGeneratePlan() {
    setLoading(true);
    try {
      const microsoftAccessToken = await getMicrosoftSessionToken();
      const browserTimeZone = getBrowserTimeZone();
      const nextToday = microsoftAccessToken
        ? await api.generatePlanWithMicrosoftSession(microsoftAccessToken, browserTimeZone)
        : await api.generatePlan(browserTimeZone);
      applyTodayResponseState(nextToday, { setToday, setAllTasks, setReminders, setAutomation });
      setLoadedViews((current) => ({ ...current, today: true, tasks: true }));
      await loadDeferredPage();
      await loadRemindersPage();
      await loadSettingsPage();
    } finally {
      setLoading(false);
    }
  }

  async function handleSyncMeetings() {
    setSyncMeetingsLoading(true);
    try {
      const microsoftAccessToken = await getMicrosoftSessionToken();
      const browserTimeZone = getBrowserTimeZone();
      const nextToday = microsoftAccessToken
        ? await api.syncMeetingsWithMicrosoftSession(microsoftAccessToken, browserTimeZone)
        : await api.syncMeetings(browserTimeZone);
      applyTodayResponseState(nextToday, { setToday, setAllTasks, setReminders, setAutomation });
    } finally {
      setSyncMeetingsLoading(false);
    }
  }

  async function handleSyncTasks() {
    setSyncTasksLoading(true);
    try {
      const microsoftAccessToken = await getMicrosoftSessionToken();
      const browserTimeZone = getBrowserTimeZone();
      const nextToday = microsoftAccessToken
        ? await api.syncTasksWithMicrosoftSession(microsoftAccessToken, browserTimeZone)
        : await api.syncTasks(browserTimeZone);
      applyTodayResponseState(nextToday, { setToday, setAllTasks, setReminders, setAutomation });
      setLoadedViews((current) => ({ ...current, tasks: true }));
      await loadDeferredPage();
      await loadRemindersPage();
      await loadSettingsPage();
    } finally {
      setSyncTasksLoading(false);
    }
  }

  return (
    <>
      <main className="app-shell">
        <AppHeader active={view} onChange={setView} />
        <div className="content-shell">
          {view === "today" ? (
            !loadedViews.today ? (
              <TodaySkeleton />
            ) : (
              <TodayView
                data={today}
                loading={loading}
                onGenerate={handleGeneratePlan}
                onSyncMeetings={handleSyncMeetings}
                onSyncTasks={handleSyncTasks}
                syncMeetingsLoading={syncMeetingsLoading}
                syncTasksLoading={syncTasksLoading}
                onTaskStatusChange={handleTaskStatusChange}
                onTaskPriorityChange={handleTaskPriorityChange}
                onOpenDetails={openTaskDetails}
              />
            )
          ) : null}

          {view === "tasks" ? (
            !loadedViews.tasks ? (
              <TasksSkeleton />
            ) : (
              <TasksView
                tasks={visibleTasks}
                loading={pageLoading.tasks}
                filter={taskFilter}
                onFilterChange={setTaskFilter}
                onCreate={async (title) => {
                  const response = await api.createTask({ title });
                  setAllTasks((current) => [response.task, ...current]);
                }}
                onUpdateStatus={handleTaskStatusChange}
                onUpdatePriority={handleTaskPriorityChange}
                onDelete={async (task) => {
                  const previousTasks = allTasks;
                  setAllTasks((current) => current.filter((item) => item.id !== task.id));
                  setToday((current) => {
                    if (!current) return current;
                    return {
                      ...current,
                      tasks: {
                        High: current.tasks.High.filter((item) => item.id !== task.id),
                        Medium: current.tasks.Medium.filter((item) => item.id !== task.id),
                        Low: current.tasks.Low.filter((item) => item.id !== task.id)
                      }
                    };
                  });
                  try {
                    await api.deleteTask(task.id);
                    await refreshTodayAndIntegrations();
                  } catch (error) {
                    console.error(error);
                    setAllTasks(previousTasks);
                    await refreshTodayAndIntegrations();
                  }
                }}
                onOpenDetails={openTaskDetails}
                onDeferUntilTomorrow={async (task) => {
                  const tomorrow = new Date();
                  tomorrow.setDate(tomorrow.getDate() + 1);
                  tomorrow.setHours(9, 0, 0, 0);
                  const response = await api.deferTask(task.id, tomorrow.toISOString());
                  setAllTasks((current) => current.filter((item) => item.id !== task.id));
                  setDeferredTasks((current) => [response.task, ...current.filter((item) => item.id !== task.id)]);
                  await refreshTodayAndIntegrations();
                }}
              />
            )
          ) : null}

          {view === "deferred" ? (
            !loadedViews.deferred ? (
              <TasksSkeleton />
            ) : (
              <DeferredView
                tasks={deferredTasks}
                loading={pageLoading.deferred}
                onBringBackNow={async (task) => {
                  const response = await api.deferTask(task.id, null);
                  setDeferredTasks((current) => current.filter((item) => item.id !== task.id));
                  setAllTasks((current) => [response.task, ...current.filter((item) => item.id !== task.id)]);
                  await refreshTodayAndIntegrations();
                }}
                onOpenDetails={openTaskDetails}
              />
            )
          ) : null}

          {view === "rejected" ? (
            !loadedViews.rejected ? (
              <TasksSkeleton />
            ) : (
              <RejectedView
                tasks={rejectedTasks}
                loading={pageLoading.rejected}
                onRestore={async (task) => {
                  const response = await api.restoreRejectedTask(task.id);
                  if (response.task) {
                    setAllTasks((current) => [response.task as Task, ...current.filter((item) => item.id !== response.task?.id)]);
                  }
                  setRejectedTasks((current) => current.filter((item) => item.id !== task.id));
                  await refreshTodayAndIntegrations();
                }}
                onAlwaysIgnore={async (task) => {
                  const response = await api.updateRejectedTask(task.id, { action: "always_ignore_similar" });
                  if (response.task) {
                    setRejectedTasks((current) => current.map((item) => (item.id === task.id ? response.task! : item)));
                  }
                  await refreshTodayAndIntegrations();
                }}
                onShouldInclude={async (task) => {
                  const response = await api.updateRejectedTask(task.id, { action: "should_have_been_included" });
                  if (response.task) {
                    setRejectedTasks((current) => current.map((item) => (item.id === task.id ? response.task! : item)));
                  }
                  await refreshTodayAndIntegrations();
                }}
              />
            )
          ) : null}

          {view === "reminders" ? (
            !loadedViews.reminders ? (
              <TasksSkeleton />
            ) : (
              <ReminderCenterView
                reminders={reminders}
                loading={pageLoading.reminders}
                onDismiss={async (reminder) => {
                  const response = await api.updateReminder(reminder.id, { status: "dismissed" });
                  setReminders((current) => current.map((item) => (item.id === reminder.id ? response.reminder : item)));
                }}
                onReactivate={async (reminder) => {
                  const response = await api.updateReminder(reminder.id, { status: "active" });
                  setReminders((current) => current.map((item) => (item.id === reminder.id ? response.reminder : item)));
                }}
              />
            )
          ) : null}

          {view === "settings" ? (
            !loadedViews.settings ? (
              <SettingsSkeleton />
            ) : (
              <SettingsView
                integrations={integrations}
                loading={pageLoading.settings}
                automation={automation}
                profile={profile}
                insights={insights}
                microsoftAccount={microsoftAccount}
                microsoftStatusText={microsoftStatusText}
                jiraStatusText={jiraStatusText}
                savingMicrosoft={savingMicrosoft}
                savingJira={savingJira}
                onUpdateSchedule={async (input) => {
                  const response = await api.updateScheduleSettings(input);
                  setAutomation(response.automation);
                }}
                onUpdateReminderSettings={async (input) => {
                  const response = await api.updateReminderSettings(input);
                  setAutomation(response.automation);
                }}
                onUpdateProfile={async (input) => {
                  const response = await api.updatePersonalizationProfile(input);
                  setProfile(response.profile);
                  await refreshTodayAndIntegrations();
                }}
                onRunCalibration={async (input) => {
                  const response = await api.calibratePersonalization(input);
                  setProfile(response.profile);
                  await refreshTodayAndIntegrations();
                }}
                onConnectMicrosoft={async () => {
                  setSavingMicrosoft(true);
                  setMicrosoftStatusText("Connecting Microsoft…");
                  try {
                    await loginWithMicrosoft();
                    setMicrosoftStatusText("Microsoft connection flow started.");
                  } catch (error) {
                    setMicrosoftStatusText(error instanceof Error ? error.message : "Failed to connect Microsoft");
                    throw error;
                  } finally {
                    setSavingMicrosoft(false);
                  }
                }}
                onDisconnectMicrosoft={async () => {
                  setSavingMicrosoft(true);
                  setMicrosoftStatusText("Disconnecting Microsoft…");
                  try {
                    await api.revokeIntegration("microsoft");
                    await logoutFromMicrosoft();
                    setMicrosoftStatusText("Microsoft disconnected.");
                    setIntegrations((current) =>
                      current
                        ? {
                            ...current,
                            microsoft: {
                              ...current.microsoft,
                              status: "disconnected",
                              accountLabel: null,
                              errorMessage: null,
                              updatedAt: null,
                              config: null
                            }
                          }
                        : current
                    );
                  } finally {
                    setSavingMicrosoft(false);
                  }
                }}
                onSaveJira={async (input) => {
                  setSavingJira(true);
                  setJiraStatusText("Saving Jira connection…");
                  try {
                    await api.saveJira(input);
                    await loadSettingsPage();
                    setJiraStatusText("Jira connection saved.");
                  } catch (error) {
                    setJiraStatusText(error instanceof Error ? error.message : "Failed to save Jira connection");
                    throw error;
                  } finally {
                    setSavingJira(false);
                  }
                }}
                onDisconnectJira={async () => {
                  setSavingJira(true);
                  setJiraStatusText("Disconnecting Jira…");
                  try {
                    await api.revokeIntegration("jira");
                    setIntegrations((current) =>
                      current
                        ? {
                            ...current,
                            jira: {
                              ...current.jira,
                              status: "disconnected",
                              accountLabel: null,
                              errorMessage: null,
                              updatedAt: null,
                              config: null
                            }
                          }
                        : current
                    );
                    setJiraStatusText("Jira disconnected.");
                  } finally {
                    setSavingJira(false);
                  }
                }}
              />
            )
          ) : null}
        </div>
      </main>

      <TaskDetailsDialog
        task={detailTask}
        detail={detailData}
        loading={detailLoading}
        error={detailError}
        onClose={() => {
          setDetailTask(null);
          setDetailData(null);
          setDetailError(null);
        }}
      />
    </>
  );
}
