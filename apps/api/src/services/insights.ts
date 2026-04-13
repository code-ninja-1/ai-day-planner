import {
  getDailyPlanSnapshot,
  getIgnoredRejectedTaskCount,
  getLatestPreferenceMemorySnapshot,
  getRejectedTaskCount,
  getTaskById,
  getUserPriorityProfile,
  listDeferredTasks,
  listIgnoredRejectedTasks,
  listPlannerRunDetails,
  listRecentDailyPlanSnapshots,
  listRecentDecisionLogs,
  listRejectedTasks,
  listReminderItems,
  listTaskStateEvents,
  listTasks
} from "../db.js";
import { getTodaySnapshot } from "./planService.js";
import type {
  BehaviorFeedbackEvent,
  DayHistoryDetail,
  DayHistorySummary,
  DayPlanBlock,
  InsightsOverview,
  InsightsTodayPayload,
  InsightsTodayTask,
  PlannerRunDetail,
  Task,
  TaskInsightsPayload
} from "../types.js";

function parseJsonValue<T>(value: unknown, fallback: T): T {
  if (typeof value !== "string" || !value.trim()) return fallback;
  try {
    return JSON.parse(value) as T;
  } catch {
    return fallback;
  }
}

function average(values: number[]) {
  if (!values.length) return null;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function toDaySummary(row: Record<string, unknown>): DayHistorySummary {
  const summary = parseJsonValue<Record<string, unknown>>(row.summary_json, {});
  const dayKey = String(row.day_key);
  const plannedTaskIds = parseJsonValue<number[]>(row.planned_task_ids_json, []);
  const plannedTasks = parseJsonValue<DayPlanBlock[]>(row.blocks_json, []).filter((block) => block.kind === "task");
  const changeEvents = listTaskStateEvents({ dayKey });
  const rejectedCount = changeEvents.filter((event) => event.eventType === "reject").length;
  const restoredCount = changeEvents.filter((event) => event.eventType === "restore").length;
  const deferredCount = changeEvents.filter((event) => event.eventType === "deferred").length;
  const demotedCount = changeEvents.filter(
    (event) => event.eventType === "priority_changed" && (event.afterJson ?? "").includes('"Low"')
  ).length;
  const disagreements = new Set(
    changeEvents
      .filter((event) => ["reject", "deferred"].includes(event.eventType) || (event.eventType === "priority_changed" && (event.afterJson ?? "").includes('"Low"')))
      .map((event) => event.taskId)
      .filter((taskId): taskId is number => taskId !== null)
  ).size + demotedCount;

  const plannedCount = plannedTaskIds.length || plannedTasks.length;
  const completedCount = listTasks(undefined, { includeDeferred: true }).filter((task) => task.completedAt?.slice(0, 10) === dayKey)
    .length;
  const scheduledMeetingCount = parseJsonValue<DayPlanBlock[]>(row.blocks_json, []).filter((block) => block.kind === "meeting").length;
  const agreementPercent =
    plannedCount > 0 ? Math.max(0, Math.min(100, Math.round(((plannedCount - disagreements) / plannedCount) * 100))) : null;
  const completionPercent =
    plannedCount > 0 ? Math.max(0, Math.min(100, Math.round((completedCount / plannedCount) * 100))) : null;

  return {
    dayKey,
    guidance: String(summary.guidance ?? ""),
    plannedTaskCount: plannedCount,
    completedTaskCount: completedCount,
    deferredTaskCount: deferredCount,
    rejectedTaskCount: rejectedCount,
    restoredTaskCount: restoredCount,
    scheduledMeetingCount,
    scheduledMeetingMinutes: Number(row.meeting_minutes ?? summary.meetingMinutes ?? 0),
    plannedTaskMinutes: Number(row.planned_task_minutes ?? summary.plannedTaskMinutes ?? 0),
    completedTaskMinutes: Number(row.completed_task_minutes ?? summary.completedTaskMinutes ?? 0),
    spilloverTaskCount: Number(row.spillover_task_count ?? summary.spilloverTaskCount ?? 0),
    agreementPercent,
    completionPercent
  };
}

function whyNotSelected(task: Task, todayTaskIds: Set<number>, guidance: string) {
  if (task.status === "Completed") {
    return null;
  }
  if (todayTaskIds.has(task.id)) {
    return null;
  }
  if (task.deferredUntil) {
    return "Deferred tasks are intentionally kept out of today’s active plan until they are due.";
  }
  return guidance || "It remained outside today’s selected blocks because current capacity was already filled by more urgent work.";
}

function whyPriority(task: Task) {
  if (task.priorityReason) return task.priorityReason;
  if (task.priority === "High") return "High because it has stronger urgency, active-work, or deadline signals.";
  if (task.priority === "Medium") return "Medium because it matters, but more urgent work is ahead of it today.";
  return "Low because it is still relevant, but not the most urgent work for today.";
}

export function getInsightsOverviewPayload(): InsightsOverview {
  const rows7 = listRecentDailyPlanSnapshots(7).map(toDaySummary);
  const rows30 = listRecentDailyPlanSnapshots(30).map(toDaySummary);
  return {
    activeTaskCount: listTasks().length,
    deferredTaskCount: listDeferredTasks().length,
    rejectedTaskCount: getRejectedTaskCount(),
    ignoredTaskCount: getIgnoredRejectedTaskCount(),
    reminderCount: listReminderItems(["active"]).length,
    completionRate7d: average(rows7.map((row) => (row.completionPercent ?? 0) / 100)),
    completionRate30d: average(rows30.map((row) => (row.completionPercent ?? 0) / 100)),
    averagePlannedMinutes7d: average(rows7.map((row) => row.plannedTaskMinutes)),
    averageCompletedMinutes7d: average(rows7.map((row) => row.completedTaskMinutes)),
    topInsights: getLatestPreferenceMemorySnapshot().insights,
    latestRun: listPlannerRunDetails(1)[0] ?? null
  };
}

export function getInsightsTodayPayload(): InsightsTodayPayload {
  const today = getTodaySnapshot();
  const dayPlanTaskBlocks = today.dayPlan.blocks.filter((block) => block.kind === "task");
  const taskIdsInPlan = new Set(dayPlanTaskBlocks.map((block) => block.taskId).filter((taskId): taskId is number => taskId !== null));
  const tasks = [...today.tasks.High, ...today.tasks.Medium, ...today.tasks.Low].map((task): InsightsTodayTask => {
    const block = dayPlanTaskBlocks.find((entry) => entry.taskId === task.id) ?? null;
    return {
      task,
      inDayPlan: taskIdsInPlan.has(task.id),
      planBlockTitle: block?.title ?? null,
      whyToday:
        task.selectionReason ??
        task.priorityExplanation ??
        (task.status === "In Progress"
          ? "This task is already underway, so the planner keeps it visible today."
          : "This task fits within today’s capacity and relevance signals."),
      whyPriority: whyPriority(task),
      whyNotHigher:
        task.priority === "Low"
          ? "Higher-priority work currently has stronger urgency, active-work, or recency signals."
          : task.priority === "Medium"
            ? "It matters, but active or more urgent work is ahead of it."
            : null,
      whyNotSelected: whyNotSelected(task, taskIdsInPlan, today.dayPlan.summary.guidance)
    };
  });

  return {
    generatedAt: today.sync.lastGeneratedAt,
    tasks,
    rejected: listRejectedTasks(),
    ignored: listIgnoredRejectedTasks()
  };
}

export function getInsightsHistoryPayload(limit = 30) {
  return {
    days: listRecentDailyPlanSnapshots(limit).map(toDaySummary)
  };
}

export function getInsightsHistoryDayPayload(dayKey: string): DayHistoryDetail | null {
  const row = getDailyPlanSnapshot(dayKey);
  if (!row) return null;

  const summary = toDaySummary(row);
  const blocks = parseJsonValue<DayPlanBlock[]>(row.blocks_json, []);
  const plannedTasks = blocks
    .filter((block) => block.kind === "task")
    .map((block) => ({
      taskId: block.taskId,
      title: block.title,
      minutes: block.durationMinutes,
      source: block.source,
      priority: block.priority,
      status: block.status
    }));

  const allTasks = listTasks(undefined, { includeDeferred: true });
  const completedTasks = allTasks.filter((task) => task.completedAt?.slice(0, 10) === dayKey);
  const deferredTasks = allTasks.filter((task) => task.deferredUntil?.slice(0, 10) === dayKey);
  const rejectedTasks = listRejectedTasks().filter((task) => task.rejectedAt.slice(0, 10) === dayKey);
  const ignoredTasks = listIgnoredRejectedTasks().filter((task) => task.updatedAt.slice(0, 10) === dayKey);
  const changeEvents = listTaskStateEvents({ dayKey });

  return {
    summary,
    plannedTasks,
    completedTasks,
    deferredTasks,
    rejectedTasks,
    ignoredTasks,
    changeEvents
  };
}

export function getTaskInsightsPayload(taskId: number): TaskInsightsPayload | null {
  const task = getTaskById(taskId);
  if (!task) return null;
  const recentEvents = listTaskStateEvents({ taskId, limit: 40 });
  const decisionEvents = listRecentDecisionLogs(120)
    .filter((event) => Number(event.task_id ?? -1) === taskId)
    .map((event): BehaviorFeedbackEvent => ({
      action: event.action as BehaviorFeedbackEvent["action"],
      source: event.source as BehaviorFeedbackEvent["source"],
      sourceRef: (event.source_ref as string | null) ?? null,
      beforePriority: (event.before_priority as Task["priority"] | null) ?? null,
      afterPriority: (event.after_priority as Task["priority"] | null) ?? null,
      inferredReason: (event.inferred_reason as string | null) ?? null,
      inferredReasonTag: (event.inferred_reason_tag as Task["decisionReasonTags"][number] | null) ?? null,
      preferencePolarity: event.preference_polarity as BehaviorFeedbackEvent["preferencePolarity"],
      createdAt: String(event.created_at)
    }));

  return {
    task,
    recentEvents,
    decisionEvents,
    reasoning: {
      selectionReason: task.selectionReason,
      priorityReason: task.priorityReason,
      scoreBreakdown: task.scoreBreakdown,
      historySignals: task.historySignals
    }
  };
}

export function getDiagnosticsPayload() {
  return {
    runs: listPlannerRunDetails(20),
    profile: getUserPriorityProfile(),
    memory: getLatestPreferenceMemorySnapshot(),
    latestDecisionEvents: listRecentDecisionLogs(20)
  };
}
