export type TaskSource = "Email" | "Jira" | "Manual";
export type TaskPriority = "High" | "Medium" | "Low";
export type TaskStatus = "Not Started" | "In Progress" | "Completed";
export type TaskStage = "Now" | "Next" | "Later" | "Review";
export type TaskEffortBucket = "15 min" | "30 min" | "1 hour" | "2+ hours";
export type ReminderStatus = "active" | "dismissed" | "resolved";
export type ReminderKind = "email_follow_up" | "jira_stale" | "deferred_due" | "meeting_prep";
export type WorkloadState = "Underloaded" | "Balanced" | "Overloaded";
export type TaskDecisionState = "accepted" | "rejected" | "uncertain" | "restored" | "ignored";
export type FilteringStyle = "conservative" | "balanced" | "aggressive";
export type PriorityBias = "focus" | "balanced" | "coverage";
export type DayPlanBlockKind = "task" | "meeting" | "buffer";
export type DayPlanBlockStatus = "planned" | "in_progress" | "up_next" | "completed" | "ended";
export type FeedbackPolarity = "positive" | "negative" | "neutral";
export type FeedbackAction =
  | "system_evaluated"
  | "reject"
  | "restore"
  | "stage_changed"
  | "priority_changed"
  | "status_changed"
  | "deferred"
  | "completed"
  | "always_ignore_similar"
  | "should_have_been_included";
export type AuditLogLevel = "debug" | "info" | "warn" | "error";
export type AuditEventStatus = "started" | "success" | "failure" | "updated" | "skipped" | "info";

export type ReasonTag =
  | "direct_request"
  | "manager_visibility"
  | "project_critical"
  | "comment_noise"
  | "newsletter_like"
  | "duplicate_signal"
  | "meeting_related"
  | "historically_rejected"
  | "historically_accepted"
  | "assigned_work"
  | "due_soon"
  | "bot_generated"
  | "review_request"
  | "blocking_work"
  | "fyi_only";

export interface JiraPlanningSubtask {
  key: string;
  title: string;
  status: string | null;
  estimateSeconds: number | null;
}

export interface ScoreBreakdownItem {
  label: string;
  value: number;
  kind: "positive" | "negative" | "neutral";
}

export interface Task {
  id: number;
  title: string;
  source: TaskSource;
  stage: TaskStage;
  stageOrder: number;
  priority: TaskPriority;
  status: TaskStatus;
  sourceLink: string | null;
  sourceRef: string | null;
  sourceThreadRef: string | null;
  jiraStatus: string | null;
  ignored: number;
  deferredUntil: string | null;
  reminderState: ReminderStatus | null;
  lastRemindedAt: string | null;
  estimatedEffortBucket: TaskEffortBucket | null;
  jiraEstimateSeconds: number | null;
  jiraSubtaskEstimateSeconds: number | null;
  jiraPlanningSubtasks: JiraPlanningSubtask[];
  priorityScore: number | null;
  priorityExplanation: string | null;
  selectionReason: string | null;
  priorityReason: string | null;
  scoreBreakdown: ScoreBreakdownItem[];
  historySignals: string[];
  taskAgeDays: number;
  carryForwardCount: number;
  completedAt: string | null;
  lastActivityAt: string | null;
  lastChangedBy: string | null;
  lastChangedAt: string | null;
  manualOverrideFlags: string[];
  decisionState: TaskDecisionState | null;
  decisionConfidence: number | null;
  decisionReason: string | null;
  decisionReasonTags: ReasonTag[];
  personalizationVersion: number | null;
  wasUserOverridden: boolean;
  restoredAt: string | null;
  rejectedAt: string | null;
  createdAt: string;
  updatedAt: string;
}

export interface Meeting {
  id: number;
  externalId: string | null;
  title: string;
  startTime: string;
  endTime: string;
  timeZone: string | null;
  durationMinutes: number;
  meetingLink: string | null;
  meetingLinkType: "join" | "calendar" | null;
  isCancelled: boolean;
  attendanceStatus: "attending" | "unattending";
  createdAt: string;
}

export interface MeetingDetail {
  id: number;
  title: string;
  startTime: string;
  endTime: string;
  timeZone: string | null;
  durationMinutes: number;
  meetingLink: string | null;
  meetingLinkType: "join" | "calendar" | null;
  isCancelled: boolean;
  attendanceStatus: "attending" | "unattending";
  organizer: string | null;
  attendees: string[];
  location: string | null;
  description: string | null;
  bodyPreview: string | null;
}

export interface IntegrationConnection {
  provider: "microsoft" | "jira";
  status: "connected" | "disconnected" | "error";
  accountLabel: string | null;
  configJson: string | null;
  accessToken: string | null;
  refreshToken: string | null;
  expiresAt: string | null;
  errorMessage: string | null;
  updatedAt: string;
}

export interface Reminder {
  id: number;
  reminderKey: string;
  taskId: number | null;
  kind: ReminderKind;
  title: string;
  reason: string;
  status: ReminderStatus;
  sourceLink: string | null;
  sourceLabel: string | null;
  scheduledFor: string | null;
  createdAt: string;
  updatedAt: string;
  dismissedAt: string | null;
  throttleUntil: string | null;
}

export interface AutomationSettings {
  scheduleEnabled: boolean;
  scheduleTimeLocal: string;
  scheduleTimezone: string;
  workdayStartLocal: string;
  workdayEndLocal: string;
  remindersEnabled: boolean;
  reminderCadenceHours: number;
  desktopNotificationsEnabled: boolean;
  lastAutoGeneratedAt: string | null;
  schedulerLastRunAt: string | null;
  schedulerLastStatus: "idle" | "ok" | "error";
  schedulerLastError: string | null;
}

export interface UserPriorityProfile {
  personalizationEnabled: boolean;
  roleFocus: string | null;
  prioritizationPrompt: string | null;
  importantWork: string[];
  noiseWork: string[];
  mustNotMiss: string[];
  importantSources: string[];
  importantPeople: string[];
  importantProjects: string[];
  positiveReasonTags: ReasonTag[];
  negativeReasonTags: ReasonTag[];
  filteringStyle: FilteringStyle;
  priorityBias: PriorityBias;
  questionnaireJson: string | null;
  exampleRankingsJson: string | null;
  lastProfileRefreshAt: string | null;
  updatedAt: string | null;
}

export interface RejectedTask {
  id: number;
  title: string;
  source: TaskSource;
  sourceLink: string | null;
  sourceRef: string | null;
  sourceThreadRef: string | null;
  jiraStatus: string | null;
  proposedPriority: TaskPriority;
  decisionState: TaskDecisionState;
  decisionConfidence: number | null;
  decisionReason: string | null;
  decisionReasonTags: ReasonTag[];
  personalizationVersion: number | null;
  candidatePayloadJson: string | null;
  rejectedAt: string;
  restoredAt: string | null;
  updatedAt: string;
}

export interface PersonalizationInsight {
  statement: string;
  confidence: number;
  source: "profile" | "history";
}

export interface BehaviorFeedbackEvent {
  action: FeedbackAction;
  source: TaskSource | "Calibration";
  sourceRef: string | null;
  beforePriority: TaskPriority | null;
  afterPriority: TaskPriority | null;
  inferredReason: string | null;
  inferredReasonTag: ReasonTag | null;
  preferencePolarity: FeedbackPolarity;
  createdAt: string;
}

export interface WorkloadSummary {
  totalMeetingMinutes: number;
  totalTaskMinutes: number;
  totalPlannedMinutes: number;
  state: WorkloadState;
}

export interface DayPlanBlock {
  id: string;
  kind: DayPlanBlockKind;
  title: string;
  startTime: string;
  endTime: string;
  timeZone: string | null;
  durationMinutes: number;
  status: DayPlanBlockStatus;
  taskId: number | null;
  meetingId: number | null;
  source: TaskSource | "Calendar" | null;
  priority: TaskPriority | null;
  link: string | null;
  note: string | null;
}

export interface DayPlanSummary {
  dayKey: string;
  baseWorkdayMinutes: number;
  adaptedTaskCapacityMinutes: number;
  remainingTaskCapacityMinutes: number;
  meetingMinutes: number;
  completedTaskMinutes: number;
  plannedTaskMinutes: number;
  remainingTaskMinutes: number;
  spilloverTaskCount: number;
  freeMinutes: number;
  focusFactor: number;
  completionRate: number;
  guidance: string;
}

export interface DayPlan {
  summary: DayPlanSummary;
  blocks: DayPlanBlock[];
  spilloverTasks: Task[];
}

export interface TaskBoardPayload {
  now: Task[];
  next: Task[];
  later: Task[];
  review: Task[];
  rejected: RejectedTask[];
  ignoredRejected: RejectedTask[];
}

export interface TodayPayload {
  meetings: Meeting[];
  tasks: Record<TaskPriority, Task[]>;
  reminders: Reminder[];
  workload: WorkloadSummary;
  dayPlan: DayPlan;
  deferredTaskCount: number;
  rejectedTaskCount: number;
  automation: AutomationSettings;
  sync: {
    microsoft: string | null;
    jira: string | null;
    lastGeneratedAt: string | null;
  };
  warnings: string[];
}

export interface HomeBanner {
  title: string;
  totalTaskCount: number;
  emailTaskCount: number;
  jiraTaskCount: number;
  meetingCount: number;
  followUpCount: number;
  inProgressTaskCount: number;
  highPriorityTaskCount: number;
  summary: string;
  upcomingMeetingLabel: string | null;
}

export interface HomeScheduleEntry {
  entryId: string;
  dayKey: string;
  taskId: number;
  startMinutes: number;
  durationMinutes: number;
  source: "planner" | "user";
  createdAt: string;
  updatedAt: string;
}

export interface HomeSchedule {
  dayKey: string;
  entries: HomeScheduleEntry[];
  hiddenMeetingIds: number[];
  hasEmptyWorkingSlots: boolean;
}

export interface HomePayload {
  banner: HomeBanner;
  tasks: Task[];
  meetings: Meeting[];
  deferredTasks: Task[];
  reminders: Reminder[];
  schedule: HomeSchedule;
}

export interface AuditEvent {
  id: number;
  timestamp: string;
  level: AuditLogLevel;
  eventType: string;
  requestId: string | null;
  runId: string | null;
  entityType: string | null;
  entityId: string | null;
  provider: string | null;
  status: AuditEventStatus;
  source: string | null;
  message: string;
  metadataJson: string | null;
}

export interface PlannerRunDetail {
  runId: string;
  triggerType: "manual" | "scheduled" | "sync";
  preferredTimeZone: string | null;
  warnings: string[];
  meetingCount: number;
  activeTaskCount: number;
  rejectedTaskCount: number;
  deferredTaskCount: number;
  workloadState: WorkloadState | null;
  createdAt: string;
  updatedAt: string;
}

export interface TaskStateEvent {
  id: number;
  taskId: number | null;
  source: TaskSource | "Calibration";
  sourceRef: string | null;
  sourceThreadRef: string | null;
  eventType: string;
  actor: "system" | "user" | "client";
  reason: string | null;
  beforeJson: string | null;
  afterJson: string | null;
  createdAt: string;
}

export interface InsightsOverview {
  activeTaskCount: number;
  deferredTaskCount: number;
  rejectedTaskCount: number;
  ignoredTaskCount: number;
  reminderCount: number;
  completionRate7d: number | null;
  completionRate30d: number | null;
  averagePlannedMinutes7d: number | null;
  averageCompletedMinutes7d: number | null;
  topInsights: PersonalizationInsight[];
  latestRun: PlannerRunDetail | null;
}

export interface InsightsTodayTask {
  task: Task;
  inDayPlan: boolean;
  planBlockTitle: string | null;
  whyToday: string;
  whyPriority: string;
  whyNotHigher: string | null;
  whyNotSelected: string | null;
}

export interface InsightsTodayPayload {
  generatedAt: string | null;
  tasks: InsightsTodayTask[];
  rejected: RejectedTask[];
  ignored: RejectedTask[];
}

export interface DayHistorySummary {
  dayKey: string;
  guidance: string;
  plannedTaskCount: number;
  completedTaskCount: number;
  removedTaskCount: number;
  deferredTaskCount: number;
  rejectedTaskCount: number;
  restoredTaskCount: number;
  stageChangedCount: number;
  scheduledMeetingCount: number;
  scheduledMeetingMinutes: number;
  plannedTaskMinutes: number;
  completedTaskMinutes: number;
  spilloverTaskCount: number;
  spilloverPercent: number | null;
  agreementPercent: number | null;
  acceptancePercent: number | null;
  completionPercent: number | null;
  emailTaskCount: number;
  jiraTaskCount: number;
  manualTaskCount: number;
}

export interface DayHistoryDetail {
  summary: DayHistorySummary;
  plannedTasks: Array<{
    taskId: number | null;
    title: string;
    minutes: number;
    source: TaskSource | "Calendar" | null;
    priority: TaskPriority | null;
    status: DayPlanBlockStatus;
  }>;
  completedTasks: Task[];
  deferredTasks: Task[];
  rejectedTasks: RejectedTask[];
  ignoredTasks: RejectedTask[];
  changeEvents: TaskStateEvent[];
}

export interface InsightsUpdatesPayload {
  startDayKey: string | null;
  endDayKey: string | null;
  totalEvents: number;
  events: TaskStateEvent[];
}

export interface TaskInsightsPayload {
  task: Task;
  recentEvents: TaskStateEvent[];
  decisionEvents: BehaviorFeedbackEvent[];
  reasoning: {
    selectionReason: string | null;
    priorityReason: string | null;
    scoreBreakdown: ScoreBreakdownItem[];
    historySignals: string[];
  };
}

export interface JiraDetail {
  type: "jira";
  key: string;
  title: string;
  status: string | null;
  statusCategory: TaskStatus;
  priority: string | null;
  transitions: Array<{
    id: string;
    name: string;
    toStatus: string | null;
    toStatusCategory: TaskStatus;
  }>;
  description: string | null;
  storyPoints: number | null;
  assignee: string | null;
  reporter: string | null;
  labels: string[];
  subtasks: Array<{
    key: string;
    title: string;
    status: string | null;
    statusCategory: TaskStatus;
    transitions: Array<{
      id: string;
      name: string;
      toStatus: string | null;
      toStatusCategory: TaskStatus;
    }>;
  }>;
  comments: Array<{
    author: string;
    createdAt: string | null;
    body: string;
  }>;
  worklogs: Array<{
    author: string;
    startedAt: string | null;
    timeSpent: string | null;
    comment: string | null;
  }>;
}

export interface EmailThreadMessage {
  id: string;
  from: string | null;
  to: string[];
  cc: string[];
  subject: string | null;
  receivedAt: string | null;
  body: string;
}

export interface EmailDetail {
  type: "email";
  from: string | null;
  to: string[];
  cc: string[];
  subject: string | null;
  receivedAt: string | null;
  body: string;
  thread: EmailThreadMessage[];
}

export type TaskDetail = JiraDetail | EmailDetail;

export interface EmailReplyDraft {
  subject: string;
  to: string[];
  cc: string[];
  body: string;
  summary: string;
  actionItems: string[];
  rationale: string;
}

export interface MeetingPrep {
  title: string;
  summary: string;
  objectives: string[];
  checklist: string[];
  talkingPoints: string[];
  questions: string[];
  risks: string[];
  notes: string;
  rationale: string;
}
