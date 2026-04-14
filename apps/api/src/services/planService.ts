import {
  clearRejectedTasksBySourceThread,
  deleteTask,
  getDecisionEventCount,
  getAutomationSettings,
  getIntegrationConnection,
  getLatestPreferenceMemorySnapshot,
  getRejectedTaskBySource,
  getTaskBySource,
  getTaskBySourceThread,
  getRejectedTaskCount,
  getSyncState,
  getUserPriorityProfile,
  groupTasksByPriority,
  listRejectedTasks,
  listTaskStateEvents,
  listDeferredTasks,
  listRecentDailyPlanSnapshots,
  listMeetings,
  listReminderItems,
  listRecentDecisionLogs,
  listTasks,
  logTaskDecisionEvent,
  recordGenerationRun,
  replaceMeetings,
  resolveStaleReminders,
  saveAutomationSettings,
  savePreferenceMemorySnapshot,
  setSyncState,
  upsertPlannerRunDetail,
  upsertDailyPlanSnapshot,
  updateRejectedTask,
  updateTask,
  upsertRejectedTask,
  upsertReminder,
  upsertTask
} from "../db.js";
import { env } from "../env.js";
import {
  buildJiraIssueBrowseUrl,
  fetchJiraIssuePlanningContext,
  fetchOpenAssignedIssues,
  getMappedJiraPriority,
  mapJiraWorkflowStatus
} from "../providers/jira.js";
import {
  fetchRecentEmails,
  fetchRecentSentEmails,
  fetchRecentSentEmailsWithAccessToken,
  fetchRecentEmailsWithAccessToken,
  fetchTodaysMeetings,
  fetchTodaysMeetingsWithAccessToken
} from "../providers/microsoft.js";
import { classifyEmail } from "./emailClassifier.js";
import {
  defaultPriorityProfile,
  distillPreferenceMemory,
  evaluateCandidateWithPersonalization
} from "./personalization.js";
import { createCorrelationId, logEvent } from "./logger.js";
import type { GraphEvent, GraphMail } from "../providers/microsoft.js";
import type {
  DayPlan,
  DayPlanBlock,
  ReminderKind,
  RejectedTask,
  ReasonTag,
  PlannerRunDetail,
  Task,
  TaskEffortBucket,
  TaskPriority,
  TaskStage,
  TaskSource,
  TaskStatus,
  TodayPayload,
  WorkloadSummary
} from "../types.js";

const DEFAULT_WORKDAY_START_LOCAL = "09:30";
const DEFAULT_WORKDAY_END_LOCAL = "18:00";

type PlanningRole = "starter" | "major" | "ender" | "review";

function isPlannableMeeting(meeting: { isCancelled: boolean; attendanceStatus?: "attending" | "unattending" }) {
  return !meeting.isCancelled && (meeting.attendanceStatus ?? "attending") !== "unattending";
}

function startAndEndOfAgendaWindow() {
  const start = new Date();
  start.setDate(start.getDate() - 2);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + 7);
  return { startIso: start.toISOString(), endIso: end.toISOString() };
}

function parseGraphCalendarDateTime(value?: string | null) {
  if (!value) return null;
  if (/[zZ]|[+-]\d{2}:\d{2}$/.test(value)) {
    return new Date(value);
  }
  return new Date(`${value}Z`);
}

function parseMeetingDateWithTimeZone(value: string | null, timeZone?: string | null) {
  if (!value) return new Date(NaN);
  if (/[zZ]|[+-]\d{2}:\d{2}$/.test(value)) {
    return new Date(value);
  }
  return new Date(timeZone?.trim().toUpperCase() === "UTC" ? `${value}Z` : value);
}

function isRecentlyUpdated(isoValue?: string | null, windowDays = 2) {
  if (!isoValue) return false;
  const updatedAt = new Date(isoValue).getTime();
  if (Number.isNaN(updatedAt)) return false;
  return Date.now() - updatedAt <= windowDays * 86_400_000;
}

function isLegacyDailyWrapUpTask(task: Task) {
  return task.source === "Manual" && task.sourceRef?.startsWith("daily-wrap-up:");
}

function cleanupLegacyDailyWrapUpTasks() {
  for (const task of listTasks(undefined, { includeDeferred: true })) {
    if (!isLegacyDailyWrapUpTask(task)) continue;
    deleteTask(task.id);
  }
}

async function getPersonalizationContext() {
  const profile = getUserPriorityProfile() ?? defaultPriorityProfile;
  const memory = (() => {
    const latest = getLatestPreferenceMemorySnapshot();
    try {
      return {
        version: Number(JSON.parse(latest.snapshotJson).version ?? latest.sourceEventCount ?? 1),
        ...JSON.parse(latest.snapshotJson)
      };
    } catch {
      return {
        version: 1,
        positiveTags: [],
        negativeTags: [],
        repeatedWins: [],
        repeatedNoise: []
      };
    }
  })();
  const recentExamples = listRecentDecisionLogs(20).map((row) => {
    let title = String(row.source_ref ?? "Task");
    try {
      if (row.feedback_payload_json) {
        const parsed = JSON.parse(String(row.feedback_payload_json)) as { title?: string };
        title = parsed.title ?? title;
      }
    } catch {
      title = String(row.source_ref ?? "Task");
    }
    return {
      title,
      source: String(row.source),
      outcome: String(row.action),
      reason: (row.decision_reason as string | null) ?? null
    };
  });
  const recentRejectedExamples = listRejectedTasks()
    .slice(0, 15)
    .map((task) => ({
      title: task.title,
      source: task.source,
      outcome: task.decisionState,
      reason: task.decisionReason
    }));
  return { profile, memory, recentExamples, recentRejectedExamples };
}

function buildFeedbackPayload(candidate: {
  title: string;
  source: TaskSource;
  sourceRef?: string | null;
  sourceThreadRef?: string | null;
  sourceLink?: string | null;
  jiraStatus?: string | null;
}) {
  return JSON.stringify(candidate);
}

function buildOutlookCalendarItemLink(eventId: string) {
  return `https://outlook.office365.com/owa/?itemid=${encodeURIComponent(eventId)}&exvsurl=1&path=/calendar/item`;
}

function normalizeOptionalString(value?: string | null) {
  const normalized = value?.trim();
  return normalized ? normalized : null;
}

function normalizeEmailText(value?: string | null) {
  return (value ?? "").toLowerCase().replace(/\s+/g, " ").trim();
}

function emailContentForDecision(email: GraphMail) {
  return `${email.bodyPreview ?? ""} ${(email.body?.content ?? "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim()}`
    .trim()
    .slice(0, 6000);
}

async function callOpenAIJson<T>(input: unknown, schema: object, name: string, systemPrompt: string): Promise<T | null> {
  if (!env.openAiApiKey) {
    return null;
  }

  try {
    const response = await fetch(`${env.openAiApiBaseUrl}/v1/responses`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${env.openAiApiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: env.openAiModel,
        input: [
          {
            role: "system",
            content: [{ type: "input_text", text: systemPrompt }]
          },
          {
            role: "user",
            content: [{ type: "input_text", text: JSON.stringify(input) }]
          }
        ],
        text: {
          format: {
            type: "json_schema",
            name,
            schema
          }
        }
      })
    });

    if (!response.ok) return null;
    const json = (await response.json()) as { output_text?: string };
    if (!json.output_text) return null;
    return JSON.parse(json.output_text) as T;
  } catch {
    return null;
  }
}

function extractIssueKeys(text: string) {
  return [...new Set((text.toUpperCase().match(/\b[A-Z][A-Z0-9]+-\d+\b/g) ?? []).map((key) => key.trim()))];
}

function buildMeetingTitleIndex(meetings: ReturnType<typeof listMeetings>) {
  return meetings
    .filter((meeting) => isPlannableMeeting(meeting))
    .map((meeting) => normalizeEmailText(meeting.title))
    .filter((title) => title.length >= 10);
}

function triageEmailForWork(input: {
  email: GraphMail;
  classification: ReturnType<typeof classifyEmail> extends Promise<infer T> ? T : never;
  jiraIssueKeys: Set<string>;
  meetingTitles: string[];
}): {
  route: "accept" | "review" | "drop";
  priority: TaskPriority;
  reason: string;
  reasonTags: ReasonTag[];
} {
  const sender = normalizeEmailText(input.email.from?.emailAddress?.address ?? "");
  const subject = normalizeEmailText(input.email.subject);
  const bodyText = normalizeEmailText(
    `${input.email.bodyPreview ?? ""} ${(input.email.body?.content ?? "").slice(0, 1200)}`
  );
  const fullText = `${subject} ${bodyText}`.trim();
  const issueKeys = extractIssueKeys(`${input.email.subject ?? ""} ${input.email.bodyPreview ?? ""}`);
  const hasMatchingJiraIssue = issueKeys.some((key) => input.jiraIssueKeys.has(key));
  const isMeetingRelated =
    /(calendar|meeting invite|invitation|accepted:|declined:|tentative:|rescheduled|reschedule|join microsoft teams|zoom meeting|webex|google meet)/.test(
      fullText
    ) || input.meetingTitles.some((title) => subject.includes(title));
  const isLikelyJiraNotification =
    issueKeys.length > 0 &&
    !/(pull request|merge request|review requested|opened a pull request|github|gitlab)/.test(fullText) &&
    /(\[jira\]|jira|comment(ed)?|comment added|mentioned you|status|open|reopened|resolved|updated)/.test(fullText);

  if (isMeetingRelated) {
    return {
      route: "drop" as const,
      priority: "Low" as TaskPriority,
      reason: "Meeting email is already represented in the calendar timeline.",
      reasonTags: ["meeting_related"] as ReasonTag[]
    };
  }

  if (isLikelyJiraNotification && hasMatchingJiraIssue) {
    return {
      route: "drop" as const,
      priority: "Low" as TaskPriority,
      reason: "Jira notification duplicates an existing Jira task.",
      reasonTags: ["duplicate_signal"] as ReasonTag[]
    };
  }

  if (input.classification.actionable && input.classification.priority === "High") {
    return {
      route: "accept" as const,
      priority: input.classification.priority,
      reason: input.classification.why,
      reasonTags: input.classification.reasonTags
    };
  }

  return {
    route: "review" as const,
    priority: input.classification.priority,
    reason: input.classification.why,
    reasonTags: input.classification.reasonTags
  };
}

function emailBodyText(email: GraphMail) {
  return normalizeEmailText(`${email.bodyPreview ?? ""} ${email.body?.content ?? ""}`);
}

function isAnnouncementLikeEmail(email: GraphMail) {
  const text = normalizeEmailText(`${email.subject ?? ""} ${email.bodyPreview ?? ""} ${email.body?.content ?? ""}`);
  const sender = normalizeEmailText(email.from?.emailAddress?.address ?? "");
  return (
    /(release highlights|release showcase|newsletter|announcement|digest|what'?s new|monthly deep dive|hackathon|optional|town hall|all hands|showcase|scam of the week|thank you and best wishes|org update|company update|release notes|webinar|event reminder|social)/.test(
      text
    ) ||
    /(newsletter|announce|communications|events|marketing)/.test(sender)
  );
}

function needsSentEmailFollowUp(sentEmail: GraphMail, inboxEmails: GraphMail[]) {
  const sentAt = sentEmail.sentDateTime ? new Date(sentEmail.sentDateTime).getTime() : 0;
  if (!sentAt || Number.isNaN(sentAt)) return false;
  const hoursSinceSent = (Date.now() - sentAt) / 3_600_000;
  if (hoursSinceSent < 18 || hoursSinceSent > 7 * 24) return false;
  if (isAnnouncementLikeEmail(sentEmail)) return false;

  const body = emailBodyText(sentEmail);
  const subject = normalizeEmailText(sentEmail.subject);
  const text = `${subject} ${body}`;
  const toCount = sentEmail.toRecipients?.length ?? 0;
  const hasActionSignal =
    /(please|can you|could you|let me know|review|approve|confirm|reply|respond|follow up|share|send|need|action required|waiting on|blocker|update me)/.test(
      text
    ) || /\?$/.test((sentEmail.bodyPreview ?? sentEmail.subject ?? "").trim());
  if (!hasActionSignal) return false;
  if (toCount === 0) return false;

  const hasReply = inboxEmails.some((mail) => {
    const sameThread =
      sentEmail.conversationId && mail.conversationId
        ? sentEmail.conversationId === mail.conversationId
        : normalizeEmailText(mail.subject).includes(subject) || subject.includes(normalizeEmailText(mail.subject));
    const receivedAt = mail.receivedDateTime ? new Date(mail.receivedDateTime).getTime() : 0;
    return sameThread && receivedAt > sentAt;
  });

  return !hasReply;
}

function persistRejectedCandidate(input: {
  title: string;
  source: TaskSource;
  sourceLink?: string | null;
  sourceRef?: string | null;
  sourceThreadRef?: string | null;
  jiraStatus?: string | null;
  proposedPriority: TaskPriority;
  decisionConfidence: number;
  decisionReason: string;
  decisionReasonTags: RejectedTask["decisionReasonTags"];
  candidatePayloadJson: string;
  personalizationVersion: number;
}) {
  const rejected = upsertRejectedTask({
    ...input,
    decisionState: "rejected"
  });
  if (rejected?.decisionState === "ignored") {
    return rejected;
  }
  logTaskDecisionEvent({
    source: input.source,
    sourceRef: input.sourceRef ?? null,
    sourceThreadRef: input.sourceThreadRef ?? null,
    action: "reject",
    afterPriority: input.proposedPriority,
    systemDecisionState: "rejected",
    decisionConfidence: input.decisionConfidence,
    decisionReason: input.decisionReason,
    decisionReasonTags: input.decisionReasonTags,
    feedbackPayloadJson: input.candidatePayloadJson,
    preferencePolarity: "negative"
  });
  return rejected;
}

function clearRejectedCandidate(source: TaskSource, sourceRef: string | null) {
  if (!sourceRef) return;
  const rejected = getRejectedTaskBySource(source, sourceRef);
  if (rejected && rejected.decisionState !== "restored") {
    updateRejectedTask(rejected.id, {
      decisionState: "restored",
      restoredAt: new Date().toISOString()
    });
  }
}

function clearRejectedCandidatesForThread(
  source: TaskSource,
  sourceRef: string | null,
  sourceThreadRef: string | null
) {
  clearRejectedCandidate(source, sourceRef);
  clearRejectedTasksBySourceThread(source, sourceThreadRef);
}

async function applyMicrosoftResults(
  emails: GraphMail[],
  meetings: GraphEvent[],
  calendarTimeZone: string | null
) {
  const personalization = await getPersonalizationContext();
  const jiraIssueKeys = new Set(
    listTasks(undefined, { includeDeferred: true })
      .filter((task) => task.source === "Jira")
      .flatMap((task) => [
        ...(task.sourceRef ? [String(task.sourceRef)] : []),
        ...task.jiraPlanningSubtasks.map((subtask) => subtask.key)
      ])
      .filter(Boolean)
  );
  const meetingTitles = buildMeetingTitleIndex(listMeetings());

  for (const email of emails) {
    const classification = await classifyEmail(email, {
      profile: personalization.profile,
      recentExamples: personalization.recentExamples,
      recentRejectedExamples: personalization.recentRejectedExamples
    });
    const now = new Date().toISOString();
    const emailContent = emailContentForDecision(email);
    const reasonTagSet = new Set(classification.reasonTags);
    const candidatePayload = {
      title: classification.title,
      source: "Email" as const,
      sourceLink: email.webLink ?? null,
      sourceRef: email.id,
      sourceThreadRef: email.conversationId ?? null,
      sender: email.from?.emailAddress?.address ?? null,
      bodyPreview: emailContent || null,
      isAssignedToUser: reasonTagSet.has("assigned_work"),
      isDirectRequest: reasonTagSet.has("direct_request"),
      dueSoon: reasonTagSet.has("due_soon"),
      isBotLike:
        reasonTagSet.has("bot_generated") ||
        /(noreply|notification|service-now|automated)/i.test(
          `${email.from?.emailAddress?.address ?? ""} ${email.subject ?? ""}`
        ),
      isDuplicate: false,
      meetingRelevant: reasonTagSet.has("meeting_related")
    };

    const strongSignal =
      classification.actionable ||
      reasonTagSet.has("review_request") ||
      reasonTagSet.has("blocking_work") ||
      reasonTagSet.has("historically_accepted");
    const triage = triageEmailForWork({
      email,
      classification,
      jiraIssueKeys,
      meetingTitles
    });

    const exactTask = getTaskBySource("Email", email.id, { includeIgnored: false });
    const threadTask =
      exactTask ??
      getTaskBySourceThread("Email", email.conversationId ?? null, {
        includeIgnored: false
      });

    const preserveVisibleEmailThread = (input: {
      priority: TaskPriority;
      decisionState: Task["decisionState"];
      decisionConfidence: number;
      decisionReason: string;
      decisionReasonTags: Task["decisionReasonTags"];
    }) => {
      clearRejectedCandidatesForThread("Email", email.id, email.conversationId ?? null);
      if (threadTask && threadTask.sourceRef !== email.id) {
        updateTask(threadTask.id, {
          title: classification.title,
          priority: threadTask.manualOverrideFlags.includes("priority") ? undefined : input.priority,
          estimatedEffortBucket: classification.estimatedEffortBucket,
          lastActivityAt: now,
          decisionState: threadTask.decisionState === "restored" ? "restored" : input.decisionState,
          decisionConfidence: input.decisionConfidence,
          decisionReason: input.decisionReason,
          decisionReasonTags: input.decisionReasonTags,
          priorityExplanation: classification.why,
          selectionReason: `Included because the email likely needs action: ${classification.why}`,
          priorityReason: `Model-estimated priority ${input.priority.toLowerCase()} and effort ${classification.estimatedEffortBucket.toLowerCase()} based on the email and your preferences.`,
          personalizationVersion: personalization.memory.version ?? 1,
          rejectedAt: null
        });
      }
    };

    if (triage.route === "drop") {
      if (threadTask) {
        preserveVisibleEmailThread({
          priority: threadTask.priority,
          decisionState: threadTask.decisionState ?? "restored",
          decisionConfidence: 0.72,
          decisionReason: threadTask.decisionReason ?? "Similar work is already active from this email thread.",
          decisionReasonTags: threadTask.decisionReasonTags
        });
      }
      continue;
    }

    if (triage.route === "review" || (!classification.actionable && !strongSignal)) {
      if (threadTask) {
        preserveVisibleEmailThread({
          priority: threadTask.priority,
          decisionState: threadTask.decisionState ?? "restored",
          decisionConfidence: 0.72,
          decisionReason:
            threadTask.decisionReason ?? "Similar work is already active from this email thread.",
          decisionReasonTags: threadTask.decisionReasonTags
        });
        continue;
      }
      persistRejectedCandidate({
        title: classification.title,
        source: "Email",
        sourceLink: email.webLink ?? null,
        sourceRef: email.id,
        sourceThreadRef: email.conversationId ?? null,
        proposedPriority: triage.priority,
        decisionConfidence: 0.72,
        decisionReason: triage.reason,
        decisionReasonTags: triage.reasonTags,
        candidatePayloadJson: buildFeedbackPayload(candidatePayload),
        personalizationVersion: personalization.memory.version ?? 1
      });
      continue;
    }

    const evaluation = await evaluateCandidateWithPersonalization({
      candidate: candidatePayload,
      profile: personalization.profile,
      memory: personalization.memory,
      recentExamples: personalization.recentExamples
    });
    const payloadJson = buildFeedbackPayload(candidatePayload);

    const finalEvaluation =
      triage.route === "accept"
        ? {
            relevance: "accept" as const,
            priority:
              triage.priority === "High" || evaluation.priority === "High"
                ? "High" as TaskPriority
                : "Medium" as TaskPriority,
            confidence: Math.max(0.82, evaluation.confidence),
            why: triage.reason,
            reasonTags: [...new Set([...triage.reasonTags, ...evaluation.reasonTags])] as ReasonTag[]
          }
        : evaluation;
    const gatedEmailEvaluation =
      finalEvaluation.priority === "High"
        ? finalEvaluation
        : {
            ...finalEvaluation,
            relevance: "uncertain" as const,
            why:
              finalEvaluation.why ||
              "Email may matter, but it was not strong enough to become a high-priority task.",
            reasonTags: [...new Set([...finalEvaluation.reasonTags, "fyi_only"])] as ReasonTag[]
          };

    if (
      gatedEmailEvaluation.relevance === "reject" ||
      (gatedEmailEvaluation.relevance === "uncertain" && gatedEmailEvaluation.confidence < 0.6)
    ) {
      if (threadTask) {
        preserveVisibleEmailThread({
          priority: threadTask.priority,
          decisionState: threadTask.decisionState ?? "restored",
          decisionConfidence: gatedEmailEvaluation.confidence,
          decisionReason: threadTask.decisionReason ?? gatedEmailEvaluation.why,
          decisionReasonTags: gatedEmailEvaluation.reasonTags
        });
        if (exactTask) {
          upsertTask({
            title: classification.title,
            source: "Email",
            priority: exactTask.priority,
            status: exactTask.status,
            sourceLink: email.webLink ?? exactTask.sourceLink ?? null,
            sourceRef: email.id,
            sourceThreadRef: email.conversationId ?? null,
            estimatedEffortBucket: classification.estimatedEffortBucket,
            decisionState: exactTask.decisionState ?? "restored",
            decisionConfidence: gatedEmailEvaluation.confidence,
            decisionReason: exactTask.decisionReason ?? gatedEmailEvaluation.why,
            decisionReasonTags: gatedEmailEvaluation.reasonTags,
            priorityExplanation: classification.why,
            selectionReason: `Included because the email likely needs action: ${classification.why}`,
            priorityReason: `Model-estimated priority ${exactTask.priority.toLowerCase()} and effort ${classification.estimatedEffortBucket.toLowerCase()} based on the email and your preferences.`,
            personalizationVersion: personalization.memory.version ?? 1
          });
        }
        continue;
      }
      persistRejectedCandidate({
        title: classification.title,
        source: "Email",
        sourceLink: email.webLink ?? null,
        sourceRef: email.id,
        sourceThreadRef: email.conversationId ?? null,
        proposedPriority: gatedEmailEvaluation.priority,
        decisionConfidence: gatedEmailEvaluation.confidence,
        decisionReason: gatedEmailEvaluation.why,
        decisionReasonTags: gatedEmailEvaluation.reasonTags,
        candidatePayloadJson: payloadJson,
        personalizationVersion: personalization.memory.version ?? 1
      });
      continue;
    }

    clearRejectedCandidatesForThread("Email", email.id, email.conversationId ?? null);
    if (threadTask && !exactTask) {
      updateTask(threadTask.id, {
        title: classification.title,
        priority: threadTask.manualOverrideFlags.includes("priority") ? undefined : gatedEmailEvaluation.priority,
        estimatedEffortBucket: classification.estimatedEffortBucket,
        lastActivityAt: now,
        decisionState: threadTask.decisionState === "restored" ? "restored" : gatedEmailEvaluation.relevance === "uncertain" ? "uncertain" : "accepted",
        decisionConfidence: gatedEmailEvaluation.confidence,
        decisionReason: gatedEmailEvaluation.why,
        decisionReasonTags: gatedEmailEvaluation.reasonTags,
        priorityExplanation: classification.why,
        selectionReason: `Included because the email likely needs action: ${classification.why}`,
        priorityReason: `Model-estimated priority ${gatedEmailEvaluation.priority.toLowerCase()} and effort ${classification.estimatedEffortBucket.toLowerCase()} based on the email and your preferences.`,
        personalizationVersion: personalization.memory.version ?? 1,
        rejectedAt: null
      });
    } else {
      upsertTask({
        title: classification.title,
        source: "Email",
        priority: gatedEmailEvaluation.priority,
        sourceLink: email.webLink ?? null,
        sourceRef: email.id,
        sourceThreadRef: email.conversationId ?? null,
        estimatedEffortBucket: classification.estimatedEffortBucket,
        decisionState: gatedEmailEvaluation.relevance === "uncertain" ? "uncertain" : "accepted",
        decisionConfidence: gatedEmailEvaluation.confidence,
        decisionReason: gatedEmailEvaluation.why,
        decisionReasonTags: gatedEmailEvaluation.reasonTags,
        priorityExplanation: classification.why,
        selectionReason: `Included because the email likely needs action: ${classification.why}`,
        priorityReason: `Model-estimated priority ${gatedEmailEvaluation.priority.toLowerCase()} and effort ${classification.estimatedEffortBucket.toLowerCase()} based on the email and your preferences.`,
        personalizationVersion: personalization.memory.version ?? 1
      });
    }
    logTaskDecisionEvent({
      source: "Email",
      sourceRef: email.id,
      sourceThreadRef: email.conversationId ?? null,
      action: "system_evaluated",
      afterPriority: threadTask && !exactTask ? (threadTask.manualOverrideFlags.includes("priority") ? threadTask.priority : gatedEmailEvaluation.priority) : gatedEmailEvaluation.priority,
      systemDecisionState: gatedEmailEvaluation.relevance === "uncertain" ? "uncertain" : "accepted",
      decisionConfidence: gatedEmailEvaluation.confidence,
      decisionReason: gatedEmailEvaluation.why,
      decisionReasonTags: gatedEmailEvaluation.reasonTags,
      feedbackPayloadJson: payloadJson,
      preferencePolarity: "neutral"
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
        const joinLink =
          normalizeOptionalString(meeting.onlineMeetingUrl) ??
          normalizeOptionalString(meeting.onlineMeeting?.joinUrl) ??
          null;
        const calendarLink =
          normalizeOptionalString(meeting.webLink) ?? buildOutlookCalendarItemLink(meeting.id);
        const meetingLink = isCancelled ? null : joinLink ?? calendarLink;
        const meetingLinkType = isCancelled ? null : joinLink ? "join" : "calendar";
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
          isCancelled,
          attendanceStatus: "attending" as const
        };
      })
    );
  }
}

async function syncMicrosoftTasks(options?: {
  microsoftGraphAccessToken?: string | null;
  microsoftWarning?: string | null;
  runId?: string | null;
}) {
  const warnings: string[] = [];
  const sinceIso = new Date(Date.now() - 5 * 24 * 60 * 60 * 1000).toISOString();
  const now = new Date().toISOString();
  logEvent({
    eventType: "sync.microsoft.tasks",
    runId: options?.runId ?? null,
    provider: "microsoft",
    status: "started",
    message: "Starting Microsoft task sync.",
    metadata: { sinceIso, sessionMode: options?.microsoftGraphAccessToken ? "obo" : "stored" }
  });

  if (options?.microsoftGraphAccessToken) {
    try {
      const emails = await fetchRecentEmailsWithAccessToken(sinceIso, options.microsoftGraphAccessToken);
      await applyMicrosoftResults(emails, [], null);
      setSyncState("microsoft", now);
      logEvent({
        eventType: "sync.microsoft.tasks",
        runId: options?.runId ?? null,
        provider: "microsoft",
        status: "success",
        message: "Microsoft task sync completed.",
        metadata: { emailCount: emails.length }
      });
    } catch (error) {
      warnings.push(
        error instanceof Error ? `Microsoft task sync failed: ${error.message}` : "Microsoft task sync failed"
      );
      logEvent({
        level: "error",
        eventType: "sync.microsoft.tasks",
        runId: options?.runId ?? null,
        provider: "microsoft",
        status: "failure",
        message: "Microsoft task sync failed.",
        metadata: { error: error instanceof Error ? error.message : String(error) }
      });
    }
  } else {
    const microsoftConnection = getIntegrationConnection("microsoft");
    if (microsoftConnection?.status === "connected" && microsoftConnection.accessToken) {
      try {
        const emails = await fetchRecentEmails(sinceIso);
        await applyMicrosoftResults(emails, [], null);
        setSyncState("microsoft", now);
        logEvent({
          eventType: "sync.microsoft.tasks",
          runId: options?.runId ?? null,
          provider: "microsoft",
          status: "success",
          message: "Microsoft task sync completed with stored session.",
          metadata: { emailCount: emails.length }
        });
      } catch (error) {
        warnings.push(
          error instanceof Error ? `Microsoft task sync failed: ${error.message}` : "Microsoft task sync failed"
        );
        logEvent({
          level: "error",
          eventType: "sync.microsoft.tasks",
          runId: options?.runId ?? null,
          provider: "microsoft",
          status: "failure",
          message: "Microsoft task sync failed with stored session.",
          metadata: { error: error instanceof Error ? error.message : String(error) }
        });
      }
    } else if (options?.microsoftWarning) {
      warnings.push(options.microsoftWarning);
      logEvent({
        level: "warn",
        eventType: "sync.microsoft.tasks",
        runId: options?.runId ?? null,
        provider: "microsoft",
        status: "skipped",
        message: "Microsoft task sync skipped because no session was available.",
        metadata: { warning: options.microsoftWarning }
      });
    }
  }

  return warnings;
}

async function syncMicrosoftMeetings(options?: {
  microsoftGraphAccessToken?: string | null;
  microsoftWarning?: string | null;
  preferredTimeZone?: string | null;
  runId?: string | null;
}) {
  const warnings: string[] = [];
  const now = new Date().toISOString();
  const { startIso, endIso } = startAndEndOfAgendaWindow();
  logEvent({
    eventType: "sync.microsoft.meetings",
    runId: options?.runId ?? null,
    provider: "microsoft",
    status: "started",
    message: "Starting Microsoft meeting sync.",
    metadata: { startIso, endIso, preferredTimeZone: options?.preferredTimeZone ?? null }
  });

  if (options?.microsoftGraphAccessToken) {
    try {
      const meetingsResult = await fetchTodaysMeetingsWithAccessToken(
        startIso,
        endIso,
        options.microsoftGraphAccessToken,
        options.preferredTimeZone
      );
      await applyMicrosoftResults([], meetingsResult.events, meetingsResult.timeZone);
      saveAutomationSettings({
        workdayStartLocal: meetingsResult.workdayStartLocal,
        workdayEndLocal: meetingsResult.workdayEndLocal
      });
      setSyncState("microsoft", now);
      logEvent({
        eventType: "sync.microsoft.meetings",
        runId: options?.runId ?? null,
        provider: "microsoft",
        status: "success",
        message: "Microsoft meeting sync completed.",
        metadata: { meetingCount: meetingsResult.events.length, timeZone: meetingsResult.timeZone }
      });
    } catch (error) {
      warnings.push(
        error instanceof Error ? `Microsoft meeting sync failed: ${error.message}` : "Microsoft meeting sync failed"
      );
      logEvent({
        level: "error",
        eventType: "sync.microsoft.meetings",
        runId: options?.runId ?? null,
        provider: "microsoft",
        status: "failure",
        message: "Microsoft meeting sync failed.",
        metadata: { error: error instanceof Error ? error.message : String(error) }
      });
    }
  } else {
    const microsoftConnection = getIntegrationConnection("microsoft");
    if (microsoftConnection?.status === "connected" && microsoftConnection.accessToken) {
      try {
        const meetingsResult = await fetchTodaysMeetings(startIso, endIso, options?.preferredTimeZone);
        await applyMicrosoftResults([], meetingsResult.events, meetingsResult.timeZone);
        saveAutomationSettings({
          workdayStartLocal: meetingsResult.workdayStartLocal,
          workdayEndLocal: meetingsResult.workdayEndLocal
        });
        setSyncState("microsoft", now);
        logEvent({
          eventType: "sync.microsoft.meetings",
          runId: options?.runId ?? null,
          provider: "microsoft",
          status: "success",
          message: "Microsoft meeting sync completed with stored session.",
          metadata: { meetingCount: meetingsResult.events.length, timeZone: meetingsResult.timeZone }
        });
      } catch (error) {
        warnings.push(
          error instanceof Error ? `Microsoft meeting sync failed: ${error.message}` : "Microsoft meeting sync failed"
        );
        logEvent({
          level: "error",
          eventType: "sync.microsoft.meetings",
          runId: options?.runId ?? null,
          provider: "microsoft",
          status: "failure",
          message: "Microsoft meeting sync failed with stored session.",
          metadata: { error: error instanceof Error ? error.message : String(error) }
        });
      }
    } else if (options?.microsoftWarning) {
      warnings.push(options.microsoftWarning);
      logEvent({
        level: "warn",
        eventType: "sync.microsoft.meetings",
        runId: options?.runId ?? null,
        provider: "microsoft",
        status: "skipped",
        message: "Microsoft meeting sync skipped because no session was available.",
        metadata: { warning: options.microsoftWarning }
      });
    }
  }

  return warnings;
}

async function syncJiraTasks(runId?: string | null) {
  const warnings: string[] = [];
  const now = new Date().toISOString();
  const sinceIso = new Date(Date.now() - 48 * 60 * 60 * 1000).toISOString();
  const personalization = await getPersonalizationContext();
  logEvent({
    eventType: "sync.jira.tasks",
    runId: runId ?? null,
    provider: "jira",
    status: "started",
    message: "Starting Jira task sync.",
    metadata: { sinceIso }
  });

  const jiraConnection = getIntegrationConnection("jira");
  if (jiraConnection?.status === "connected") {
    try {
      const issues = await fetchOpenAssignedIssues(sinceIso);
      const jiraConfig = jiraConnection.configJson
        ? (JSON.parse(jiraConnection.configJson) as { baseUrl?: string })
        : null;
      for (const issue of issues) {
        let planningContext: Awaited<ReturnType<typeof fetchJiraIssuePlanningContext>> | null = null;
        try {
          planningContext = await fetchJiraIssuePlanningContext(issue.key);
        } catch {
          planningContext = null;
        }
        const candidatePayload = {
          title: `${issue.key} ${issue.fields.summary}`,
          source: "Jira" as const,
          sourceLink: jiraConfig?.baseUrl ? buildJiraIssueBrowseUrl(jiraConfig.baseUrl, issue.key) : issue.self,
          sourceRef: issue.key,
          jiraStatus: issue.fields.status?.name ?? null,
          jiraPriority: issue.fields.priority?.name ?? null,
          projectKey: issue.key.split("-")[0] ?? null,
          isAssignedToUser: true,
          isDirectRequest: /review|blocked|urgent|prod/i.test(
            `${issue.fields.summary ?? ""} ${issue.fields.status?.name ?? ""}`
          ),
          dueSoon:
            /today|urgent|asap/i.test(issue.fields.summary ?? "") || isRecentlyUpdated(issue.fields.updated),
          meetingRelevant: /meeting|review|prep/i.test(issue.fields.summary ?? ""),
          isDuplicate: false
        };
        const evaluation = await evaluateCandidateWithPersonalization({
          candidate: candidatePayload,
          profile: personalization.profile,
          memory: personalization.memory,
          recentExamples: personalization.recentExamples
        });
        const payloadJson = buildFeedbackPayload(candidatePayload);
        const existingTask = getTaskBySource("Jira", issue.key, { includeIgnored: false });
        if (evaluation.relevance === "reject" || (evaluation.relevance === "uncertain" && evaluation.confidence < 0.6)) {
          if (existingTask) {
            clearRejectedCandidate("Jira", issue.key);
            upsertTask({
              title: `${issue.key} ${issue.fields.summary}`,
              source: "Jira",
              priority: existingTask.priority,
              status: mapJiraWorkflowStatus(issue.fields.status?.name),
              sourceLink: candidatePayload.sourceLink,
              sourceRef: issue.key,
              jiraStatus: issue.fields.status?.name ?? null,
              lastActivityAt: issue.fields.updated ?? null,
              jiraEstimateSeconds:
                issue.fields.timeoriginalestimate ?? issue.fields.timetracking?.originalEstimateSeconds ?? null,
              jiraSubtaskEstimateSeconds: planningContext?.openSubtaskEstimateSeconds ?? null,
              jiraPlanningSubtasks: planningContext?.subtasks ?? [],
              decisionState: existingTask.decisionState ?? "restored",
              decisionConfidence: evaluation.confidence,
              decisionReason: existingTask.decisionReason ?? evaluation.why,
              decisionReasonTags: evaluation.reasonTags,
              personalizationVersion: personalization.memory.version ?? 1
            });
            continue;
          }
          persistRejectedCandidate({
            title: `${issue.key} ${issue.fields.summary}`,
            source: "Jira",
            sourceLink: candidatePayload.sourceLink,
            sourceRef: issue.key,
            jiraStatus: issue.fields.status?.name ?? null,
            proposedPriority: evaluation.priority,
            decisionConfidence: evaluation.confidence,
            decisionReason: evaluation.why,
            decisionReasonTags: evaluation.reasonTags,
            candidatePayloadJson: payloadJson,
            personalizationVersion: personalization.memory.version ?? 1
          });
          continue;
        }

        clearRejectedCandidate("Jira", issue.key);
        upsertTask({
          title: `${issue.key} ${issue.fields.summary}`,
          source: "Jira",
          priority: evaluation.priority,
          status: mapJiraWorkflowStatus(issue.fields.status?.name),
          sourceLink: jiraConfig?.baseUrl ? buildJiraIssueBrowseUrl(jiraConfig.baseUrl, issue.key) : issue.self,
          sourceRef: issue.key,
          jiraStatus: issue.fields.status?.name ?? null,
          lastActivityAt: issue.fields.updated ?? null,
          jiraEstimateSeconds:
            issue.fields.timeoriginalestimate ?? issue.fields.timetracking?.originalEstimateSeconds ?? null,
          jiraSubtaskEstimateSeconds: planningContext?.openSubtaskEstimateSeconds ?? null,
          jiraPlanningSubtasks: planningContext?.subtasks ?? [],
          decisionState: evaluation.relevance === "uncertain" ? "uncertain" : "accepted",
          decisionConfidence: evaluation.confidence,
          decisionReason: evaluation.why,
          decisionReasonTags: evaluation.reasonTags,
          personalizationVersion: personalization.memory.version ?? 1
        });
        logTaskDecisionEvent({
          source: "Jira",
          sourceRef: issue.key,
          action: "system_evaluated",
          afterPriority: evaluation.priority,
          systemDecisionState: evaluation.relevance === "uncertain" ? "uncertain" : "accepted",
          decisionConfidence: evaluation.confidence,
          decisionReason: evaluation.why,
          decisionReasonTags: evaluation.reasonTags,
          feedbackPayloadJson: payloadJson,
          preferencePolarity: "neutral"
        });
      }
      setSyncState("jira", now);
      logEvent({
        eventType: "sync.jira.tasks",
        runId: runId ?? null,
        provider: "jira",
        status: "success",
        message: "Jira task sync completed.",
        metadata: { issueCount: issues.length }
      });
    } catch (error) {
      warnings.push(error instanceof Error ? `Jira sync failed: ${error.message}` : "Jira sync failed");
      logEvent({
        level: "error",
        eventType: "sync.jira.tasks",
        runId: runId ?? null,
        provider: "jira",
        status: "failure",
        message: "Jira task sync failed.",
        metadata: { error: error instanceof Error ? error.message : String(error) }
      });
    }
  } else {
    logEvent({
      level: "warn",
      eventType: "sync.jira.tasks",
      runId: runId ?? null,
      provider: "jira",
      status: "skipped",
      message: "Jira task sync skipped because Jira is not connected."
    });
  }

  return warnings;
}

function daysBetween(fromIso: string, toIso: string) {
  return Math.max(0, Math.floor((new Date(toIso).getTime() - new Date(fromIso).getTime()) / 86_400_000));
}

function inferEffortBucket(task: Task): TaskEffortBucket {
  const jiraPlanningMinutes = effectiveJiraPlanningMinutes(task);
  if (jiraPlanningMinutes !== null) {
    const minutes = jiraPlanningMinutes;
    if (minutes <= 15) return "15 min";
    if (minutes <= 30) return "30 min";
    if (minutes <= 60) return "1 hour";
    return "2+ hours";
  }
  if (task.estimatedEffortBucket) {
    return task.estimatedEffortBucket;
  }
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

function minutesForTask(task: Task) {
  const jiraPlanningMinutes = effectiveJiraPlanningMinutes(task);
  if (jiraPlanningMinutes !== null) {
    return jiraPlanningMinutes;
  }
  return minutesForEffort(task.estimatedEffortBucket);
}

function effectiveJiraPlanningMinutes(task: Task) {
  if (task.source !== "Jira") return null;
  const parentMinutes =
    task.jiraEstimateSeconds && task.jiraEstimateSeconds > 0
      ? Math.max(15, Math.ceil(task.jiraEstimateSeconds / 60))
      : 0;
  const subtaskMinutes =
    task.jiraSubtaskEstimateSeconds && task.jiraSubtaskEstimateSeconds > 0
      ? Math.max(15, Math.ceil(task.jiraSubtaskEstimateSeconds / 60))
      : 0;
  const effective = subtaskMinutes > 0 ? subtaskMinutes : parentMinutes;
  return effective > 0 ? effective : null;
}

function preferredJiraSubtask(task: Task) {
  return (
    task.jiraPlanningSubtasks.find((subtask) =>
      /(progress|coding|review|testing|qa|blocked|in dev|development)/i.test(subtask.status ?? "")
    ) ??
    task.jiraPlanningSubtasks[0] ??
    null
  );
}

function hasActiveJiraSubtask(task: Task) {
  return task.jiraPlanningSubtasks.some((subtask) =>
    /(progress|coding|review|testing|qa|blocked|in dev|development)/i.test(subtask.status ?? "")
  );
}

function deriveJiraPlanningPriority(task: Task): TaskPriority {
  if (task.source !== "Jira") {
    return task.priority;
  }
  if (task.status === "In Progress") {
    return "High";
  }
  if (hasActiveJiraSubtask(task)) {
    return "Medium";
  }
  return "Low";
}

const DEFAULT_WORKDAY_MINUTES = 510;

function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value));
}

function average(values: number[]) {
  if (!values.length) return null;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function localDayKey(value = new Date()) {
  const year = value.getFullYear();
  const month = `${value.getMonth() + 1}`.padStart(2, "0");
  const day = `${value.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function startOfLocalDay(value = new Date()) {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0, 0);
}

function endOfLocalDay(value = new Date()) {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 23, 59, 59, 999);
}

function isSameLocalDay(isoValue: string | null, reference = new Date()) {
  if (!isoValue) return false;
  const target = new Date(isoValue);
  return (
    target.getFullYear() === reference.getFullYear() &&
    target.getMonth() === reference.getMonth() &&
    target.getDate() === reference.getDate()
  );
}

function addMinutes(value: Date, minutes: number) {
  return new Date(value.getTime() + minutes * 60_000);
}

function parseLocalTime(value: string | null | undefined, fallbackHours: number, fallbackMinutes: number) {
  const match = value?.match(/^(\d{2}):(\d{2})$/);
  if (!match) {
    return { hours: fallbackHours, minutes: fallbackMinutes };
  }
  return {
    hours: Number(match[1]),
    minutes: Number(match[2])
  };
}

function diffMinutes(start: Date, end: Date) {
  return Math.max(0, Math.round((end.getTime() - start.getTime()) / 60_000));
}

function roundUpToNextFiveMinutes(value: Date) {
  const rounded = new Date(value);
  rounded.setSeconds(0, 0);
  const minutes = rounded.getMinutes();
  const remainder = minutes % 5;
  if (remainder !== 0) {
    rounded.setMinutes(minutes + (5 - remainder));
  }
  return rounded;
}

function meetingStart(meeting: ReturnType<typeof listMeetings>[number]) {
  return parseMeetingDateWithTimeZone(meeting.startTime, meeting.timeZone);
}

function meetingEnd(meeting: ReturnType<typeof listMeetings>[number]) {
  return parseMeetingDateWithTimeZone(meeting.endTime, meeting.timeZone);
}

function estimateRemainingTaskMinutes(task: Task) {
  const total = minutesForTask(task);
  if (task.status === "Completed") return 0;
  if (task.source === "Jira" && (task.jiraSubtaskEstimateSeconds ?? 0) > 0) {
    if (task.status === "In Progress") {
      return Math.max(30, Math.round(total * 0.8));
    }
    return total;
  }
  if (task.status === "In Progress") {
    return Math.max(15, Math.round(total * 0.6));
  }
  return total;
}

function mergeAdjacentTaskBlocks(blocks: DayPlanBlock[]) {
  const sorted = [...blocks].sort((left, right) => new Date(left.startTime).getTime() - new Date(right.startTime).getTime());
  const merged: DayPlanBlock[] = [];

  for (const block of sorted) {
    const previous = merged[merged.length - 1];
    if (
      previous &&
      previous.kind === "task" &&
      block.kind === "task" &&
      previous.taskId !== null &&
      previous.taskId === block.taskId &&
      previous.source === block.source &&
      previous.priority === block.priority &&
      previous.link === block.link
    ) {
      const gapMinutes = diffMinutes(new Date(previous.endTime), new Date(block.startTime));
      if (gapMinutes > 15) {
        merged.push({ ...block });
        continue;
      }
      previous.endTime = block.endTime;
      previous.durationMinutes += block.durationMinutes + gapMinutes;
      previous.note = block.note ?? previous.note;
      previous.status =
        previous.status === "in_progress" || block.status === "in_progress"
          ? "in_progress"
          : previous.status === "up_next" || block.status === "up_next"
            ? "up_next"
            : previous.status;
      continue;
    }
    merged.push({ ...block });
  }

  return merged;
}

function isPrimaryPlanningTask(task: Task) {
  const planningPriority = task.source === "Jira" ? deriveJiraPlanningPriority(task) : task.priority;
  return (
    task.status === "In Progress" ||
    task.carryForwardCount > 0 ||
    task.source === "Jira" ||
    planningPriority === "High"
  );
}

function isLightEmailPlanningTask(task: Task) {
  const emailSignals = `${task.title} ${task.selectionReason ?? ""} ${task.priorityReason ?? ""}`.toLowerCase();
  return (
    task.source === "Email" &&
    !isPrimaryPlanningTask(task) &&
    estimateRemainingTaskMinutes(task) <= 60 &&
    !/announcement|release|showcase|newsletter|optional|hackathon|webinar|training|scam|highlights|monthly deep dive/.test(emailSignals) &&
    (/reply|follow up|approval|review|action|respond|check in|comment|mentioned/.test(emailSignals) || task.priority !== "Low")
  );
}

function buildTaskPlanningTitle(task: Task, scheduledMinutes: number) {
  if (task.source !== "Jira") return task.title;

  const subtask = preferredJiraSubtask(task);
  const subtaskEstimateMinutes = subtask?.estimateSeconds ? Math.max(15, Math.ceil(subtask.estimateSeconds / 60)) : null;
  const hasLargeSubtaskLoad =
    Boolean(task.jiraSubtaskEstimateSeconds && task.jiraSubtaskEstimateSeconds >= 90 * 60) ||
    task.jiraPlanningSubtasks.length >= 2;
  const hasActiveSubtask = Boolean(subtask && /(progress|coding|review|testing|qa|blocked|in dev|development)/i.test(subtask.status ?? ""));

  if (subtask && subtaskEstimateMinutes && scheduledMinutes >= subtaskEstimateMinutes) {
    return `Complete subtask ${subtask.key} in ${task.sourceRef ?? "this story"}`;
  }

  if (subtask && (hasActiveSubtask || hasLargeSubtaskLoad || task.status === "In Progress")) {
    return `Progress ${subtask.key} in ${task.sourceRef ?? "this story"}`;
  }

  if (subtask && scheduledMinutes <= Math.min(90, Math.max(30, subtaskEstimateMinutes ?? 45))) {
    return `Progress ${subtask.key} in ${task.sourceRef ?? "this story"}`;
  }

  return task.title;
}

function buildTaskPlanningNote(task: Task) {
  if (task.source === "Jira" && task.jiraPlanningSubtasks.length) {
    const subtask = preferredJiraSubtask(task);
    const openCount = task.jiraPlanningSubtasks.length;
    const totalMinutes = task.jiraSubtaskEstimateSeconds ? Math.ceil(task.jiraSubtaskEstimateSeconds / 60) : null;
    if (subtask && totalMinutes && totalMinutes >= 90) {
      return `${openCount} open subtasks remain (${Math.ceil(totalMinutes / 60)}h total). Start with ${subtask.key}${subtask.status ? ` • ${subtask.status}` : ""}.`;
    }
    if (subtask) {
      return `Next useful slice: ${subtask.key}${subtask.status ? ` • ${subtask.status}` : ""}.`;
    }
  }
  return task.priorityExplanation ?? null;
}

function shiftDayKey(base: Date, dayDelta: number) {
  const shifted = new Date(base);
  shifted.setDate(shifted.getDate() + dayDelta);
  return localDayKey(shifted);
}

function isStarterPlanningEmailTask(task: Task) {
  const text = `${task.title} ${task.selectionReason ?? ""} ${task.priorityReason ?? ""}`.toLowerCase();
  return (
    task.source === "Email" &&
    estimateRemainingTaskMinutes(task) <= 30 &&
    !/announcement|release|showcase|newsletter|optional|hackathon|webinar|training|scam|highlights|monthly deep dive/.test(text) &&
    (/reply|follow up|approval|review|action|respond|check in|comment|mentioned|password|pull request/.test(text) ||
      task.priority !== "Low")
  );
}

function isClosingPlanningTask(task: Task) {
  const remaining = estimateRemainingTaskMinutes(task);
  if (task.source === "Email") {
    return isLightEmailPlanningTask(task) && remaining <= 30;
  }
  return remaining <= 45 && (task.priority === "Low" || task.stage === "Later" || task.source === "Manual");
}

function planningRoleForTask(task: Task): PlanningRole {
  const explicit = task.historySignals.find((entry) => /^Planning role:/i.test(entry));
  if (explicit) {
    const normalized = explicit.split(":")[1]?.trim().toLowerCase();
    if (normalized === "starter" || normalized === "major" || normalized === "ender" || normalized === "review") {
      return normalized;
    }
  }
  if (task.stage === "Review" || task.decisionState === "uncertain") return "review";
  if (isStarterPlanningEmailTask(task)) return "starter";
  if (isClosingPlanningTask(task)) return "ender";
  return "major";
}

async function applyPlanningCategories() {
  const activeTasks = listTasks(undefined, { includeDeferred: true }).filter((task) => task.status !== "Completed");
  if (!activeTasks.length) return;

  const fallbackRoles = new Map<number, PlanningRole>();
  for (const task of activeTasks) {
    fallbackRoles.set(task.id, planningRoleForTask(task));
  }

  const aiResult = await callOpenAIJson<
    { assignments: Array<{ taskId: number; role: PlanningRole; reason?: string }> }
  >(
    {
      tasks: activeTasks.slice(0, 30).map((task) => ({
        taskId: task.id,
        title: task.title,
        source: task.source,
        priority: task.priority,
        status: task.status,
        stage: task.stage,
        selectionReason: task.selectionReason,
        priorityReason: task.priorityReason,
        priorityExplanation: task.priorityExplanation,
        remainingMinutes: estimateRemainingTaskMinutes(task),
        jiraStatus: task.jiraStatus,
        carryForwardCount: task.carryForwardCount
      }))
    },
    {
      type: "object",
      additionalProperties: false,
      required: ["assignments"],
      properties: {
        assignments: {
          type: "array",
          items: {
            type: "object",
            additionalProperties: false,
            required: ["taskId", "role"],
            properties: {
              taskId: { type: "number" },
              role: { type: "string", enum: ["starter", "major", "ender", "review"] },
              reason: { type: "string" }
            }
          }
        }
      }
    },
    "day_plan_task_roles",
    [
      "Categorize work tasks for a day planner into starter, major, ender, or review.",
      "starter: simple easy opener tasks, usually 15-30 minutes, often email/reply/review/admin.",
      "major: the main focused work blocks, usually Jira, in-progress work, spillover, substantial work.",
      "ender: simple closing tasks near the end of the day, like reviews, follow-ups, and light email/admin.",
      "review: uncertain items that should not be confidently scheduled into the active day.",
      "Prefer only 1-2 starter tasks and 1-2 ender tasks across the whole set.",
      "Do not mark many tasks as starter or ender unless they are genuinely simple.",
      "Use the actual title, source, status, and priority signals. Return JSON only."
    ].join(" ")
  );

  const roleMap = new Map<number, PlanningRole>(fallbackRoles);
  for (const assignment of aiResult?.assignments ?? []) {
    if (!roleMap.has(assignment.taskId)) continue;
    roleMap.set(assignment.taskId, assignment.role);
  }

  for (const task of activeTasks) {
    const role = roleMap.get(task.id) ?? fallbackRoles.get(task.id) ?? "major";
    const nextHistorySignals = [
      ...task.historySignals.filter((entry) => !/^Planning role:/i.test(entry)),
      `Planning role: ${role}`
    ];
    updateTask(task.id, {
      historySignals: nextHistorySignals,
      lastChangedBy: "system",
      lastChangedAt: new Date().toISOString()
    });
  }
}

function planningBehaviorAdjustment(task: Task) {
  const recentEvents = listTaskStateEvents({
    startDayKey: shiftDayKey(new Date(), -21),
    endDayKey: localDayKey(new Date()),
    limit: 500
  }).filter((event) => event.taskId === task.id);

  const removedFromTimelineCount = recentEvents.filter((event) => event.eventType === "timeline_slot_removed").length;
  const addedToTimelineCount = recentEvents.filter((event) => event.eventType === "timeline_slot_added").length;
  const updatedInTimelineCount = recentEvents.filter((event) => event.eventType === "timeline_slot_updated").length;
  const completedCount = recentEvents.filter((event) => {
    if (event.eventType !== "task_updated" || !event.afterJson) return false;
    try {
      const parsed = JSON.parse(event.afterJson) as { status?: string };
      return parsed.status === "Completed";
    } catch {
      return false;
    }
  }).length;

  return {
    scoreDelta:
      Math.min(10, addedToTimelineCount * 2) +
      Math.min(8, updatedInTimelineCount) +
      Math.min(12, completedCount * 3) -
      Math.min(18, removedFromTimelineCount * 6),
    removedFromTimelineCount
  };
}

function buildTaskStoryContext(task: Task) {
  if (task.source !== "Jira") return null;
  const storyKey = task.sourceRef ?? null;
  const nextSubtask = preferredJiraSubtask(task);
  const openCount = task.jiraPlanningSubtasks.length;
  const openHours =
    task.jiraSubtaskEstimateSeconds && task.jiraSubtaskEstimateSeconds > 0
      ? `${(task.jiraSubtaskEstimateSeconds / 3600).toFixed(task.jiraSubtaskEstimateSeconds % 3600 === 0 ? 0 : 1)}h`
      : null;

  const parts = [`Story ${storyKey ?? "Jira item"}`];
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

function taskExecutionBucket(task: Task) {
  if (task.source === "Jira" && task.status === "In Progress") {
    return 0;
  }
  if (task.source === "Jira" && hasActiveJiraSubtask(task)) {
    return 1;
  }
  if (task.source === "Jira") {
    return 2;
  }
  if (task.status === "In Progress" && task.priority === "High") {
    return 3;
  }
  if (task.carryForwardCount > 0) {
    return 4;
  }
  return 5;
}

function deriveAdaptiveFocusFactor(todayKey: string, weekday: number) {
  const rows = listRecentDailyPlanSnapshots(28).filter((row) => String(row.day_key) !== todayKey);
  const ratios = rows
    .map((row) => {
      const planned = Number(row.planned_task_minutes ?? 0);
      const completed = Number(row.completed_task_minutes ?? 0);
      if (planned < 30) return null;
      return clamp(completed / planned, 0.45, 1.15);
    })
    .filter((value): value is number => value !== null);

  const weekdayRatios = rows
    .filter((row) => Number(row.weekday ?? -1) === weekday)
    .map((row) => {
      const planned = Number(row.planned_task_minutes ?? 0);
      const completed = Number(row.completed_task_minutes ?? 0);
      if (planned < 30) return null;
      return clamp(completed / planned, 0.45, 1.15);
    })
    .filter((value): value is number => value !== null);

  const learned = average(weekdayRatios.length >= 2 ? weekdayRatios : ratios) ?? 0.92;
  return clamp(learned, 0.72, 1.08);
}

function plannedTaskIdsFromSnapshot(row: Record<string, unknown>) {
  try {
    const parsed = JSON.parse(String(row.planned_task_ids_json ?? "[]")) as unknown;
    return Array.isArray(parsed) ? parsed.filter((value): value is number => typeof value === "number") : [];
  } catch {
    return [];
  }
}

function deriveCarryForwardCount(task: Task, todayKey: string) {
  if (task.status === "Completed") {
    return 0;
  }

  const snapshotCarryForward = listRecentDailyPlanSnapshots(14)
    .filter((row) => String(row.day_key) !== todayKey)
    .filter((row) => plannedTaskIdsFromSnapshot(row).includes(task.id)).length;

  if (snapshotCarryForward > 0) {
    return snapshotCarryForward;
  }

  const ageDays = daysBetween(task.createdAt, new Date().toISOString());
  return Math.max(0, ageDays - 1);
}

function buildDayPlan(tasks: Task[], meetings: ReturnType<typeof listMeetings>): DayPlan {
  const now = new Date();
  const dayKey = localDayKey(now);
  const weekday = now.getDay();
  const automation = getAutomationSettings();
  const workdayStartParts = parseLocalTime(automation.workdayStartLocal, 9, 30);
  const workdayEndParts = parseLocalTime(automation.workdayEndLocal, 18, 0);
  const todayStart = startOfLocalDay(now);
  const todayEnd = endOfLocalDay(now);
  const workdayStart = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
    workdayStartParts.hours,
    workdayStartParts.minutes,
    0,
    0
  );
  const workdayEnd = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
    workdayEndParts.hours,
    workdayEndParts.minutes,
    0,
    0
  );
  const effectiveBaseWorkdayMinutes = Math.max(60, diffMinutes(workdayStart, workdayEnd));

  const todaysMeetings = meetings
    .filter((meeting) => isPlannableMeeting(meeting) && meetingStart(meeting) < todayEnd && meetingEnd(meeting) > todayStart)
    .sort((left, right) => meetingStart(left).getTime() - meetingStart(right).getTime());
  const remainingMeetings = todaysMeetings.filter((meeting) => meetingEnd(meeting).getTime() > now.getTime());
  const meetingMinutes = todaysMeetings.reduce((sum, meeting) => sum + meeting.durationMinutes, 0);

  const baseTaskCapacityMinutes = Math.max(0, effectiveBaseWorkdayMinutes - meetingMinutes);
  const focusFactor = deriveAdaptiveFocusFactor(dayKey, weekday);
  const adaptedTaskCapacityMinutes = Math.max(0, Math.round(baseTaskCapacityMinutes * focusFactor));
  const completedTodayTasks = tasks.filter((task) => task.status === "Completed" && isSameLocalDay(task.completedAt, now));
  const completedTaskMinutes = completedTodayTasks.reduce((sum, task) => sum + minutesForTask(task), 0);
  const remainingTaskCapacityMinutes = Math.max(0, adaptedTaskCapacityMinutes - completedTaskMinutes);

  const activeTasks = tasks.filter((task) => task.status !== "Completed");
  const rankedTasks = activeTasks
    .map((task) => {
      const remainingMinutes = estimateRemainingTaskMinutes(task);
      const executionBucket = taskExecutionBucket(task);
      const behaviorAdjustment = planningBehaviorAdjustment(task);
      const rank =
        (task.priorityScore ?? 0) +
        (task.status === "In Progress" ? 22 : 0) +
        ((task.source === "Jira" ? deriveJiraPlanningPriority(task) : task.priority) === "High"
          ? 18
          : (task.source === "Jira" ? deriveJiraPlanningPriority(task) : task.priority) === "Medium"
            ? 9
            : 3) +
        Math.min(20, task.carryForwardCount * 5) +
        (task.source === "Email" && /follow up|reply|approval|action/i.test(task.title) ? 6 : 0) +
        (task.source === "Jira" && isRecentlyUpdated(task.lastActivityAt) ? 8 : 0) +
        behaviorAdjustment.scoreDelta;
      return {
        task,
        remainingMinutes,
        remainingEstimate: remainingMinutes,
        executionBucket,
        rank,
        behaviorAdjustment
      };
    })
    .filter((entry) => entry.remainingMinutes > 0)
    .sort((left, right) => {
      if (left.executionBucket !== right.executionBucket) {
        return left.executionBucket - right.executionBucket;
      }
      if (right.rank !== left.rank) return right.rank - left.rank;
      if (left.remainingMinutes !== right.remainingMinutes) return left.remainingMinutes - right.remainingMinutes;
      return left.task.updatedAt < right.task.updatedAt ? 1 : -1;
    });

  const planningStart = roundUpToNextFiveMinutes(new Date(Math.max(now.getTime(), workdayStart.getTime())));
  const schedulingEnd = new Date(
    Math.max(
      workdayEnd.getTime(),
      ...remainingMeetings.map((meeting) => meetingEnd(meeting).getTime()),
      planningStart.getTime()
    )
  );
  const taskSchedulingEnd = schedulingEnd;

  const blocks: DayPlanBlock[] = [];
  const taskBlocks: DayPlanBlock[] = [];
  const plannedTaskIds = new Set<number>();
  const currentMeeting = remainingMeetings.find(
    (meeting) => meetingStart(meeting).getTime() <= now.getTime() && meetingEnd(meeting).getTime() > now.getTime()
  );
  const nextMeetingId = currentMeeting
    ? currentMeeting.id
    : remainingMeetings.find((meeting) => meetingStart(meeting).getTime() > now.getTime())?.id ?? null;

  const scheduleTaskBlock = (
    slotStart: Date,
    slotMinutes: number,
    focusStreak: number
  ): { block: DayPlanBlock | null; consumedMinutes: number; nextFocusStreak: number } => {
    if (slotMinutes < 15) {
      return { block: null, consumedMinutes: 0, nextFocusStreak: focusStreak };
    }

    const preferQuickWin = focusStreak >= 2 || slotMinutes <= 45;
    const fallbackCandidates = rankedTasks.filter((entry) => entry.remainingEstimate > 0);
    const quickCandidates = fallbackCandidates.filter((entry) => entry.remainingEstimate <= 45);
    const primaryCandidates = fallbackCandidates.filter((entry) => planningRoleForTask(entry.task) === "major");
    const starterCandidates = fallbackCandidates.filter((entry) => planningRoleForTask(entry.task) === "starter");
    const closingCandidates = fallbackCandidates.filter((entry) => planningRoleForTask(entry.task) === "ender");
    const nowTaskMinutes = taskBlocks.reduce((sum, block) => sum + block.durationMinutes, 0);
    const totalSchedulingWindowMinutes = Math.max(15, diffMinutes(planningStart, taskSchedulingEnd));
    const progressRatio = clamp(nowTaskMinutes / totalSchedulingWindowMinutes, 0, 1);
    const remainingDayMinutes = Math.max(0, diffMinutes(slotStart, taskSchedulingEnd));
    const lastTaskId = taskBlocks[taskBlocks.length - 1]?.taskId ?? null;
    const lastTaskSource = taskBlocks.length
      ? tasks.find((candidate) => candidate.id === taskBlocks[taskBlocks.length - 1]!.taskId)?.source ?? null
      : null;
    const taskById = (taskId: number | null) =>
      taskId === null ? null : tasks.find((candidate) => candidate.id === taskId) ?? null;
    const plannedMinutesForTask = (taskId: number) =>
      taskBlocks.reduce((sum, block) => sum + (block.taskId === taskId ? block.durationMinutes : 0), 0);
    const plannedBlockCountForTask = (taskId: number) =>
      taskBlocks.reduce((sum, block) => sum + (block.taskId === taskId ? 1 : 0), 0);
    const starterTaskIds = new Set(
      taskBlocks
        .map((block) => block.taskId)
        .filter((taskId): taskId is number => taskId !== null)
        .filter((taskId) => {
          const task = taskById(taskId);
          return task ? planningRoleForTask(task) === "starter" : false;
        })
    );
    const closingTaskIds = new Set(
      taskBlocks
        .map((block) => block.taskId)
        .filter((taskId): taskId is number => taskId !== null)
        .filter((taskId) => {
          const task = taskById(taskId);
          return task ? planningRoleForTask(task) === "ender" : false;
        })
    );
    const alreadyPlannedTaskMinutes = taskBlocks.reduce((sum, block) => sum + block.durationMinutes, 0);
    const plannedPrimaryMinutes = taskBlocks.reduce((sum, block) => {
      const task = block.taskId !== null ? tasks.find((candidate) => candidate.id === block.taskId) : null;
      return sum + (task && isPrimaryPlanningTask(task) ? block.durationMinutes : 0);
    }, 0);
    const shouldOpenWithStarter = starterCandidates.length > 0 && starterTaskIds.size < 2 && alreadyPlannedTaskMinutes < 45;
    const shouldCloseWithLight =
      closingCandidates.length > 0 &&
      closingTaskIds.size < 2 &&
      (remainingDayMinutes <= 90 || progressRatio >= 0.78);
    const continuePreviousMajor =
      lastTaskId !== null
        ? fallbackCandidates.find(
            (entry) =>
              entry.task.id === lastTaskId &&
              planningRoleForTask(entry.task) === "major" &&
              entry.task.source !== "Email" &&
              entry.remainingEstimate >= 30
          ) ?? null
        : null;

    const chooseCandidate = (
      pool: typeof fallbackCandidates,
      options?: { avoidTaskId?: number | null; preferFresh?: boolean }
    ) => {
      if (!pool.length) return null;
      let filtered = [...pool];
      if (options?.avoidTaskId !== null && options?.avoidTaskId !== undefined && filtered.some((entry) => entry.task.id !== options.avoidTaskId)) {
        filtered = filtered.filter((entry) => entry.task.id !== options.avoidTaskId);
      }
      filtered = filtered.filter((entry) => !(entry.task.source === "Email" && plannedBlockCountForTask(entry.task.id) > 0));
      const notOverFragmented = filtered.filter((entry) => {
        const plannedBlocks = plannedBlockCountForTask(entry.task.id);
        if (plannedBlocks === 0) return true;
        if (entry.task.source === "Email") return false;
        if (plannedBlocks >= 2 && entry.task.source !== "Jira") return false;
        return true;
      });
      if (notOverFragmented.length) {
        filtered = notOverFragmented;
      }
      if (options?.preferFresh) {
        const fresh = filtered.filter((entry) => plannedMinutesForTask(entry.task.id) === 0);
        if (fresh.length) filtered = fresh;
      }
      filtered.sort((left, right) => {
        const leftPlanned = plannedMinutesForTask(left.task.id);
        const rightPlanned = plannedMinutesForTask(right.task.id);
        if (leftPlanned !== rightPlanned) return leftPlanned - rightPlanned;
        if (left.behaviorAdjustment.removedFromTimelineCount !== right.behaviorAdjustment.removedFromTimelineCount) {
          return left.behaviorAdjustment.removedFromTimelineCount - right.behaviorAdjustment.removedFromTimelineCount;
        }
        if (right.rank !== left.rank) return right.rank - left.rank;
        return left.remainingEstimate - right.remainingEstimate;
      });
      return filtered[0] ?? null;
    };

    const chosen =
      (shouldOpenWithStarter ? chooseCandidate(starterCandidates, { avoidTaskId: lastTaskId, preferFresh: true }) : null) ??
      (shouldCloseWithLight ? chooseCandidate(closingCandidates, { avoidTaskId: lastTaskId, preferFresh: true }) : null) ??
      continuePreviousMajor ??
      chooseCandidate(primaryCandidates, { avoidTaskId: lastTaskId, preferFresh: true }) ??
      (preferQuickWin ? chooseCandidate(quickCandidates, { avoidTaskId: lastTaskId, preferFresh: true }) : null) ??
      chooseCandidate(fallbackCandidates, { avoidTaskId: lastTaskId, preferFresh: true }) ??
      chooseCandidate(fallbackCandidates, { avoidTaskId: lastTaskId });

    if (!chosen) {
      return { block: null, consumedMinutes: 0, nextFocusStreak: focusStreak };
    }

    const capacityLeft = Math.max(
      0,
      remainingTaskCapacityMinutes - taskBlocks.reduce((sum, block) => sum + block.durationMinutes, 0)
    );
    const canForceVisibilityBlock =
      taskBlocks.length === 0 &&
      slotMinutes >= 15 &&
      (chosen.task.status === "In Progress" || chosen.executionBucket <= 1) &&
      slotStart.getTime() < workdayEnd.getTime();
    const effectiveCapacityLeft = capacityLeft >= 15 ? capacityLeft : canForceVisibilityBlock ? 15 : 0;

    if (effectiveCapacityLeft < 15) {
      const stretchCandidate =
        fallbackCandidates.find((entry) => entry.executionBucket <= 2 && entry.remainingEstimate > 0) ??
        fallbackCandidates.find((entry) => entry.remainingEstimate > 0);
      if (!stretchCandidate) {
        return { block: null, consumedMinutes: 0, nextFocusStreak: focusStreak };
      }

      const stretchDuration = Math.max(
        15,
        Math.min(
          slotMinutes,
          stretchCandidate.remainingEstimate,
          stretchCandidate.executionBucket <= 1 ? 45 : 30
        )
      );
      const stretchEnd = addMinutes(slotStart, stretchDuration);
      const stretchBlock: DayPlanBlock = {
        id: `task-stretch-${stretchCandidate.task.id}-${slotStart.getTime()}`,
        kind: "task",
        title: buildTaskPlanningTitle(stretchCandidate.task, stretchDuration),
        startTime: slotStart.toISOString(),
        endTime: stretchEnd.toISOString(),
        timeZone: null,
        durationMinutes: stretchDuration,
        status: "planned",
        taskId: stretchCandidate.task.id,
        meetingId: null,
        source: stretchCandidate.task.source,
        priority:
          stretchCandidate.task.source === "Jira"
            ? deriveJiraPlanningPriority(stretchCandidate.task)
            : stretchCandidate.task.priority,
        link: stretchCandidate.task.sourceLink,
        note: `${buildTaskPlanningNote(stretchCandidate.task) ?? "Optional stretch work."} Use this block if you finish earlier than expected or want to pull work forward.`
      };
      stretchCandidate.remainingEstimate -= stretchDuration;
      plannedTaskIds.add(stretchCandidate.task.id);
      return {
        block: stretchBlock,
        consumedMinutes: stretchDuration,
        nextFocusStreak: 0
      };
    }

    const planningPriority = chosen.task.source === "Jira" ? deriveJiraPlanningPriority(chosen.task) : chosen.task.priority;
    const plannedAlreadyForChosen = plannedMinutesForTask(chosen.task.id);
    const idealChunk =
      planningRoleForTask(chosen.task) === "starter" || planningRoleForTask(chosen.task) === "ender"
        ? Math.min(30, chosen.remainingEstimate)
        : chosen.task.source === "Email"
          ? Math.min(30, chosen.remainingEstimate)
          : plannedAlreadyForChosen > 0
            ? Math.min(90, chosen.remainingEstimate)
            : chosen.executionBucket <= 1 || planningPriority === "High" || chosen.task.status === "In Progress"
              ? Math.min(120, chosen.remainingEstimate)
              : Math.min(90, chosen.remainingEstimate);
    const minimumChunk =
      chosen.task.source === "Email"
        ? 15
        : lastTaskId === chosen.task.id
          ? 45
          : lastTaskSource === chosen.task.source
          ? 30
          : 60;
    const maxChunk = Math.min(slotMinutes, effectiveCapacityLeft, idealChunk);
    const durationMinutes = Math.max(Math.min(minimumChunk, maxChunk), maxChunk >= 15 ? 15 : maxChunk);
    const endTime = addMinutes(slotStart, durationMinutes);
    const status: DayPlanBlock["status"] =
      chosen.task.status === "In Progress" && slotStart.getTime() <= now.getTime() ? "in_progress" : "planned";
    const block: DayPlanBlock = {
      id: `task-${chosen.task.id}-${slotStart.getTime()}`,
      kind: "task",
      title: buildTaskPlanningTitle(chosen.task, durationMinutes),
      startTime: slotStart.toISOString(),
      endTime: endTime.toISOString(),
      timeZone: null,
      durationMinutes,
      status,
      taskId: chosen.task.id,
      meetingId: null,
      source: chosen.task.source,
      priority: chosen.task.source === "Jira" ? deriveJiraPlanningPriority(chosen.task) : chosen.task.priority,
      link: chosen.task.sourceLink,
      note: buildTaskPlanningNote(chosen.task)
    };
    chosen.remainingEstimate -= durationMinutes;
    plannedTaskIds.add(chosen.task.id);
    return {
      block,
      consumedMinutes: durationMinutes,
      nextFocusStreak: durationMinutes >= 45 ? focusStreak + 1 : 0
    };
  };

  let cursor = planningStart;
  let focusStreak = 0;

  for (const meeting of remainingMeetings) {
    const start = meetingStart(meeting);
    const end = meetingEnd(meeting);
    const effectiveMeetingStart = new Date(Math.max(start.getTime(), planningStart.getTime()));

    while (cursor.getTime() < effectiveMeetingStart.getTime()) {
      const slotMinutes = diffMinutes(cursor, effectiveMeetingStart);
      const scheduled = scheduleTaskBlock(cursor, slotMinutes, focusStreak);
      if (!scheduled.block) break;
      taskBlocks.push(scheduled.block);
      blocks.push(scheduled.block);
      cursor = addMinutes(cursor, scheduled.consumedMinutes);
      focusStreak = scheduled.nextFocusStreak;
    }

    if (cursor.getTime() < effectiveMeetingStart.getTime()) {
      const bufferEnd = effectiveMeetingStart;
      const bufferMinutes = diffMinutes(cursor, bufferEnd);
      if (bufferMinutes >= 15) {
        blocks.push({
          id: `buffer-${cursor.getTime()}`,
          kind: "buffer",
          title: "Flex buffer",
          startTime: cursor.toISOString(),
          endTime: bufferEnd.toISOString(),
          timeZone: null,
          durationMinutes: bufferMinutes,
          status: "planned",
          taskId: null,
          meetingId: null,
          source: null,
          priority: null,
          link: null,
          note: "Use this gap for prep, notes, or a quick reset."
        });
      }
    }

    blocks.push({
      id: `meeting-${meeting.id}`,
      kind: "meeting",
      title: meeting.title,
      startTime: meeting.startTime,
      endTime: meeting.endTime,
      timeZone: meeting.timeZone,
      durationMinutes: meeting.durationMinutes,
      status:
        start.getTime() <= now.getTime() && end.getTime() > now.getTime()
          ? "in_progress"
          : meeting.id === nextMeetingId
            ? "up_next"
            : "planned",
      taskId: null,
      meetingId: meeting.id,
      source: "Calendar",
      priority: null,
      link: meeting.meetingLink,
      note: meeting.meetingLinkType === "join" ? "Join meeting when it starts." : "Open in calendar for details."
    });
    cursor = new Date(Math.max(cursor.getTime(), end.getTime()));
    focusStreak = 0;
  }

  while (cursor.getTime() < taskSchedulingEnd.getTime()) {
    const slotMinutes = diffMinutes(cursor, taskSchedulingEnd);
    const scheduled = scheduleTaskBlock(cursor, slotMinutes, focusStreak);
    if (!scheduled.block) break;
    taskBlocks.push(scheduled.block);
    blocks.push(scheduled.block);
    cursor = addMinutes(cursor, scheduled.consumedMinutes);
    focusStreak = scheduled.nextFocusStreak;
  }

  const displayBlocks = mergeAdjacentTaskBlocks(blocks);
  const plannedTaskMinutes = taskBlocks.reduce((sum, block) => sum + block.durationMinutes, 0);
  const spilloverTasks = rankedTasks
    .filter((entry) => entry.remainingEstimate > 0)
    .map((entry) => entry.task);
  const remainingTaskMinutes = rankedTasks.reduce((sum, entry) => sum + Math.max(0, entry.remainingEstimate), 0);
  const freeMinutes = Math.max(0, remainingTaskCapacityMinutes - plannedTaskMinutes);
  const completionRate = clamp(
    completedTaskMinutes / Math.max(30, completedTaskMinutes + plannedTaskMinutes),
    0,
    1
  );

  const guidance = (() => {
    if (currentMeeting) {
      return "You are in a meeting now. The next focused task starts right after it ends.";
    }
    if (spilloverTasks.length > 0) {
      return "Prioritize the scheduled blocks first and pull spillover only if meetings finish on time.";
    }
    if (freeMinutes >= 45) {
      return "You have meaningful slack later today for review, follow-ups, or one extra low-priority task.";
    }
    if (meetingMinutes >= 240) {
      return "Meeting-heavy day. Keep transitions tight and protect your first focus block.";
    }
    if (focusFactor < 0.9) {
      return "This weekday usually runs tight, so the plan is intentionally conservative on task load.";
    }
    return "The day balances focused work with lighter follow-through around your meetings.";
  })();

  return {
    summary: {
      dayKey,
      baseWorkdayMinutes: effectiveBaseWorkdayMinutes,
      adaptedTaskCapacityMinutes,
      remainingTaskCapacityMinutes,
      meetingMinutes,
      completedTaskMinutes,
      plannedTaskMinutes,
      remainingTaskMinutes,
      spilloverTaskCount: spilloverTasks.length,
      freeMinutes,
      focusFactor,
      completionRate,
      guidance
    },
    blocks: displayBlocks,
    spilloverTasks
  };
}

function computeTaskSignals(task: Task, meetings: ReturnType<typeof listMeetings>) {
  let score = 0;
  const reasons: string[] = [];
  const scoreBreakdown: Array<{ label: string; value: number; kind: "positive" | "negative" | "neutral" }> = [];
  const ageDays = daysBetween(task.createdAt, new Date().toISOString());
  const carryForwardCount = deriveCarryForwardCount(task, localDayKey(new Date()));
  const effort = inferEffortBucket(task);
  const jiraPlanningPriority = task.source === "Jira" ? deriveJiraPlanningPriority(task) : task.priority;

  if (jiraPlanningPriority === "High") {
    score += 34;
    scoreBreakdown.push({ label: "High execution priority", value: 34, kind: "positive" });
  } else if (jiraPlanningPriority === "Medium") {
    score += 18;
    scoreBreakdown.push({ label: "Medium execution priority", value: 18, kind: "positive" });
  } else {
    score += 8;
    scoreBreakdown.push({ label: "Low execution priority", value: 8, kind: "neutral" });
  }

  if (task.source === "Jira") {
    score += 10;
    scoreBreakdown.push({ label: "Jira task baseline", value: 10, kind: "positive" });
  }
  if (task.source === "Email") {
    score += 6;
    scoreBreakdown.push({ label: "Email task baseline", value: 6, kind: "neutral" });
  }
  if (task.source === "Jira" && isRecentlyUpdated(task.lastActivityAt)) {
    score += 16;
    reasons.push("Recent Jira activity");
    scoreBreakdown.push({ label: "Recent Jira changes", value: 16, kind: "positive" });
  } else if (task.source === "Jira") {
    score -= 4;
    reasons.push("Open Jira issue ready after current work");
    scoreBreakdown.push({ label: "No recent Jira change", value: -4, kind: "negative" });
  }
  if (task.source === "Jira" && jiraPlanningPriority === "High") {
    reasons.unshift("Current in-progress Jira story");
  } else if (task.source === "Jira" && jiraPlanningPriority === "Medium") {
    reasons.unshift("In-progress Jira subtask work");
  } else if (task.source === "Jira") {
    reasons.unshift("Pending Jira story or subtask");
  }
  if (task.source === "Jira" && task.jiraSubtaskEstimateSeconds && task.jiraSubtaskEstimateSeconds > 0) {
    const subtaskLoadBoost = Math.min(12, Math.ceil(task.jiraSubtaskEstimateSeconds / 3600) * 2);
    score += subtaskLoadBoost;
    scoreBreakdown.push({ label: "Subtask workload remaining", value: subtaskLoadBoost, kind: "positive" });
    if (task.jiraSubtaskEstimateSeconds >= 90 * 60) {
      reasons.push("Open subtasks drive the real workload");
    }
  }
  if (task.status === "In Progress") {
    score += 12;
    reasons.push("Already in progress");
    scoreBreakdown.push({ label: "Already in progress", value: 12, kind: "positive" });
  }
  if (task.deferredUntil) {
    const deferredTime = new Date(task.deferredUntil).getTime();
    const hoursUntil = (deferredTime - Date.now()) / 3_600_000;
    if (hoursUntil <= 0) {
      score += 20;
      reasons.push("Deferred task is due again");
      scoreBreakdown.push({ label: "Deferred task due again", value: 20, kind: "positive" });
    } else {
      score -= 18;
      scoreBreakdown.push({ label: "Deferred into the future", value: -18, kind: "negative" });
    }
  }
  if (ageDays >= 3 && task.status !== "Completed") {
    score += 10;
    reasons.push(`Unfinished for ${ageDays} days`);
    scoreBreakdown.push({ label: "Task age", value: 10, kind: "positive" });
  }
  if (carryForwardCount > 0) {
    const carryForwardBoost = Math.min(16, carryForwardCount * 4);
    score += carryForwardBoost;
    reasons.push(`Carried forward ${carryForwardCount} day${carryForwardCount === 1 ? "" : "s"}`);
    scoreBreakdown.push({ label: "Carry-forward pressure", value: carryForwardBoost, kind: "positive" });
  }
  if (task.source === "Email" && /(action required|follow up|approval|urgent)/i.test(task.title)) {
    score += 12;
    reasons.push("Urgency signal detected");
    scoreBreakdown.push({ label: "Urgency signal", value: 12, kind: "positive" });
  }
  if (task.source === "Jira" && task.jiraStatus && /blocked|review|qa/i.test(task.jiraStatus)) {
    score += 8;
    reasons.push(task.jiraStatus);
    scoreBreakdown.push({ label: "Workflow state pressure", value: 8, kind: "positive" });
  }
  const nextMeeting = meetings.find(
    (meeting) => isPlannableMeeting(meeting) && new Date(meeting.startTime).getTime() > Date.now()
  );
  if (nextMeeting && /prep|agenda|review/i.test(task.title)) {
    score += 10;
    reasons.push("Relevant to an upcoming meeting");
  }

  const explanation =
    task.source === "Jira"
      ? buildTaskStoryContext(task) ?? reasons[0] ?? "Assigned Jira issue is available as next work."
      : task.source === "Email"
        ? task.decisionReason ?? task.priorityExplanation ?? reasons[0] ?? "Recent email needs attention"
        : reasons[0] ?? "Manual task still active";

  const selectionReason =
    task.status === "In Progress"
      ? "Included because this work is already in progress and should stay visible until finished."
      : task.source === "Jira"
        ? "Included because it is assigned Jira work that fits your current capacity and execution order."
        : task.source === "Email"
          ? task.decisionReason
            ? `Included because the email was judged actionable: ${task.decisionReason}`
            : "Included because it looks like actionable email work that still needs review or follow-through."
          : "Included because you created it manually and it remains active.";

  const priorityReason =
    jiraPlanningPriority === "High"
      ? "High priority because active story work, urgency, or recent signals outweigh other tasks."
      : jiraPlanningPriority === "Medium"
        ? "Medium priority because it matters, but more active work is ahead of it."
        : task.source === "Email"
          ? "Priority reflects the email content, urgency, and your saved preferences."
          : "Low priority because it remains relevant, but higher-urgency work is already in motion.";

  const historySignals = [
    ageDays >= 3 ? `Open for ${ageDays} days` : null,
    carryForwardCount > 0 ? `Carried forward ${carryForwardCount} times` : null,
    task.lastChangedBy ? `Last changed by ${task.lastChangedBy}` : null,
    task.lastChangedAt ? `Last changed at ${task.lastChangedAt}` : null
  ].filter((entry): entry is string => Boolean(entry));

  return {
    priorityScore: Math.round(score),
    priorityExplanation: explanation,
    selectionReason,
    priorityReason,
    scoreBreakdown,
    historySignals,
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

function priorityForStage(stage: TaskStage): TaskPriority {
  switch (stage) {
    case "Now":
      return "High";
    case "Next":
      return "Medium";
    default:
      return "Low";
  }
}

function recommendedStageForTask(task: Task, plannedTodayTaskIds: Set<number>): TaskStage {
  if (task.status === "Completed") {
    return task.stage ?? "Later";
  }
  if (task.decisionState === "uncertain") {
    return "Review";
  }
  if (plannedTodayTaskIds.has(task.id)) {
    return "Now";
  }
  if (task.source === "Jira" && (task.status === "In Progress" || hasActiveJiraSubtask(task) || task.carryForwardCount > 0)) {
    return "Next";
  }
  if (task.priority === "High" || (task.priorityScore ?? 0) >= 42 || task.carryForwardCount > 0 || task.status === "In Progress") {
    return "Next";
  }
  if (task.source === "Email" && /(review|comment|mentioned|approval|action required)/i.test(task.title)) {
    return "Review";
  }
  return "Later";
}

function assignTaskStages(tasks: Task[], dayPlan: DayPlan) {
  const plannedTodayTaskIds = new Set(
    dayPlan.blocks.map((block) => block.taskId).filter((taskId): taskId is number => taskId !== null)
  );
  const grouped = new Map<TaskStage, Task[]>();
  for (const task of tasks) {
    if (task.status === "Completed") continue;
    const stage = task.manualOverrideFlags.includes("stage")
      ? task.stage
      : recommendedStageForTask(task, plannedTodayTaskIds);
    const existing = grouped.get(stage) ?? [];
    existing.push(task);
    grouped.set(stage, existing);
  }

  for (const stage of ["Now", "Next", "Later", "Review"] as TaskStage[]) {
    const ordered = (grouped.get(stage) ?? []).sort((left, right) => {
      if (left.manualOverrideFlags.includes("stageOrder") && right.manualOverrideFlags.includes("stageOrder")) {
        return left.stageOrder - right.stageOrder;
      }
      if (right.priorityScore !== left.priorityScore) return (right.priorityScore ?? 0) - (left.priorityScore ?? 0);
      if (left.status !== right.status) return left.status === "In Progress" ? -1 : right.status === "In Progress" ? 1 : 0;
      return new Date(right.updatedAt).getTime() - new Date(left.updatedAt).getTime();
    });

    ordered.forEach((task, index) => {
      const nextStage = task.manualOverrideFlags.includes("stage") ? task.stage : stage;
      const nextStageOrder = task.manualOverrideFlags.includes("stageOrder") ? task.stageOrder : index;
      const nextPriority = task.manualOverrideFlags.includes("priority") ? task.priority : priorityForStage(nextStage);
      if (task.stage === nextStage && task.stageOrder === nextStageOrder && task.priority === nextPriority) return;
      updateTask(task.id, {
        stage: nextStage,
        stageOrder: nextStageOrder,
        priority: nextPriority,
        lastChangedBy: "system",
        lastChangedAt: new Date().toISOString()
      });
    });
  }
}

function applyTaskIntelligence() {
  const meetings = listMeetings();
  const tasks = listTasks(undefined, { includeDeferred: true });
  for (const task of tasks) {
    const intelligence = computeTaskSignals(task, meetings);
    const overridePriority = task.manualOverrideFlags.includes("priority");
    updateTask(task.id, {
      priority:
        overridePriority
          ? task.priority
          : task.source === "Jira"
            ? deriveJiraPlanningPriority(task)
            : derivePriorityFromScore(intelligence.priorityScore),
      priorityScore: intelligence.priorityScore,
      priorityExplanation: intelligence.priorityExplanation,
      selectionReason: intelligence.selectionReason,
      priorityReason: intelligence.priorityReason,
      scoreBreakdown: intelligence.scoreBreakdown,
      historySignals: intelligence.historySignals,
      estimatedEffortBucket: intelligence.estimatedEffortBucket,
      taskAgeDays: intelligence.taskAgeDays,
      carryForwardCount: intelligence.carryForwardCount,
      lastChangedBy: "system",
      lastChangedAt: new Date().toISOString()
    });
  }
}

async function syncReminders(options?: {
  microsoftGraphAccessToken?: string | null;
}) {
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
    if (!isPlannableMeeting(meeting)) continue;
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

  const followUpSinceIso = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();
  try {
    const [sentEmails, inboxEmails] = options?.microsoftGraphAccessToken
      ? await Promise.all([
          fetchRecentSentEmailsWithAccessToken(followUpSinceIso, options.microsoftGraphAccessToken),
          fetchRecentEmailsWithAccessToken(followUpSinceIso, options.microsoftGraphAccessToken)
        ])
      : await Promise.all([fetchRecentSentEmails(followUpSinceIso), fetchRecentEmails(followUpSinceIso)]);

    for (const sentEmail of sentEmails) {
      if (!needsSentEmailFollowUp(sentEmail, inboxEmails)) continue;
      const reminderKey = `sent_follow_up:${sentEmail.id}`;
      activeKeys.push(reminderKey);
      upsertReminder({
        reminderKey,
        taskId: null,
        kind: "email_follow_up",
        title: sentEmail.subject?.trim() ? `Follow up: ${sentEmail.subject.trim()}` : "Follow up on sent email",
        reason: "You sent this email and there has not been a reply yet. Check if a follow-up is needed.",
        status: "active",
        sourceLink: sentEmail.webLink ?? null,
        sourceLabel: "Sent email",
        scheduledFor: sentEmail.sentDateTime ?? null,
        throttleUntil: new Date(Date.now() + cadenceMs).toISOString()
      });
    }
  } catch {
    // Reminder generation should stay resilient even if follow-up mail inspection fails.
  }

  resolveStaleReminders(activeKeys);
}

async function refreshPreferenceMemory() {
  const recentEvents = listRecentDecisionLogs(120).map((row) => ({
    action: String(row.action ?? ""),
    title: (() => {
      try {
        if (row.feedback_payload_json) {
          const parsed = JSON.parse(String(row.feedback_payload_json)) as { title?: string };
          if (parsed.title) return parsed.title;
        }
      } catch {}
      return String(row.source_ref ?? "Task");
    })(),
    source: String(row.source ?? ""),
    inferredReason: (row.inferred_reason as string | null) ?? null,
    inferredReasonTag: (row.inferred_reason_tag as string | null) ?? null,
    preferencePolarity: String(row.preference_polarity ?? "neutral")
  }));
  const profile = getUserPriorityProfile();
  const result = await distillPreferenceMemory({
    profile,
    recentEvents,
    sourceEventCount: getDecisionEventCount()
  });
  savePreferenceMemorySnapshot({
    snapshotJson: JSON.stringify(result.snapshot),
    insights: result.insights,
    sourceEventCount: getDecisionEventCount()
  });
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
        isPlannableMeeting(meeting)
      );
    })
    .reduce((sum, meeting) => sum + meeting.durationMinutes, 0);

  const taskMinutes = tasks
    .filter((task) => task.status !== "Completed")
    .reduce((sum, task) => sum + minutesForTask(task), 0);

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
  cleanupLegacyDailyWrapUpTasks();
  const tasks = listTasks();
  const meetings = listMeetings().filter((meeting) => isPlannableMeeting(meeting));
  const dayPlan = buildDayPlan(tasks, meetings);
  assignTaskStages(tasks, dayPlan);
  const stageAwareTasks = listTasks().filter((task) => task.status !== "Completed");
  upsertDailyPlanSnapshot({
    dayKey: dayPlan.summary.dayKey,
    weekday: new Date().getDay(),
    baseWorkdayMinutes: dayPlan.summary.baseWorkdayMinutes,
    adaptedTaskCapacityMinutes: dayPlan.summary.adaptedTaskCapacityMinutes,
    remainingTaskCapacityMinutes: dayPlan.summary.remainingTaskCapacityMinutes,
    meetingMinutes: dayPlan.summary.meetingMinutes,
    plannedTaskMinutes: dayPlan.summary.plannedTaskMinutes,
    completedTaskMinutes: dayPlan.summary.completedTaskMinutes,
    remainingTaskMinutes: dayPlan.summary.remainingTaskMinutes,
    spilloverTaskCount: dayPlan.summary.spilloverTaskCount,
    freeMinutes: dayPlan.summary.freeMinutes,
    focusFactor: dayPlan.summary.focusFactor,
    completionRate: dayPlan.summary.completionRate,
    plannedTaskIds: dayPlan.blocks
      .map((block) => block.taskId)
      .filter((taskId): taskId is number => taskId !== null),
    summaryJson: JSON.stringify(dayPlan.summary),
    blocksJson: JSON.stringify(dayPlan.blocks)
  });
  return {
    meetings,
    tasks: groupTasksByPriority(stageAwareTasks),
    reminders: listReminderItems(["active", "dismissed"]),
    workload: buildWorkloadSummary(stageAwareTasks),
    dayPlan,
    deferredTaskCount: listDeferredTasks().length,
    rejectedTaskCount: getRejectedTaskCount(),
    automation: getAutomationSettings(),
    sync: {
      microsoft: getSyncState("microsoft"),
      jira: getSyncState("jira"),
      lastGeneratedAt: getSyncState("plan")
    },
    warnings
  };
}

function persistPlannerRun(runId: string, triggerType: PlannerRunDetail["triggerType"], preferredTimeZone: string | null, payload: TodayPayload) {
  upsertPlannerRunDetail({
    runId,
    triggerType,
    preferredTimeZone,
    warnings: payload.warnings,
    meetingCount: payload.meetings.length,
    activeTaskCount: [...payload.tasks.High, ...payload.tasks.Medium, ...payload.tasks.Low].length,
    rejectedTaskCount: payload.rejectedTaskCount,
    deferredTaskCount: payload.deferredTaskCount,
    workloadState: payload.workload.state
  });
}

export function getTodaySnapshot() {
  return buildPayload([]);
}

async function postSyncMaintenance(
  triggerType: "manual" | "scheduled",
  warnings: string[],
  options?: { microsoftGraphAccessToken?: string | null }
) {
  applyTaskIntelligence();
  await applyPlanningCategories();
  await syncReminders(options);
  await refreshPreferenceMemory();
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
    runId?: string | null;
  },
  triggerType: "manual" | "scheduled" = "manual"
): Promise<TodayPayload> {
  const runId = options?.runId ?? createCorrelationId();
  logEvent({
    eventType: "planner.generate",
    runId,
    status: "started",
    message: "Planner generation started.",
    metadata: { triggerType, preferredTimeZone: options?.preferredTimeZone ?? null }
  });
  const jiraWarnings = await syncJiraTasks(runId);
  const microsoftMeetingWarnings = await syncMicrosoftMeetings({ ...options, runId });
  const microsoftTaskWarnings = await syncMicrosoftTasks({ ...options, runId });
  const warnings = [...microsoftTaskWarnings, ...microsoftMeetingWarnings, ...jiraWarnings];
  await postSyncMaintenance(triggerType, warnings, { microsoftGraphAccessToken: options?.microsoftGraphAccessToken ?? null });
  const payload = buildPayload(warnings);
  persistPlannerRun(runId, triggerType, options?.preferredTimeZone ?? null, payload);
  logEvent({
    eventType: "planner.generate",
    runId,
    status: "success",
    message: "Planner generation completed.",
    metadata: {
      triggerType,
      warningsCount: warnings.length,
      meetingCount: payload.meetings.length,
      activeTaskCount: [...payload.tasks.High, ...payload.tasks.Medium, ...payload.tasks.Low].length
    }
  });
  return payload;
}

export async function syncMeetingsOnly(options?: {
  microsoftGraphAccessToken?: string | null;
  microsoftWarning?: string | null;
  preferredTimeZone?: string | null;
  runId?: string | null;
}) {
  const runId = options?.runId ?? createCorrelationId();
  const warnings = await syncMicrosoftMeetings({ ...options, runId });
  await syncReminders({ microsoftGraphAccessToken: options?.microsoftGraphAccessToken ?? null });
  const payload = buildPayload(warnings);
  persistPlannerRun(runId, "sync", options?.preferredTimeZone ?? null, payload);
  return payload;
}

export async function syncTasksOnly(options?: {
  microsoftGraphAccessToken?: string | null;
  microsoftWarning?: string | null;
  preferredTimeZone?: string | null;
  runId?: string | null;
}) {
  const runId = options?.runId ?? createCorrelationId();
  const jiraWarnings = await syncJiraTasks(runId);
  const microsoftMeetingWarnings = await syncMicrosoftMeetings({ ...options, runId });
  const microsoftTaskWarnings = await syncMicrosoftTasks({ ...options, runId });
  applyTaskIntelligence();
  await syncReminders({ microsoftGraphAccessToken: options?.microsoftGraphAccessToken ?? null });
  const payload = buildPayload([...microsoftTaskWarnings, ...microsoftMeetingWarnings, ...jiraWarnings]);
  persistPlannerRun(runId, "sync", options?.preferredTimeZone ?? null, payload);
  return payload;
}

export function getDeferredTasksPayload() {
  return { tasks: listDeferredTasks() };
}

export function getReminderCenterPayload() {
  return { reminders: listReminderItems(["active", "dismissed", "resolved"]) };
}
