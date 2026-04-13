import {
  clearRejectedTasksBySourceThread,
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
import {
  buildJiraIssueBrowseUrl,
  fetchJiraIssuePlanningContext,
  fetchOpenAssignedIssues,
  getMappedJiraPriority,
  mapJiraWorkflowStatus
} from "../providers/jira.js";
import {
  fetchRecentEmails,
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
  TaskSource,
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
  return { profile, memory, recentExamples };
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

function extractIssueKeys(text: string) {
  return [...new Set((text.toUpperCase().match(/\b[A-Z][A-Z0-9]+-\d+\b/g) ?? []).map((key) => key.trim()))];
}

function buildMeetingTitleIndex(meetings: ReturnType<typeof listMeetings>) {
  return meetings
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
  const isOrgNoise =
    /(release highlights|release showcase|showcase|what'?s new|monthly deep dive|getting started with chatgpt|part \d+ of \d+|newsletter|optional|thank you and best wishes|scam of the week|badminton|hackathon|transport requests)/.test(
      fullText
    );
  const isDevWorkflow =
    /(pull request|merge request|review requested|opened a pull request|github|gitlab)/.test(fullText);
  const isSecurityOrAccountAction =
    /(password is about to expire|password.*expire|credential.*expire|certificate.*expire|token.*expire|security alert|mfa)/.test(
      fullText
    );
  const isCommentOrMention =
    /(mentioned you|comment(ed)?|comment added|requested review|requested changes|tagged you|needs your review|your input|reply needed)/.test(
      fullText
    );
  const isGenericAction =
    /(action required|approval required|please review|approve|follow up|required from you|your input needed|missing lifecycle rules)/.test(
      fullText
    );
  const isLikelyJiraNotification =
    issueKeys.length > 0 &&
    !isDevWorkflow &&
    /(\[jira\]|jira|comment(ed)?|comment added|mentioned you|status|open|reopened|resolved|updated)/.test(fullText);
  const explicitHumanRequest =
    !/(noreply|notification|service-now|automated)/.test(sender) &&
    /(please|can you|need you|follow up|reply needed|your input|please review|approve)/.test(fullText);

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

  if (isOrgNoise) {
    return {
      route: "drop" as const,
      priority: "Low" as TaskPriority,
      reason: "Newsletter, announcement, optional, or social email with no clear task.",
      reasonTags: ["newsletter_like"] as ReasonTag[]
    };
  }

  if (isSecurityOrAccountAction) {
    return {
      route: "accept" as const,
      priority: "Medium" as TaskPriority,
      reason: "Security or account maintenance email requires action.",
      reasonTags: ["direct_request"] as ReasonTag[]
    };
  }

  if (isDevWorkflow || explicitHumanRequest) {
    return {
      route: "accept" as const,
      priority: /(urgent|asap|today|immediately|by eod)/.test(fullText) ? "High" : "Medium",
      reason: isDevWorkflow
        ? "Code workflow email likely needs review or follow-up."
        : "Direct request likely needs action.",
      reasonTags: ["direct_request"] as ReasonTag[]
    };
  }

  if (isCommentOrMention || isGenericAction || input.classification.actionable) {
    return {
      route: "review" as const,
      priority: input.classification.priority,
      reason:
        isCommentOrMention
          ? "Comment, mention, or review activity may matter and should stay reviewable."
          : "Potential work item is not confirmed enough for the active task list.",
      reasonTags: isCommentOrMention
        ? (["historically_accepted"] as ReasonTag[])
        : (["fyi_only"] as ReasonTag[])
    };
  }

  return {
    route: "drop" as const,
    priority: "Low" as TaskPriority,
    reason: "No clear work signal found.",
    reasonTags: ["fyi_only"] as ReasonTag[]
  };
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
    const classification = await classifyEmail(email);
    const now = new Date().toISOString();
    const emailText = `${email.subject ?? ""} ${email.bodyPreview ?? ""}`;
    const hasCommentSignal = /(comment(ed)?|mentioned you|requested review|requested changes|assigned to you|reply needed|your input|needs your review|tagged you)/i.test(
      emailText
    );
    const hasUserRelevantSignal = /(you|your|assigned|review|approve|follow up|reply|blocker|due|owner)/i.test(emailText);
    const candidatePayload = {
      title: classification.title,
      source: "Email" as const,
      sourceLink: email.webLink ?? null,
      sourceRef: email.id,
      sourceThreadRef: email.conversationId ?? null,
      sender: email.from?.emailAddress?.address ?? null,
      bodyPreview: email.bodyPreview ?? null,
      isDirectRequest: /(please|can you|need you|action required|review|approve)/i.test(
        `${email.subject ?? ""} ${email.bodyPreview ?? ""}`
      ),
      dueSoon: /(today|asap|urgent|eod|tomorrow)/i.test(`${email.subject ?? ""} ${email.bodyPreview ?? ""}`),
      isBotLike:
        /(noreply|notification|service-now|automated)/i.test(
          `${email.from?.emailAddress?.address ?? ""} ${email.subject ?? ""}`
        ) && !(hasCommentSignal && hasUserRelevantSignal),
      isDuplicate: false,
      meetingRelevant: /(meeting|agenda|prep)/i.test(`${email.subject ?? ""} ${email.bodyPreview ?? ""}`)
    };

    const strongSignal =
      candidatePayload.isDirectRequest ||
      candidatePayload.dueSoon ||
      candidatePayload.meetingRelevant ||
      (hasCommentSignal && hasUserRelevantSignal) ||
      !candidatePayload.isBotLike;
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
          lastActivityAt: now,
          decisionState: threadTask.decisionState === "restored" ? "restored" : input.decisionState,
          decisionConfidence: input.decisionConfidence,
          decisionReason: input.decisionReason,
          decisionReasonTags: input.decisionReasonTags,
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

    if (finalEvaluation.relevance === "reject" || (finalEvaluation.relevance === "uncertain" && finalEvaluation.confidence < 0.6)) {
      if (threadTask) {
        preserveVisibleEmailThread({
          priority: threadTask.priority,
          decisionState: threadTask.decisionState ?? "restored",
          decisionConfidence: finalEvaluation.confidence,
          decisionReason: threadTask.decisionReason ?? finalEvaluation.why,
          decisionReasonTags: finalEvaluation.reasonTags
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
            decisionState: exactTask.decisionState ?? "restored",
            decisionConfidence: finalEvaluation.confidence,
            decisionReason: exactTask.decisionReason ?? finalEvaluation.why,
            decisionReasonTags: finalEvaluation.reasonTags,
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
        proposedPriority: finalEvaluation.priority,
        decisionConfidence: finalEvaluation.confidence,
        decisionReason: finalEvaluation.why,
        decisionReasonTags: finalEvaluation.reasonTags,
        candidatePayloadJson: payloadJson,
        personalizationVersion: personalization.memory.version ?? 1
      });
      continue;
    }

    clearRejectedCandidatesForThread("Email", email.id, email.conversationId ?? null);
    if (threadTask && !exactTask) {
      updateTask(threadTask.id, {
        title: classification.title,
        priority: threadTask.manualOverrideFlags.includes("priority") ? undefined : finalEvaluation.priority,
        lastActivityAt: now,
        decisionState: threadTask.decisionState === "restored" ? "restored" : finalEvaluation.relevance === "uncertain" ? "uncertain" : "accepted",
        decisionConfidence: finalEvaluation.confidence,
        decisionReason: finalEvaluation.why,
        decisionReasonTags: finalEvaluation.reasonTags,
        personalizationVersion: personalization.memory.version ?? 1,
        rejectedAt: null
      });
    } else {
      upsertTask({
        title: classification.title,
        source: "Email",
        priority: finalEvaluation.priority,
        sourceLink: email.webLink ?? null,
        sourceRef: email.id,
        sourceThreadRef: email.conversationId ?? null,
        decisionState: finalEvaluation.relevance === "uncertain" ? "uncertain" : "accepted",
        decisionConfidence: finalEvaluation.confidence,
        decisionReason: finalEvaluation.why,
        decisionReasonTags: finalEvaluation.reasonTags,
        personalizationVersion: personalization.memory.version ?? 1
      });
    }
    logTaskDecisionEvent({
      source: "Email",
      sourceRef: email.id,
      sourceThreadRef: email.conversationId ?? null,
      action: "system_evaluated",
      afterPriority: threadTask && !exactTask ? (threadTask.manualOverrideFlags.includes("priority") ? threadTask.priority : finalEvaluation.priority) : finalEvaluation.priority,
      systemDecisionState: finalEvaluation.relevance === "uncertain" ? "uncertain" : "accepted",
      decisionConfidence: finalEvaluation.confidence,
      decisionReason: finalEvaluation.why,
      decisionReasonTags: finalEvaluation.reasonTags,
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
          isCancelled
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
  const effective = Math.max(parentMinutes, subtaskMinutes);
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
const DEFAULT_WORKDAY_START_HOUR = 9;
const DEFAULT_WORKDAY_START_MINUTE = 30;

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
  if (task.status === "In Progress") {
    return Math.max(15, Math.round(total * 0.6));
  }
  return total;
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
  return 4;
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

function buildDayPlan(tasks: Task[], meetings: ReturnType<typeof listMeetings>): DayPlan {
  const now = new Date();
  const dayKey = localDayKey(now);
  const weekday = now.getDay();
  const todayStart = startOfLocalDay(now);
  const todayEnd = endOfLocalDay(now);
  const workdayStart = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
    DEFAULT_WORKDAY_START_HOUR,
    DEFAULT_WORKDAY_START_MINUTE,
    0,
    0
  );
  const workdayEnd = addMinutes(workdayStart, DEFAULT_WORKDAY_MINUTES);

  const todaysMeetings = meetings
    .filter((meeting) => !meeting.isCancelled && meetingStart(meeting) < todayEnd && meetingEnd(meeting) > todayStart)
    .sort((left, right) => meetingStart(left).getTime() - meetingStart(right).getTime());
  const remainingMeetings = todaysMeetings.filter((meeting) => meetingEnd(meeting).getTime() > now.getTime());
  const meetingMinutes = todaysMeetings.reduce((sum, meeting) => sum + meeting.durationMinutes, 0);

  const baseTaskCapacityMinutes = Math.max(0, DEFAULT_WORKDAY_MINUTES - meetingMinutes);
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
      const rank =
        (task.priorityScore ?? 0) +
        (task.status === "In Progress" ? 22 : 0) +
        ((task.source === "Jira" ? deriveJiraPlanningPriority(task) : task.priority) === "High"
          ? 18
          : (task.source === "Jira" ? deriveJiraPlanningPriority(task) : task.priority) === "Medium"
            ? 9
            : 3) +
        Math.min(10, task.carryForwardCount * 2) +
        (task.source === "Email" && /follow up|reply|approval|action/i.test(task.title) ? 6 : 0) +
        (task.source === "Jira" && isRecentlyUpdated(task.lastActivityAt) ? 8 : 0);
      return {
        task,
        remainingMinutes,
        remainingEstimate: remainingMinutes,
        executionBucket,
        rank
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
    const quickCandidates = rankedTasks.filter((entry) => entry.remainingEstimate > 0 && entry.remainingEstimate <= 45);
    const fallbackCandidates = rankedTasks.filter((entry) => entry.remainingEstimate > 0);
    const inFlightJiraCandidate = fallbackCandidates.find(
      (entry) => entry.task.source === "Jira" && entry.executionBucket <= 1
    );
    const chosen =
      inFlightJiraCandidate ??
      (preferQuickWin && fallbackCandidates[0]?.executionBucket > 1 ? quickCandidates[0] : null) ??
      fallbackCandidates.find((entry) => entry.remainingEstimate <= slotMinutes + 15) ??
      fallbackCandidates[0];

    if (!chosen) {
      return { block: null, consumedMinutes: 0, nextFocusStreak: focusStreak };
    }

    const capacityLeft = Math.max(
      0,
      remainingTaskCapacityMinutes -
        taskBlocks.reduce((sum, block) => sum + block.durationMinutes, 0)
    );
    const canForceVisibilityBlock =
      taskBlocks.length === 0 &&
      slotMinutes >= 15 &&
      (chosen.task.status === "In Progress" || chosen.executionBucket <= 1);
    const effectiveCapacityLeft = capacityLeft >= 15 ? capacityLeft : canForceVisibilityBlock ? 15 : 0;

    if (effectiveCapacityLeft < 15) {
      return { block: null, consumedMinutes: 0, nextFocusStreak: focusStreak };
    }

    const planningPriority = chosen.task.source === "Jira" ? deriveJiraPlanningPriority(chosen.task) : chosen.task.priority;
    const idealChunk =
      chosen.remainingEstimate <= 30
        ? chosen.remainingEstimate
        : chosen.executionBucket <= 1 || planningPriority === "High" || chosen.task.status === "In Progress"
          ? Math.min(120, chosen.remainingEstimate)
          : Math.min(60, chosen.remainingEstimate);
    const durationMinutes = Math.max(15, Math.min(slotMinutes, effectiveCapacityLeft, idealChunk));
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

  while (cursor.getTime() < schedulingEnd.getTime()) {
    const slotMinutes = diffMinutes(cursor, schedulingEnd);
    const scheduled = scheduleTaskBlock(cursor, slotMinutes, focusStreak);
    if (!scheduled.block) break;
    taskBlocks.push(scheduled.block);
    blocks.push(scheduled.block);
    cursor = addMinutes(cursor, scheduled.consumedMinutes);
    focusStreak = scheduled.nextFocusStreak;
  }

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
      baseWorkdayMinutes: DEFAULT_WORKDAY_MINUTES,
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
    blocks: blocks.sort((left, right) => new Date(left.startTime).getTime() - new Date(right.startTime).getTime()),
    spilloverTasks
  };
}

function computeTaskSignals(task: Task, meetings: ReturnType<typeof listMeetings>) {
  let score = 0;
  const reasons: string[] = [];
  const scoreBreakdown: Array<{ label: string; value: number; kind: "positive" | "negative" | "neutral" }> = [];
  const ageDays = daysBetween(task.createdAt, new Date().toISOString());
  const carryForwardCount = task.status === "Completed" ? 0 : Math.max(0, ageDays - 1);
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
    (meeting) => !meeting.isCancelled && new Date(meeting.startTime).getTime() > Date.now()
  );
  if (nextMeeting && /prep|agenda|review/i.test(task.title)) {
    score += 10;
    reasons.push("Relevant to an upcoming meeting");
  }

  const explanation =
    task.source === "Jira"
      ? buildTaskStoryContext(task) ?? reasons[0] ?? "Assigned Jira issue is available as next work."
      : task.source === "Email"
        ? reasons[0] ?? "Recent email needs attention"
        : reasons[0] ?? "Manual task still active";

  const selectionReason =
    task.status === "In Progress"
      ? "Included because this work is already in progress and should stay visible until finished."
      : task.source === "Jira"
        ? "Included because it is assigned Jira work that fits your current capacity and execution order."
        : task.source === "Email"
          ? "Included because it looks like actionable email work that still needs review or follow-through."
          : "Included because you created it manually and it remains active.";

  const priorityReason =
    jiraPlanningPriority === "High"
      ? "High priority because active story work, urgency, or recent signals outweigh other tasks."
      : jiraPlanningPriority === "Medium"
        ? "Medium priority because it matters, but more active work is ahead of it."
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
        !meeting.isCancelled
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
  const tasks = listTasks();
  const meetings = listMeetings();
  const dayPlan = buildDayPlan(tasks, meetings);
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
    tasks: groupTasksByPriority(tasks),
    reminders: listReminderItems(["active", "dismissed"]),
    workload: buildWorkloadSummary(tasks),
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

async function postSyncMaintenance(triggerType: "manual" | "scheduled", warnings: string[]) {
  applyTaskIntelligence();
  syncReminders();
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
  await postSyncMaintenance(triggerType, warnings);
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
  syncReminders();
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
  const microsoftWarnings = await syncMicrosoftTasks({ ...options, runId });
  applyTaskIntelligence();
  syncReminders();
  const payload = buildPayload([...microsoftWarnings, ...jiraWarnings]);
  persistPlannerRun(runId, "sync", options?.preferredTimeZone ?? null, payload);
  return payload;
}

export function getDeferredTasksPayload() {
  return { tasks: listDeferredTasks() };
}

export function getReminderCenterPayload() {
  return { reminders: listReminderItems(["active", "dismissed", "resolved"]) };
}
