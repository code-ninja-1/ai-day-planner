import { env } from "../env.js";
import type {
  FeedbackAction,
  FeedbackPolarity,
  PersonalizationInsight,
  PriorityBias,
  ReasonTag,
  TaskPriority,
  TaskSource,
  TaskDecisionState,
  UserPriorityProfile
} from "../types.js";

export interface CandidateInput {
  title: string;
  source: TaskSource;
  sourceLink?: string | null;
  sourceRef?: string | null;
  sourceThreadRef?: string | null;
  jiraStatus?: string | null;
  jiraPriority?: string | null;
  projectKey?: string | null;
  sender?: string | null;
  bodyPreview?: string | null;
  dueSoon?: boolean;
  isAssignedToUser?: boolean;
  isDirectRequest?: boolean;
  isBotLike?: boolean;
  isDuplicate?: boolean;
  meetingRelevant?: boolean;
}

export interface CandidateEvaluation {
  relevance: "accept" | "reject" | "uncertain";
  priority: TaskPriority;
  confidence: number;
  why: string;
  reasonTags: ReasonTag[];
}

export interface FeedbackReasonResult {
  likelyReason: string;
  reasonTag: ReasonTag | null;
  positiveOrNegativePreference: FeedbackPolarity;
  confidence: number;
}

export interface PreferenceMemorySnapshot {
  version: number;
  positiveTags: ReasonTag[];
  negativeTags: ReasonTag[];
  repeatedWins: string[];
  repeatedNoise: string[];
}

interface CandidateContentAnalysis {
  actionNeeded: boolean;
  urgency: "low" | "medium" | "high";
  relevance: "low" | "medium" | "high";
  noise: "low" | "medium" | "high";
  summary: string;
  reasonTags: ReasonTag[];
}

const reasonTagEnum: ReasonTag[] = [
  "direct_request",
  "manager_visibility",
  "project_critical",
  "comment_noise",
  "newsletter_like",
  "duplicate_signal",
  "meeting_related",
  "historically_rejected",
  "historically_accepted",
  "assigned_work",
  "due_soon",
  "bot_generated",
  "review_request",
  "blocking_work",
  "fyi_only"
];

export const defaultPriorityProfile: UserPriorityProfile = {
  personalizationEnabled: true,
  roleFocus: null,
  prioritizationPrompt:
    "Prioritize direct requests, assigned Jira work, blockers, urgent items, and work tied to important projects or people. Deprioritize newsletters, automated digests, noisy comment notifications, and FYI-only updates unless I explicitly restore them.",
  importantWork: [],
  noiseWork: [],
  mustNotMiss: [],
  importantSources: ["Jira"],
  importantPeople: [],
  importantProjects: [],
  positiveReasonTags: ["direct_request", "assigned_work", "blocking_work", "due_soon"],
  negativeReasonTags: ["newsletter_like", "comment_noise", "bot_generated", "fyi_only"],
  filteringStyle: "conservative",
  priorityBias: "balanced",
  questionnaireJson: null,
  exampleRankingsJson: null,
  lastProfileRefreshAt: null,
  updatedAt: null
};

function normalized(value: string | null | undefined) {
  return (value ?? "").trim().toLowerCase();
}

function uniqueReasonTags(tags: Array<ReasonTag | null | undefined>) {
  return [...new Set(tags.filter((tag): tag is ReasonTag => Boolean(tag)))];
}

function listMatches(haystack: string, values: string[]) {
  const normalizedHaystack = normalized(haystack);
  return values.filter((value) => {
    const token = normalized(value);
    return token.length >= 2 && normalizedHaystack.includes(token);
  });
}

function computeContentSignals(candidate: CandidateInput) {
  const text = `${candidate.title} ${candidate.bodyPreview ?? ""} ${candidate.jiraStatus ?? ""} ${candidate.jiraPriority ?? ""}`.toLowerCase();
  const sender = normalized(candidate.sender);
  const tags = uniqueReasonTags([
    candidate.isDirectRequest || /(please|can you|need you|action required|reply needed|follow up needed)/.test(text)
      ? "direct_request"
      : null,
    candidate.isAssignedToUser ? "assigned_work" : null,
    candidate.dueSoon || /(today|asap|urgent|eod|tomorrow|by end of day|before standup)/.test(text) ? "due_soon" : null,
    candidate.meetingRelevant || /(prep|agenda|meeting|standup|sync-up|office hours)/.test(text) ? "meeting_related" : null,
    /(blocked|blocker|incident|prod|production|escalation|sev|critical)/.test(text) ? "blocking_work" : null,
    /(review request|please review|needs review|approval required|approve|requested review|requested changes)/.test(text)
      ? "review_request"
      : null,
    /(mentioned you|tagged you|for you|assigned to you|your input|reply needed)/.test(text)
      ? "historically_accepted"
      : null,
    /(newsletter|digest|announcement|weekly update|roundup)/.test(text) ? "newsletter_like" : null,
    candidate.isBotLike || /(noreply|notification|service-now|automated|mailer-daemon)/.test(`${sender} ${text}`)
      ? "bot_generated"
      : null,
    /(fyi|for your information|no action required)/.test(text) ? "fyi_only" : null,
    candidate.isDuplicate ? "duplicate_signal" : null,
    /(commented|comment on|watcher|new comment|comment added)/.test(text) &&
    !/(mentioned you|tagged you|requested review|assigned to you|your input|reply needed)/.test(text)
      ? "comment_noise"
      : null
  ]);

  return {
    text,
    sender,
    tags,
    hasStrongUserSignal:
      candidate.isAssignedToUser ||
      candidate.isDirectRequest ||
      candidate.dueSoon ||
      tags.includes("review_request") ||
      /(mentioned you|tagged you|reply needed|your input)/.test(text),
    hasHighNoiseSignal:
      tags.includes("newsletter_like") ||
      tags.includes("bot_generated") ||
      tags.includes("comment_noise") ||
      tags.includes("fyi_only")
  };
}

function extractLikelyReasonTags(candidate: CandidateInput): ReasonTag[] {
  return computeContentSignals(candidate).tags;
}

function scoreWithProfile(
  candidate: CandidateInput,
  profile: UserPriorityProfile,
  memory: PreferenceMemorySnapshot,
  contentAnalysis?: CandidateContentAnalysis | null
) {
  const signals = computeContentSignals(candidate);
  const llmTags = contentAnalysis?.reasonTags ?? [];
  const tags = uniqueReasonTags([...signals.tags, ...llmTags]);
  let importanceScore = 0;
  let noiseScore = 0;
  let urgencyScore = 0;
  const reasons: string[] = [];
  const sender = normalized(candidate.sender);
  const title = normalized(`${candidate.title} ${candidate.bodyPreview ?? ""}`);
  const prompt = normalized(profile.prioritizationPrompt);
  const projectKey = normalized(candidate.projectKey);
  const importantSources = profile.importantSources ?? [];
  const importantPeople = profile.importantPeople ?? [];
  const importantProjects = profile.importantProjects ?? [];
  const focusAreas = profile.importantWork ?? [];
  const ignorePatterns = profile.noiseWork ?? [];
  const neverHide = profile.mustNotMiss ?? [];
  const positiveMemoryTags = memory.positiveTags ?? [];
  const negativeMemoryTags = memory.negativeTags ?? [];
  const repeatedNoise = memory.repeatedNoise ?? [];
  const repeatedWins = memory.repeatedWins ?? [];

  if (candidate.isDuplicate || tags.includes("duplicate_signal")) return {
    evaluation: {
      relevance: "reject" as const,
      priority: "Low" as TaskPriority,
      confidence: 0.99,
      why: "Duplicate source already represented in your plan.",
      reasonTags: ["duplicate_signal"] as ReasonTag[]
      }
  };

  if (candidate.isAssignedToUser || tags.includes("assigned_work")) {
    importanceScore += 30;
    reasons.push("Assigned directly to you");
  }
  if (tags.includes("direct_request")) {
    importanceScore += 24;
    reasons.push("Contains a direct ask");
  }
  if (tags.includes("blocking_work")) {
    importanceScore += 22;
    reasons.push("Looks blocking or urgent");
  }
  if (tags.includes("due_soon")) {
    urgencyScore += 16;
    reasons.push("Time-sensitive signal detected");
  }
  if (tags.includes("review_request")) {
    importanceScore += 14;
    reasons.push("Needs review or approval");
  }
  if (candidate.source === "Jira") importanceScore += 10;
  if (candidate.source === "Email" && signals.hasStrongUserSignal) importanceScore += 6;

  if (importantSources.map(normalized).includes(normalized(candidate.source))) {
    importanceScore += 10;
    reasons.push(`You usually prioritize ${candidate.source}`);
  }
  const priorityPeopleMatches = sender ? listMatches(sender, importantPeople) : [];
  if (priorityPeopleMatches.length) {
    importanceScore += 16;
    reasons.push("From a person or sender you care about");
  }
  if (projectKey && importantProjects.map(normalized).includes(projectKey)) {
    importanceScore += 15;
    reasons.push("Matches a priority project");
  }
  const focusMatches = listMatches(title, focusAreas);
  if (focusMatches.length) {
    importanceScore += 12;
    reasons.push("Matches one of your focus areas");
  }
  const neverHideMatches = listMatches(title, neverHide);
  if (neverHideMatches.length) {
    importanceScore += 26;
    urgencyScore += 8;
    reasons.push("Marked as something you never want hidden");
  }

  const ignoreMatches = listMatches(title, ignorePatterns);
  if (ignoreMatches.length) {
    noiseScore += 18;
    reasons.push("Matches patterns you usually ignore");
  }

  if (prompt) {
    if (
      /(assigned|direct request|blocker|urgent|important project|important people|review request|mention)/.test(prompt) &&
      (tags.includes("assigned_work") ||
        tags.includes("direct_request") ||
        tags.includes("blocking_work") ||
        tags.includes("due_soon") ||
        tags.includes("review_request"))
    ) {
      importanceScore += 8;
    }
    if (
      /(newsletter|digest|comment notification|fyi|bot|automated)/.test(prompt) &&
      (tags.includes("newsletter_like") ||
        tags.includes("comment_noise") ||
        tags.includes("fyi_only") ||
        tags.includes("bot_generated"))
    ) {
      noiseScore += 10;
    }
  }

  if (tags.includes("newsletter_like")) noiseScore += 24;
  if (tags.includes("bot_generated")) noiseScore += signals.hasStrongUserSignal ? 2 : 10;
  if (tags.includes("comment_noise")) noiseScore += signals.hasStrongUserSignal ? 0 : 14;
  if (tags.includes("fyi_only")) noiseScore += signals.hasStrongUserSignal ? 0 : 10;

  for (const tag of tags) {
    if (positiveMemoryTags.includes(tag)) importanceScore += 6;
    if (negativeMemoryTags.includes(tag)) noiseScore += 9;
  }

  if (repeatedNoise.some((pattern) => title.includes(normalized(pattern)))) {
    noiseScore += 12;
    reasons.push("Looks similar to items you often reject");
  }
  if (repeatedWins.some((pattern) => title.includes(normalized(pattern)))) {
    importanceScore += 10;
    reasons.push("Looks similar to items you usually keep");
  }

  if (contentAnalysis) {
    importanceScore +=
      contentAnalysis.relevance === "high" ? 14 : contentAnalysis.relevance === "medium" ? 7 : 0;
    urgencyScore += contentAnalysis.urgency === "high" ? 12 : contentAnalysis.urgency === "medium" ? 6 : 0;
    noiseScore += contentAnalysis.noise === "high" ? 12 : contentAnalysis.noise === "medium" ? 5 : 0;
    if (contentAnalysis.actionNeeded) {
      importanceScore += 10;
      reasons.unshift(contentAnalysis.summary);
    }
  }

  const preferenceBias: Record<PriorityBias, { accept: number; reject: number; priority: number }> = {
    focus: { accept: 14, reject: -10, priority: 4 },
    balanced: { accept: 10, reject: -12, priority: 0 },
    coverage: { accept: 6, reject: -15, priority: -3 }
  };
  const styleThresholds: Record<
    UserPriorityProfile["filteringStyle"],
    { accept: number; reject: number }
  > = {
    conservative: { accept: 14, reject: -8 },
    balanced: { accept: 10, reject: -12 },
    aggressive: { accept: 6, reject: -16 }
  };
  const netScore = importanceScore + urgencyScore - noiseScore;
  const thresholds = styleThresholds[profile.filteringStyle];
  const bias = preferenceBias[profile.priorityBias];
  const adjustedScore = netScore + bias.priority;

  const priority: TaskPriority =
    adjustedScore >= 34 ? "High" : adjustedScore >= 14 ? "Medium" : "Low";
  const relevance =
    adjustedScore >= thresholds.accept + bias.accept
      ? "accept"
      : adjustedScore <= thresholds.reject + bias.reject && !signals.hasStrongUserSignal && !neverHideMatches.length
        ? "reject"
        : "uncertain";

  const mainWhy =
    reasons[0] ??
    contentAnalysis?.summary ??
    (relevance === "reject"
      ? "Low-signal item based on your preferences."
      : "Matches your current work preferences.");
  return {
    evaluation: {
      relevance,
      priority,
      confidence: Math.max(0.38, Math.min(0.97, 0.58 + Math.abs(adjustedScore) / 70)),
      why: mainWhy,
      reasonTags: tags
    }
  };
}

async function analyzeCandidateContent(
  candidate: CandidateInput,
  profile: UserPriorityProfile
): Promise<CandidateContentAnalysis | null> {
  try {
    return await callResponsesApi<CandidateContentAnalysis>(
      [
        {
          role: "system",
          content: [
            {
              type: "input_text",
              text:
                "Analyze a work item for a daily planner. Focus on whether the content likely needs action, how urgent it is, how relevant it is to the user, and how noisy it looks. Do not decide final accept or reject. Return strict JSON only."
            }
          ]
        },
        {
          role: "user",
          content: [
            {
              type: "input_text",
              text: JSON.stringify({
                candidate,
                profile: {
                  roleFocus: profile.roleFocus,
                  prioritizationPrompt: profile.prioritizationPrompt,
                  focusAreas: profile.importantWork,
                  ignorePatterns: profile.noiseWork,
                  neverHide: profile.mustNotMiss,
                  importantPeople: profile.importantPeople,
                  importantProjects: profile.importantProjects
                }
              })
            }
          ]
        }
      ],
      {
        type: "object",
        additionalProperties: false,
        required: ["actionNeeded", "urgency", "relevance", "noise", "summary", "reasonTags"],
        properties: {
          actionNeeded: { type: "boolean" },
          urgency: { type: "string", enum: ["low", "medium", "high"] },
          relevance: { type: "string", enum: ["low", "medium", "high"] },
          noise: { type: "string", enum: ["low", "medium", "high"] },
          summary: { type: "string" },
          reasonTags: {
            type: "array",
            items: { type: "string", enum: reasonTagEnum },
            minItems: 0,
            maxItems: 6
          }
        }
      },
      "candidate_content_analysis"
    );
  } catch {
    return null;
  }
}

async function callResponsesApi<T>(input: unknown, schema: Record<string, unknown>, name: string) {
  if (!env.openAiApiKey) return null;

  const response = await fetch(`${env.openAiApiBaseUrl}/v1/responses`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${env.openAiApiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: env.openAiModel,
      input,
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
}

export async function evaluateCandidateWithPersonalization(input: {
  candidate: CandidateInput;
  profile: UserPriorityProfile;
  memory: PreferenceMemorySnapshot;
  recentExamples: Array<{ title: string; source: string; outcome: string; reason: string | null }>;
}) {
  const contentAnalysis = await analyzeCandidateContent(input.candidate, input.profile);
  const heuristic = scoreWithProfile(input.candidate, input.profile, input.memory, contentAnalysis).evaluation;
  if (!input.profile.personalizationEnabled) return heuristic;

  const signals = computeContentSignals(input.candidate);
  const tags = uniqueReasonTags([...signals.tags, ...(contentAnalysis?.reasonTags ?? [])]);
  const neverHideMatches = listMatches(
    normalized(`${input.candidate.title} ${input.candidate.bodyPreview ?? ""}`),
    input.profile.mustNotMiss ?? []
  );
  const hardAccept =
    input.candidate.isAssignedToUser ||
    input.candidate.isDirectRequest ||
    input.candidate.dueSoon ||
    tags.includes("blocking_work") ||
    tags.includes("review_request") ||
    tags.includes("historically_accepted") ||
    neverHideMatches.length > 0;
  const hardReject =
    input.candidate.isDuplicate ||
    (tags.includes("newsletter_like") && !signals.hasStrongUserSignal) ||
    (tags.includes("bot_generated") && !signals.hasStrongUserSignal && contentAnalysis?.actionNeeded !== true);

  if (hardAccept) {
    return {
      relevance: "accept" as const,
      priority: heuristic.priority === "Low" ? "Medium" : heuristic.priority,
      confidence: 0.98,
      why: heuristic.why,
      reasonTags: uniqueReasonTags([
        ...heuristic.reasonTags,
        input.candidate.isAssignedToUser ? "assigned_work" : null
      ])
    };
  }

  if (hardReject) {
    return {
      relevance: "reject" as const,
      priority: "Low" as TaskPriority,
      confidence: 0.98,
      why: heuristic.why,
      reasonTags: heuristic.reasonTags
      };
  }

  return heuristic;
}

export async function synthesizePriorityProfile(input: {
  roleFocus: string;
  prioritizationPrompt: string;
  importantWork: string[];
  noiseWork: string[];
  mustNotMiss: string[];
  importantPeople: string[];
  importantProjects: string[];
  filteringStyle: UserPriorityProfile["filteringStyle"];
  priorityBias: UserPriorityProfile["priorityBias"];
  exampleRankings: Array<{ title: string; source: TaskSource; decision: "show_today" | "keep_low" | "reject_noise" }>;
}) {
  const fallback: Partial<UserPriorityProfile> = {
    roleFocus: input.roleFocus,
    prioritizationPrompt: input.prioritizationPrompt || defaultPriorityProfile.prioritizationPrompt,
    importantWork: input.importantWork,
    noiseWork: input.noiseWork,
    mustNotMiss: input.mustNotMiss,
    importantPeople: input.importantPeople,
    importantProjects: input.importantProjects,
    filteringStyle: input.filteringStyle,
    priorityBias: input.priorityBias,
    importantSources: [...new Set(input.exampleRankings.filter((item) => item.decision !== "reject_noise").map((item) => item.source))],
    positiveReasonTags: ["direct_request", "assigned_work", "due_soon"],
    negativeReasonTags: ["newsletter_like", "comment_noise", "bot_generated"],
    lastProfileRefreshAt: new Date().toISOString()
  };

  try {
    const parsed = await callResponsesApi<Partial<UserPriorityProfile>>(
      [
        {
          role: "system",
          content: [
            {
              type: "input_text",
              text:
                "Convert calibration answers into a compact structured work-priority profile. Return JSON only and keep values practical, not verbose."
            }
          ]
        },
        {
          role: "user",
          content: [{ type: "input_text", text: JSON.stringify(input) }]
        }
      ],
      {
        type: "object",
        additionalProperties: false,
        required: [
          "roleFocus",
          "importantWork",
          "noiseWork",
          "mustNotMiss",
          "importantSources",
          "importantPeople",
          "importantProjects",
          "positiveReasonTags",
          "negativeReasonTags",
          "filteringStyle",
          "priorityBias"
        ],
        properties: {
          roleFocus: { type: "string" },
          prioritizationPrompt: { type: "string" },
          importantWork: { type: "array", items: { type: "string" } },
          noiseWork: { type: "array", items: { type: "string" } },
          mustNotMiss: { type: "array", items: { type: "string" } },
          importantSources: { type: "array", items: { type: "string" } },
          importantPeople: { type: "array", items: { type: "string" } },
          importantProjects: { type: "array", items: { type: "string" } },
          positiveReasonTags: { type: "array", items: { type: "string", enum: reasonTagEnum } },
          negativeReasonTags: { type: "array", items: { type: "string", enum: reasonTagEnum } },
          filteringStyle: { type: "string", enum: ["conservative", "balanced", "aggressive"] },
          priorityBias: { type: "string", enum: ["focus", "balanced", "coverage"] }
        }
      },
      "priority_profile"
    );

    return {
      ...fallback,
      ...parsed,
      lastProfileRefreshAt: new Date().toISOString()
    };
  } catch {
    return fallback;
  }
}

export async function analyzeFeedbackReason(input: {
  action: FeedbackAction;
  taskTitle: string;
  source: TaskSource | "Calibration";
  beforePriority?: TaskPriority | null;
  afterPriority?: TaskPriority | null;
  decisionReason?: string | null;
  decisionTags?: ReasonTag[];
  context?: string | null;
}): Promise<FeedbackReasonResult> {
  const fallback: FeedbackReasonResult = (() => {
    if (input.action === "reject" || input.action === "always_ignore_similar") {
      return {
        likelyReason: "User treated this as low-signal work.",
        reasonTag: "historically_rejected" as ReasonTag,
        positiveOrNegativePreference: "negative" as FeedbackPolarity,
        confidence: 0.72
      };
    }
    if (input.action === "restore" || input.action === "should_have_been_included") {
      return {
        likelyReason: "User felt this task should have stayed visible.",
        reasonTag: "historically_accepted" as ReasonTag,
        positiveOrNegativePreference: "positive" as FeedbackPolarity,
        confidence: 0.78
      };
    }
    if (input.action === "priority_changed" && input.beforePriority !== input.afterPriority) {
      return {
        likelyReason: "User manually corrected the priority ordering.",
        reasonTag: input.afterPriority === "High" ? "historically_accepted" : "historically_rejected",
        positiveOrNegativePreference: input.afterPriority === "High" ? "positive" : "negative",
        confidence: 0.7
      };
    }
    return {
      likelyReason: "User interaction updated the planner's understanding of importance.",
      reasonTag: null,
      positiveOrNegativePreference: "neutral",
      confidence: 0.55
    };
  })();

  try {
    const parsed = await callResponsesApi<FeedbackReasonResult>(
      [
        {
          role: "system",
          content: [
            {
              type: "input_text",
              text:
                "Infer the most likely reason behind a user's task feedback action. Return compact JSON only. Use neutral if unsure."
            }
          ]
        },
        { role: "user", content: [{ type: "input_text", text: JSON.stringify(input) }] }
      ],
      {
        type: "object",
        additionalProperties: false,
        required: ["likelyReason", "reasonTag", "positiveOrNegativePreference", "confidence"],
        properties: {
          likelyReason: { type: "string" },
          reasonTag: { anyOf: [{ type: "string", enum: reasonTagEnum }, { type: "null" }] },
          positiveOrNegativePreference: { type: "string", enum: ["positive", "negative", "neutral"] },
          confidence: { type: "number", minimum: 0, maximum: 1 }
        }
      },
      "feedback_reason"
    );
    return parsed
      ? {
          ...parsed,
          reasonTag: (parsed.reasonTag as ReasonTag | null) ?? null
        }
      : fallback;
  } catch {
    return fallback;
  }
}

export async function distillPreferenceMemory(input: {
  profile: UserPriorityProfile;
  recentEvents: Array<{
    action: string;
    title: string;
    source: string;
    inferredReason: string | null;
    inferredReasonTag: string | null;
    preferencePolarity: string;
  }>;
  sourceEventCount: number;
}): Promise<{ snapshot: PreferenceMemorySnapshot; insights: PersonalizationInsight[] }> {
  const positiveCounts = new Map<ReasonTag, number>();
  const negativeCounts = new Map<ReasonTag, number>();

  for (const event of input.recentEvents) {
    const tag = event.inferredReasonTag as ReasonTag | null;
    if (!tag) continue;
    if (event.preferencePolarity === "positive") {
      positiveCounts.set(tag, (positiveCounts.get(tag) ?? 0) + 1);
    } else if (event.preferencePolarity === "negative") {
      negativeCounts.set(tag, (negativeCounts.get(tag) ?? 0) + 1);
    }
  }

  const positiveTags = [...positiveCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([tag]) => tag);
  const negativeTags = [...negativeCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([tag]) => tag);

  const snapshot: PreferenceMemorySnapshot = {
    version: Math.max(1, input.sourceEventCount),
    positiveTags,
    negativeTags,
    repeatedWins: input.recentEvents
      .filter((event) => event.preferencePolarity === "positive")
      .map((event) => event.title)
      .slice(0, 6),
    repeatedNoise: input.recentEvents
      .filter((event) => event.preferencePolarity === "negative")
      .map((event) => event.title)
      .slice(0, 6)
  };

  const insights: PersonalizationInsight[] = [
    ...positiveTags.map((tag) => ({
      statement: `You often keep work tagged ${tag.replace(/_/g, " ")} in your active plan.`,
      confidence: Math.min(0.95, 0.55 + (positiveCounts.get(tag) ?? 0) * 0.08),
      source: "history" as const
    })),
    ...negativeTags.map((tag) => ({
      statement: `You often reject work tagged ${tag.replace(/_/g, " ")}.`,
      confidence: Math.min(0.95, 0.55 + (negativeCounts.get(tag) ?? 0) * 0.08),
      source: "history" as const
    }))
  ].slice(0, 6);

  return { snapshot, insights };
}
