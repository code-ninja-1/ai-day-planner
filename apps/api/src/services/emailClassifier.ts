import { env } from "../env.js";
import type { ReasonTag, TaskEffortBucket, TaskPriority, UserPriorityProfile } from "../types.js";
import type { GraphMail } from "../providers/microsoft.js";

export interface EmailClassification {
  actionable: boolean;
  priority: TaskPriority;
  estimatedEffortBucket: TaskEffortBucket;
  title: string;
  why: string;
  reasonTags: ReasonTag[];
}

interface EmailClassificationContext {
  profile: UserPriorityProfile;
  recentExamples?: Array<{ title: string; source: string; outcome: string; reason: string | null }>;
  recentRejectedExamples?: Array<{ title: string; source: string; outcome: string; reason: string | null }>;
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

function stripHtml(value?: string | null) {
  return (value ?? "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, " ")
    .trim();
}

function emailContentForModel(email: GraphMail) {
  const combined = [email.subject ?? "", email.bodyPreview ?? "", stripHtml(email.body?.content ?? "")]
    .filter(Boolean)
    .join("\n\n");
  return combined.slice(0, 12_000);
}

function heuristicClassification(email: GraphMail, context?: EmailClassificationContext): EmailClassification {
  const text = emailContentForModel(email).toLowerCase();
  const sender = (email.from?.emailAddress?.address ?? "").toLowerCase();
  const commentOrThreadSignal =
    /(comment(ed)?|mentioned you|requested review|requested changes|assigned to you|reply needed|your input|needs your review|tagged you)/.test(
      text
    );
  const importantPeople = context?.profile.importantPeople?.map((value) => value.toLowerCase()) ?? [];
  const importantProjects = context?.profile.importantProjects?.map((value) => value.toLowerCase()) ?? [];
  const mustNotMiss = context?.profile.mustNotMiss?.map((value) => value.toLowerCase()) ?? [];
  const senderPreferred = importantPeople.some((value) => value.length >= 2 && sender.includes(value));
  const projectPreferred = importantProjects.some((value) => value.length >= 2 && text.includes(value));
  const mustNotMissMatch = mustNotMiss.some((value) => value.length >= 2 && text.includes(value));
  const actionable =
    /(please|action|required|follow up|review|approve|can you|need you|todo|by eod|by tomorrow)/.test(
      text
    ) || commentOrThreadSignal || senderPreferred || projectPreferred || mustNotMissMatch;
  const priority: TaskPriority =
    /(urgent|asap|today|blocking|immediately)/.test(text)
      ? "High"
      : actionable && (senderPreferred || projectPreferred || mustNotMissMatch || commentOrThreadSignal)
        ? "Medium"
        : "Low";
  const estimatedEffortBucket: TaskEffortBucket =
    !actionable
      ? "15 min"
      : /(investigate|analysis|write up|draft|prepare|prep|review changes|multiple items)/.test(text)
        ? "1 hour"
        : /(reply|approve|review|comment|follow up|triage|check)/.test(text)
          ? "30 min"
          : "15 min";
  const reasonTags: ReasonTag[] = [
    ...(commentOrThreadSignal ? (["review_request"] as ReasonTag[]) : []),
    ...(/(newsletter|digest|announcement|roundup)/.test(text) ? (["newsletter_like"] as ReasonTag[]) : []),
    ...(/(fyi|for your information|no action required)/.test(text) ? (["fyi_only"] as ReasonTag[]) : []),
    ...(/(urgent|asap|today|blocking|immediately|eod)/.test(text) ? (["due_soon"] as ReasonTag[]) : []),
    ...(/(please|can you|need you|action required|approve|reply needed)/.test(text)
      ? (["direct_request"] as ReasonTag[])
      : []),
    ...(senderPreferred || projectPreferred || mustNotMissMatch ? (["historically_accepted"] as ReasonTag[]) : [])
  ];

  return {
    actionable,
    priority,
    estimatedEffortBucket,
    title: (email.subject || "Email follow-up").trim(),
    why: actionable
      ? "Email likely needs your response, review, or follow-through."
      : "Email looks informational and does not appear to require action from you.",
    reasonTags
  };
}

export async function classifyEmail(
  email: GraphMail,
  context?: EmailClassificationContext
): Promise<EmailClassification> {
  if (!env.openAiApiKey) {
    return heuristicClassification(email, context);
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
            content: [
              {
                type: "input_text",
                text:
                  "Decide whether this email truly needs to be addressed by the user. Use the full email subject and content, the user's saved preferences, recent accepted examples, and especially recent rejected examples. Be conservative. actionable=true only when the user clearly needs to reply, review, approve, decide, prepare, follow up, or do meaningful work because of this email. actionable=false for meetings, calendar invites, announcements, newsletters, digests, passive notifications, automated alerts, thread noise, status updates, and FYI emails unless they contain a direct and important ask for the user. Only mark priority High when the email is important enough to deserve becoming a planner task. Medium and Low usually mean it should stay out of the task list. Return a concise task title, priority, estimated effort, a short why explanation, and reason tags."
              }
            ]
          },
          {
            role: "user",
            content: [
              {
                type: "input_text",
                text: JSON.stringify({
                  email: {
                    subject: email.subject,
                    from: email.from?.emailAddress?.address,
                    to: (email.toRecipients ?? []).map((recipient) => recipient.emailAddress?.address).filter(Boolean),
                    cc: (email.ccRecipients ?? []).map((recipient) => recipient.emailAddress?.address).filter(Boolean),
                    bodyPreview: email.bodyPreview,
                    content: emailContentForModel(email)
                  },
                  userPreferences: context
                    ? {
                        roleFocus: context.profile.roleFocus,
                        prioritizationPrompt: context.profile.prioritizationPrompt,
                        importantWork: context.profile.importantWork,
                        noiseWork: context.profile.noiseWork,
                        mustNotMiss: context.profile.mustNotMiss,
                        importantSources: context.profile.importantSources,
                        importantPeople: context.profile.importantPeople,
                        importantProjects: context.profile.importantProjects,
                        positiveReasonTags: context.profile.positiveReasonTags,
                        negativeReasonTags: context.profile.negativeReasonTags,
                        filteringStyle: context.profile.filteringStyle,
                        priorityBias: context.profile.priorityBias
                      }
                    : null,
                  recentExamples: (context?.recentExamples ?? []).slice(0, 10),
                  recentRejectedExamples: (context?.recentRejectedExamples ?? []).slice(0, 15)
                })
              }
            ]
          }
        ],
        text: {
          format: {
            type: "json_schema",
            name: "email_classification",
            schema: {
              type: "object",
              additionalProperties: false,
              required: ["actionable", "priority", "estimatedEffortBucket", "title", "why", "reasonTags"],
              properties: {
                actionable: { type: "boolean" },
                priority: { type: "string", enum: ["High", "Medium", "Low"] },
                estimatedEffortBucket: { type: "string", enum: ["15 min", "30 min", "1 hour", "2+ hours"] },
                title: { type: "string" },
                why: { type: "string" },
                reasonTags: {
                  type: "array",
                  items: { type: "string", enum: reasonTagEnum },
                  minItems: 0,
                  maxItems: 6
                }
              }
            }
          }
        }
      })
    });

    if (!response.ok) {
      return heuristicClassification(email, context);
    }

    const json = (await response.json()) as {
      output_text?: string;
    };

    if (!json.output_text) {
      return heuristicClassification(email, context);
    }

    const parsed = JSON.parse(json.output_text) as EmailClassification;
    return {
      actionable: Boolean(parsed.actionable),
      priority: parsed.priority,
      estimatedEffortBucket: parsed.estimatedEffortBucket,
      title: parsed.title?.trim() || email.subject || "Email follow-up",
      why: parsed.why?.trim() || "Email was evaluated for actionability.",
      reasonTags: Array.isArray(parsed.reasonTags) ? parsed.reasonTags : []
    };
  } catch {
    return heuristicClassification(email, context);
  }
}
