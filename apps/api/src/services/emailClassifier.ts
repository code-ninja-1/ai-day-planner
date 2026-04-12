import { env } from "../env.js";
import type { TaskPriority } from "../types.js";
import type { GraphMail } from "../providers/microsoft.js";

export interface EmailClassification {
  actionable: boolean;
  priority: TaskPriority;
  title: string;
}

function heuristicClassification(email: GraphMail): EmailClassification {
  const text = `${email.subject ?? ""} ${email.bodyPreview ?? ""}`.toLowerCase();
  const actionable =
    /(please|action|required|follow up|review|approve|can you|need you|todo|by eod|by tomorrow)/.test(
      text
    );
  const priority: TaskPriority =
    /(urgent|asap|today|blocking|immediately)/.test(text)
      ? "High"
      : actionable
        ? "Medium"
        : "Low";

  return {
    actionable,
    priority,
    title: (email.subject || "Email follow-up").trim()
  };
}

export async function classifyEmail(email: GraphMail): Promise<EmailClassification> {
  if (!env.openAiApiKey) {
    return heuristicClassification(email);
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
                  "Classify whether an email should become a work task. Return compact JSON with actionable(boolean), priority(High|Medium|Low), and title(string). Prefer false for newsletters or FYIs."
              }
            ]
          },
          {
            role: "user",
            content: [
              {
                type: "input_text",
                text: JSON.stringify({
                  subject: email.subject,
                  preview: email.bodyPreview,
                  from: email.from?.emailAddress?.address
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
              required: ["actionable", "priority", "title"],
              properties: {
                actionable: { type: "boolean" },
                priority: { type: "string", enum: ["High", "Medium", "Low"] },
                title: { type: "string" }
              }
            }
          }
        }
      })
    });

    if (!response.ok) {
      return heuristicClassification(email);
    }

    const json = (await response.json()) as {
      output_text?: string;
    };

    if (!json.output_text) {
      return heuristicClassification(email);
    }

    const parsed = JSON.parse(json.output_text) as EmailClassification;
    return {
      actionable: Boolean(parsed.actionable),
      priority: parsed.priority,
      title: parsed.title?.trim() || email.subject || "Email follow-up"
    };
  } catch {
    return heuristicClassification(email);
  }
}

