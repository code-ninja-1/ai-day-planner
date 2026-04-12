import cors from "cors";
import express from "express";
import { z } from "zod";
import { acquireGraphTokenOnBehalfOf, getOptionalMicrosoftSession } from "./auth/microsoftAuth.js";
import {
  createManualTask,
  deleteIntegrationConnection,
  deleteTask,
  getAutomationSettings,
  getReminderById,
  getTaskById,
  getSyncState,
  listIntegrationConnections,
  listMeetings,
  listReminderItems,
  listTasks,
  listDeferredTasks,
  normalizeTask,
  saveAutomationSettings,
  saveIntegrationConnection,
  updateTask,
  updateReminder,
  groupTasksByPriority
} from "./db.js";
import { env } from "./env.js";
import { fetchJiraIssueDetail, normalizeJiraBaseUrl, validateJiraCredentials } from "./providers/jira.js";
import {
  exchangeMicrosoftCode,
  fetchEmailDetailWithAccessToken,
  fetchMicrosoftProfileWithAccessToken,
  getMicrosoftAuthUrl
} from "./providers/microsoft.js";
import {
  generatePlan,
  getDeferredTasksPayload,
  getReminderCenterPayload,
  getTodaySnapshot,
  syncMeetingsOnly,
  syncTasksOnly
} from "./services/planService.js";
import { scheduleAutomation } from "./services/scheduler.js";

const app = express();
app.use(cors({ origin: env.appOrigin, credentials: true }));
app.use(express.json());

const taskCreateSchema = z.object({
  title: z.string().min(1),
  priority: z.enum(["High", "Medium", "Low"]).optional(),
  status: z.enum(["Not Started", "In Progress", "Completed"]).optional()
});

const taskUpdateSchema = taskCreateSchema
  .partial()
  .extend({ deferredUntil: z.string().datetime().nullable().optional() });

const reminderUpdateSchema = z.object({
  status: z.enum(["active", "dismissed", "resolved"]).optional(),
  reason: z.string().min(1).optional(),
  scheduledFor: z.string().datetime().nullable().optional(),
  throttleUntil: z.string().datetime().nullable().optional()
});

const automationSettingsSchema = z.object({
  scheduleEnabled: z.boolean().optional(),
  scheduleTimeLocal: z.string().regex(/^\d{2}:\d{2}$/).optional(),
  scheduleTimezone: z.string().min(1).optional(),
  remindersEnabled: z.boolean().optional(),
  reminderCadenceHours: z.number().int().min(1).max(72).optional(),
  desktopNotificationsEnabled: z.boolean().optional()
});

app.get("/api/health", (_req, res) => {
  res.json({ ok: true });
});

app.get("/api/today", (_req, res) => {
  res.json(getTodaySnapshot());
});

app.post("/api/plan/generate", async (req, res) => {
  try {
    let microsoftGraphAccessToken: string | null = null;
    let microsoftWarning: string | null = null;
    const preferredTimeZone =
      typeof req.body?.timeZone === "string" && req.body.timeZone.trim() ? req.body.timeZone.trim() : null;

    try {
      const session = await getOptionalMicrosoftSession(req);
      if (session) {
        microsoftGraphAccessToken = await acquireGraphTokenOnBehalfOf(session);
        saveIntegrationConnection({
          provider: "microsoft",
          status: "connected",
          accountLabel: session.accountLabel ?? session.displayName,
          configJson: JSON.stringify({ mode: "msal-obo", oid: session.oid }),
          accessToken: null,
          refreshToken: null,
          expiresAt: null,
          errorMessage: null
        });
      } else {
        microsoftWarning = "Microsoft is not connected for this browser session.";
      }
    } catch (error) {
      microsoftWarning =
        error instanceof Error
          ? `Microsoft session is unavailable: ${error.message}`
          : "Microsoft session is unavailable.";
    }

    const payload = await generatePlan({
      microsoftGraphAccessToken,
      microsoftWarning,
      preferredTimeZone
    }, "manual");
    res.json(payload);
  } catch (error) {
    res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to generate plan"
    });
  }
});

app.post("/api/plan/generate-now", async (req, res) => {
  try {
    let microsoftGraphAccessToken: string | null = null;
    let microsoftWarning: string | null = null;
    const preferredTimeZone =
      typeof req.body?.timeZone === "string" && req.body.timeZone.trim() ? req.body.timeZone.trim() : null;

    try {
      const session = await getOptionalMicrosoftSession(req);
      if (session) {
        microsoftGraphAccessToken = await acquireGraphTokenOnBehalfOf(session);
      } else {
        microsoftWarning = "Microsoft is not connected for this browser session.";
      }
    } catch (error) {
      microsoftWarning =
        error instanceof Error
          ? `Microsoft session is unavailable: ${error.message}`
          : "Microsoft session is unavailable.";
    }

    const payload = await generatePlan(
      {
        microsoftGraphAccessToken,
        microsoftWarning,
        preferredTimeZone
      },
      "manual"
    );
    res.json(payload);
  } catch (error) {
    res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to generate plan"
    });
  }
});

app.post("/api/sync/meetings", async (req, res) => {
  try {
    let microsoftGraphAccessToken: string | null = null;
    let microsoftWarning: string | null = null;
    const preferredTimeZone =
      typeof req.body?.timeZone === "string" && req.body.timeZone.trim() ? req.body.timeZone.trim() : null;

    try {
      const session = await getOptionalMicrosoftSession(req);
      if (session) {
        microsoftGraphAccessToken = await acquireGraphTokenOnBehalfOf(session);
      } else {
        microsoftWarning = "Microsoft is not connected for this browser session.";
      }
    } catch (error) {
      microsoftWarning =
        error instanceof Error
          ? `Microsoft session is unavailable: ${error.message}`
          : "Microsoft session is unavailable.";
    }

    res.json(
      await syncMeetingsOnly({
        microsoftGraphAccessToken,
        microsoftWarning,
        preferredTimeZone
      })
    );
  } catch (error) {
    res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to sync meetings"
    });
  }
});

app.post("/api/sync/tasks", async (req, res) => {
  try {
    let microsoftGraphAccessToken: string | null = null;
    let microsoftWarning: string | null = null;
    const preferredTimeZone =
      typeof req.body?.timeZone === "string" && req.body.timeZone.trim() ? req.body.timeZone.trim() : null;

    try {
      const session = await getOptionalMicrosoftSession(req);
      if (session) {
        microsoftGraphAccessToken = await acquireGraphTokenOnBehalfOf(session);
      } else {
        microsoftWarning = "Microsoft is not connected for this browser session.";
      }
    } catch (error) {
      microsoftWarning =
        error instanceof Error
          ? `Microsoft session is unavailable: ${error.message}`
          : "Microsoft session is unavailable.";
    }

    res.json(
      await syncTasksOnly({
        microsoftGraphAccessToken,
        microsoftWarning,
        preferredTimeZone
      })
    );
  } catch (error) {
    res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to sync tasks"
    });
  }
});

app.get("/api/tasks", (req, res) => {
  const status = req.query.status as "Not Started" | "In Progress" | "Completed" | undefined;
  res.json({ tasks: listTasks(status) });
});

app.get("/api/tasks/deferred", (_req, res) => {
  res.json(getDeferredTasksPayload());
});

app.get("/api/tasks/:id/details", async (req, res) => {
  const task = getTaskById(Number(req.params.id));
  if (!task) {
    return res.status(404).json({ message: "Task not found" });
  }

  if (task.source === "Manual") {
    return res.status(400).json({ message: "Manual tasks do not have source details" });
  }

  try {
    if (task.source === "Jira") {
      if (!task.sourceRef) {
        return res.status(400).json({ message: "Jira task is missing source reference" });
      }
      return res.json({ detail: await fetchJiraIssueDetail(task.sourceRef) });
    }

    if (!task.sourceRef) {
      return res.status(400).json({ message: "Email task is missing source reference" });
    }

    const session = await getOptionalMicrosoftSession(req);
    if (!session) {
      return res.status(401).json({ message: "Microsoft is not connected for this browser session." });
    }

    const graphToken = await acquireGraphTokenOnBehalfOf(session);
    return res.json({
      detail: await fetchEmailDetailWithAccessToken(task.sourceRef, task.sourceThreadRef, graphToken)
    });
  } catch (error) {
    return res.status(500).json({
      message: error instanceof Error ? error.message : "Failed to fetch task details"
    });
  }
});

app.post("/api/tasks", (req, res) => {
  const parsed = taskCreateSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid task payload" });
  }
  const row = createManualTask(parsed.data);
  return res.status(201).json({ task: normalizeTask(row as Record<string, unknown>) });
});

app.patch("/api/tasks/:id", (req, res) => {
  const parsed = taskUpdateSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid task payload" });
  }
  const row = updateTask(Number(req.params.id), parsed.data);
  if (!row) {
    return res.status(404).json({ message: "Task not found" });
  }
  return res.json({ task: normalizeTask(row as Record<string, unknown>) });
});

app.patch("/api/tasks/:id/defer", (req, res) => {
  const schema = z.object({
    deferredUntil: z.string().datetime().nullable()
  });
  const parsed = schema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid defer payload" });
  }
  const row = updateTask(Number(req.params.id), {
    deferredUntil: parsed.data.deferredUntil,
    manualOverrideFlags: ["deferredUntil"]
  });
  if (!row) {
    return res.status(404).json({ message: "Task not found" });
  }
  return res.json({ task: normalizeTask(row as Record<string, unknown>) });
});

app.delete("/api/tasks/:id", (req, res) => {
  const deleted = deleteTask(Number(req.params.id));
  if (!deleted) {
    return res.status(404).json({ message: "Task not found" });
  }
  return res.status(204).send();
});

app.get("/api/settings/integrations", async (req, res) => {
  const rows = listIntegrationConnections().map((row) => ({
    provider: row.provider,
    status: row.status,
    accountLabel: row.account_label,
    errorMessage: row.error_message,
    updatedAt: row.updated_at,
    lastSyncAt: getSyncState(String(row.provider)),
    config:
      row.provider === "jira" && row.config_json
        ? (() => {
            try {
              const parsed = JSON.parse(String(row.config_json)) as {
                baseUrl?: string;
                email?: string;
                apiToken?: string;
              };
              return {
                baseUrl: parsed.baseUrl ?? "",
                email: parsed.email ?? "",
                apiToken: parsed.apiToken ?? ""
              };
            } catch {
              return null;
            }
          })()
        : null
  }));

  const microsoft = rows.find((row) => row.provider === "microsoft") ?? {
    provider: "microsoft",
    status: "disconnected",
    accountLabel: null,
    errorMessage: null,
    updatedAt: null,
    lastSyncAt: null,
    config: null
  };
  const jira = rows.find((row) => row.provider === "jira") ?? {
    provider: "jira",
    status: "disconnected",
    accountLabel: null,
    errorMessage: null,
    updatedAt: null,
    lastSyncAt: null,
    config: null
  };

  try {
    const session = await getOptionalMicrosoftSession(req);
    if (session) {
      const graphToken = await acquireGraphTokenOnBehalfOf(session);
      const profile = await fetchMicrosoftProfileWithAccessToken(graphToken);
      microsoft.status = "connected";
      microsoft.accountLabel =
        profile.userPrincipalName ?? profile.displayName ?? session.accountLabel ?? session.displayName;
      microsoft.errorMessage = null;
    }
  } catch (error) {
    microsoft.status = "error";
    microsoft.errorMessage =
      error instanceof Error
        ? `Microsoft session is unavailable: ${error.message}`
        : "Microsoft session is unavailable";
  }

  res.json({ integrations: { microsoft, jira } });
});

app.get("/api/settings/automation", (_req, res) => {
  res.json({
    automation: getAutomationSettings(),
    reminders: listReminderItems(["active", "dismissed", "resolved"])
  });
});

app.patch("/api/settings/schedule", (req, res) => {
  const parsed = automationSettingsSchema.pick({
    scheduleEnabled: true,
    scheduleTimeLocal: true,
    scheduleTimezone: true
  }).safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid schedule settings" });
  }
  const automation = saveAutomationSettings(parsed.data);
  scheduleAutomation();
  return res.json({ automation });
});

app.patch("/api/settings/reminders", (req, res) => {
  const parsed = automationSettingsSchema.pick({
    remindersEnabled: true,
    reminderCadenceHours: true,
    desktopNotificationsEnabled: true
  }).safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid reminder settings" });
  }
  const automation = saveAutomationSettings(parsed.data);
  return res.json({ automation });
});

app.get("/api/reminders", (_req, res) => {
  res.json(getReminderCenterPayload());
});

app.patch("/api/reminders/:id", (req, res) => {
  const parsed = reminderUpdateSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid reminder payload" });
  }
  const reminder = getReminderById(Number(req.params.id));
  if (!reminder) {
    return res.status(404).json({ message: "Reminder not found" });
  }
  const nextStatus = parsed.data.status ?? reminder.status;
  const updated = updateReminder(reminder.id, {
    ...parsed.data,
    status: nextStatus,
    dismissedAt: nextStatus === "dismissed" ? new Date().toISOString() : null
  });
  return res.json({ reminder: updated });
});

app.post("/api/settings/integrations/jira", async (req, res) => {
  const schema = z.object({
    baseUrl: z.string().url(),
    email: z.string().email(),
    apiToken: z.string().min(1)
  });
  const parsed = schema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ message: "Invalid Jira settings" });
  }

  try {
    const normalizedInput = {
      ...parsed.data,
      baseUrl: normalizeJiraBaseUrl(parsed.data.baseUrl),
      email: parsed.data.email.trim(),
      apiToken: parsed.data.apiToken.trim()
    };
    const validation = await validateJiraCredentials(normalizedInput);
    saveIntegrationConnection({
      provider: "jira",
      status: "connected",
      accountLabel:
        validation.profile.emailAddress ?? validation.profile.displayName ?? normalizedInput.email,
      configJson: JSON.stringify({
        ...normalizedInput,
        authType: validation.authType
      }),
      accessToken: null,
      refreshToken: null,
      expiresAt: null,
      errorMessage: null
    });
    return res.status(201).json({ ok: true });
  } catch (error) {
    saveIntegrationConnection({
      provider: "jira",
      status: "error",
      accountLabel: parsed.data.email.trim(),
      configJson: JSON.stringify({
        ...parsed.data,
        baseUrl: (() => {
          try {
            return normalizeJiraBaseUrl(parsed.data.baseUrl);
          } catch {
            return parsed.data.baseUrl.trim();
          }
        })()
      }),
      accessToken: null,
      refreshToken: null,
      expiresAt: null,
      errorMessage: error instanceof Error ? error.message : "Jira validation failed"
    });
    return res.status(400).json({
      message: error instanceof Error ? error.message : "Jira validation failed"
    });
  }
});

app.delete("/api/settings/integrations/:provider", (req, res) => {
  const provider = req.params.provider;
  if (provider !== "microsoft" && provider !== "jira") {
    return res.status(400).json({ message: "Unsupported integration provider" });
  }

  deleteIntegrationConnection(provider);
  return res.status(204).send();
});

app.get("/api/auth/microsoft/start", (_req, res) => {
  if (!env.microsoftClientId || !env.microsoftClientSecret) {
    return res.status(400).json({ message: "Microsoft OAuth is not configured in apps/api/.env" });
  }
  res.json({ url: getMicrosoftAuthUrl() });
});

app.get("/api/auth/microsoft/callback", async (req, res) => {
  const code = String(req.query.code ?? "");
  if (!code) {
    return res.status(400).send("Missing code");
  }
  try {
    await exchangeMicrosoftCode(code);
    return res.redirect(`${env.appOrigin}/settings?connected=microsoft`);
  } catch (error) {
    return res
      .status(400)
      .send(error instanceof Error ? error.message : "Microsoft connection failed");
  }
});

app.listen(env.port, () => {
  console.log(`API listening on http://localhost:${env.port}`);
});

scheduleAutomation();
