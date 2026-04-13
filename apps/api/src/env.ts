import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, "../.env") });

function resolveFromApiRoot(value: string) {
  return path.isAbsolute(value) ? value : path.resolve(__dirname, "..", value);
}

const defaultDatabasePath = process.env.DATABASE_PATH
  ? resolveFromApiRoot(process.env.DATABASE_PATH)
  : path.resolve(__dirname, "../data/ai-day-planner.db");

const defaultLogDirectory = process.env.LOG_DIRECTORY
  ? resolveFromApiRoot(process.env.LOG_DIRECTORY)
  : path.resolve(path.dirname(defaultDatabasePath), "logs");

export const env = {
  port: Number(process.env.PORT ?? 4000),
  appOrigin: process.env.APP_ORIGIN ?? "http://localhost:5173",
  databasePath: defaultDatabasePath,
  microsoftTenantId: process.env.MICROSOFT_TENANT_ID ?? "common",
  microsoftClientId: process.env.MICROSOFT_CLIENT_ID ?? "",
  microsoftClientSecret: process.env.MICROSOFT_CLIENT_SECRET ?? "",
  microsoftApiAudience:
    process.env.MICROSOFT_API_AUDIENCE ??
    (process.env.MICROSOFT_CLIENT_ID ? `api://${process.env.MICROSOFT_CLIENT_ID}` : ""),
  microsoftRedirectUri:
    process.env.MICROSOFT_REDIRECT_URI ??
    "http://localhost:4000/api/auth/microsoft/callback",
  openAiApiKey: process.env.OPENAI_API_KEY ?? "",
  openAiApiBaseUrl: process.env.OPENAI_API_BASE_URL ?? "https://api.openai.com",
  openAiModel: process.env.OPENAI_MODEL ?? "gpt-4.1-mini",
  jiraAllowSelfSignedTls: process.env.JIRA_ALLOW_SELF_SIGNED_TLS === "true",
  logLevel: process.env.LOG_LEVEL ?? "info",
  logDirectory: defaultLogDirectory,
  enableSensitiveDebugLogs: process.env.ENABLE_SENSITIVE_DEBUG_LOGS === "true",
  logRetentionDays: Number(process.env.LOG_RETENTION_DAYS ?? 14)
};

export const microsoftAuthority = `https://login.microsoftonline.com/${env.microsoftTenantId}`;
export const microsoftIssuer = `${microsoftAuthority}/v2.0`;
