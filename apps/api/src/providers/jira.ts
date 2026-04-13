import { env } from "../env.js";
import { getIntegrationConnection } from "../db.js";

export interface JiraIssue {
  id: string;
  key: string;
  self: string;
  fields: {
    summary: string;
    status?: { name?: string };
    priority?: { name?: string };
    updated?: string;
    issuetype?: { name?: string; subtask?: boolean };
    timeoriginalestimate?: number;
    timetracking?: { originalEstimateSeconds?: number };
  };
}

interface JiraIssueDetailResponse {
  key: string;
  fields: Record<string, unknown> & {
    summary?: string;
    status?: { name?: string };
    priority?: { name?: string };
    description?: unknown;
    assignee?: { displayName?: string };
    reporter?: { displayName?: string };
    labels?: string[];
    subtasks?: Array<{
      key: string;
      fields?: {
        summary?: string;
        status?: { name?: string };
      };
    }>;
    comment?: {
      comments?: Array<{
        author?: { displayName?: string };
        created?: string;
        body?: unknown;
      }>;
    };
    worklog?: {
      worklogs?: Array<{
        author?: { displayName?: string };
        started?: string;
        timeSpent?: string;
        comment?: unknown;
      }>;
    };
  };
}

export interface JiraCredentials {
  baseUrl: string;
  email: string;
  apiToken: string;
  authType?: "basic" | "bearer";
}

export interface JiraPlanningContext {
  openSubtaskEstimateSeconds: number | null;
  subtasks: Array<{
    key: string;
    title: string;
    status: string | null;
    estimateSeconds: number | null;
  }>;
}

function mapJiraPriority(name?: string) {
  const value = (name ?? "").toLowerCase();
  if (value.includes("highest") || value.includes("high") || value.includes("blocker")) {
    return "High" as const;
  }
  if (value.includes("low") || value.includes("lowest") || value.includes("minor")) {
    return "Low" as const;
  }
  return "Medium" as const;
}

export function getMappedJiraPriority(name?: string) {
  return mapJiraPriority(name);
}

export function buildJiraIssueBrowseUrl(baseUrl: string, issueKey: string) {
  return buildJiraUrl(baseUrl, `/browse/${issueKey}`);
}

export function normalizeJiraBaseUrl(input: string) {
  const parsed = new URL(input.trim());
  if (!["http:", "https:"].includes(parsed.protocol)) {
    throw new Error("Jira URL must start with http:// or https://");
  }

  return parsed.origin;
}

function buildJiraUrl(baseUrl: string, path: string) {
  return new URL(path, `${normalizeJiraBaseUrl(baseUrl)}/`).toString();
}

function formatNetworkError(error: unknown) {
  if (error instanceof Error && "cause" in error && error.cause) {
    const cause = error.cause as { code?: string; message?: string };
    return new Error(
      cause.code
        ? `Jira network error (${cause.code}): ${cause.message ?? error.message}. Check VPN/network access to your Jira host.`
        : `Jira network error: ${error.message}`
    );
  }

  return error instanceof Error ? error : new Error("Unknown Jira network error");
}

async function buildJiraHttpError(prefix: string, response: Response) {
  let details = "";
  try {
    details = (await response.text()).trim();
  } catch {
    details = "";
  }

  return new Error(details ? `${prefix}: ${response.status} ${details}` : `${prefix}: ${response.status}`);
}

async function jiraFetchWithAuthType(
  url: string,
  credentials: JiraCredentials,
  authType: "basic" | "bearer"
) {
  const previousTls = process.env.NODE_TLS_REJECT_UNAUTHORIZED;
  try {
    if (env.jiraAllowSelfSignedTls) {
      process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
    }

    const authorization =
      authType === "bearer"
        ? `Bearer ${credentials.apiToken}`
        : `Basic ${Buffer.from(`${credentials.email}:${credentials.apiToken}`).toString("base64")}`;

    return await fetch(url, {
      headers: {
        Authorization: authorization,
        Accept: "application/json"
      }
    });
  } catch (error) {
    throw formatNetworkError(error);
  } finally {
    if (env.jiraAllowSelfSignedTls) {
      if (previousTls === undefined) {
        delete process.env.NODE_TLS_REJECT_UNAUTHORIZED;
      } else {
        process.env.NODE_TLS_REJECT_UNAUTHORIZED = previousTls;
      }
    }
  }
}

async function jiraFetchAuthed(url: string, credentials: JiraCredentials) {
  if (credentials.authType) {
    return jiraFetchWithAuthType(url, credentials, credentials.authType);
  }

  const basicResponse = await jiraFetchWithAuthType(url, credentials, "basic");
  if (basicResponse.status !== 401 && basicResponse.status !== 403) {
    return basicResponse;
  }

  return jiraFetchWithAuthType(url, credentials, "bearer");
}

export async function validateJiraCredentials(input: {
  baseUrl: string;
  email: string;
  apiToken: string;
}) {
  const credentials = {
    baseUrl: normalizeJiraBaseUrl(input.baseUrl),
    email: input.email.trim(),
    apiToken: input.apiToken.trim()
  };
  const url = buildJiraUrl(credentials.baseUrl, "/rest/api/2/myself");
  const basicResponse = await jiraFetchWithAuthType(url, credentials, "basic");

  if (basicResponse.ok) {
    return {
      profile: (await basicResponse.json()) as { displayName?: string; emailAddress?: string },
      authType: "basic" as const
    };
  }

  const bearerResponse = await jiraFetchWithAuthType(url, credentials, "bearer");
  if (bearerResponse.ok) {
    return {
      profile: (await bearerResponse.json()) as { displayName?: string; emailAddress?: string },
      authType: "bearer" as const
    };
  }

  throw await buildJiraHttpError("Jira validation failed", bearerResponse);
}

export async function fetchOpenAssignedIssuesForCredentials(
  credentials: JiraCredentials,
  _sinceIso: string
) {
  const jql = "assignee = currentUser() AND resolution = Unresolved ORDER BY updated DESC";
  const searchUrl = new URL(buildJiraUrl(credentials.baseUrl, "/rest/api/2/search"));
  searchUrl.searchParams.set("jql", jql);
  searchUrl.searchParams.set("maxResults", "50");
  searchUrl.searchParams.set("fields", "summary,status,priority,updated,issuetype,timeoriginalestimate,timetracking");

  const response = await jiraFetchAuthed(searchUrl.toString(), credentials);

  if (!response.ok) {
    throw await buildJiraHttpError("Jira search failed", response);
  }

  const json = (await response.json()) as { issues: JiraIssue[] };
  return json.issues.filter((issue) => issue.fields.issuetype?.subtask !== true);
}

export async function fetchOpenAssignedIssues(sinceIso: string) {
  const connection = getIntegrationConnection("jira");
  if (!connection?.configJson) {
    throw new Error("Jira integration is not connected");
  }

  const config = JSON.parse(connection.configJson) as JiraCredentials;
  return fetchOpenAssignedIssuesForCredentials(config, sinceIso);
}

function isDoneLikeStatus(status?: string | null) {
  const value = (status ?? "").toLowerCase();
  return /(done|closed|resolved|complete|completed|cancelled|canceled)/.test(value);
}

function choosePlanningSubtasks(
  subtasks: Array<{
    key: string;
    title: string;
    status: string | null;
    estimateSeconds: number | null;
    updated: string | null;
  }>
) {
  return [...subtasks].sort((left, right) => {
    const leftInProgress = /(progress|coding|review|testing|qa|blocked|in dev|development)/i.test(left.status ?? "");
    const rightInProgress = /(progress|coding|review|testing|qa|blocked|in dev|development)/i.test(right.status ?? "");
    if (leftInProgress !== rightInProgress) return leftInProgress ? -1 : 1;
    const leftEstimate = left.estimateSeconds ?? 0;
    const rightEstimate = right.estimateSeconds ?? 0;
    if (rightEstimate !== leftEstimate) return rightEstimate - leftEstimate;
    return new Date(right.updated ?? 0).getTime() - new Date(left.updated ?? 0).getTime();
  });
}

export async function fetchJiraIssuePlanningContextForCredentials(
  credentials: JiraCredentials,
  issueKey: string
): Promise<JiraPlanningContext> {
  const searchUrl = new URL(buildJiraUrl(credentials.baseUrl, "/rest/api/2/search"));
  searchUrl.searchParams.set("jql", `parent = ${issueKey} AND resolution = Unresolved ORDER BY updated DESC`);
  searchUrl.searchParams.set("maxResults", "50");
  searchUrl.searchParams.set("fields", "summary,status,updated,timeoriginalestimate,timetracking");

  const response = await jiraFetchAuthed(searchUrl.toString(), credentials);
  if (!response.ok) {
    throw await buildJiraHttpError("Jira subtask lookup failed", response);
  }

  const json = (await response.json()) as { issues?: Array<JiraIssue & { fields: JiraIssue["fields"] & { updated?: string } }> };
  const subtasks = choosePlanningSubtasks(
    (json.issues ?? [])
      .map((issue) => ({
        key: issue.key,
        title: issue.fields.summary ?? issue.key,
        status: issue.fields.status?.name ?? null,
        estimateSeconds: extractOriginalEstimateSeconds(issue.fields as unknown as Record<string, unknown>),
        updated: issue.fields.updated ?? null
      }))
      .filter((subtask) => !isDoneLikeStatus(subtask.status))
  );

  const summed = subtasks.reduce((total, subtask) => total + Math.max(0, subtask.estimateSeconds ?? 0), 0);

  return {
    openSubtaskEstimateSeconds: summed > 0 ? summed : null,
    subtasks: subtasks.map(({ key, title, status, estimateSeconds }) => ({
      key,
      title,
      status,
      estimateSeconds
    }))
  };
}

export async function fetchJiraIssuePlanningContext(issueKey: string) {
  const connection = getIntegrationConnection("jira");
  if (!connection?.configJson) {
    throw new Error("Jira integration is not connected");
  }

  const config = JSON.parse(connection.configJson) as JiraCredentials;
  return fetchJiraIssuePlanningContextForCredentials(config, issueKey);
}

function adfToText(node: unknown): string {
  if (!node) return "";
  if (typeof node === "string") return node;
  if (Array.isArray(node)) return node.map(adfToText).filter(Boolean).join("\n");
  if (typeof node !== "object") return "";

  const record = node as {
    type?: string;
    text?: string;
    content?: unknown[];
    attrs?: Record<string, unknown>;
  };

  if (record.text) return record.text;

  const content = (record.content ?? []).map(adfToText).filter(Boolean);

  switch (record.type) {
    case "paragraph":
      return content.join("");
    case "bulletList":
      return content.map((item) => `• ${item}`).join("\n");
    case "orderedList":
      return content.map((item, index) => `${index + 1}. ${item}`).join("\n");
    case "listItem":
      return content.join(" ");
    case "hardBreak":
      return "\n";
    case "heading":
      return content.join("");
    case "mention":
      return String(record.attrs?.text ?? record.attrs?.id ?? "");
    case "codeBlock":
      return content.join("\n");
    default:
      return content.join(record.type === "doc" ? "\n\n" : "");
  }
}

function extractStoryPoints(fields: Record<string, unknown>) {
  const known = fields.customfield_10016;
  if (typeof known === "number") {
    return known;
  }

  for (const [key, value] of Object.entries(fields)) {
    if (!key.startsWith("customfield_")) continue;
    if (typeof value === "number") {
      return value;
    }
  }
  return null;
}

function extractOriginalEstimateSeconds(fields: Record<string, unknown>) {
  const direct = fields.timeoriginalestimate;
  if (typeof direct === "number") {
    return direct;
  }

  const timetracking = fields.timetracking as { originalEstimateSeconds?: unknown } | undefined;
  if (typeof timetracking?.originalEstimateSeconds === "number") {
    return timetracking.originalEstimateSeconds;
  }

  return null;
}

export async function fetchJiraIssueDetail(issueKey: string) {
  const connection = getIntegrationConnection("jira");
  if (!connection?.configJson) {
    throw new Error("Jira integration is not connected");
  }

  const config = JSON.parse(connection.configJson) as JiraCredentials;
  const detailUrl = new URL(buildJiraUrl(config.baseUrl, `/rest/api/2/issue/${issueKey}`));
  detailUrl.searchParams.set(
    "fields",
    [
      "summary",
      "status",
      "priority",
      "description",
      "subtasks",
      "comment",
      "worklog",
      "assignee",
      "reporter",
      "labels",
      "customfield_10016"
    ].join(",")
  );

  const response = await jiraFetchAuthed(detailUrl.toString(), config);
  if (!response.ok) {
    throw await buildJiraHttpError("Jira issue detail failed", response);
  }

  const json = (await response.json()) as JiraIssueDetailResponse;
  const fields = json.fields ?? {};

  return {
    type: "jira" as const,
    key: json.key,
    title: String(fields.summary ?? issueKey),
    status: fields.status?.name ?? null,
    priority: fields.priority?.name ?? null,
    description: adfToText(fields.description) || null,
    storyPoints: extractStoryPoints(fields),
    assignee: fields.assignee?.displayName ?? null,
    reporter: fields.reporter?.displayName ?? null,
    labels: fields.labels ?? [],
    subtasks: (fields.subtasks ?? []).map((subtask) => ({
      key: subtask.key,
      title: subtask.fields?.summary ?? subtask.key,
      status: subtask.fields?.status?.name ?? null
    })),
    comments: (fields.comment?.comments ?? []).map((comment) => ({
      author: comment.author?.displayName ?? "Unknown",
      createdAt: comment.created ?? null,
      body: adfToText(comment.body) || ""
    })),
    worklogs: (fields.worklog?.worklogs ?? []).map((worklog) => ({
      author: worklog.author?.displayName ?? "Unknown",
      startedAt: worklog.started ?? null,
      timeSpent: worklog.timeSpent ?? null,
      comment: adfToText(worklog.comment) || null
    }))
  };
}
