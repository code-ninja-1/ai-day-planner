import { env } from "../env.js";
import { getIntegrationConnection } from "../db.js";
import type { TaskStatus } from "../types.js";

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

interface JiraTransitionResponse {
  transitions?: Array<{
    id?: string;
    name?: string;
    to?: {
      name?: string;
    };
  }>;
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

export interface JiraTransition {
  id: string;
  name: string;
  toStatus: string | null;
  toStatusCategory: TaskStatus;
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

export function mapJiraWorkflowStatus(status?: string | null): TaskStatus {
  const value = (status ?? "").toLowerCase();
  if (/(done|closed|resolved|complete|completed|cancelled|canceled)/.test(value)) {
    return "Completed";
  }
  if (/(progress|coding|review|testing|qa|blocked|in dev|development)/.test(value)) {
    return "In Progress";
  }
  return "Not Started";
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

function getJiraCredentials() {
  const connection = getIntegrationConnection("jira");
  if (!connection?.configJson) {
    throw new Error("Jira integration is not connected");
  }
  return JSON.parse(connection.configJson) as JiraCredentials;
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
  authType: "basic" | "bearer",
  init?: RequestInit
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
      ...init,
      headers: {
        Authorization: authorization,
        Accept: "application/json",
        ...(init?.headers ?? {})
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

async function jiraFetchAuthed(url: string, credentials: JiraCredentials, init?: RequestInit) {
  if (credentials.authType) {
    return jiraFetchWithAuthType(url, credentials, credentials.authType, init);
  }

  const basicResponse = await jiraFetchWithAuthType(url, credentials, "basic", init);
  if (basicResponse.status !== 401 && basicResponse.status !== 403) {
    return basicResponse;
  }

  return jiraFetchWithAuthType(url, credentials, "bearer", init);
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
  return fetchOpenAssignedIssuesForCredentials(getJiraCredentials(), sinceIso);
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
  return fetchJiraIssuePlanningContextForCredentials(getJiraCredentials(), issueKey);
}

async function fetchJiraTransitionsForCredentials(credentials: JiraCredentials, issueKey: string): Promise<JiraTransition[]> {
  const url = buildJiraUrl(credentials.baseUrl, `/rest/api/2/issue/${issueKey}/transitions`);
  const response = await jiraFetchAuthed(url, credentials);
  if (!response.ok) {
    throw await buildJiraHttpError("Jira transitions fetch failed", response);
  }
  const json = (await response.json()) as JiraTransitionResponse;
  return (json.transitions ?? [])
    .map((transition) => ({
      id: String(transition.id ?? "").trim(),
      name: String(transition.name ?? "").trim(),
      toStatus: transition.to?.name ?? null,
      toStatusCategory: mapJiraWorkflowStatus(transition.to?.name ?? transition.name ?? null)
    }))
    .filter((transition) => transition.id && transition.name);
}

function pickTransitionForTaskStatus(
  transitions: JiraTransition[],
  targetStatus: TaskStatus,
  currentStatus?: string | null
) {
  const exactCategory = transitions.filter((transition) => transition.toStatusCategory === targetStatus);
  if (!exactCategory.length) return null;

  const currentCategory = mapJiraWorkflowStatus(currentStatus);
  const nonSameCategory = exactCategory.find((transition) => transition.toStatusCategory !== currentCategory);
  return nonSameCategory ?? exactCategory[0] ?? null;
}

async function transitionJiraIssueForCredentials(credentials: JiraCredentials, issueKey: string, transitionId: string) {
  const url = buildJiraUrl(credentials.baseUrl, `/rest/api/2/issue/${issueKey}/transitions`);
  const response = await jiraFetchAuthed(url, credentials, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      transition: {
        id: transitionId
      }
    })
  });
  if (!response.ok) {
    throw await buildJiraHttpError("Jira transition failed", response);
  }
}

export async function fetchJiraTransitions(issueKey: string) {
  return fetchJiraTransitionsForCredentials(getJiraCredentials(), issueKey);
}

export async function transitionJiraIssue(issueKey: string, transitionId: string) {
  const credentials = getJiraCredentials();
  await transitionJiraIssueForCredentials(credentials, issueKey, transitionId);
}

export async function transitionJiraIssueToTaskStatus(issueKey: string, targetStatus: TaskStatus, currentStatus?: string | null) {
  const credentials = getJiraCredentials();
  const transitions = await fetchJiraTransitionsForCredentials(credentials, issueKey);
  const transition = pickTransitionForTaskStatus(transitions, targetStatus, currentStatus);
  if (!transition) {
    throw new Error(`No Jira transition is available that maps to ${targetStatus}.`);
  }
  await transitionJiraIssueForCredentials(credentials, issueKey, transition.id);
  return transition;
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

export async function fetchJiraIssueForSync(issueKey: string) {
  const credentials = getJiraCredentials();
  const detailUrl = new URL(buildJiraUrl(credentials.baseUrl, `/rest/api/2/issue/${issueKey}`));
  detailUrl.searchParams.set("fields", "summary,status,priority,updated,issuetype,timeoriginalestimate,timetracking");

  const response = await jiraFetchAuthed(detailUrl.toString(), credentials);
  if (!response.ok) {
    throw await buildJiraHttpError("Jira issue fetch failed", response);
  }

  return (await response.json()) as JiraIssue;
}

export async function fetchJiraIssueDetail(issueKey: string) {
  const config = getJiraCredentials();
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
  const transitions = await fetchJiraTransitionsForCredentials(config, issueKey).catch(() => []);
  const subtasks = await Promise.all(
    (fields.subtasks ?? []).map(async (subtask) => {
      const subtaskTransitions = await fetchJiraTransitionsForCredentials(config, subtask.key).catch(() => []);
      const currentStatus = subtask.fields?.status?.name ?? null;
      return {
        key: subtask.key,
        title: subtask.fields?.summary ?? subtask.key,
        status: currentStatus,
        statusCategory: mapJiraWorkflowStatus(currentStatus),
        transitions: subtaskTransitions
      };
    })
  );

  return {
    type: "jira" as const,
    key: json.key,
    title: String(fields.summary ?? issueKey),
    status: fields.status?.name ?? null,
    statusCategory: mapJiraWorkflowStatus(fields.status?.name ?? null),
    priority: fields.priority?.name ?? null,
    transitions,
    description: adfToText(fields.description) || null,
    storyPoints: extractStoryPoints(fields),
    assignee: fields.assignee?.displayName ?? null,
    reporter: fields.reporter?.displayName ?? null,
    labels: fields.labels ?? [],
    subtasks,
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
