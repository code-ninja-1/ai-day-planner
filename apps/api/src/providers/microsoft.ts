import { env } from "../env.js";
import { getIntegrationConnection, saveIntegrationConnection } from "../db.js";

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const scopes = ["offline_access", "Mail.Read", "Calendars.Read", "User.Read"];

function microsoftAuthBase() {
  return `https://login.microsoftonline.com/${env.microsoftTenantId}/oauth2/v2.0`;
}

export function getMicrosoftAuthUrl() {
  const params = new URLSearchParams({
    client_id: env.microsoftClientId,
    response_type: "code",
    redirect_uri: env.microsoftRedirectUri,
    response_mode: "query",
    scope: scopes.join(" "),
    prompt: "select_account"
  });
  return `${microsoftAuthBase()}/authorize?${params.toString()}`;
}

export async function exchangeMicrosoftCode(code: string) {
  const response = await fetch(`${microsoftAuthBase()}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: env.microsoftClientId,
      client_secret: env.microsoftClientSecret,
      grant_type: "authorization_code",
      code,
      redirect_uri: env.microsoftRedirectUri,
      scope: scopes.join(" ")
    })
  });

  if (!response.ok) {
    throw new Error(`Microsoft token exchange failed: ${response.status}`);
  }

  const json = (await response.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };

  const me = await fetch(`${GRAPH_ROOT}/me`, {
    headers: { Authorization: `Bearer ${json.access_token}` }
  });

  const meJson = (await me.json()) as { userPrincipalName?: string; displayName?: string };
  saveIntegrationConnection({
    provider: "microsoft",
    status: "connected",
    accountLabel: meJson.userPrincipalName ?? meJson.displayName ?? "Microsoft account",
    configJson: JSON.stringify({ scopes }),
    accessToken: json.access_token,
    refreshToken: json.refresh_token ?? null,
    expiresAt: new Date(Date.now() + json.expires_in * 1000).toISOString(),
    errorMessage: null
  });
}

async function refreshIfNeeded() {
  const connection = getIntegrationConnection("microsoft");
  if (!connection?.accessToken) {
    throw new Error("Microsoft integration is not connected");
  }

  if (!connection.expiresAt || new Date(connection.expiresAt).getTime() > Date.now() + 60_000) {
    return connection.accessToken;
  }

  if (!connection.refreshToken) {
    return connection.accessToken;
  }

  const response = await fetch(`${microsoftAuthBase()}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: env.microsoftClientId,
      client_secret: env.microsoftClientSecret,
      grant_type: "refresh_token",
      refresh_token: connection.refreshToken,
      redirect_uri: env.microsoftRedirectUri,
      scope: scopes.join(" ")
    })
  });

  if (!response.ok) {
    throw new Error(`Microsoft token refresh failed: ${response.status}`);
  }

  const json = (await response.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };

  saveIntegrationConnection({
    ...connection,
    status: "connected",
    accessToken: json.access_token,
    refreshToken: json.refresh_token ?? connection.refreshToken,
    expiresAt: new Date(Date.now() + json.expires_in * 1000).toISOString(),
    errorMessage: null
  });

  return json.access_token;
}

async function graph<T>(path: string, extraHeaders?: Record<string, string>) {
  const token = await refreshIfNeeded();
  return graphWithAccessToken<T>(path, token, extraHeaders);
}

function buildGraphPath(pathname: string, params?: Record<string, string>) {
  const search = new URLSearchParams();
  for (const [key, value] of Object.entries(params ?? {})) {
    search.set(key, value);
  }
  const query = search.toString();
  return query ? `${pathname}?${query}` : pathname;
}

export async function graphWithAccessToken<T>(
  path: string,
  accessToken: string,
  extraHeaders?: Record<string, string>
) {
  const response = await fetch(`${GRAPH_ROOT}${path}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(extraHeaders ?? {})
    }
  });
  if (!response.ok) {
    let details = "";
    try {
      details = await response.text();
    } catch {
      details = "";
    }
    throw new Error(
      details
        ? `Microsoft Graph error: ${response.status} ${details}`
        : `Microsoft Graph error: ${response.status}`
    );
  }
  return (await response.json()) as T;
}

export interface GraphMail {
  id: string;
  conversationId?: string;
  subject?: string;
  webLink?: string;
  bodyPreview?: string;
  body?: {
    contentType?: string;
    content?: string;
  };
  receivedDateTime?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  toRecipients?: Array<{ emailAddress?: { name?: string; address?: string } }>;
  ccRecipients?: Array<{ emailAddress?: { name?: string; address?: string } }>;
}

export interface GraphEvent {
  id: string;
  subject?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  onlineMeetingUrl?: string;
  webLink?: string;
  isCancelled?: boolean;
}

interface GraphMailboxSettings {
  timeZone?: string;
}

async function tryFetchMailboxTimeZone(accessToken?: string) {
  try {
    const settings = accessToken
      ? await graphWithAccessToken<GraphMailboxSettings>("/me/mailboxSettings?$select=timeZone", accessToken)
      : await graph<GraphMailboxSettings>("/me/mailboxSettings?$select=timeZone");
    return settings.timeZone ?? null;
  } catch {
    return null;
  }
}

function sortMailsByReceivedDate<T extends { receivedDateTime?: string }>(mails: T[], order: "asc" | "desc") {
  const direction = order === "asc" ? 1 : -1;
  return [...mails].sort((left, right) => {
    const leftTime = left.receivedDateTime ? Date.parse(left.receivedDateTime) : 0;
    const rightTime = right.receivedDateTime ? Date.parse(right.receivedDateTime) : 0;
    return (leftTime - rightTime) * direction;
  });
}

export async function fetchRecentEmails(sinceIso: string) {
  const query = buildGraphPath("/me/messages", {
    $top: "25",
    $select: "id,conversationId,subject,webLink,bodyPreview,receivedDateTime,from",
    $filter: `receivedDateTime ge ${sinceIso}`
  });
  const data = await graph<{ value: GraphMail[] }>(query);
  return sortMailsByReceivedDate(data.value, "desc");
}

export async function fetchRecentEmailsWithAccessToken(sinceIso: string, accessToken: string) {
  const query = buildGraphPath("/me/messages", {
    $top: "25",
    $select: "id,conversationId,subject,webLink,bodyPreview,receivedDateTime,from",
    $filter: `receivedDateTime ge ${sinceIso}`
  });
  const data = await graphWithAccessToken<{ value: GraphMail[] }>(query, accessToken);
  return sortMailsByReceivedDate(data.value, "desc");
}

export async function fetchTodaysMeetings(startIso: string, endIso: string, preferredTimeZone?: string | null) {
  const timeZone = preferredTimeZone ?? (await tryFetchMailboxTimeZone());
  const query = buildGraphPath("/me/calendarView", {
    startDateTime: startIso,
    endDateTime: endIso,
    $orderby: "start/dateTime",
    $top: "25"
  });
  const data = await graph<{ value: GraphEvent[] }>(
    query,
    timeZone ? { Prefer: `outlook.timezone="${timeZone}"` } : undefined
  );
  return { events: data.value, timeZone };
}

export async function fetchTodaysMeetingsWithAccessToken(
  startIso: string,
  endIso: string,
  accessToken: string,
  preferredTimeZone?: string | null
) {
  const timeZone = preferredTimeZone ?? (await tryFetchMailboxTimeZone(accessToken));
  const query = buildGraphPath("/me/calendarView", {
    startDateTime: startIso,
    endDateTime: endIso,
    $orderby: "start/dateTime",
    $top: "25"
  });
  const data = await graphWithAccessToken<{ value: GraphEvent[] }>(
    query,
    accessToken,
    timeZone ? { Prefer: `outlook.timezone="${timeZone}"` } : undefined
  );
  return { events: data.value, timeZone };
}

export async function fetchMicrosoftProfileWithAccessToken(accessToken: string) {
  return graphWithAccessToken<{ userPrincipalName?: string; displayName?: string }>("/me", accessToken);
}

function formatAddress(entry?: { emailAddress?: { name?: string; address?: string } }) {
  if (!entry?.emailAddress) return null;
  return entry.emailAddress.name
    ? `${entry.emailAddress.name} <${entry.emailAddress.address ?? ""}>`
    : (entry.emailAddress.address ?? null);
}

function htmlToText(content?: string) {
  if (!content) return "";
  return content
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .trim();
}

function bodyToText(mail: GraphMail) {
  if (!mail.body?.content) return mail.bodyPreview ?? "";
  if (mail.body.contentType?.toLowerCase() === "html") {
    return htmlToText(mail.body.content);
  }
  return mail.body.content;
}

export async function fetchEmailDetailWithAccessToken(
  messageId: string,
  conversationId: string | null | undefined,
  accessToken: string
) {
  const fields = "id,conversationId,subject,webLink,bodyPreview,body,receivedDateTime,from,toRecipients,ccRecipients";
  const messagePath = buildGraphPath(`/me/messages/${encodeURIComponent(messageId)}`, {
    $select: fields
  });
  const message = await graphWithAccessToken<GraphMail>(messagePath, accessToken);

  let thread: GraphMail[] = [];
  if (conversationId) {
    const threadPath = buildGraphPath("/me/messages", {
      $top: "10",
      $select: fields,
      $filter: `conversationId eq '${conversationId}'`
    });
    const threadResponse = await graphWithAccessToken<{ value: GraphMail[] }>(
      threadPath,
      accessToken
    );
    thread = sortMailsByReceivedDate(
      threadResponse.value.filter((mail) => mail.id !== messageId),
      "asc"
    );
  }

  return {
    type: "email" as const,
    from: formatAddress(message.from),
    to: (message.toRecipients ?? []).map((recipient) => formatAddress(recipient)).filter(Boolean) as string[],
    cc: (message.ccRecipients ?? []).map((recipient) => formatAddress(recipient)).filter(Boolean) as string[],
    subject: message.subject ?? null,
    receivedAt: message.receivedDateTime ?? null,
    body: bodyToText(message),
    thread: thread.map((mail) => ({
      id: mail.id,
      from: formatAddress(mail.from),
      to: (mail.toRecipients ?? []).map((recipient) => formatAddress(recipient)).filter(Boolean) as string[],
      cc: (mail.ccRecipients ?? []).map((recipient) => formatAddress(recipient)).filter(Boolean) as string[],
      subject: mail.subject ?? null,
      receivedAt: mail.receivedDateTime ?? null,
      body: bodyToText(mail)
    }))
  };
}
