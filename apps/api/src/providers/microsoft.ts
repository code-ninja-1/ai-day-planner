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

function buildOutlookCalendarItemLink(eventId: string) {
  return `https://outlook.office365.com/owa/?itemid=${encodeURIComponent(eventId)}&exvsurl=1&path=/calendar/item`;
}

function normalizeOptionalString(value?: string | null) {
  const normalized = value?.trim();
  return normalized ? normalized : null;
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
  sentDateTime?: string;
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
  onlineMeeting?: { joinUrl?: string };
  isOnlineMeeting?: boolean;
  onlineMeetingProvider?: string;
  webLink?: string;
  bodyPreview?: string;
  body?: {
    contentType?: string;
    content?: string;
  };
  organizer?: { emailAddress?: { name?: string; address?: string } };
  attendees?: Array<{ emailAddress?: { name?: string; address?: string } }>;
  location?: { displayName?: string };
  isCancelled?: boolean;
}

interface GraphMailboxSettings {
  timeZone?: string;
  workingHours?: {
    startTime?: string;
    endTime?: string;
  };
}

function normalizeMicrosoftLocalTime(value?: string | null, fallback = "09:30") {
  const raw = value?.trim();
  if (!raw) return fallback;
  const match = raw.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return fallback;
  const hours = match[1].padStart(2, "0");
  const minutes = match[2];
  return `${hours}:${minutes}`;
}

async function tryFetchMailboxSettings(accessToken?: string) {
  try {
    const settings = accessToken
      ? await graphWithAccessToken<GraphMailboxSettings>("/me/mailboxSettings?$select=timeZone,workingHours", accessToken)
      : await graph<GraphMailboxSettings>("/me/mailboxSettings?$select=timeZone,workingHours");
    return {
      timeZone: settings.timeZone ?? null,
      workdayStartLocal: normalizeMicrosoftLocalTime(settings.workingHours?.startTime, "09:30"),
      workdayEndLocal: normalizeMicrosoftLocalTime(settings.workingHours?.endTime, "18:00")
    };
  } catch {
    return {
      timeZone: null,
      workdayStartLocal: "09:30",
      workdayEndLocal: "18:00"
    };
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

function decodeHtmlEntities(content: string) {
  return content
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&nbsp;/gi, " ");
}

function cleanExtractedUrl(url: string) {
  return decodeHtmlEntities(url).replace(/^[("'[\s]+/, "").replace(/[)"'\].,\s>]+$/, "");
}

function isLikelyMeetingJoinUrl(url: string) {
  return /(teams\.microsoft\.com\/l\/meetup-join|teams\.live\.com\/meet|meet\.google\.com|(?:[\w-]+\.)?zoom\.us\/(?:j|wc\/join)\/|(?:[\w-]+\.)?webex\.com\/(?:meet|join|j\.php)|join\.skype\.com|skype\.com\/meet|gotomeeting\.com\/join|bluejeans\.com\/|chime\.aws\/|ringcentral\.com\/join)/i.test(
    url
  );
}

function extractMeetingJoinUrlFromContent(content?: string | null) {
  if (!content) return null;
  const decoded = decodeHtmlEntities(content);
  const hrefMatches = Array.from(decoded.matchAll(/href=["']([^"']+)["']/gi)).map((match) => match[1]);
  const textMatches = Array.from(decoded.matchAll(/https?:\/\/[^\s"'<>]+/gi)).map((match) => match[0]);
  const candidates = [...hrefMatches, ...textMatches].map(cleanExtractedUrl);
  return candidates.find((candidate) => isLikelyMeetingJoinUrl(candidate)) ?? null;
}

function extractMeetingJoinUrl(event: GraphEvent) {
  return (
    normalizeOptionalString(event.onlineMeeting?.joinUrl) ??
    normalizeOptionalString(event.onlineMeetingUrl) ??
    extractMeetingJoinUrlFromContent(event.body?.content) ??
    extractMeetingJoinUrlFromContent(event.bodyPreview) ??
    null
  );
}

function eventLooksJoinable(event: GraphEvent) {
  return Boolean(
    event.isOnlineMeeting ||
      event.onlineMeetingProvider ||
      extractMeetingJoinUrlFromContent(event.body?.content) ||
      extractMeetingJoinUrlFromContent(event.bodyPreview)
  );
}

function eventNeedsLinkFallback(event: GraphEvent) {
  return !normalizeOptionalString(event.webLink) || (eventLooksJoinable(event) && !extractMeetingJoinUrl(event));
}

async function fetchEventJoinInfo(eventId: string, accessToken?: string) {
  const query = buildGraphPath(`/me/events/${encodeURIComponent(eventId)}`, {
    $select: "id,isOnlineMeeting,onlineMeeting,onlineMeetingUrl,onlineMeetingProvider,webLink,bodyPreview,body"
  });

  const event = accessToken
    ? await graphWithAccessToken<GraphEvent>(query, accessToken)
    : await graph<GraphEvent>(query);

  return event;
}

async function enrichEventsWithJoinInfo(events: GraphEvent[], accessToken?: string) {
  const enriched = await Promise.all(
    events.map(async (event) => {
      if (!eventNeedsLinkFallback(event)) {
        return {
          ...event,
          webLink: normalizeOptionalString(event.webLink) ?? buildOutlookCalendarItemLink(event.id)
        };
      }

      try {
        const detail = await fetchEventJoinInfo(event.id, accessToken);
        const joinUrl = extractMeetingJoinUrl(detail) ?? extractMeetingJoinUrl(event);
        return {
          ...event,
          bodyPreview: detail.bodyPreview ?? event.bodyPreview,
          body: detail.body ?? event.body,
          onlineMeeting: detail.onlineMeeting ?? event.onlineMeeting,
          onlineMeetingUrl:
            normalizeOptionalString(detail.onlineMeetingUrl) ??
            normalizeOptionalString(event.onlineMeetingUrl) ??
            joinUrl ??
            undefined,
          onlineMeetingProvider: detail.onlineMeetingProvider ?? event.onlineMeetingProvider,
          webLink:
            normalizeOptionalString(detail.webLink) ??
            normalizeOptionalString(event.webLink) ??
            buildOutlookCalendarItemLink(event.id),
          isOnlineMeeting: detail.isOnlineMeeting ?? event.isOnlineMeeting ?? Boolean(joinUrl)
        };
      } catch {
        const joinUrl = extractMeetingJoinUrl(event);
        return {
          ...event,
          onlineMeetingUrl: normalizeOptionalString(event.onlineMeetingUrl) ?? joinUrl ?? undefined,
          webLink: normalizeOptionalString(event.webLink) ?? buildOutlookCalendarItemLink(event.id)
        };
      }
    })
  );

  return enriched;
}

export async function fetchRecentEmails(sinceIso: string) {
  const query = buildGraphPath("/me/messages", {
    $top: "200",
    $select: "id,conversationId,subject,webLink,bodyPreview,body,receivedDateTime,sentDateTime,from,toRecipients,ccRecipients",
    $filter: `receivedDateTime ge ${sinceIso}`
  });
  const data = await graph<{ value: GraphMail[] }>(query);
  return sortMailsByReceivedDate(data.value, "desc");
}

export async function fetchRecentEmailsWithAccessToken(sinceIso: string, accessToken: string) {
  const query = buildGraphPath("/me/messages", {
    $top: "200",
    $select: "id,conversationId,subject,webLink,bodyPreview,body,receivedDateTime,sentDateTime,from,toRecipients,ccRecipients",
    $filter: `receivedDateTime ge ${sinceIso}`
  });
  const data = await graphWithAccessToken<{ value: GraphMail[] }>(query, accessToken);
  return sortMailsByReceivedDate(data.value, "desc");
}

export async function fetchRecentSentEmails(sinceIso: string) {
  const query = buildGraphPath("/me/mailFolders/SentItems/messages", {
    $top: "150",
    $select: "id,conversationId,subject,webLink,bodyPreview,body,receivedDateTime,sentDateTime,from,toRecipients,ccRecipients",
    $filter: `sentDateTime ge ${sinceIso}`
  });
  const data = await graph<{ value: GraphMail[] }>(query);
  return [...data.value].sort((left, right) => {
    const leftTime = left.sentDateTime ? Date.parse(left.sentDateTime) : 0;
    const rightTime = right.sentDateTime ? Date.parse(right.sentDateTime) : 0;
    return rightTime - leftTime;
  });
}

export async function fetchRecentSentEmailsWithAccessToken(sinceIso: string, accessToken: string) {
  const query = buildGraphPath("/me/mailFolders/SentItems/messages", {
    $top: "150",
    $select: "id,conversationId,subject,webLink,bodyPreview,body,receivedDateTime,sentDateTime,from,toRecipients,ccRecipients",
    $filter: `sentDateTime ge ${sinceIso}`
  });
  const data = await graphWithAccessToken<{ value: GraphMail[] }>(query, accessToken);
  return [...data.value].sort((left, right) => {
    const leftTime = left.sentDateTime ? Date.parse(left.sentDateTime) : 0;
    const rightTime = right.sentDateTime ? Date.parse(right.sentDateTime) : 0;
    return rightTime - leftTime;
  });
}

export async function fetchTodaysMeetings(startIso: string, endIso: string, preferredTimeZone?: string | null) {
  const mailboxSettings = await tryFetchMailboxSettings();
  const timeZone = preferredTimeZone ?? mailboxSettings.timeZone;
  const query = buildGraphPath("/me/calendarView", {
    startDateTime: startIso,
    endDateTime: endIso,
    $select: "id,subject,start,end,onlineMeetingUrl,onlineMeeting,isOnlineMeeting,onlineMeetingProvider,webLink,bodyPreview,body,isCancelled",
    $orderby: "start/dateTime",
    $top: "25"
  });
  const data = await graph<{ value: GraphEvent[] }>(
    query,
    timeZone ? { Prefer: `outlook.timezone="${timeZone}"` } : undefined
  );
  return {
    events: await enrichEventsWithJoinInfo(data.value),
    timeZone,
    workdayStartLocal: mailboxSettings.workdayStartLocal,
    workdayEndLocal: mailboxSettings.workdayEndLocal
  };
}

export async function fetchTodaysMeetingsWithAccessToken(
  startIso: string,
  endIso: string,
  accessToken: string,
  preferredTimeZone?: string | null
) {
  const mailboxSettings = await tryFetchMailboxSettings(accessToken);
  const timeZone = preferredTimeZone ?? mailboxSettings.timeZone;
  const query = buildGraphPath("/me/calendarView", {
    startDateTime: startIso,
    endDateTime: endIso,
    $select: "id,subject,start,end,onlineMeetingUrl,onlineMeeting,isOnlineMeeting,onlineMeetingProvider,webLink,bodyPreview,body,isCancelled",
    $orderby: "start/dateTime",
    $top: "25"
  });
  const data = await graphWithAccessToken<{ value: GraphEvent[] }>(
    query,
    accessToken,
    timeZone ? { Prefer: `outlook.timezone="${timeZone}"` } : undefined
  );
  return {
    events: await enrichEventsWithJoinInfo(data.value, accessToken),
    timeZone,
    workdayStartLocal: mailboxSettings.workdayStartLocal,
    workdayEndLocal: mailboxSettings.workdayEndLocal
  };
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

function normalizeMailRecipient(value: string) {
  const trimmed = value.trim();
  const match = trimmed.match(/<([^>]+)>/);
  const address = (match?.[1] ?? trimmed).replace(/^mailto:/i, "").trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address) ? address : null;
}

function htmlToText(content?: string) {
  if (!content) return "";
  const text = decodeHtmlEntities(
    content
      .replace(/<!--[\s\S]*?-->/g, "")
      .replace(/<head[\s\S]*?<\/head>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<meta[^>]*>/gi, "")
      .replace(/<title[\s\S]*?<\/title>/gi, "")
      .replace(/<(div|p|li|tr|h\d)[^>]*>/gi, "\n")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(div|p|li|tr|table|section|article|h\d)>/gi, "\n")
      .replace(/<a [^>]*href=["']([^"']+)["'][^>]*>(.*?)<\/a>/gi, "$2")
      .replace(/<[^>]+>/g, "")
  );

  return text
    .split("\n")
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter((line) => line && !/^content-type:|^charset=|^font-family:|^color:|^http-equiv=|^<meta/i.test(line))
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function bodyToText(mail: GraphMail) {
  if (!mail.body?.content) return htmlToText(mail.bodyPreview ?? "");
  if (mail.body.contentType?.toLowerCase() === "html") {
    return htmlToText(mail.body.content);
  }
  return htmlToText(mail.body.content);
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

export async function sendMailWithAccessToken(
  accessToken: string,
  input: {
    subject: string;
    body: string;
    to: string[];
    cc: string[];
  }
) {
  const toRecipients = [...new Set(input.to.map((value) => normalizeMailRecipient(value)).filter(Boolean))];
  const ccRecipients = [...new Set(input.cc.map((value) => normalizeMailRecipient(value)).filter(Boolean))];
  if (!toRecipients.length) {
    throw new Error("At least one valid To recipient is required before sending.");
  }

  const response = await fetch(`${GRAPH_ROOT}/me/sendMail`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      message: {
        subject: input.subject,
        body: {
          contentType: "Text",
          content: input.body
        },
        toRecipients: toRecipients.map((address) => ({
          emailAddress: {
            address
          }
        })),
        ccRecipients: ccRecipients.map((address) => ({
          emailAddress: {
            address
          }
        }))
      },
      saveToSentItems: true
    })
  });

  if (!response.ok) {
    let details = "";
    try {
      details = await response.text();
    } catch {
      details = "";
    }
    throw new Error(details ? `Microsoft send mail failed: ${response.status} ${details}` : `Microsoft send mail failed: ${response.status}`);
  }
}

export async function fetchMeetingDetailWithAccessToken(eventId: string, accessToken: string) {
  const query = buildGraphPath(`/me/events/${encodeURIComponent(eventId)}`, {
    $select:
      "id,subject,start,end,bodyPreview,body,webLink,onlineMeeting,onlineMeetingUrl,organizer,attendees,location,isCancelled"
  });
  return graphWithAccessToken<GraphEvent>(query, accessToken);
}
