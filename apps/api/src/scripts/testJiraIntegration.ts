import { env } from "../env.js";
import {
  fetchRecentAssignedIssuesForCredentials,
  getMappedJiraPriority,
  normalizeJiraBaseUrl,
  validateJiraCredentials
} from "../providers/jira.js";

function mask(value: string) {
  if (value.length <= 6) {
    return "*".repeat(value.length);
  }
  return `${value.slice(0, 3)}***${value.slice(-3)}`;
}

async function main() {
  const rawBaseUrl = process.env.JIRA_BASE_URL ?? "";
  const email = (process.env.JIRA_EMAIL ?? "").trim();
  const apiToken = (process.env.JIRA_API_TOKEN ?? "").trim();

  if (!rawBaseUrl || !email || !apiToken) {
    throw new Error(
      "Missing Jira test credentials. Set JIRA_BASE_URL, JIRA_EMAIL, and JIRA_API_TOKEN in apps/api/.env."
    );
  }

  const credentials = {
    baseUrl: normalizeJiraBaseUrl(rawBaseUrl),
    email,
    apiToken
  };

  console.log("Testing Jira integration...");
  console.log(`Base URL: ${credentials.baseUrl}`);
  console.log(`Email: ${credentials.email}`);
  console.log(`Token: ${mask(credentials.apiToken)}`);

  const validation = await validateJiraCredentials(credentials);
  console.log("Authentication: OK");
  console.log(`Auth type: ${validation.authType}`);
  console.log(
    `Resolved account: ${validation.profile.emailAddress ?? validation.profile.displayName ?? "Unknown"}`
  );

  const sinceIso = new Date(Date.now() - 48 * 60 * 60 * 1000).toISOString();
  const issues = await fetchRecentAssignedIssuesForCredentials(
    {
      ...credentials,
      authType: validation.authType
    },
    sinceIso
  );

  console.log(`Assigned issues updated since ${sinceIso}: ${issues.length}`);
  for (const issue of issues.slice(0, 5)) {
    console.log(
      `- ${issue.key}: ${issue.fields.summary} | status=${issue.fields.status?.name ?? "Unknown"} | mappedPriority=${getMappedJiraPriority(issue.fields.priority?.name)}`
    );
  }

  if (!issues.length) {
    console.log("No recently updated assigned issues found. Authentication still looks healthy.");
  }
}

main().catch((error) => {
  console.error("Jira integration test failed.");
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
