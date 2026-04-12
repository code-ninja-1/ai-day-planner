import { getIntegrationConnection } from "../db.js";
import { fetchRecentEmails, fetchTodaysMeetings } from "../providers/microsoft.js";
import { classifyEmail } from "../services/emailClassifier.js";

async function main() {
  const microsoftConnection = getIntegrationConnection("microsoft");
  if (!microsoftConnection?.accessToken) {
    throw new Error(
      "Microsoft is not connected in the local app database yet. Connect Microsoft in the app first, then rerun this script."
    );
  }

  console.log("Testing Microsoft email integration...");
  console.log(`Connected account: ${microsoftConnection.accountLabel ?? "Unknown"}`);

  const sinceIso = new Date(Date.now() - 48 * 60 * 60 * 1000).toISOString();
  const emails = await fetchRecentEmails(sinceIso);
  console.log(`Recent emails fetched since ${sinceIso}: ${emails.length}`);

  for (const email of emails.slice(0, 3)) {
    const classification = await classifyEmail(email);
    console.log(`- Subject: ${email.subject ?? "(no subject)"}`);
    console.log(`  From: ${email.from?.emailAddress?.address ?? "Unknown sender"}`);
    console.log(`  Actionable: ${classification.actionable}`);
    console.log(`  Priority: ${classification.priority}`);
    console.log(`  Suggested title: ${classification.title}`);
  }

  const start = new Date();
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + 1);
  const meetings = await fetchTodaysMeetings(start.toISOString(), end.toISOString());
  console.log(`Today's meetings fetched: ${meetings.events.length}`);
  console.log(`Calendar timezone: ${meetings.timeZone ?? "Unavailable"}`);

  for (const meeting of meetings.events.slice(0, 5)) {
    console.log(`- ${meeting.subject ?? "Untitled meeting"} | ${meeting.start?.dateTime ?? "Unknown start"}`);
  }

  if (!emails.length) {
    console.log("No recent emails found. Microsoft connectivity still looks healthy.");
  }
}

main().catch((error) => {
  console.error("Microsoft email integration test failed.");
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
