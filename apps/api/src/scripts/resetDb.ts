import { clearHardDevelopmentData, clearSoftDevelopmentData, db } from "../db.js";
import { env } from "../env.js";

const mode = process.argv[2] === "hard" ? "hard" : "soft";

if (mode === "hard") {
  clearHardDevelopmentData();
} else {
  clearSoftDevelopmentData();
}

const counts = {
  tasks: Number((db.prepare("SELECT COUNT(*) as count FROM tasks").get() as { count: number }).count ?? 0),
  meetings: Number((db.prepare("SELECT COUNT(*) as count FROM meetings").get() as { count: number }).count ?? 0),
  reminders: Number((db.prepare("SELECT COUNT(*) as count FROM reminders").get() as { count: number }).count ?? 0),
  rejected: Number((db.prepare("SELECT COUNT(*) as count FROM rejected_tasks").get() as { count: number }).count ?? 0),
  auditEvents: Number((db.prepare("SELECT COUNT(*) as count FROM audit_events").get() as { count: number }).count ?? 0)
};

console.log(
  JSON.stringify(
    {
      ok: true,
      mode,
      databasePath: env.databasePath,
      counts
    },
    null,
    2
  )
);
