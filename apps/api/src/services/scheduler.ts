import { generatePlan } from "./planService.js";
import { getAutomationSettings, saveAutomationSettings } from "../db.js";

let timer: NodeJS.Timeout | null = null;

function offsetMinutesFor(timeZone: string, date: Date) {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone,
    timeZoneName: "shortOffset",
    hour: "2-digit"
  }).formatToParts(date);
  const zone = parts.find((part) => part.type === "timeZoneName")?.value ?? "GMT+0";
  const match = zone.match(/GMT([+-])(\d{1,2})(?::?(\d{2}))?/i);
  if (!match) return 0;
  const sign = match[1] === "-" ? -1 : 1;
  const hours = Number(match[2] ?? 0);
  const minutes = Number(match[3] ?? 0);
  return sign * (hours * 60 + minutes);
}

function zonedDateToUtc(year: number, month: number, day: number, hours: number, minutes: number, timeZone: string) {
  const baseUtc = Date.UTC(year, month - 1, day, hours, minutes, 0, 0);
  const offset = offsetMinutesFor(timeZone, new Date(baseUtc));
  return new Date(baseUtc - offset * 60_000);
}

function nowInTimeZone(timeZone: string) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(new Date());
  const year = Number(parts.find((part) => part.type === "year")?.value ?? 0);
  const month = Number(parts.find((part) => part.type === "month")?.value ?? 0);
  const day = Number(parts.find((part) => part.type === "day")?.value ?? 0);
  return { year, month, day };
}

function nextRunAt(scheduleTimeLocal: string, timeZone: string) {
  const [hourText, minuteText] = scheduleTimeLocal.split(":");
  const hours = Number(hourText ?? 8);
  const minutes = Number(minuteText ?? 30);
  const localNow = nowInTimeZone(timeZone);
  let next = zonedDateToUtc(localNow.year, localNow.month, localNow.day, hours, minutes, timeZone);
  if (next.getTime() <= Date.now()) {
    const tomorrow = new Date(Date.UTC(localNow.year, localNow.month - 1, localNow.day, 12, 0, 0, 0));
    tomorrow.setUTCDate(tomorrow.getUTCDate() + 1);
    next = zonedDateToUtc(
      tomorrow.getUTCFullYear(),
      tomorrow.getUTCMonth() + 1,
      tomorrow.getUTCDate(),
      hours,
      minutes,
      timeZone
    );
  }
  return next;
}

async function runScheduledPlan() {
  try {
    await generatePlan(undefined, "scheduled");
    saveAutomationSettings({
      schedulerLastRunAt: new Date().toISOString(),
      schedulerLastStatus: "ok",
      schedulerLastError: null
    });
  } catch (error) {
    saveAutomationSettings({
      schedulerLastRunAt: new Date().toISOString(),
      schedulerLastStatus: "error",
      schedulerLastError: error instanceof Error ? error.message : "Scheduled generation failed"
    });
  } finally {
    scheduleAutomation();
  }
}

export function scheduleAutomation() {
  if (timer) {
    clearTimeout(timer);
    timer = null;
  }

  const settings = getAutomationSettings();
  if (!settings.scheduleEnabled) return;

  const next = nextRunAt(settings.scheduleTimeLocal, settings.scheduleTimezone);
  const delay = Math.max(10_000, next.getTime() - Date.now());
  timer = setTimeout(() => {
    void runScheduledPlan();
  }, delay);
}
