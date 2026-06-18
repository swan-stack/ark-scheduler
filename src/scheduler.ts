import { BusyBlock, OrgUser } from "./graph";

export type SchedulerRule = {
  startDate: Date;
  days: number;
  workdayStartHour: number;
  workdayEndHour: number;
  meetingMinutes: number;
  slotMinutes: number;
  excludeLunch: boolean;
  lunchStartHour: number;
  lunchEndHour: number;
  includeWeekends: boolean;
  proximityMinutes: number;
  requireAll: boolean;
  minAvailable: number;
};

export type ProximityReason = {
  userId: string;
  userName: string;
  direction: "before" | "after";
  gapMinutes: number;
  eventStart: Date;
  eventEnd: Date;
};

export type FreeSlot = {
  start: Date;
  end: Date;
  availableUsers: string[];
  unavailableUsers: string[];
  proximityLevel: "normal" | "near" | "tight";
  proximityReasons: ProximityReason[];
};

const MS = 60 * 1000;

export function computeFreeSlots(users: OrgUser[], busyBlocks: BusyBlock[], rule: SchedulerRule): FreeSlot[] {
  const selectedUserIds = new Set(users.map(u => u.id));
  const normalizedRule = normalizeRule(rule, users.length);
  const baseSlots = makeBaseSlots(normalizedRule);
  const meetingSlotCount = Math.max(1, Math.ceil(normalizedRule.meetingMinutes / normalizedRule.slotMinutes));
  const output: FreeSlot[] = [];

  for (let i = 0; i <= baseSlots.length - meetingSlotCount; i++) {
    if (!isContiguous(baseSlots, i, meetingSlotCount)) continue;

    const start = baseSlots[i].start;
    const end = baseSlots[i + meetingSlotCount - 1].end;
    const unavailable = new Set<string>();

    for (const block of busyBlocks) {
      if (!selectedUserIds.has(block.userId)) continue;
      if (overlaps(start, end, block.start, block.end)) unavailable.add(block.userId);
    }

    const available = users.filter(u => !unavailable.has(u.id));
    const passes = normalizedRule.requireAll ? unavailable.size === 0 : available.length >= normalizedRule.minAvailable;
    if (!passes) continue;

    const availableIds = new Set(available.map(u => u.id));
    const proximityReasons = computeProximity(start, end, busyBlocks, availableIds, normalizedRule.proximityMinutes);
    const beforeCount = proximityReasons.filter(r => r.direction === "before").length;
    const afterCount = proximityReasons.filter(r => r.direction === "after").length;

    output.push({
      start,
      end,
      availableUsers: available.map(u => u.displayName),
      unavailableUsers: users.filter(u => unavailable.has(u.id)).map(u => u.displayName),
      proximityLevel: beforeCount > 0 && afterCount > 0 ? "tight" : proximityReasons.length > 0 ? "near" : "normal",
      proximityReasons,
    });
  }

  return dedupeSlots(output).sort((a, b) => a.start.getTime() - b.start.getTime());
}

function normalizeRule(rule: SchedulerRule, selectedCount: number): SchedulerRule {
  const slotMinutes = clampNumber(rule.slotMinutes, 5, 120, 30);
  const meetingMinutes = clampNumber(rule.meetingMinutes, slotMinutes, 8 * 60, 60);
  const minAvailable = Math.min(Math.max(1, rule.minAvailable || 1), Math.max(1, selectedCount));
  return {
    ...rule,
    days: clampNumber(rule.days, 1, 90, 14),
    workdayStartHour: clampNumber(rule.workdayStartHour, 0, 23, 9),
    workdayEndHour: clampNumber(rule.workdayEndHour, 1, 24, 18),
    lunchStartHour: clampNumber(rule.lunchStartHour, 0, 23, 12),
    lunchEndHour: clampNumber(rule.lunchEndHour, 1, 24, 13),
    proximityMinutes: clampNumber(rule.proximityMinutes, 0, 180, 30),
    slotMinutes,
    meetingMinutes,
    minAvailable,
  };
}

function makeBaseSlots(rule: SchedulerRule) {
  const slots: { start: Date; end: Date }[] = [];
  const cursor = startOfDay(rule.startDate);
  const endDate = addDays(cursor, rule.days);

  while (cursor < endDate) {
    const day = cursor.getDay();
    const weekend = day === 0 || day === 6;

    if (rule.includeWeekends || !weekend) {
      const dayStart = setHour(cursor, Math.min(rule.workdayStartHour, rule.workdayEndHour));
      const dayEnd = setHour(cursor, Math.max(rule.workdayStartHour, rule.workdayEndHour));
      const lunchStart = setHour(cursor, Math.min(rule.lunchStartHour, rule.lunchEndHour));
      const lunchEnd = setHour(cursor, Math.max(rule.lunchStartHour, rule.lunchEndHour));

      for (let s = new Date(dayStart); s < dayEnd; s = addMinutes(s, rule.slotMinutes)) {
        const e = addMinutes(s, rule.slotMinutes);
        if (e > dayEnd) continue;
        if (rule.excludeLunch && overlaps(s, e, lunchStart, lunchEnd)) continue;
        slots.push({ start: new Date(s), end: new Date(e) });
      }
    }

    cursor.setDate(cursor.getDate() + 1);
  }

  return slots;
}

function computeProximity(start: Date, end: Date, blocks: BusyBlock[], availableUserIds: Set<string>, threshold: number): ProximityReason[] {
  const reasons: ProximityReason[] = [];
  const thresholdMs = threshold * MS;

  for (const block of blocks) {
    if (!availableUserIds.has(block.userId)) continue;

    if (block.end <= start) {
      const gap = start.getTime() - block.end.getTime();
      if (gap <= thresholdMs) reasons.push(reason(block, "before", gap));
    }

    if (block.start >= end) {
      const gap = block.start.getTime() - end.getTime();
      if (gap <= thresholdMs) reasons.push(reason(block, "after", gap));
    }
  }

  return closestOnly(reasons);
}

function closestOnly(reasons: ProximityReason[]) {
  const map = new Map<string, ProximityReason>();
  for (const r of reasons) {
    const key = `${r.userId}-${r.direction}`;
    const existing = map.get(key);
    if (!existing || r.gapMinutes < existing.gapMinutes) map.set(key, r);
  }
  return [...map.values()].sort((a, b) => a.gapMinutes - b.gapMinutes);
}

function reason(block: BusyBlock, direction: "before" | "after", gapMs: number): ProximityReason {
  return {
    userId: block.userId,
    userName: block.userName,
    direction,
    gapMinutes: Math.round(gapMs / MS),
    eventStart: block.start,
    eventEnd: block.end,
  };
}

function overlaps(aStart: Date, aEnd: Date, bStart: Date, bEnd: Date) {
  return aStart < bEnd && bStart < aEnd;
}

function isContiguous(slots: { start: Date; end: Date }[], i: number, count: number) {
  for (let j = i; j < i + count - 1; j++) {
    if (slots[j].end.getTime() !== slots[j + 1].start.getTime()) return false;
  }
  return true;
}

function dedupeSlots(slots: FreeSlot[]) {
  const seen = new Set<string>();
  return slots.filter(s => {
    const key = `${s.start.getTime()}-${s.end.getTime()}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function clampNumber(value: number, min: number, max: number, fallback: number) {
  if (!Number.isFinite(value)) return fallback;
  return Math.min(max, Math.max(min, value));
}

function startOfDay(d: Date) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function setHour(d: Date, hour: number) {
  const x = new Date(d);
  x.setHours(hour, 0, 0, 0);
  return x;
}

function addMinutes(d: Date, minutes: number) {
  return new Date(d.getTime() + minutes * MS);
}

function addDays(d: Date, days: number) {
  const x = new Date(d);
  x.setDate(x.getDate() + days);
  return x;
}
