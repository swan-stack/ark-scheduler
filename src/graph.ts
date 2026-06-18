import { AccountInfo, IPublicClientApplication } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import { loginRequest } from "./authConfig";

export type OrgUser = {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName: string;
  userType?: string;
  accountEnabled?: boolean;
};

export type BusyBlock = {
  userId: string;
  userName: string;
  start: Date;
  end: Date;
  status: string;
};

export async function ensureGraphConsent(instance: IPublicClientApplication, account: AccountInfo) {
  await instance.acquireTokenPopup({ ...loginRequest, account, prompt: "select_account" });
}

async function getAccessToken(instance: IPublicClientApplication, account: AccountInfo) {
  try {
    const token = await withTimeout(
      instance.acquireTokenSilent({ ...loginRequest, account, forceRefresh: false }),
      12000,
      "MSAL silent token",
    );
    return token.accessToken;
  } catch (silentError: any) {
    const code = String(silentError?.errorCode || silentError?.name || silentError?.message || "");

    // Cached logins from the older build may not have User.ReadBasic.All consent yet.
    // In that case, or when MSAL hidden iframe times out, recover with an explicit popup.
    if (code.includes("user_cancelled")) throw silentError;
    const token = await instance.acquireTokenPopup({ ...loginRequest, account });
    return token.accessToken;
  }
}

export function createGraphClient(instance: IPublicClientApplication, account: AccountInfo) {
  return Client.init({
    authProvider: async done => {
      try {
        done(null, await getAccessToken(instance, account));
      } catch (error) {
        done(error as any, null);
      }
    },
  });
}

function normalizeUser(u: any): OrgUser {
  return {
    id: String(u.id || u.userPrincipalName || u.mail),
    displayName: String(u.displayName || u.mail || u.userPrincipalName || "Unknown"),
    mail: u.mail || undefined,
    userPrincipalName: String(u.userPrincipalName || u.mail || ""),
    userType: u.userType,
    accountEnabled: u.accountEnabled,
  };
}

async function withTimeout<T>(promise: Promise<T>, ms: number, label: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed_out`)), ms);
  });
  try {
    return await Promise.race([promise, timeout]);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

async function graphGetWithRetry(client: Client, url: string, label: string) {
  try {
    return await withTimeout(client.api(url).version("v1.0").get(), 30000, label);
  } catch (firstError) {
    await new Promise(resolve => setTimeout(resolve, 800));
    return await withTimeout(client.api(url).version("v1.0").get(), 45000, label);
  }
}

const DEFAULT_EXCLUDED_EMAILS = [
  "",
];

function getExcludedEmails() {
  const fromEnv = String(import.meta.env.VITE_EXCLUDED_EMAILS || "")
    .split(",")
    .map(v => v.trim().toLowerCase())
    .filter(Boolean);
  return new Set([...DEFAULT_EXCLUDED_EMAILS, ...fromEnv].map(v => v.toLowerCase()).filter(Boolean));
}

function isExcludedUser(u: OrgUser, excludedEmails: Set<string>) {
  const mail = String(u.mail || "").toLowerCase();
  const upn = String(u.userPrincipalName || "").toLowerCase();
  if (excludedEmails.has(mail) || excludedEmails.has(upn)) return true;

  const address = mail || upn;
  if (!address) return true;

  const local = address.split("@")[0];
  return ["admin", "administrator", "noreply", "no-reply", "postmaster"].includes(local);
}

export async function listOrgUsers(client: Client): Promise<OrgUser[]> {
  const users = new Map<string, OrgUser>();
  const excludedEmails = getExcludedEmails();
  let url = "/users?$select=id,displayName,mail,userPrincipalName,userType,accountEnabled&$top=100";

  for (let guard = 0; guard < 100 && url; guard++) {
    const res: any = await graphGetWithRetry(client, url, "Graph /users");
    for (const raw of res.value || []) {
      const u = normalizeUser(raw);
      if (!u.userPrincipalName && !u.mail) continue;
      if (u.accountEnabled === false) continue;
      if (u.userType && u.userType !== "Member") continue;
      if (isExcludedUser(u, excludedEmails)) continue;
      users.set(u.id, u);
    }
    url = res["@odata.nextLink"] ? String(res["@odata.nextLink"]).replace("https://graph.microsoft.com/v1.0", "") : "";
  }

  return Array.from(users.values()).sort((a, b) => a.displayName.localeCompare(b.displayName));
}

export async function searchUsers(client: Client, query: string): Promise<OrgUser[]> {
  const q = query.trim().toLowerCase();
  const all = await listOrgUsers(client);
  if (!q) return all.slice(0, 100);

  return all
    .filter(u => [u.displayName, u.mail, u.userPrincipalName].filter(Boolean).some(v => String(v).toLowerCase().includes(q)))
    .slice(0, 100);
}

export async function getSchedules(client: Client, users: OrgUser[], start: Date, end: Date, intervalMinutes = 30): Promise<BusyBlock[]> {
  const schedules = users.map(u => u.mail || u.userPrincipalName).filter(Boolean);
  if (!schedules.length) return [];

  const res: any = await withTimeout(
    client.api("/me/calendar/getSchedule").post({
      schedules,
      startTime: { dateTime: toGraphLocalDateTime(start), timeZone: "Asia/Seoul" },
      endTime: { dateTime: toGraphLocalDateTime(end), timeZone: "Asia/Seoul" },
      availabilityViewInterval: intervalMinutes,
    }),
    45000,
    "Graph getSchedule",
  );

  const blocks: BusyBlock[] = [];
  (res.value || []).forEach((item: any, idx: number) => {
    const scheduleId = String(item.scheduleId || "").toLowerCase();
    const matched = users[idx] || users.find(u => [u.mail, u.userPrincipalName].filter(Boolean).some(v => String(v).toLowerCase() === scheduleId));
    if (!matched) return;

    for (const scheduleItem of item.scheduleItems || []) {
      const status = String(scheduleItem.status || "busy").toLowerCase();
      if (!["busy", "oof", "tentative", "workingelsewhere"].includes(status)) continue;

      const blockStart = parseGraphDateTime(scheduleItem.start);
      const blockEnd = parseGraphDateTime(scheduleItem.end);
      if (Number.isNaN(blockStart.getTime()) || Number.isNaN(blockEnd.getTime()) || blockEnd <= blockStart) continue;

      blocks.push({
        userId: matched.id,
        userName: matched.displayName,
        status,
        start: blockStart,
        end: blockEnd,
      });
    }
  });

  return blocks;
}

function pad(n: number) {
  return String(n).padStart(2, "0");
}

function toGraphLocalDateTime(d: Date) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function parseGraphDateTime(value: { dateTime: string; timeZone?: string }) {
  if (!value?.dateTime) return new Date("");

  const raw = String(value.dateTime);
  const tz = String(value.timeZone || "").toLowerCase();

  if (raw.endsWith("Z") || /[+-]\d{2}:?\d{2}$/.test(raw)) {
    return new Date(raw);
  }

  if (tz === "utc" || tz === "etc/utc") {
    return new Date(raw + "Z");
  }

  return new Date(raw);
}
