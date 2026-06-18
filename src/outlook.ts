import { OrgUser } from "./graph";

function pad(n: number) {
  return String(n).padStart(2, "0");
}

function fmtLocal(d: Date) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:00`;
}

export function buildOutlookComposeUrl(users: OrgUser[], start: Date, end: Date) {
  const to = users.map(u => u.mail || u.userPrincipalName).filter(Boolean).join(";");

  const params = new URLSearchParams({
    path: "/calendar/action/compose",
    rru: "addevent",
    subject: "Meeting",
    startdt: fmtLocal(start),
    enddt: fmtLocal(end),
    to,
  });

  return `https://outlook.office.com/calendar/0/deeplink/compose?${params.toString()}`;
}
