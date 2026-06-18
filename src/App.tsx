import { useEffect, useMemo, useState } from "react";
import { useIsAuthenticated, useMsal } from "@azure/msal-react";
import FullCalendar from "@fullcalendar/react";
import timeGridPlugin from "@fullcalendar/timegrid";
import dayGridPlugin from "@fullcalendar/daygrid";
import interactionPlugin from "@fullcalendar/interaction";
import { loginRequest } from "./authConfig";
import { BusyBlock, createGraphClient, getSchedules, listOrgUsers, OrgUser } from "./graph";
import { computeFreeSlots, FreeSlot, SchedulerRule } from "./scheduler";
import { buildOutlookComposeUrl } from "./outlook";
import { customGroups, visibleUserEmails } from "./adminSettings";
import "./styles.css";

const defaultRule: SchedulerRule = {
  startDate: new Date(),
  days: 14,
  workdayStartHour: 9,
  workdayEndHour: 18,
  meetingMinutes: 60,
  slotMinutes: 30,
  excludeLunch: true,
  lunchStartHour: 12,
  lunchEndHour: 13,
  includeWeekends: false,
  proximityMinutes: 30,
  requireAll: true,
  minAvailable: 2,
};

export default function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const account = accounts[0];
  const client = useMemo(() => (account ? createGraphClient(instance, account) : null), [instance, account]);

  const [query, setQuery] = useState("");
  const [orgUsers, setOrgUsers] = useState<OrgUser[]>([]);
  const [selectedUsers, setSelectedUsers] = useState<OrgUser[]>([]);
  const [busyBlocks, setBusyBlocks] = useState<BusyBlock[]>([]);
  const [freeSlots, setFreeSlots] = useState<FreeSlot[]>([]);
  const [rule, setRule] = useState<SchedulerRule>(defaultRule);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");
  const [activeSlot, setActiveSlot] = useState<FreeSlot | null>(null);

  const selectedIds = useMemo(() => new Set(selectedUsers.map(u => u.id)), [selectedUsers]);
  const sortedPeople = useMemo(() => {
    const map = new Map<string, OrgUser>();
    orgUsers.forEach(u => map.set(u.id, u));
    selectedUsers.forEach(u => map.set(u.id, u));
    const q = query.trim().toLowerCase();
    const visibleSet = new Set(visibleUserEmails.map(v => v.toLowerCase()));
    const people = [...map.values()].filter(u => {
      const email = String(u.mail || u.userPrincipalName || "").toLowerCase();
      if (!visibleSet.has(email)) return false;
      if (!q) return true;
      return [u.displayName, u.mail, u.userPrincipalName].filter(Boolean).some(v => String(v).toLowerCase().includes(q));
    });
    return people.sort((a, b) => {
      const as = selectedIds.has(a.id) ? 0 : 1;
      const bs = selectedIds.has(b.id) ? 0 : 1;
      if (as !== bs) return as - bs;
      return a.displayName.localeCompare(b.displayName);
    });
  }, [orgUsers, query, selectedIds, selectedUsers]);

  async function login() {
    try {
      setMessage("Microsoft 로그인 팝업을 여는 중입니다.");
      await instance.loginRedirect(loginRequest);
      setMessage("");
    } catch (e: any) {
      setMessage(`로그인 실패: ${e.message || e}`);
    }
  }

  async function loadOrgUsers() {
    if (!client) return;
    setLoading(true);
    setMessage("조직 사용자 목록을 불러오는 중입니다.");
    try {
      const users = await listOrgUsers(client);
      setOrgUsers(users);
      setMessage(`조직 사용자 ${users.length}명을 불러왔습니다. 참석자를 클릭해 선택하세요.`);
    } catch (e: any) {
      setMessage(`조직 사용자 조회 실패: ${e.message || e}. 아래 권한 재승인 후 다시 조회하세요.`);
    } finally {
      setLoading(false);
    }
  }

  async function reconnectAndLoadUsers() {
    if (!client) return;
    await loadOrgUsers();
  }

  useEffect(() => {
    if (!client) return;
    loadOrgUsers();
  }, [client]);

  function toggleUser(user: OrgUser) {
    setSelectedUsers(prev => (prev.some(u => u.id === user.id) ? prev.filter(u => u.id !== user.id) : [...prev, user]));
  }

  function clearSelected() {
    setSelectedUsers([]);
    setFreeSlots([]);
    setBusyBlocks([]);
  }

  function selectGroup(emails: string[]) {
    const emailSet = new Set(emails.map(v => v.toLowerCase()));
    const matched = orgUsers.filter(u => {
      const email = String(u.mail || u.userPrincipalName || "").toLowerCase();
      return emailSet.has(email);
    });
    setSelectedUsers(matched);
    setFreeSlots([]);
    setBusyBlocks([]);
  }
  async function findFreeTimes() {
    if (!client || selectedUsers.length === 0) return;
    setLoading(true);
    setMessage("");
    try {
      const start = new Date(rule.startDate);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + rule.days);

      const fixedRule = {
        ...rule,
        slotMinutes: 30,
        proximityMinutes: 30,
        lunchStartHour: 12,
        lunchEndHour: 13,
      };

      const blocks = await getSchedules(client, selectedUsers, start, end, 30);
      const slots = computeFreeSlots(selectedUsers, blocks, fixedRule);

      setBusyBlocks(blocks);
      setFreeSlots(slots);
      setMessage(`${selectedUsers.length}명 기준 일정 ${blocks.length}개 조회, 가능 슬롯 ${slots.length}개 산출.`);
    } catch (e: any) {
      setMessage(`free/busy 조회 실패: ${e.message || e}`);
    } finally {
      setLoading(false);
    }
  }

  const events = freeSlots.map(slot => ({
    title: label(slot),
    start: slot.start,
    end: slot.end,
    classNames: [`slot-${slot.proximityLevel}`],
    extendedProps: { slot },
  }));

  if (!isAuthenticated) {
    return (
      <main className="login-page">
        <section className="card narrow">
          <h1>Ark Scheduler</h1>
          <p>Microsoft 365 계정으로 로그인해 조직 사용자의 공통 빈 시간을 찾습니다.</p>
          <button onClick={login}>Microsoft 로그인</button>
          <p className="hint">필요 권한: User.Read, User.ReadBasic.All, Calendars.Read, Calendars.Read.Shared</p>
          {message && <p className="message">{message}</p>}
        </section>
      </main>
    );
  }

  return (
    <main className="app">
      <header className="app-header">
        <div>
          <h1>Ark Scheduler</h1>
          <p>{account?.username}</p>
        </div>
        <button className="secondary" onClick={() => instance.logoutRedirect()}>로그아웃</button>
      </header>

      <section className="layout-three">
        <aside className="panel conditions-panel">
          <h2>조건</h2>
          <RuleEditor rule={rule} setRule={setRule} selectedCount={selectedUsers.length} />

          {customGroups.length > 0 && (
            <div className="group-buttons">
              {customGroups.map(group => (
                <div className="group-button-pair" key={group.name}>
                  <button className="secondary small" onClick={() => selectGroup(group.emails)}>
                    {group.name}
                  </button>                </div>
              ))}
            </div>
          )}

          <button className="primary full" onClick={findFreeTimes} disabled={loading || selectedUsers.length === 0}>
            {loading ? "조회 중" : "공통 빈 시간 찾기"}
          </button>
          {message && <p className="message">{message}</p>}
          {orgUsers.length === 0 && (
            <button className="secondary full" onClick={reconnectAndLoadUsers} disabled={loading}>
              권한 재승인 / 조직 사용자 다시 조회
            </button>
          )}

          <h2>요약</h2>
          <div className="metric"><span>선택</span><strong>{selectedUsers.length}명</strong></div>
          <div className="metric"><span>일정</span><strong>{busyBlocks.length}개</strong></div>
          <div className="metric"><span>가능 슬롯</span><strong>{freeSlots.length}개</strong></div>

          <h2>표시 기준</h2>
          <div className="legend"><span className="box normal" /> 여유</div>
          <div className="legend"><span className="box near" /> 앞/뒤 {rule.proximityMinutes}분 이내 일정</div>
          <div className="legend"><span className="box tight" /> 앞/뒤 모두 30분 내 근접일정 있음</div>
        </aside>

        <aside className="panel people-panel">
          <div className="panel-title-row">
            <h2>참석자 <span className="count-badge">{sortedPeople.length}/{orgUsers.length}</span></h2>
            <button className="text-button" onClick={clearSelected} disabled={selectedUsers.length === 0}>전체 해제</button>
          </div>
          <div className="search-row">
            <input value={query} onChange={e => setQuery(e.target.value)} placeholder="목록 내 필터: 이름, 이메일, UPN" />
            <button onClick={() => setQuery("")} disabled={!query}>해제</button>
          </div>

          <div className="selected-strip">
            {selectedUsers.length === 0 ? (
              <div className="placeholder">선택된 참석자가 여기에 표시됩니다.</div>
            ) : (
              selectedUsers.map(u => (
                <button key={u.id} className="selected-chip" onClick={() => toggleUser(u)} title="클릭하면 선택 해제">
                  {u.displayName} ×
                </button>
              ))
            )}
          </div>

          <div className="people-list">
            {sortedPeople.length === 0 ? (
              <div className="large-placeholder">조직 사용자 목록을 불러오는 중입니다.</div>
            ) : (
              sortedPeople.map(u => {
                const selected = selectedIds.has(u.id);
                return (
                  <button key={u.id} className={`person-card ${selected ? "selected" : ""}`} onClick={() => toggleUser(u)}>
                    <span className="check">{selected ? "✓" : "+"}</span>
                    <span className="person-text">
                      <strong>{u.displayName}</strong>
                      <small>{u.mail || u.userPrincipalName}</small>
                    </span>
                  </button>
                );
              })
            )}
          </div>
        </aside>

        <section className="calendar-panel">
          <FullCalendar
            plugins={[timeGridPlugin, dayGridPlugin, interactionPlugin]}
            initialView="dayGridMonth"
            headerToolbar={{ left: "prev,next today", center: "title", right: "timeGridWeek,dayGridMonth" }}
            slotMinTime={`${String(rule.workdayStartHour).padStart(2, "0")}:00:00`}
            slotMaxTime={`${String(rule.workdayEndHour).padStart(2, "0")}:00:00`}
            weekends={rule.includeWeekends}
            events={events}
            height="auto"
            nowIndicator
            eventClick={info => {
              const slot = info.event.extendedProps.slot as FreeSlot;
              setActiveSlot(slot);
            }}
            eventDidMount={info => {
              const slot = info.event.extendedProps.slot as FreeSlot;
              info.el.title = tooltip(slot);
            }}
          />

          {activeSlot && (
            <div className="slot-modal">
              <div className="slot-modal-card">
                <div className="slot-modal-header">
                  <h3>슬롯 상세</h3>
                  <button className="text-button" onClick={() => setActiveSlot(null)}>닫기</button>
                </div>

                <p className="slot-time">
                  {fmtDate(activeSlot.start)} {fmtTime(activeSlot.start)}-{fmtTime(activeSlot.end)}
                </p>

                <h4>참석 가능</h4>
                <ul>
                  {activeSlot.availableUsers.map(name => (
                    <li key={name}>{name}</li>
                  ))}
                </ul>

                {activeSlot.unavailableUsers.length > 0 && (
                  <>
                    <h4>참석 불가</h4>
                    <ul>
                      {activeSlot.unavailableUsers.map(name => (
                        <li key={name}>{name}</li>
                      ))}
                    </ul>
                  </>
                )}

                {activeSlot.proximityReasons.length > 0 && (
                  <>
                    <h4>근접 일정</h4>
                    <ul>
                      {activeSlot.proximityReasons.map((r, idx) => (
                        <li key={idx}>
                          {r.userName}: {r.direction === "before" ? "직전" : "직후"} 일정과 {r.gapMinutes}분
                        </li>
                      ))}
                    </ul>
                  </>
                )}

                <button
                  className="primary full"
                  onClick={() =>
                    window.open(
                      buildOutlookComposeUrl(selectedUsers, activeSlot.start, activeSlot.end),
                      "_blank",
                      "noopener,noreferrer"
                    )
                  }
                >
                  Outlook 일정 만들기
                </button>
              </div>
            </div>
          )}
        </section>
      </section>
    </main>
  );
}

function RuleEditor({ rule, setRule, selectedCount }: { rule: SchedulerRule; setRule: (r: SchedulerRule) => void; selectedCount: number }) {
  function patch(p: Partial<SchedulerRule>) { setRule({ ...rule, ...p }); }
  return (
    <div className="rules">
      <label>탐색 Weeks
        <select value={Math.round(rule.days / 7)} onChange={e => patch({ days: Number(e.target.value) * 7 })}>
          <option value={1}>1주</option>
          <option value={2}>2주</option>
          <option value={3}>3주</option>
          <option value={4}>4주</option>
          <option value={8}>8주</option>
        </select>
      </label>
      <label>회의시간 <input type="number" min={5} step={5} value={rule.meetingMinutes} onChange={e => patch({ meetingMinutes: Number(e.target.value) })} /></label>
      <label>시작시간
        <select value={rule.workdayStartHour} onChange={e => patch({ workdayStartHour: Number(e.target.value) })}>
          <option value={7}>07:00</option>
          <option value={8}>08:00</option>
          <option value={9}>09:00</option>
          <option value={10}>10:00</option>
          <option value={11}>11:00</option>
        </select>
      </label>
      <label>종료시간
        <select value={rule.workdayEndHour} onChange={e => patch({ workdayEndHour: Number(e.target.value) })}>
          <option value={17}>17:00</option>
          <option value={18}>18:00</option>
          <option value={19}>19:00</option>
          <option value={20}>20:00</option>
          <option value={21}>21:00</option>
        </select>
      </label>
      <label className="check-row"><input type="checkbox" checked={rule.excludeLunch} onChange={e => patch({ excludeLunch: e.target.checked })} /> 점심제외 (12-13:00)</label>
      <label className="check-row"><input type="checkbox" checked={rule.includeWeekends} onChange={e => patch({ includeWeekends: e.target.checked })} /> 주말 포함</label>
      <label className="check-row wide"><input type="checkbox" checked={rule.requireAll} onChange={e => patch({ requireAll: e.target.checked })} /> 전원 가능만</label>
      {!rule.requireAll && (
        <label>최소 가능 인원 <input type="number" min={1} max={selectedCount || 1} value={rule.minAvailable} onChange={e => patch({ minAvailable: Number(e.target.value) })} /></label>
      )}
    </div>
  );
}

function label(slot: FreeSlot) {
  const time = `${fmtTime(slot.start)}-${fmtTime(slot.end)}`;
  if (slot.proximityLevel === "tight") return `${time} 매우 타이트`;
  if (slot.proximityLevel === "near") return `${time} 타이트`;
  return `${time} 여유`;
}

function tooltip(slot: FreeSlot) {
  const base = `${fmtDate(slot.start)} ${fmtTime(slot.start)}-${fmtTime(slot.end)}\n가능: ${slot.availableUsers.join(", ")}`;
  if (slot.unavailableUsers.length > 0) return `${base}\n불가: ${slot.unavailableUsers.join(", ")}`;
  if (slot.proximityReasons.length === 0) return `${base}\n앞뒤 근접 일정 없음`;
  return `${base}\n${slot.proximityReasons.map(r => `${r.userName}: ${r.direction === "before" ? "직전" : "직후"} 일정과 ${r.gapMinutes}분`).join("\n")}`;
}

function fmtDate(d: Date) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function fmtTime(d: Date) {
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}
