import React, { useState, useRef, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";

/* ============================================================
   VP Education Members Dashboard
   All data lives in memory + the user's own JSON file.
   No network calls. Export JSON / Excel anytime.
   ============================================================ */

// ---------- Brand tokens (Toastmasters-inspired) ----------
const C = {
  blue: "#004165",      // Loyal Blue — chrome, headers
  blueDeep: "#00314D",
  maroon: "#772432",    // True Maroon — signature accent
  gold: "#F2DF74",      // Happy Yellow — active highlights
  paper: "#F6F5F0",     // page background
  ink: "#1C2A33",
  grayLine: "#E3E1D8",
  green: "#2E7D32",
  greenBg: "#E8F3E9",
  amber: "#B45309",
  amberBg: "#FCF3E3",
  red: "#B3261E",
  redBg: "#FBEAE8",
};

const SERIF = "Georgia, 'Times New Roman', serif";

// ---------- Constants ----------
const PATHS = [
  "Presentation Mastery", "Dynamic Leadership", "Visionary Communication",
  "Engaging Humor", "Leadership Development", "Motivational Strategies",
  "Persuasive Influence", "Strategic Relationships", "Team Collaboration",
  "Innovative Planning", "Effective Coaching",
];

const COMMON_ROLES = [
  "Timekeeper", "Listener", "General Evaluator", "Table Topics Master",
  "Evaluator", "Toastmaster of the Day",
  "Speaker", "Table Topics Speaker",
];


const ONBOARDING_STAGES = [
  { from: 1, to: 14, label: "Attend & observe" },
  { from: 15, to: 30, label: "Table Topics speaker" },
  { from: 31, to: 50, label: "Timekeeper" },
  { from: 51, to: 70, label: "Listener" },
  { from: 71, to: 85, label: "Table Topics Master" },
  { from: 86, to: 100, label: "Deliver Icebreaker speech" },
];

const RECOGNITION_TYPES = [
  "Best Table Topics",
  "Best Evaluator",
  "First time in a role",
  "Pathways level completion",
  "Contest winner",
  "Other",
];



const CLUB_LINK = "https://toastmasterclub.org";

const CLS_NOTE_P = "text-xs italic";
const MUTED = "#5B6B73";

// Cycle definitions: months are 0-indexed (Jan = 0)
const DEFAULT_CYCLES = [
  { id: "c1", name: "Storytelling", span: "Jul–Aug", months: [6, 7] },
  { id: "c2", name: "Humour", span: "Sep–Oct", months: [8, 9] },
  { id: "c3", name: "Vocal Variety", span: "Nov–Dec", months: [10, 11] },
  { id: "c4", name: "Structure & Clarity", span: "Jan–Feb", months: [0, 1] },
  { id: "c5", name: "Persuasion", span: "Mar–Apr", months: [2, 3] },
  { id: "c6", name: "Member's Choice", span: "May–Jun", months: [4, 5] },
].map((c) => ({ ...c, tips: [], prNotes: "", actions: [] }));

const emptyData = () => ({
  version: 2,
  members: [],
  cycles: DEFAULT_CYCLES.map((c) => ({ ...c, tips: [], actions: [] })),
  recognitions: [],
  weeks: [],
  educationGoals: [],
  onboardingTemplate: null, // null = use ONBOARDING_STAGES default
  buddyGroups: [],          // [{ id, name, memberIds: [] }]
  meetingRota: [],          // [{ id, date, theme, roles: {...}, agendaItems: [{ id, memberName, topic }] }]
  events: [],               // [{ id, title, detail, date, dismissed: [] }]  — pop-up announcements for members
});

// ---------- Auth ----------
const ADMIN_PASSWORD = "unicorn"; // change this to update the shared committee password
const SS_UNLOCKED_KEY = "vpe-admin-unlocked";

// ---------- Helpers ----------
const uid = () => Math.random().toString(36).slice(2, 10);

const memberPaths = (m) => m.paths || (m.path ? [m.path] : []);

const daysSince = (iso) => {
  if (!iso) return null;
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d)) return null;
  return Math.floor((Date.now() - d.getTime()) / 86400000);
};

const onboardingDay = (iso) => {
  const ds = daysSince(iso);
  return ds === null ? null : ds + 1; // start date = day 1
};

const fmtDate = (iso) => {
  if (!iso) return "—";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d)) return "—";
  return d.toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
};

const memberStatus = (m) => {
  const ds = daysSince(m.lastAttended);
  if (ds === null) return { key: "dormant", label: "No attendance recorded", fg: C.red, bg: C.redBg };
  if (ds > 60) return { key: "dormant", label: `Dormant — ${ds} days`, fg: C.red, bg: C.redBg };
  if (ds > 30) return { key: "nudge", label: `Needs a nudge — ${ds} days`, fg: C.amber, bg: C.amberBg };
  return { key: "ok", label: `On track — ${ds}d ago`, fg: C.green, bg: C.greenBg };
};

const currentCycleId = (cycles) => {
  const m = new Date().getMonth();
  const found = cycles.find((c) => c.months.includes(m));
  return found ? found.id : null;
};

const stageForDay = (day, plan = ONBOARDING_STAGES) => {
  if (day === null) return null;
  const total = plan.length ? plan[plan.length - 1].to : 100;
  if (day > total) return { done: true };
  return plan.find((s) => day >= s.from && day <= s.to) || plan[0];
};

// ---------- Small UI atoms ----------
function Badge({ fg, bg, children }) {
  return (
    <span
      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold whitespace-nowrap"
      style={{ color: fg, backgroundColor: bg }}
    >
      {children}
    </span>
  );
}

function SectionTitle({ children, sub }) {
  return (
    <div className="mb-5">
      <h2 className="text-2xl" style={{ fontFamily: SERIF, color: C.blue }}>{children}</h2>
      {sub && <p className="text-sm mt-1" style={{ color: "#5B6B73" }}>{sub}</p>}
      <div className="mt-2 h-0.5 w-12" style={{ backgroundColor: C.maroon }} />
    </div>
  );
}

function Card({ children, className = "", accent }) {
  return (
    <div
      className={`bg-white rounded-lg shadow-sm ${className}`}
      style={{
        border: `1px solid ${C.grayLine}`,
        borderLeft: accent ? `4px solid ${accent}` : `1px solid ${C.grayLine}`,
      }}
    >
      {children}
    </div>
  );
}

function Btn({ children, onClick, kind = "primary", className = "", title, disabled }) {
  const styles = {
    primary: { backgroundColor: C.blue, color: "white" },
    maroon: { backgroundColor: C.maroon, color: "white" },
    ghost: { backgroundColor: "white", color: C.blue, border: `1px solid ${C.grayLine}` },
    danger: { backgroundColor: "white", color: C.red, border: `1px solid ${C.red}` },
  };
  return (
    <button
      title={title}
      disabled={disabled}
      onClick={onClick}
      className={`px-3 py-1.5 rounded-md text-sm font-semibold transition-opacity hover:opacity-85 focus:outline-none focus:ring-2 disabled:opacity-40 ${className}`}
      style={styles[kind]}
    >
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label className="block text-sm">
      <span className="block mb-1 font-semibold" style={{ color: C.blueDeep }}>{label}</span>
      {children}
    </label>
  );
}

const inputCls =
  "w-full px-2.5 py-1.5 rounded-md text-sm bg-white focus:outline-none focus:ring-2";
const inputStyle = { border: `1px solid ${C.grayLine}`, color: C.ink };

// ---------- Welcome screen ----------
function Welcome({ onStartFresh, onLoadFile, onMemberMode }) {
  const fileRef = useRef(null);
  return (
    <div className="min-h-screen flex items-center justify-center p-6" style={{ backgroundColor: C.blue }}>
      <div className="max-w-md w-full bg-white rounded-xl shadow-lg p-8 text-center"
        style={{ borderTop: `6px solid ${C.maroon}` }}>
        <div className="text-xs tracking-widest font-bold mb-2" style={{ color: C.maroon }}>
          VP EDUCATION
        </div>
        <h1 className="text-3xl mb-3" style={{ fontFamily: SERIF, color: C.blue }}>
          Members Dashboard
        </h1>
        <p className="text-sm mb-6" style={{ color: "#5B6B73" }}>
          Your data is saved automatically in this browser. Load a JSON backup to restore a previous session, or start fresh.
        </p>
        <div className="flex flex-col gap-3">
          <Btn onClick={() => fileRef.current?.click()}>Load JSON file</Btn>
          <Btn kind="ghost" onClick={onStartFresh}>Start fresh</Btn>
          <div className="pt-3" style={{ borderTop: `1px solid ${C.grayLine}` }}>
            <Btn kind="ghost" onClick={onMemberMode} className="w-full">
              I'm a member — View my progress →
            </Btn>
          </div>
        </div>
        <input
          ref={fileRef}
          type="file"
          accept=".json,application/json"
          className="hidden"
          onChange={(e) => onLoadFile(e.target.files?.[0])}
        />
        <p className="text-xs mt-6" style={{ color: "#8A958F" }}>
          Remember to Export JSON before closing — nothing is saved automatically.
        </p>
        <a href={CLUB_LINK} target="_blank" rel="noopener noreferrer"
          className="inline-block mt-3 text-xs font-semibold underline" style={{ color: C.blue }}>
          toastmasterclub.org ↗
        </a>
      </div>
    </div>
  );
}

// ---------- Admin login ----------
function AdminLogin({ onUnlock, onMemberMode }) {
  const [pw, setPw] = useState("");
  const [error, setError] = useState(false);

  const attempt = () => {
    if (pw === ADMIN_PASSWORD) {
      sessionStorage.setItem(SS_UNLOCKED_KEY, "1");
      onUnlock();
    } else {
      setError(true);
      setPw("");
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-6" style={{ backgroundColor: C.blue }}>
      <div className="max-w-sm w-full bg-white rounded-xl shadow-lg p-8"
        style={{ borderTop: `6px solid ${C.maroon}` }}>
        <div className="text-xs tracking-widest font-bold mb-2" style={{ color: C.maroon }}>VP EDUCATION</div>
        <h1 className="text-2xl mb-2" style={{ fontFamily: SERIF, color: C.blue }}>Admin login</h1>
        <p className="text-sm mb-5" style={{ color: "#5B6B73" }}>
          Enter the committee password to access the dashboard.
        </p>
        <div className="space-y-3">
          <Field label="Password">
            <input type="password" className={inputCls} style={inputStyle} value={pw}
              onChange={(e) => { setPw(e.target.value); setError(false); }}
              placeholder="Password"
              onKeyDown={(e) => { if (e.key === "Enter") attempt(); }}
              autoFocus />
          </Field>
          {error && <p className="text-xs" style={{ color: C.red }}>Incorrect password. Try again.</p>}
          <Btn className="w-full" disabled={!pw} onClick={attempt}>Login</Btn>
          <div className="pt-3" style={{ borderTop: `1px solid ${C.grayLine}` }}>
            <Btn kind="ghost" className="w-full" onClick={onMemberMode}>
              I'm a member — View my progress →
            </Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- Member form modal ----------
function MemberModal({ initial, onboardingTemplate, onSave, onClose }) {
  const [m, setM] = useState(() => {
    if (!initial) {
      return {
        id: uid(), name: "", paths: [], level: 1,
        lastAttended: "", totalMeetings: 0, roles: [], isNew: false,
        onboardingStart: "", notes: "", roleLog: [], devFeeling: "", devNextStep: "", levelDates: {},
        customPlan: onboardingTemplate || null,
      };
    }
    return { ...initial, paths: memberPaths(initial), roleLog: initial.roleLog || [], devFeeling: initial.devFeeling || "", devNextStep: initial.devNextStep || "", levelDates: initial.levelDates || {} };
  });
  const [roleLogDraft, setRoleLogDraft] = useState({ role: "", date: "" });
  const set = (k, v) => setM((p) => ({ ...p, [k]: v }));

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center pt-10 px-3 pb-6 sm:pt-16 sm:px-6 overflow-y-auto"
      style={{ backgroundColor: "rgba(0,49,77,0.55)" }}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-2xl my-6"
        style={{ borderTop: `5px solid ${C.maroon}` }}>
        <div className="px-5 py-4 flex items-center justify-between" style={{ borderBottom: `1px solid ${C.grayLine}` }}>
          <h3 className="text-lg" style={{ fontFamily: SERIF, color: C.blue }}>
            {initial ? "Edit member" : "Add member"}
          </h3>
          <button onClick={onClose} className="text-xl leading-none px-2" style={{ color: "#8A958F" }} aria-label="Close">×</button>
        </div>
        <div className="p-5 grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="Name">
            <input className={inputCls} style={inputStyle} value={m.name}
              onChange={(e) => set("name", e.target.value)} placeholder="Full name" />
          </Field>
          <div className="sm:col-span-2">
            <span className="block mb-1 text-sm font-semibold" style={{ color: C.blueDeep }}>Pathways paths</span>
            <div className="flex flex-wrap gap-1.5">
              {PATHS.map((p) => {
                const on = (m.paths || []).includes(p);
                return (
                  <button key={p} type="button"
                    onClick={() => set("paths", on ? m.paths.filter((x) => x !== p) : [...(m.paths || []), p])}
                    className="px-2 py-1 rounded-full text-xs font-semibold"
                    style={{
                      backgroundColor: on ? C.blue : "white",
                      color: on ? "white" : C.blue,
                      border: `1px solid ${on ? C.blue : C.grayLine}`,
                    }}>
                    {p}
                  </button>
                );
              })}
            </div>
            {(m.paths || []).length === 0 && (
              <p className="text-xs mt-1" style={{ color: C.amber }}>No path selected — click to add one or more.</p>
            )}
          </div>
          <Field label="Current level (working on)">
            <select className={inputCls} style={inputStyle} value={m.level}
              onChange={(e) => set("level", Number(e.target.value))}>
              {[1, 2, 3, 4, 5].map((l) => <option key={l} value={l}>Level {l}</option>)}
            </select>
          </Field>
          <Field label="Last attended meeting">
            <input type="date" className={inputCls} style={inputStyle} value={m.lastAttended}
              onChange={(e) => set("lastAttended", e.target.value)} />
          </Field>

          <div className="sm:col-span-2 flex flex-wrap items-end gap-4 p-3 rounded-md"
            style={{ backgroundColor: C.paper, border: `1px dashed ${C.grayLine}` }}>
            <label className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.blueDeep }}>
              <input type="checkbox" checked={!!m.isCommittee} onChange={(e) => set("isCommittee", e.target.checked)} />
              Committee member
            </label>
          </div>

          <div className="sm:col-span-2 flex flex-wrap items-end gap-4 p-3 rounded-md"
            style={{ backgroundColor: m.onboardingStart ? C.paper : "transparent", border: `1px dashed ${C.grayLine}` }}>
            <label className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.blueDeep }}>
              <input type="checkbox"
                checked={!!m.onboardingStart}
                onChange={(e) => set("onboardingStart", e.target.checked ? new Date().toISOString().slice(0, 10) : "")} />
              100-day plan (opt-in for any member)
            </label>
            {m.onboardingStart && (
              <Field label="Plan start date">
                <input type="date" className={inputCls} style={inputStyle} value={m.onboardingStart}
                  onChange={(e) => set("onboardingStart", e.target.value)} />
              </Field>
            )}
          </div>

          <div className="sm:col-span-2">
            <span className="block mb-1 text-sm font-semibold" style={{ color: C.blueDeep }}>Role history (dated)</span>
            {(m.roleLog || []).length > 0 && (
              <ul className="mb-2 space-y-1">
                {(m.roleLog || []).map((entry) => (
                  <li key={entry.id} className="flex items-center gap-2 text-xs px-2 py-1 rounded"
                    style={{ backgroundColor: C.paper, border: `1px solid ${C.grayLine}` }}>
                    <span className="font-semibold" style={{ color: C.blueDeep }}>{entry.role}</span>
                    <span style={{ color: "#5B6B73" }}>{fmtDate(entry.date)}</span>
                    <button className="ml-auto text-xs" style={{ color: "#8A958F" }}
                      onClick={() => set("roleLog", m.roleLog.filter((e) => e.id !== entry.id))}
                      aria-label="Remove">✕</button>
                  </li>
                ))}
              </ul>
            )}
            <div className="flex gap-2">
              <select className={`${inputCls} flex-1`} style={inputStyle} value={roleLogDraft.role}
                onChange={(e) => setRoleLogDraft((p) => ({ ...p, role: e.target.value }))}>
                <option value="">Role…</option>
                {COMMON_ROLES.map((r) => <option key={r}>{r}</option>)}
              </select>
              <input type="date" className={`${inputCls} flex-1`} style={inputStyle}
                value={roleLogDraft.date}
                onChange={(e) => setRoleLogDraft((p) => ({ ...p, date: e.target.value }))} />
              <Btn kind="ghost" onClick={() => {
                const { role, date } = roleLogDraft;
                if (role.trim() && date) {
                  const existing = (m.roleLog || []).find((e) => e.role === role.trim());
                  if (existing) {
                    set("roleLog", m.roleLog.map((e) => e.role === role.trim() ? { ...e, date } : e));
                  } else {
                    set("roleLog", [...(m.roleLog || []), { id: uid(), role: role.trim(), date }]);
                  }
                  setRoleLogDraft({ role: "", date: "" });
                }
              }}>Log</Btn>
            </div>
          </div>

          <div className="sm:col-span-2">
            <Field label="Main goal for joining Toastmasters">
              <input className={inputCls} style={inputStyle} value={m.goal || ""}
                onChange={(e) => set("goal", e.target.value)}
                placeholder="e.g. Build public speaking confidence, improve leadership skills…" />
            </Field>
          </div>

          <div className="sm:col-span-2">
            <Field label="Notes">
              <textarea rows={3} className={inputCls} style={inputStyle} value={m.notes}
                onChange={(e) => set("notes", e.target.value)}
                placeholder="Mentoring pairings, things to remember…" />
            </Field>
          </div>
        </div>
        <div className="px-5 py-4 flex justify-end gap-2" style={{ borderTop: `1px solid ${C.grayLine}` }}>
          <Btn kind="ghost" onClick={onClose}>Cancel</Btn>
          <Btn onClick={() => { if (m.name.trim()) onSave(m); }} disabled={!m.name.trim()}>
            Save member
          </Btn>
        </div>
      </div>
    </div>
  );
}

// ---------- Onboarding progress bar ----------
function OnboardingBar({ startISO, compact, stages, stagesDone, onToggleStage }) {
  const plan = stages && stages.length ? stages : ONBOARDING_STAGES;
  const total = plan[plan.length - 1].to;
  const day = onboardingDay(startISO);

  const isChecked = (i) => stagesDone && stagesDone[i];
  const isDone = (s, i) => isChecked(i);
  const isActive = (s, i) => !isChecked(i) && day !== null && day >= s.from && day <= s.to;

  const checkedCount = stagesDone ? stagesDone.filter(Boolean).length : 0;
  const pct = Math.min(100, Math.round((checkedCount / plan.length) * 100));
  const allDone = checkedCount >= plan.length;

  const currentStage = day !== null ? stageForDay(day, plan) : null;

  return (
    <div>
      <div className="flex justify-between text-xs mb-1">
        <span className="font-semibold" style={{ color: C.blueDeep }}>
          {allDone ? "Plan complete 🎉"
            : day !== null && currentStage && !currentStage.done ? `Day ${day} of ${total} — ${currentStage.label}`
            : day === null ? `${checkedCount}/${plan.length} stages complete`
            : `Day ${day} of ${total}`}
        </span>
        <span style={{ color: "#5B6B73" }}>{pct}%</span>
      </div>
      <div className="h-2.5 rounded-full overflow-hidden" style={{ backgroundColor: C.grayLine }}>
        <div className="h-full rounded-full" style={{ width: `${pct}%`, backgroundColor: allDone ? C.green : C.maroon }} />
      </div>
      {!compact && (
        <ol className="mt-3 grid grid-cols-1 sm:grid-cols-2 gap-1.5">
          {plan.map((s, i) => {
            const done = isDone(s, i);
            const active = isActive(s, i);
            const checked = isChecked(i);
            return (
              <li key={`${s.from}-${s.label}`} className="flex items-center gap-2 text-xs px-2 py-1 rounded"
                style={{
                  backgroundColor: active ? C.gold : done ? C.greenBg : "transparent",
                  color: active ? C.blueDeep : done ? C.green : "#5B6B73",
                  fontWeight: active ? 700 : 500,
                }}>
                {onToggleStage ? (
                  <input type="checkbox" checked={checked} className="shrink-0"
                    onChange={() => onToggleStage(i)}
                    title="Mark stage complete" />
                ) : (
                  <span>{done ? "✓" : active ? "▶" : "○"}</span>
                )}
                <span>Days {s.from}–{s.to}: {s.label}</span>
              </li>
            );
          })}
        </ol>
      )}
    </div>
  );
}

// ---------- Custom 100-day plan editor ----------
function PlanModal({ member, onSave, onClose }) {
  const [stages, setStages] = useState(
    (member.customPlan && member.customPlan.length ? member.customPlan : ONBOARDING_STAGES).map((s) => ({ ...s }))
  );
  const set = (i, k, v) =>
    setStages((p) => p.map((s, idx) => (idx === i ? { ...s, [k]: k === "label" ? v : Math.max(1, Number(v) || 1) } : s)));
  const remove = (i) => setStages((p) => p.filter((_, idx) => idx !== i));
  const add = () => {
    const last = stages[stages.length - 1];
    const from = last ? last.to + 1 : 1;
    setStages((p) => [...p, { from, to: from + 13, label: "New stage" }]);
  };
  const valid = stages.length > 0 && stages.every((s, i) =>
    s.label.trim() && s.from <= s.to && (i === 0 || s.from > stages[i - 1].to));

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center pt-10 px-3 pb-6 sm:pt-16 sm:px-6 overflow-y-auto"
      style={{ backgroundColor: "rgba(0,49,77,0.55)" }}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-xl my-6" style={{ borderTop: `5px solid ${C.maroon}` }}>
        <div className="px-5 py-4 flex items-center justify-between" style={{ borderBottom: `1px solid ${C.grayLine}` }}>
          <h3 className="text-lg" style={{ fontFamily: SERIF, color: C.blue }}>
            Onboarding plan — {member.name}
          </h3>
          <button onClick={onClose} className="text-xl leading-none px-2" style={{ color: "#8A958F" }} aria-label="Close">×</button>
        </div>
        <div className="p-5 space-y-2">
          <div className="grid grid-cols-12 gap-2 text-xs font-bold" style={{ color: C.blueDeep }}>
            <span className="col-span-2">From day</span><span className="col-span-2">To day</span>
            <span className="col-span-7">Stage</span><span />
          </div>
          {stages.map((s, i) => (
            <div key={i} className="grid grid-cols-12 gap-2 items-center">
              <input type="number" min="1" className={`${inputCls} col-span-2`} style={inputStyle}
                value={s.from} onChange={(e) => set(i, "from", e.target.value)} />
              <input type="number" min="1" className={`${inputCls} col-span-2`} style={inputStyle}
                value={s.to} onChange={(e) => set(i, "to", e.target.value)} />
              <input className={`${inputCls} col-span-7`} style={inputStyle}
                value={s.label} onChange={(e) => set(i, "label", e.target.value)} />
              <button className="text-sm" style={{ color: C.red }} onClick={() => remove(i)} aria-label="Remove stage">✕</button>
            </div>
          ))}
          {!valid && (
            <p className="text-xs" style={{ color: C.red }}>
              Each stage needs a name, and day ranges must be in order without overlapping.
            </p>
          )}
          <div className="flex gap-2 pt-2">
            <Btn kind="ghost" onClick={add}>+ Add stage</Btn>
            <Btn kind="ghost" onClick={() => setStages(ONBOARDING_STAGES.map((s) => ({ ...s })))}>
              Reset to standard plan
            </Btn>
          </div>
        </div>
        <div className="px-5 py-4 flex justify-end gap-2" style={{ borderTop: `1px solid ${C.grayLine}` }}>
          <Btn kind="ghost" onClick={onClose}>Cancel</Btn>
          <Btn disabled={!valid} onClick={() => onSave(stages)}>Save plan</Btn>
        </div>
      </div>
    </div>
  );
}

// ---------- Member role logger (used in member portal profile) ----------
function MemberRoleLogger({ memberId, setData }) {
  const today = new Date().toISOString().slice(0, 10);
  const [role, setRole] = useState("");
  const [date, setDate] = useState(today);

  const log = () => {
    const r = role.trim();
    if (!r || !date) return;
    setData((d) => ({
      ...d,
      members: d.members.map((m) => {
        if (m.id !== memberId) return m;
        const existing = (m.roleLog || []).find((e) => e.role === r);
        const newLog = existing
          ? (m.roleLog || []).map((e) => e.role === r ? { ...e, date } : e)
          : [...(m.roleLog || []), { id: uid(), role: r, date }];
        return { ...m, roleLog: newLog };
      }),
    }));
    setRole("");
    setDate(today);
  };

  return (
    <div>
      <div className="text-xs font-semibold mb-2" style={{ color: C.blueDeep }}>Log a role</div>
      <datalist id="member-role-suggestions">
        {COMMON_ROLES.map((r) => <option key={r} value={r} />)}
      </datalist>
      <div className="flex flex-wrap gap-2 items-end">
        <div className="flex-1 min-w-36">
          <input className={inputCls} list="member-role-suggestions" style={inputStyle}
            value={role} onChange={(e) => setRole(e.target.value)}
            placeholder="Role name…" />
        </div>
        <div>
          <input type="date" className={inputCls} style={inputStyle}
            value={date} onChange={(e) => setDate(e.target.value)} />
        </div>
        <Btn kind="primary" onClick={log} disabled={!role.trim() || !date}>Add</Btn>
      </div>
    </div>
  );
}

// ---------- Member profile card (read-only self-view) ----------
function MemberProfileCard({ member: m, recognitions, educationGoals, allMembers, setData, onboardingTemplate, rota, buddyGroups }) {
  const st = memberStatus(m);
  const daysAgo = daysSince(m.lastAttended);
  const myRecognitions = (recognitions || []).filter((r) => r.member === m.name);
  const paths = memberPaths(m);
  const currentYear = new Date().getFullYear();
  const goals = (educationGoals || []).filter((g) => g.year === currentYear);
  const myLevel = m.level;

  const [newGoalLabel, setNewGoalLabel] = useState("");
  const [editingPlan, setEditingPlan] = useState(false);

  const effectivePlan = (m.customPlan && m.customPlan.length)
    ? m.customPlan
    : (onboardingTemplate && onboardingTemplate.length ? onboardingTemplate : ONBOARDING_STAGES);

  const toggleStage = setData
    ? (i) => {
        setData((d) => ({
          ...d,
          members: d.members.map((mem) => {
            if (mem.id !== m.id) return mem;
            const current = mem.stagesDone || Array(effectivePlan.length).fill(false);
            return { ...mem, stagesDone: current.map((v, idx) => (idx === i ? !v : v)) };
          }),
        }));
      }
    : undefined;

  const addGoal = () => {
    const label = newGoalLabel.trim();
    if (!label || !setData) return;
    const last = effectivePlan[effectivePlan.length - 1];
    const from = last ? last.to + 1 : 101;
    setData((d) => ({
      ...d,
      members: d.members.map((mem) => {
        if (mem.id !== m.id) return mem;
        const base = mem.customPlan && mem.customPlan.length ? mem.customPlan : effectivePlan;
        const newPlan = [...base, { from, to: from + 13, label }];
        const current = mem.stagesDone || Array(base.length).fill(false);
        return { ...mem, customPlan: newPlan, stagesDone: [...current, false] };
      }),
    }));
    setNewGoalLabel("");
  };

  return (
    <div className="space-y-4 max-w-2xl mx-auto">
      <div className="rounded-xl p-6" style={{ backgroundColor: C.blue, borderBottom: `5px solid ${C.maroon}` }}>
        <div className="flex items-center justify-between mb-1">
          <div className="text-xs tracking-widest font-bold" style={{ color: C.gold }}>MEMBER PROFILE</div>
          {m.isCommittee && (
            <span className="text-xs font-bold px-2 py-0.5 rounded-full"
              style={{ backgroundColor: "rgba(255,255,255,0.15)", color: C.gold }}>★ Committee</span>
          )}
        </div>
        <h1 className="text-3xl text-white mb-1" style={{ fontFamily: SERIF }}>{m.name}</h1>
        {paths.length > 0 && (
          <div className="text-sm" style={{ color: "rgba(255,255,255,0.75)" }}>
            {paths.join(" · ")} — Level {m.level}
          </div>
        )}
      </div>

      <Card className="p-4" accent={C.gold}>
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: C.blueDeep }}>CLUB LINKS</div>
        <div className="flex flex-wrap gap-2 mb-3">
          <a href="https://toastmasterclub.org/view_meeting.php" target="_blank" rel="noopener noreferrer"
            className="flex-1 text-center text-sm font-semibold px-4 py-2.5 rounded-lg"
            style={{ backgroundColor: C.blue, color: "white", minWidth: 180 }}>
            Sign up for meetings / request a speech ↗
          </a>
          <a href="https://apps.basecamp.toastmasters.org/dashboard" target="_blank" rel="noopener noreferrer"
            className="flex-1 text-center text-sm font-semibold px-4 py-2.5 rounded-lg"
            style={{ backgroundColor: C.maroon, color: "white", minWidth: 180 }}>
            Mark speech as complete on Basecamp ↗
          </a>
        </div>
        <p className="text-xs" style={{ color: MUTED }}>
          After your speech or role, remember to mark it as complete on Basecamp.
        </p>
      </Card>

      <Card className="p-4" accent={st.fg}>
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: C.blueDeep }}>ATTENDANCE</div>
        <div className="mb-3">
          <div className="text-xs mb-1" style={{ color: MUTED }}>Last attended</div>
          {setData ? (
            <input
              type="date"
              className={inputCls}
              style={inputStyle}
              value={m.lastAttended || ""}
              onChange={(e) => setData((d) => ({
                ...d,
                members: d.members.map((mem) =>
                  mem.id === m.id ? { ...mem, lastAttended: e.target.value } : mem
                ),
              }))}
            />
          ) : (
            <div className="font-bold" style={{ color: C.ink }}>{fmtDate(m.lastAttended)}</div>
          )}
          {daysAgo !== null && (
            <div className="text-xs mt-1" style={{ color: st.fg }}>{daysAgo} days ago</div>
          )}
        </div>
        <Badge fg={st.fg} bg={st.bg}>{st.label}</Badge>
      </Card>

      <Card className="p-4" accent={C.maroon}>
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: C.blueDeep }}>PATHWAYS PROGRESS</div>
        {paths.length === 0 ? (
          <p className="text-sm" style={{ color: C.amber }}>No Pathways path selected yet — speak to your VPE.</p>
        ) : (
          <div className="space-y-3">
            <div>
              <div className="text-xs mb-0.5" style={{ color: MUTED }}>Path(s)</div>
              <div className="font-semibold" style={{ color: C.blueDeep }}>{paths.join(", ")}</div>
            </div>
            <div>
              <div className="text-xs mb-0.5" style={{ color: MUTED }}>Currently working on</div>
              <div className="font-bold" style={{ color: C.blueDeep }}>Level {m.level}</div>
            </div>
          </div>
        )}
      </Card>

      {!m.onboardingStart && setData && (
        <Card className="p-4" accent={C.gold}>
          <div className="text-xs font-bold tracking-wide mb-2" style={{ color: C.blueDeep }}>100-DAY PLAN</div>
          <p className="text-sm mb-4" style={{ color: MUTED }}>
            Track your progress with a structured 100-day plan. You can tick off stages, add personal goals, and customise your checklist.
          </p>
          <Btn kind="maroon" onClick={() => {
            const today = new Date().toISOString().slice(0, 10);
            setData((d) => ({
              ...d,
              members: d.members.map((mem) =>
                mem.id === m.id ? { ...mem, onboardingStart: today } : mem
              ),
            }));
          }}>
            Enroll in 100-day plan
          </Btn>
        </Card>
      )}

      {m.onboardingStart && (
        <Card className="p-4" accent={C.gold}>
          <div className="flex items-center justify-between mb-1">
            <div className="text-xs font-bold tracking-wide" style={{ color: C.blueDeep }}>100-DAY PLAN</div>
            {setData && (
              <Btn kind="ghost" onClick={() => setEditingPlan(true)}>Edit checklist</Btn>
            )}
          </div>
          {setData && (
            <div className="flex items-center gap-3 mb-3 text-xs" style={{ color: MUTED }}>
              <span>Started:</span>
              <input
                type="date"
                className={inputCls}
                style={{ ...inputStyle, padding: "1px 6px", fontSize: "0.75rem" }}
                value={m.onboardingStart}
                onChange={(e) => {
                  const val = e.target.value;
                  setData((d) => ({
                    ...d,
                    members: d.members.map((mem) =>
                      mem.id === m.id ? { ...mem, onboardingStart: val } : mem
                    ),
                  }));
                }}
              />
              <button
                className="underline"
                style={{ color: MUTED }}
                onClick={() => {
                  if (window.confirm("Opt out of the 100-day plan? Your checklist progress will be kept if you re-enrol.")) {
                    setData((d) => ({
                      ...d,
                      members: d.members.map((mem) =>
                        mem.id === m.id ? { ...mem, onboardingStart: "" } : mem
                      ),
                    }));
                  }
                }}>
                Opt out
              </button>
            </div>
          )}
          <OnboardingBar
            startISO={m.onboardingStart}
            stages={effectivePlan}
            stagesDone={m.stagesDone}
            onToggleStage={toggleStage}
          />
          {setData && (
            <div className="mt-4 pt-3 flex gap-2" style={{ borderTop: `1px solid ${C.grayLine}` }}>
              <input
                className={`${inputCls} flex-1`}
                style={inputStyle}
                placeholder="Add a personal goal…"
                value={newGoalLabel}
                onChange={(e) => setNewGoalLabel(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") addGoal(); }}
              />
              <Btn kind="ghost" onClick={addGoal} disabled={!newGoalLabel.trim()}>Add goal</Btn>
            </div>
          )}
        </Card>
      )}

      {editingPlan && setData && (
        <PlanModal
          member={{ ...m, customPlan: effectivePlan }}
          onSave={(stages) => {
            setData((d) => ({
              ...d,
              members: d.members.map((mem) =>
                mem.id === m.id ? { ...mem, customPlan: stages } : mem
              ),
            }));
            setEditingPlan(false);
          }}
          onClose={() => setEditingPlan(false)}
        />
      )}

      <Card className="p-4" accent={C.blue}>
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: C.blueDeep }}>MY ROLES</div>
        {/* Role history */}
        {(m.roleLog && m.roleLog.length > 0) ? (
          <div className="space-y-1 mb-4">
            {[...m.roleLog].sort((a, b) => b.date.localeCompare(a.date)).map((entry) => {
              const sixMonthsAgo = new Date();
              sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
              const stale = entry.date < sixMonthsAgo.toISOString().slice(0, 10);
              return (
                <div key={entry.id} className="flex items-center gap-2 text-sm py-1"
                  style={{ borderBottom: `1px solid ${C.grayLine}` }}>
                  <span className="flex-1 font-semibold" style={{ color: stale ? C.amber : C.blueDeep }}>
                    {entry.role}
                    {stale && <span className="ml-1 text-xs font-normal" style={{ color: C.amber }}>· time to refresh!</span>}
                  </span>
                  <span className="text-xs shrink-0" style={{ color: stale ? C.amber : "#8A958F" }}>{fmtDate(entry.date)}</span>
                  {setData && (
                    <button className="text-xs shrink-0" style={{ color: MUTED }}
                      onClick={() => setData((d) => ({
                        ...d,
                        members: d.members.map((mem) =>
                          mem.id === m.id
                            ? { ...mem, roleLog: (mem.roleLog || []).filter((e) => e.id !== entry.id) }
                            : mem
                        ),
                      }))} title="Remove">✕</button>
                  )}
                </div>
              );
            })}
          </div>
        ) : (
          <p className="text-xs mb-4" style={{ color: "#8A958F" }}>No roles logged yet.</p>
        )}
        {/* Log a role */}
        {setData && (
          <MemberRoleLogger memberId={m.id} setData={setData} />
        )}
      </Card>

      {m.goal && (
        <Card className="p-4" accent={C.blue}>
          <div className="text-xs font-bold tracking-wide mb-2" style={{ color: C.blueDeep }}>MY TOASTMASTERS GOAL</div>
          <p className="text-sm" style={{ color: C.ink }}>{m.goal}</p>
        </Card>
      )}

      {myRecognitions.length > 0 && (
        <Card className="p-4" accent={C.gold}>
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: C.blueDeep }}>RECOGNITIONS</div>
          <div className="space-y-1.5">
            {myRecognitions.map((r) => (
              <div key={r.id} className="flex items-center gap-2 text-sm">
                <span style={{ color: C.gold }}>★</span>
                <span className="font-semibold" style={{ color: C.blueDeep }}>{r.type}</span>
                {r.detail && <span style={{ color: MUTED }}>— {r.detail}</span>}
              </div>
            ))}
          </div>
        </Card>
      )}

      {goals.length > 0 && (
        <Card className="p-4" accent={C.maroon}>
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: C.blueDeep }}>
            HOW YOU CAN HELP — EDUCATION GOALS {currentYear}
          </div>
          <p className="text-xs mb-3" style={{ color: MUTED }}>
            The club has set level-completion targets for this year. Completing your next Pathways level directly helps us hit these goals.
          </p>
          <div className="space-y-3">
            {goals.map((g) => {
              const completed = g.completed || 0;
              const pct = g.target > 0 ? Math.min(100, Math.round((completed / g.target) * 100)) : 0;
              const hit = completed >= g.target;
              const isMyLevel = g.level === myLevel;
              return (
                <div key={g.id}>
                  <div className="flex items-center justify-between text-xs mb-1">
                    <span className="font-semibold" style={{ color: isMyLevel ? C.maroon : C.blueDeep }}>
                      Level {g.level} — {completed}/{g.target} completed
                      {isMyLevel && (
                        <span className="ml-2 font-normal" style={{ color: C.maroon }}>← you're working on this!</span>
                      )}
                    </span>
                    <span style={{ color: hit ? C.green : C.ink }}>{pct}%</span>
                  </div>
                  <div className="h-2 rounded-full overflow-hidden" style={{ backgroundColor: C.grayLine }}>
                    <div className="h-full rounded-full"
                      style={{ width: `${pct}%`, backgroundColor: hit ? C.green : C.maroon }} />
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
      )}

      {/* Committee badge */}
      {m.isCommittee && (
        <Card className="p-4" accent={C.blueDeep}>
          <span className="text-xs font-bold tracking-wide" style={{ color: C.gold }}>★ COMMITTEE MEMBER</span>
        </Card>
      )}
    </div>
  );
}

// ---------- Member self-registration form ----------
function MemberSelfAdd({ name, onSave }) {
  const [m, setM] = useState({
    id: uid(), name, paths: [], level: 1,
    lastAttended: "", totalMeetings: 0, roles: [], isNew: false,
    onboardingStart: "", notes: "", roleLog: [], goal: "", levelDates: {},
  });
  const set = (k, v) => setM((p) => ({ ...p, [k]: v }));

  return (
    <div className="space-y-4">
      <div>
        <span className="block mb-1 text-sm font-semibold" style={{ color: C.blueDeep }}>Pathways path(s)</span>
        <div className="flex flex-wrap gap-1.5">
          {PATHS.map((p) => {
            const on = m.paths.includes(p);
            return (
              <button key={p} type="button"
                onClick={() => set("paths", on ? m.paths.filter((x) => x !== p) : [...m.paths, p])}
                className="px-2 py-1 rounded-full text-xs font-semibold"
                style={{
                  backgroundColor: on ? C.blue : "white",
                  color: on ? "white" : C.blue,
                  border: `1px solid ${on ? C.blue : C.grayLine}`,
                }}>
                {p}
              </button>
            );
          })}
        </div>
      </div>
      <Field label="Current level (working on)">
        <select className={inputCls} style={inputStyle} value={m.level}
          onChange={(e) => set("level", Number(e.target.value))}>
          {[1, 2, 3, 4, 5].map((l) => <option key={l} value={l}>Level {l}</option>)}
        </select>
      </Field>
      <label className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.blueDeep }}>
        <input type="checkbox"
          checked={!!m.onboardingStart}
          onChange={(e) => set("onboardingStart", e.target.checked ? new Date().toISOString().slice(0, 10) : "")} />
        Opt into 100-day plan
      </label>
      {m.onboardingStart && (
        <Field label="Plan start date">
          <input type="date" className={inputCls} style={inputStyle} value={m.onboardingStart}
            onChange={(e) => set("onboardingStart", e.target.value)} />
        </Field>
      )}
      <Field label="Main goal for joining Toastmasters (optional)">
        <input className={inputCls} style={inputStyle} value={m.goal}
          onChange={(e) => set("goal", e.target.value)}
          placeholder="e.g. Build confidence speaking in public" />
      </Field>
      <Btn kind="maroon" className="w-full" onClick={() => onSave(m)}>Save my info</Btn>
    </div>
  );
}

// ---------- Member portal (name entry → profile or self-registration) ----------
function MemberPortalView({ data, setData, onSwitchToAdmin }) {
  const [nameInput, setNameInput] = useState("");
  const [searched, setSearched] = useState(false);
  const trimmed = nameInput.trim();
  const member = (data && trimmed)
    ? data.members.find((m) => m.name.toLowerCase() === trimmed.toLowerCase()) || null
    : null;

  const today = new Date().toISOString().slice(0, 10);
  const upcomingEvents = (data?.events || [])
    .filter((ev) => ev.date >= today)
    .sort((a, b) => a.date.localeCompare(b.date));


  const handleSearch = (e) => {
    e && e.preventDefault();
    if (trimmed) setSearched(true);
  };

  const handleAddSelf = (newMember) => {
    if (!data) {
      setData({ ...emptyData(), members: [newMember] });
    } else {
      setData((d) => ({ ...d, members: [...d.members, newMember] }));
    }
  };

  const reset = () => { setSearched(false); setNameInput(""); };

  return (
    <div className="min-h-screen" style={{ backgroundColor: C.paper, color: C.ink }}>
      <header className="sticky top-0 z-40 px-4 py-3 flex items-center justify-between"
        style={{ backgroundColor: C.blue }}>
        <div>
          <div className="text-xs tracking-widest font-bold" style={{ color: C.gold }}>TOASTMASTERS</div>
          <div className="text-base text-white" style={{ fontFamily: SERIF }}>
            {member ? member.name : "Member Portal"}
          </div>
        </div>
        <div className="flex items-center gap-3">
          {searched && (
            <button className="text-sm underline" style={{ color: "rgba(255,255,255,0.75)" }} onClick={reset}>
              ← Back
            </button>
          )}
          <button className="text-xs px-2 py-1 rounded font-semibold"
            style={{ backgroundColor: "rgba(255,255,255,0.12)", color: "rgba(255,255,255,0.75)" }}
            onClick={onSwitchToAdmin}>
            Admin →
          </button>
        </div>
      </header>

      <div className="p-4 sm:p-6">
        {!searched ? (
          <div className="max-w-sm mx-auto mt-12">
            <div className="text-center mb-8">
              <h2 className="text-2xl mb-2" style={{ fontFamily: SERIF, color: C.blue }}>Welcome</h2>
              <p className="text-sm" style={{ color: MUTED }}>Enter your name to view your progress.</p>
            </div>
            <form onSubmit={handleSearch} className="flex gap-2">
              <input
                className={`${inputCls} flex-1`}
                style={inputStyle}
                placeholder="Your name…"
                value={nameInput}
                onChange={(e) => setNameInput(e.target.value)}
                autoFocus
              />
              <Btn kind="maroon" onClick={handleSearch} disabled={!trimmed}>Go</Btn>
            </form>
          </div>
        ) : member ? (
          <>
            {/* Upcoming events panel */}
            {upcomingEvents.length > 0 && (
              <div className="mb-4 rounded-xl overflow-hidden" style={{ border: `1px solid ${C.grayLine}` }}>
                <div className="px-4 py-2" style={{ backgroundColor: C.blueDeep }}>
                  <span className="text-xs font-bold tracking-wide" style={{ color: C.gold }}>UPCOMING IN THE CLUB</span>
                </div>
                <div style={{ backgroundColor: "#FFFDF5" }}>
                  {upcomingEvents.map((ev, i) => (
                    <div key={ev.id} className="px-4 py-3 flex items-start gap-4"
                      style={{ borderTop: i > 0 ? `1px solid ${C.grayLine}` : "none" }}>
                      <div className="shrink-0 text-xs font-semibold pt-0.5 w-24" style={{ color: C.maroon }}>
                        {fmtDate(ev.date)}
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="font-semibold text-sm" style={{ color: C.blueDeep }}>{ev.title}</div>
                        {ev.detail && <p className="text-xs mt-0.5" style={{ color: MUTED }}>{ev.detail}</p>}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            <MemberProfileCard
              member={member}
              recognitions={data?.recognitions || []}
              educationGoals={data?.educationGoals || []}
              allMembers={data?.members || []}
              setData={setData}
              onboardingTemplate={data?.onboardingTemplate}
              rota={data?.meetingRota || []}
              buddyGroups={data?.buddyGroups || []}
            />
          </>
        ) : (
          <div className="max-w-sm mx-auto mt-8">
            <Card className="p-5" accent={C.maroon}>
              <h3 className="text-lg mb-1" style={{ fontFamily: SERIF, color: C.blue }}>
                Hi, {trimmed}!
              </h3>
              <p className="text-sm mb-4" style={{ color: MUTED }}>
                We don't have a record for you yet. Fill in a few basics and your VPE can add the rest.
              </p>
              <MemberSelfAdd name={trimmed} onSave={handleAddSelf} />
            </Card>
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Bulk member import modal ----------
// ---------- Bulk member editor (full-screen spreadsheet) ----------
// ---------- Compact multi-select cells for BulkEditModal ----------
function PathsCell({ value, onChange }) {
  const [open, setOpen] = useState(false);
  const [dropUp, setDropUp] = useState(false);
  const triggerRef = useRef(null);
  const toggle = (p) =>
    onChange(value.includes(p) ? value.filter((x) => x !== p) : [...value, p]);
  const handleOpen = () => {
    if (triggerRef.current) {
      const rect = triggerRef.current.getBoundingClientRect();
      setDropUp(window.innerHeight - rect.bottom < 220);
    }
    setOpen((v) => !v);
  };
  const dropPos = dropUp
    ? { bottom: "100%", top: "auto", marginBottom: 2 }
    : { top: "100%", bottom: "auto", marginTop: 2 };
  return (
    <div style={{ position: "relative" }} ref={triggerRef}>
      <div onClick={handleOpen}
        className="cursor-pointer rounded border px-1.5 py-1 flex flex-wrap gap-0.5 items-center"
        style={{ minHeight: "28px", borderColor: C.grayLine, backgroundColor: "white" }}>
        {value.length === 0
          ? <span className="text-xs" style={{ color: MUTED }}>Select paths…</span>
          : value.map((p) => (
            <span key={p} className="text-xs px-1 py-0.5 rounded"
              style={{ backgroundColor: C.paper, color: C.blueDeep, border: `1px solid ${C.grayLine}` }}>
              {p}
            </span>
          ))}
      </div>
      {open && (
        <>
          <div style={{ position: "fixed", inset: 0, zIndex: 99 }} onClick={() => setOpen(false)} />
          <div style={{ position: "absolute", left: 0, zIndex: 100, backgroundColor: "white",
            border: `1px solid ${C.grayLine}`, borderRadius: 6, boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
            minWidth: 220, padding: "6px 0", ...dropPos }}>
            {PATHS.map((p) => (
              <label key={p} className="flex items-center gap-2 px-3 py-1 cursor-pointer text-xs"
                style={{ color: C.ink }}
                onMouseOver={(e) => e.currentTarget.style.backgroundColor = C.paper}
                onMouseOut={(e) => e.currentTarget.style.backgroundColor = "transparent"}>
                <input type="checkbox" checked={value.includes(p)}
                  onChange={() => toggle(p)} onClick={(e) => e.stopPropagation()} />
                {p}
              </label>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

function BuddyGroupCell({ value, groups, onChange }) {
  const [open, setOpen] = useState(false);
  const [dropUp, setDropUp] = useState(false);
  const triggerRef = useRef(null);
  const toggle = (id) =>
    onChange(value.includes(id) ? value.filter((x) => x !== id) : [...value, id]);
  const selectedNames = groups.filter((g) => value.includes(g.id)).map((g) => g.name);
  const handleOpen = () => {
    if (triggerRef.current) {
      const rect = triggerRef.current.getBoundingClientRect();
      setDropUp(window.innerHeight - rect.bottom < 160);
    }
    setOpen((v) => !v);
  };
  const dropPos = dropUp
    ? { bottom: "100%", top: "auto", marginBottom: 2 }
    : { top: "100%", bottom: "auto", marginTop: 2 };
  return (
    <div style={{ position: "relative" }} ref={triggerRef}>
      <div onClick={handleOpen}
        className="cursor-pointer rounded border px-1.5 py-1 flex flex-wrap gap-0.5 items-center"
        style={{ minHeight: "28px", borderColor: C.grayLine, backgroundColor: "white" }}>
        {selectedNames.length === 0
          ? <span className="text-xs" style={{ color: MUTED }}>None</span>
          : selectedNames.map((n) => (
            <span key={n} className="text-xs px-1 py-0.5 rounded"
              style={{ backgroundColor: C.paper, color: C.maroon, border: `1px solid ${C.grayLine}` }}>
              {n}
            </span>
          ))}
      </div>
      {open && (
        <>
          <div style={{ position: "fixed", inset: 0, zIndex: 99 }} onClick={() => setOpen(false)} />
          <div style={{ position: "absolute", left: 0, zIndex: 100, backgroundColor: "white",
            border: `1px solid ${C.grayLine}`, borderRadius: 6, boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
            minWidth: 180, padding: "6px 0", ...dropPos }}>
            {groups.length === 0
              ? <p className="text-xs px-3 py-1.5" style={{ color: MUTED }}>No groups created yet.</p>
              : groups.map((g) => (
                <label key={g.id} className="flex items-center gap-2 px-3 py-1 cursor-pointer text-xs"
                  style={{ color: C.ink }}
                  onMouseOver={(e) => e.currentTarget.style.backgroundColor = C.paper}
                  onMouseOut={(e) => e.currentTarget.style.backgroundColor = "transparent"}>
                  <input type="checkbox" checked={value.includes(g.id)}
                    onChange={() => toggle(g.id)} onClick={(e) => e.stopPropagation()} />
                  {g.name}
                </label>
              ))}
          </div>
        </>
      )}
    </div>
  );
}

function BulkEditModal({ data, onSave, onClose }) {
  const [search, setSearch] = useState("");

  const getGroupIds = (memberId) =>
    (data.buddyGroups || []).filter((g) => g.memberIds.includes(memberId)).map((g) => g.id);

  const [rows, setRows] = useState(() =>
    data.members.map((m) => ({
      ...m,
      paths: memberPaths(m),
      _buddyGroupIds: getGroupIds(m.id),
    }))
  );

  const trimSearch = search.trim().toLowerCase();
  const visibleRows = trimSearch
    ? rows.filter((r) => r.name.toLowerCase().includes(trimSearch))
    : rows;

  const update = (id, k, v) =>
    setRows((prev) => prev.map((r) => (r.id === id ? { ...r, [k]: v } : r)));

  const addRow = () => {
    setRows((prev) => [
      ...prev,
      {
        id: uid(), name: "", paths: [], level: 1,
        lastAttended: "", totalMeetings: 0, roles: [], isNew: false,
        onboardingStart: "", notes: "", roleLog: [], devFeeling: "", devNextStep: "", levelDates: {},
        _buddyGroupIds: [],
      },
    ]);
    setSearch("");
  };

  const removeRow = (id) => setRows((prev) => prev.filter((r) => r.id !== id));

  const handleSave = () => {
    const cleanRows = rows.filter((r) => r.name.trim());
    const members = cleanRows.map(({ _buddyGroupIds, ...r }) => r);
    const buddyGroups = (data.buddyGroups || []).map((g) => ({
      ...g,
      memberIds: cleanRows
        .filter((r) => (r._buddyGroupIds || []).includes(g.id))
        .map((r) => r.id),
    }));
    onSave({ members, buddyGroups });
  };

  const th = "px-2 py-2 text-left text-xs font-bold whitespace-nowrap sticky top-0";
  const td = "px-1 py-1 align-middle";
  const cell = `${inputCls} text-xs`;

  return (
    <div className="fixed inset-0 z-50 flex flex-col" style={{ backgroundColor: C.paper, color: C.ink }}>
      {/* Header */}
      <div className="flex flex-wrap items-center gap-2 px-4 py-3 shrink-0"
        style={{ backgroundColor: C.blue, borderBottom: `3px solid ${C.maroon}` }}>
        <span className="text-white font-semibold" style={{ fontFamily: SERIF, fontSize: "1.05rem" }}>
          Bulk member editor
        </span>
        <div className="flex items-center gap-1.5 ml-2">
          <input
            className={inputCls}
            style={{ ...inputStyle, width: "200px", backgroundColor: "rgba(255,255,255,0.12)", color: "white", borderColor: "rgba(255,255,255,0.25)" }}
            placeholder="Search members…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          <span className="text-xs" style={{ color: "rgba(255,255,255,0.55)" }}>
            {visibleRows.length}/{rows.length}
          </span>
        </div>
        <div className="ml-auto flex items-center gap-2">
          <Btn kind="ghost" onClick={addRow}>+ Add member</Btn>
          <Btn kind="maroon" onClick={handleSave}>Save all</Btn>
          <button onClick={onClose} className="text-white text-2xl leading-none px-2 opacity-70 hover:opacity-100" aria-label="Close">×</button>
        </div>
      </div>

      {/* Table */}
      <div className="flex-1 overflow-auto">
        <table className="border-collapse text-sm" style={{ minWidth: "1300px", width: "100%" }}>
          <thead style={{ backgroundColor: C.blueDeep }}>
            <tr>
              <th className={th} style={{ color: C.gold, minWidth: "160px" }}>Name *</th>
              <th className={th} style={{ color: C.gold, minWidth: "220px" }}>Pathways path(s)</th>
              <th className={th} style={{ color: C.gold, minWidth: "70px" }}>Level</th>
              <th className={th} style={{ color: C.gold, minWidth: "130px" }}>Last attended</th>
              <th className={th} style={{ color: C.gold, minWidth: "130px" }}>100-day plan start</th>
              <th className={th} style={{ color: C.gold, minWidth: "200px" }}>Toastmasters goal</th>
              <th className={th} style={{ color: C.gold, minWidth: "170px" }}>Notes</th>
              <th className={th} style={{ color: C.gold, minWidth: "160px" }}>Buddy group(s)</th>
              <th className={th} style={{ color: C.gold, minWidth: "36px" }}></th>
            </tr>
          </thead>
          <tbody>
            {visibleRows.length === 0 && (
              <tr>
                <td colSpan={12} className="px-4 py-10 text-center" style={{ color: MUTED }}>
                  {rows.length === 0
                    ? "No members yet — click '+ Add member' to start."
                    : "No members match your search."}
                </td>
              </tr>
            )}
            {visibleRows.map((r, i) => (
              <tr key={r.id}
                style={{
                  backgroundColor: i % 2 === 0 ? "white" : C.paper,
                  borderBottom: `1px solid ${C.grayLine}`,
                }}>
                <td className={td}>
                  <input className={cell} style={inputStyle} value={r.name}
                    onChange={(e) => update(r.id, "name", e.target.value)}
                    placeholder="Full name" />
                </td>
                <td className={td}>
                  <PathsCell
                    value={r.paths || []}
                    onChange={(v) => update(r.id, "paths", v)}
                  />
                </td>
                <td className={td}>
                  <select className={cell} style={inputStyle} value={r.level}
                    onChange={(e) => update(r.id, "level", Number(e.target.value))}>
                    {[1, 2, 3, 4, 5].map((l) => <option key={l} value={l}>Level {l}</option>)}
                  </select>
                </td>
                <td className={td}>
                  <input type="date" className={cell} style={inputStyle} value={r.lastAttended || ""}
                    onChange={(e) => update(r.id, "lastAttended", e.target.value)} />
                </td>
                <td className={td}>
                  <input type="date" className={cell} style={inputStyle} value={r.onboardingStart || ""}
                    onChange={(e) => update(r.id, "onboardingStart", e.target.value)} />
                </td>
                <td className={td}>
                  <input className={cell} style={inputStyle} value={r.goal || ""}
                    onChange={(e) => update(r.id, "goal", e.target.value)}
                    placeholder="Main Toastmasters goal…" />
                </td>
                <td className={td}>
                  <input className={cell} style={inputStyle} value={r.notes || ""}
                    onChange={(e) => update(r.id, "notes", e.target.value)}
                    placeholder="Notes" />
                </td>
                <td className={td}>
                  <BuddyGroupCell
                    value={r._buddyGroupIds || []}
                    groups={data.buddyGroups || []}
                    onChange={(v) => update(r.id, "_buddyGroupIds", v)}
                  />
                </td>
                <td className={td} style={{ textAlign: "center" }}>
                  <button className="text-xs px-1" style={{ color: C.red }}
                    onClick={() => removeRow(r.id)} title="Remove member">✕</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Footer */}
      <div className="px-4 py-3 flex flex-wrap items-center gap-4 shrink-0"
        style={{ borderTop: `1px solid ${C.grayLine}`, backgroundColor: "white" }}>
        <span className="text-xs" style={{ color: MUTED }}>
          Click Pathways or Buddy group cells to open a multi-select dropdown. Buddy group changes apply to all group members.
        </span>
        <div className="ml-auto flex gap-2">
          <Btn kind="ghost" onClick={onClose}>Cancel</Btn>
          <Btn kind="maroon" onClick={handleSave}>Save all changes</Btn>
        </div>
      </div>
    </div>
  );
}

// ---------- Buddy groups admin panel ----------
function BuddyGroupsAdmin({ data, setData }) {
  const [open, setOpen] = useState(false);
  const [newName, setNewName] = useState("");
  const groups = data.buddyGroups || [];

  const addGroup = () => {
    const name = newName.trim();
    if (!name) return;
    setData((d) => ({ ...d, buddyGroups: [...(d.buddyGroups || []), { id: uid(), name, memberIds: [] }] }));
    setNewName("");
  };

  const updateGroup = (id, fn) =>
    setData((d) => ({ ...d, buddyGroups: (d.buddyGroups || []).map((g) => g.id === id ? fn(g) : g) }));

  const removeGroup = (id) => {
    if (window.confirm("Delete this buddy group?"))
      setData((d) => ({ ...d, buddyGroups: (d.buddyGroups || []).filter((g) => g.id !== id) }));
  };

  return (
    <div className="mt-6">
      <button
        onClick={() => setOpen((v) => !v)}
        className="flex items-center gap-2 text-sm font-semibold mb-3 w-full text-left"
        style={{ color: C.blueDeep }}>
        <span style={{ color: C.maroon }}>{open ? "▲" : "▼"}</span>
        Buddy Groups {open ? "(hide)" : `(${groups.length} group${groups.length !== 1 ? "s" : ""})`}
      </button>

      {open && (
        <div className="space-y-3">
          <Card className="p-4" accent={C.blue}>
            <h3 className="text-sm font-bold mb-3" style={{ color: C.blueDeep }}>Create a new group</h3>
            <div className="flex gap-2">
              <input className={`${inputCls} flex-1`} style={inputStyle}
                placeholder="Group name (e.g. Group A, Team Spark…)"
                value={newName}
                onChange={(e) => setNewName(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") addGroup(); }} />
              <Btn kind="maroon" onClick={addGroup} disabled={!newName.trim()}>Create</Btn>
            </div>
          </Card>

          {groups.length === 0 && (
            <p className="text-sm" style={{ color: MUTED }}>No buddy groups yet — create one above.</p>
          )}

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {groups.map((g) => {
              const groupMembers = g.memberIds
                .map((id) => data.members.find((m) => m.id === id))
                .filter(Boolean);
              const unassigned = data.members.filter((m) => !g.memberIds.includes(m.id));
              return (
                <Card key={g.id} className="p-4" accent={C.maroon}>
                  <div className="flex items-center justify-between mb-3">
                    <input
                      className="font-bold text-sm bg-transparent focus:outline-none focus:ring-1 rounded px-1"
                      style={{ color: C.blueDeep, border: "1px solid transparent", fontFamily: SERIF }}
                      value={g.name}
                      onChange={(e) => updateGroup(g.id, (grp) => ({ ...grp, name: e.target.value }))}
                    />
                    <button className="text-xs" style={{ color: "#8A958F" }} onClick={() => removeGroup(g.id)}>
                      Delete group
                    </button>
                  </div>

                  <div className="flex flex-wrap gap-1.5 mb-3 min-h-6">
                    {groupMembers.length === 0 && (
                      <span className="text-xs" style={{ color: MUTED }}>No members yet.</span>
                    )}
                    {groupMembers.map((m) => (
                      <span key={m.id} className="flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold"
                        style={{ backgroundColor: C.paper, color: C.blueDeep, border: `1px solid ${C.grayLine}` }}>
                        {m.name}
                        <button
                          onClick={() => updateGroup(g.id, (grp) => ({ ...grp, memberIds: grp.memberIds.filter((id) => id !== m.id) }))}
                          className="ml-0.5 leading-none" style={{ color: "#8A958F" }} aria-label="Remove">×</button>
                      </span>
                    ))}
                  </div>

                  {unassigned.length > 0 && (
                    <select
                      className={`${inputCls} text-xs`} style={inputStyle}
                      value=""
                      onChange={(e) => {
                        if (e.target.value)
                          updateGroup(g.id, (grp) => ({ ...grp, memberIds: [...grp.memberIds, e.target.value] }));
                      }}>
                      <option value="">Add member…</option>
                      {unassigned.sort((a, b) => a.name.localeCompare(b.name)).map((m) => (
                        <option key={m.id} value={m.id}>{m.name}</option>
                      ))}
                    </select>
                  )}
                </Card>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

// ---------- Views ----------
function HomeView({ data, go }) {
  const today = new Date().toISOString().slice(0, 10);
  const dormant = data.members.filter((m) => memberStatus(m).key === "dormant");
  const newbies = data.members.filter((m) => m.onboardingStart);
  const activeCycle = data.cycles.find((c) => c.id === currentCycleId(data.cycles));
  const unposted = data.recognitions.filter((r) => !r.posted);
  const upcomingMeetings = (data.meetingRota || [])
    .filter((r) => r.date >= today)
    .sort((a, b) => a.date.localeCompare(b.date))
    .slice(0, 5);

  const stats = [
    { label: "Active members", value: data.members.length, onClick: () => go("members") },
    { label: "Dormant / no attendance", value: dormant.length, color: dormant.length ? C.red : C.green, onClick: () => go("members", "dormant") },
    { label: "In 100-day plan", value: newbies.length, onClick: () => go("onboarding") },
    { label: "Unposted recognitions", value: unposted.length, color: unposted.length ? C.amber : C.green, onClick: () => go("recognition") },
  ];

  return (
    <div>
      <SectionTitle sub={`Today is ${new Date().toLocaleDateString(undefined, { weekday: "long", day: "numeric", month: "long", year: "numeric" })}`}>
        At a glance
      </SectionTitle>

      {/* North Star */}
      <div className="mb-6 rounded-lg p-5"
        style={{ backgroundColor: C.blue, borderLeft: `5px solid ${C.gold}` }}>
        <div className="text-xs tracking-widest font-bold mb-2" style={{ color: C.gold }}>★ NORTH STAR</div>
        <ol className="space-y-2">
          <li className="flex gap-3 text-sm text-white">
            <span className="font-bold" style={{ color: C.gold, fontFamily: SERIF }}>1</span>
            <span><span className="font-semibold">Every meeting feels worth attending</span> — members get something out of it each time.</span>
          </li>
          <li className="flex gap-3 text-sm text-white">
            <span className="font-bold" style={{ color: C.gold, fontFamily: SERIF }}>2</span>
            <span><span className="font-semibold">Members actively progress through Pathways</span> — not just showing up, but moving forward.</span>
          </li>
        </ol>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-6">
        {stats.map((s) => (
          <button key={s.label} onClick={s.onClick} className="text-left">
            <Card className="p-4 h-full hover:shadow-md transition-shadow">
              <div className="text-3xl font-bold" style={{ fontFamily: SERIF, color: s.color || C.blue }}>
                {s.value}
              </div>
              <div className="text-xs mt-1 font-semibold" style={{ color: "#5B6B73" }}>{s.label}</div>
            </Card>
          </button>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Card className="p-4" accent={C.gold}>
          <h3 className="font-bold text-sm mb-1" style={{ color: C.blueDeep }}>Current skill cycle</h3>
          {activeCycle ? (
            <>
              <div className="text-xl" style={{ fontFamily: SERIF, color: C.maroon }}>
                {activeCycle.name} <span className="text-sm" style={{ color: "#5B6B73" }}>({activeCycle.span})</span>
              </div>
              <div className="text-xs mt-2" style={{ color: "#5B6B73" }}>
                {activeCycle.tips.filter((t) => !t.posted).length} tip(s) still to post
              </div>
              <div className="mt-3"><Btn kind="ghost" onClick={() => go("cycles")}>Open cycle</Btn></div>
            </>
          ) : <p className="text-sm" style={{ color: "#5B6B73" }}>No cycle matches this month.</p>}
        </Card>

        <Card className="p-4" accent={C.blue}>
          <h3 className="font-bold text-sm mb-3" style={{ color: C.blueDeep }}>Upcoming Toastmasters</h3>
          {upcomingMeetings.length === 0 ? (
            <p className="text-sm" style={{ color: MUTED }}>No upcoming meetings in the rota yet.</p>
          ) : (
            <div className="space-y-2">
              {upcomingMeetings.map((mtg, i) => {
                const tm = (mtg.roles || {}).toastmaster;
                return (
                  <div key={mtg.id} className="flex items-center gap-3 py-1.5"
                    style={{ borderBottom: i < upcomingMeetings.length - 1 ? `1px solid ${C.grayLine}` : "none" }}>
                    <div className="shrink-0 text-center" style={{ minWidth: 56 }}>
                      <div className="text-xs font-bold" style={{ color: C.blueDeep }}>{fmtDate(mtg.date)}</div>
                    </div>
                    <div className="flex-1 text-sm font-semibold" style={{ color: tm ? C.maroon : MUTED }}>
                      {tm || <span className="italic font-normal">TBC</span>}
                    </div>
                    {i === 0 && <Badge fg={C.blueDeep} bg={C.gold}>Next</Badge>}
                  </div>
                );
              })}
            </div>
          )}
          <div className="mt-3"><Btn kind="ghost" onClick={() => go("rota")}>Manage rota</Btn></div>
        </Card>

        {dormant.length > 0 && (
          <Card className="p-4 lg:col-span-2" accent={C.red}>
            <h3 className="font-bold text-sm mb-2" style={{ color: C.red }}>Members to reach out to</h3>
            <div className="flex flex-wrap gap-2">
              {dormant.map((m) => (
                <Badge key={m.id} fg={C.red} bg={C.redBg}>
                  {m.name} · {daysSince(m.lastAttended) !== null ? `${daysSince(m.lastAttended)}d` : "no attendance"}
                </Badge>
              ))}
            </div>
          </Card>
        )}
      </div>
    </div>
  );
}

function MembersView({ data, setData, initialFilter }) {
  const [filter, setFilter] = useState(initialFilter || "all");
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null); // member object or "new"
  const [showBulk, setShowBulk] = useState(false);

  const filtered = useMemo(() => {
    return data.members
      .filter((m) => {
        const st = memberStatus(m).key;
        if (filter === "dormant" && st !== "dormant") return false;
        if (filter === "nudge" && st !== "nudge") return false;
        if (filter === "active" && st === "dormant") return false;
        if (filter === "nopath" && memberPaths(m).length > 0) return false;
        if (search && !m.name.toLowerCase().includes(search.toLowerCase())) return false;
        return true;
      })
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [data.members, filter, search]);

  const saveMember = (m) => {
    setData((d) => {
      const exists = d.members.some((x) => x.id === m.id);
      return { ...d, members: exists ? d.members.map((x) => (x.id === m.id ? m : x)) : [...d.members, m] };
    });
    setEditing(null);
  };

  const bulkSave = ({ members, buddyGroups }) => {
    setData((d) => ({ ...d, members, buddyGroups }));
    setShowBulk(false);
  };

  const deleteMember = (id) => {
    if (window.confirm("Remove this member from the dashboard?")) {
      setData((d) => ({ ...d, members: d.members.filter((m) => m.id !== id) }));
    }
  };

  const filters = [
    ["all", "All"], ["active", "Active"], ["nudge", "Needs nudge"], ["dormant", "Dormant"], ["nopath", "No path selected"],
  ];

  return (
    <div>
      <SectionTitle sub="Pathways progress, attendance, and roles for every member.">Members</SectionTitle>

      <div className="flex flex-wrap items-center gap-2 mb-4 px-4 py-3 rounded-lg"
        style={{ backgroundColor: C.paper, border: `1px solid ${C.grayLine}` }}>
        <span className="text-xs font-semibold" style={{ color: C.blueDeep }}>Quick links</span>
        <a href="https://toastmasterclub.org/view_meeting.php" target="_blank" rel="noopener noreferrer"
          className="text-xs px-3 py-1.5 rounded-full font-semibold"
          style={{ backgroundColor: C.blue, color: "white" }}>
          Sign up for meetings / request a speech ↗
        </a>
        <a href="https://apps.basecamp.toastmasters.org/dashboard" target="_blank" rel="noopener noreferrer"
          className="text-xs px-3 py-1.5 rounded-full font-semibold"
          style={{ backgroundColor: C.maroon, color: "white" }}>
          Mark speech as complete on Basecamp ↗
        </a>
        <span className="text-xs ml-auto" style={{ color: MUTED }}>
          After your speech or role, remember to mark it as complete on Basecamp.
        </span>
      </div>

      {(() => {
        const todayStr = new Date().toISOString().slice(0, 10);
        const nextMtg = (data.meetingRota || [])
          .filter((r) => r.date >= todayStr)
          .sort((a, b) => a.date.localeCompare(b.date))[0] || null;
        const upcomingEvts = (data.events || [])
          .filter((e) => e.date >= todayStr)
          .sort((a, b) => a.date.localeCompare(b.date));
        if (!nextMtg && upcomingEvts.length === 0) return null;
        return (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
            {nextMtg && (
              <Card className="p-4" accent={C.gold} style={{ backgroundColor: "#FBF5DA" }}>
                <div className="text-xs font-bold tracking-wide mb-1" style={{ color: C.blueDeep }}>NEXT MEETING</div>
                <div className="text-base font-bold mb-1" style={{ color: C.maroon, fontFamily: SERIF }}>
                  {fmtDate(nextMtg.date)}
                </div>
                <div className="text-xs" style={{ color: C.blueDeep }}>
                  Toastmaster of the Day:{" "}
                  <span className="font-bold" style={{ color: (nextMtg.roles || {}).toastmaster ? C.ink : MUTED }}>
                    {(nextMtg.roles || {}).toastmaster || "TBC"}
                  </span>
                </div>
              </Card>
            )}
            {upcomingEvts.length > 0 && (
              <Card className="p-0 overflow-hidden" style={{ gridColumn: nextMtg ? undefined : "1 / -1" }}>
                <div className="px-4 py-2" style={{ backgroundColor: C.blueDeep }}>
                  <span className="text-xs font-bold tracking-wide" style={{ color: "white" }}>UPCOMING IN THE CLUB</span>
                </div>
                <div className="divide-y" style={{ borderColor: C.grayLine }}>
                  {upcomingEvts.map((ev) => (
                    <div key={ev.id} className="px-4 py-3 flex items-start gap-4">
                      <div className="shrink-0 text-xs font-semibold pt-0.5 w-24" style={{ color: C.maroon }}>
                        {fmtDate(ev.date)}
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="font-semibold text-sm" style={{ color: C.blueDeep }}>{ev.title}</div>
                        {ev.detail && <p className="text-xs mt-0.5" style={{ color: MUTED }}>{ev.detail}</p>}
                      </div>
                    </div>
                  ))}
                </div>
              </Card>
            )}
          </div>
        );
      })()}

      <div className="flex flex-wrap items-center gap-2 mb-4">
        {filters.map(([k, label]) => (
          <button key={k} onClick={() => setFilter(k)}
            className="px-3 py-1.5 rounded-full text-xs font-semibold"
            style={{
              backgroundColor: filter === k ? C.blue : "white",
              color: filter === k ? "white" : C.blue,
              border: `1px solid ${filter === k ? C.blue : C.grayLine}`,
            }}>
            {label}
          </button>
        ))}
        <input className={`${inputCls} max-w-xs ml-auto`} style={inputStyle}
          placeholder="Search by name…" value={search} onChange={(e) => setSearch(e.target.value)} />
        <Btn kind="ghost" onClick={() => setShowBulk(true)}>Bulk edit</Btn>
        <Btn kind="maroon" onClick={() => setEditing("new")}>+ Add member</Btn>
      </div>

      {filtered.length === 0 ? (
        <Card className="p-8 text-center">
          <p className="text-sm" style={{ color: "#5B6B73" }}>
            {data.members.length === 0
              ? "No members yet. Add your first member to start tracking."
              : "No members match this filter."}
          </p>
        </Card>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-3">
          {filtered.map((m) => {
            const st = memberStatus(m);
            const daysAgo = daysSince(m.lastAttended);
            return (
              <Card key={m.id} className="p-4 flex flex-col gap-2"
                accent={st.key === "dormant" ? C.red : st.key === "ok" ? C.green : C.amber}>
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <div className="font-bold" style={{ color: C.blueDeep }}>{m.name}</div>
                    <div className="text-xs" style={{ color: "#5B6B73" }}>
                      {memberPaths(m).length > 0 ? (
                        `${memberPaths(m).join(", ")} · Level ${m.level}`
                      ) : (
                        <span className="font-semibold" style={{ color: C.amber }}>⚠ No Pathways path selected</span>
                      )}
                    </div>
                  </div>
                  <div className="flex flex-col items-end gap-1">
                    {m.isCommittee && (
                      <Badge fg={C.gold} bg={C.blueDeep}>★ Committee</Badge>
                    )}
                    <Badge fg={st.fg} bg={st.bg}>{st.key === "dormant" ? "● DORMANT" : st.label.split(" — ")[0]}</Badge>
                    {memberPaths(m).length === 0 && <Badge fg={C.amber} bg={C.amberBg}>No path</Badge>}
                  </div>
                </div>

                <div className="text-xs space-y-1" style={{ color: C.ink }}>
                  <div><span className="font-semibold">Last attended:</span> {fmtDate(m.lastAttended)}
                    {st.key === "dormant" && daysAgo !== null && <span style={{ color: C.red }}> ({daysAgo} days ago)</span>}
                  </div>
                </div>

                {m.roleLog && m.roleLog.length > 0 && (
                  <div className="text-xs space-y-0.5">
                    <span className="font-semibold" style={{ color: C.blueDeep }}>Recent roles:</span>
                    {[...m.roleLog].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 3).map((entry) => (
                      <div key={entry.id} className="flex gap-1">
                        <span style={{ color: C.ink }}>{entry.role}</span>
                        <span style={{ color: "#8A958F" }}>· {fmtDate(entry.date)}</span>
                      </div>
                    ))}
                  </div>
                )}

                {m.onboardingStart && (
                  <div className="mt-1 p-2 rounded" style={{ backgroundColor: C.paper }}>
                    <div className="text-xs font-bold mb-1" style={{ color: C.maroon }}>100-day plan</div>
                    <OnboardingBar startISO={m.onboardingStart} stages={m.customPlan} compact />
                  </div>
                )}

                {m.goal && (
                  <div className="text-xs italic" style={{ color: MUTED }}>
                    Goal: {m.goal}
                  </div>
                )}
                {m.notes && <p className={CLS_NOTE_P} style={{ color: MUTED }}>”{m.notes}”</p>}

                <div className="flex gap-2 mt-auto pt-2 flex-wrap">
                  <button
                    className="text-xs px-2 py-1 rounded-full font-semibold border"
                    style={{
                      backgroundColor: m.isCommittee ? C.blueDeep : "transparent",
                      color: m.isCommittee ? C.gold : MUTED,
                      borderColor: m.isCommittee ? C.blueDeep : C.grayLine,
                    }}
                    onClick={() => setData((d) => ({
                      ...d,
                      members: d.members.map((x) => x.id === m.id ? { ...x, isCommittee: !x.isCommittee } : x),
                    }))}>
                    {m.isCommittee ? "★ Committee" : "○ Committee"}
                  </button>
                  <Btn kind="ghost" onClick={() => setEditing(m)}>Edit</Btn>
                  <Btn kind="danger" onClick={() => deleteMember(m.id)}>Remove</Btn>
                </div>
              </Card>
            );
          })}
        </div>
      )}

      {editing && (
        <MemberModal
          initial={editing === "new" ? null : editing}
          onboardingTemplate={data.onboardingTemplate}
          onSave={saveMember}
          onClose={() => setEditing(null)}
        />
      )}
      {showBulk && (
        <BulkEditModal data={data} onSave={bulkSave} onClose={() => setShowBulk(false)} />
      )}

      <BuddyGroupsAdmin data={data} setData={setData} />
    </div>
  );
}

function OnboardingView({ data, setData }) {
  const newbies = data.members.filter((m) => m.onboardingStart);
  const [editingPlan, setEditingPlan] = useState(null); // member object or "template"
  const template = data.onboardingTemplate || ONBOARDING_STAGES;

  const savePlan = (stages) => {
    if (editingPlan === "template") {
      setData((d) => ({ ...d, onboardingTemplate: stages }));
    } else {
      setData((d) => ({
        ...d,
        members: d.members.map((m) => (m.id === editingPlan.id ? { ...m, customPlan: stages } : m)),
      }));
    }
    setEditingPlan(null);
  };

  const toggleStage = (memberId, stageIdx, planLength) => {
    setData((d) => ({
      ...d,
      members: d.members.map((m) => {
        if (m.id !== memberId) return m;
        const current = m.stagesDone || Array(planLength).fill(false);
        return { ...m, stagesDone: current.map((v, i) => (i === stageIdx ? !v : v)) };
      }),
    }));
  };

  return (
    <div>
      <SectionTitle sub="Each new member's journey from first visit to Icebreaker. Every plan can be tailored to the member.">
        100-Day Onboarding
      </SectionTitle>

      {/* Global template */}
      <Card className="p-4 mb-5" accent={C.blue}>
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="text-sm font-bold" style={{ color: C.blueDeep }}>Club default template</div>
            <div className="text-xs mt-0.5" style={{ color: MUTED }}>
              {data.onboardingTemplate ? "Custom template active" : "Using built-in Toastmasters stages"} — applied to all new members automatically.
            </div>
          </div>
          <div className="flex gap-2">
            {data.onboardingTemplate && (
              <Btn kind="danger" onClick={() => {
                if (window.confirm("Reset to the built-in Toastmasters plan?")) {
                  setData((d) => ({ ...d, onboardingTemplate: null }));
                }
              }}>Reset to default</Btn>
            )}
            <Btn kind="ghost" onClick={() => setEditingPlan("template")}>Edit template</Btn>
          </div>
        </div>
        <div className="mt-3">
          <OnboardingBar startISO={null} stages={template} compact />
        </div>
      </Card>

      {newbies.length === 0 ? (
        <Card className="p-8 text-center">
          <p className="text-sm" style={{ color: "#5B6B73" }}>
            No members are enrolled in a 100-day plan yet. Open any member record and enable “100-day plan” to get started.
          </p>
        </Card>
      ) : (
        <div className="space-y-4">
          {newbies.map((m) => (
            <Card key={m.id} className="p-4" accent={C.maroon}>
              <div className="flex flex-wrap items-baseline justify-between gap-2 mb-3">
                <div className="font-bold" style={{ color: C.blueDeep, fontFamily: SERIF }}>
                  {m.name}
                  {m.customPlan && m.customPlan.length > 0 && (
                    <span className="ml-2 align-middle"><Badge fg={C.maroon} bg="#F6E8EB">custom plan</Badge></span>
                  )}
                </div>
                <div className="flex items-center gap-3">
                  <span className="text-xs" style={{ color: "#5B6B73" }}>Started {fmtDate(m.onboardingStart)}</span>
                  <Btn kind="ghost" onClick={() => setEditingPlan(m)}>Edit plan</Btn>
                </div>
              </div>
              <OnboardingBar
                startISO={m.onboardingStart}
                stages={m.customPlan}
                stagesDone={m.stagesDone}
                onToggleStage={(i) => {
                  const plan = (m.customPlan && m.customPlan.length) ? m.customPlan : ONBOARDING_STAGES;
                  toggleStage(m.id, i, plan.length);
                }}
              />
            </Card>
          ))}
        </div>
      )}
      {editingPlan && (
        <PlanModal
          member={editingPlan === "template" ? { name: "Club default", customPlan: data.onboardingTemplate } : editingPlan}
          onSave={savePlan}
          onClose={() => setEditingPlan(null)}
        />
      )}
    </div>
  );
}

function CyclesView({ data, setData }) {
  const activeId = currentCycleId(data.cycles);
  const [open, setOpen] = useState(activeId);
  const [drafts, setDrafts] = useState({}); // {cycleId: {tip, action}}

  const setCycle = (id, fn) =>
    setData((d) => ({ ...d, cycles: d.cycles.map((c) => (c.id === id ? fn(c) : c)) }));

  const draft = (id, k) => drafts[id]?.[k] || "";
  const setDraft = (id, k, v) => setDrafts((p) => ({ ...p, [id]: { ...p[id], [k]: v } }));

  return (
    <div>
      <SectionTitle sub="Six-week skill cycles across the Toastmasters year. The active cycle is highlighted in gold.">
        Skill Cycles
      </SectionTitle>
      <div className="space-y-3">
        {data.cycles.map((c) => {
          const isActive = c.id === activeId;
          const isOpen = open === c.id;
          const unpostedTips = c.tips.filter((t) => !t.posted).length;
          const openActions = c.actions.filter((a) => !a.done).length;
          return (
            <Card key={c.id} accent={isActive ? C.gold : undefined}>
              <button onClick={() => setOpen(isOpen ? null : c.id)}
                className="w-full px-4 py-3 flex flex-wrap items-center gap-2 text-left"
                style={{ backgroundColor: isActive ? "#FBF5DA" : "white", borderRadius: "0.5rem" }}>
                <span className="text-lg" style={{ fontFamily: SERIF, color: C.maroon }}>{c.name}</span>
                <span className="text-xs font-semibold" style={{ color: "#5B6B73" }}>{c.span}</span>
                {isActive && <Badge fg={C.blueDeep} bg={C.gold}>ACTIVE NOW</Badge>}
                <span className="ml-auto text-xs" style={{ color: "#5B6B73" }}>
                  {unpostedTips > 0 && <span style={{ color: C.amber }}>{unpostedTips} tip(s) to post · </span>}
                  {openActions > 0 && <span>{openActions} open action(s) · </span>}
                  {isOpen ? "▲" : "▼"}
                </span>
              </button>

              {isOpen && (
                <div className="px-4 pb-4"
                  style={{ borderTop: `1px solid ${C.grayLine}` }}>
                  {/* Editable cycle name + span */}
                  <div className="flex flex-wrap gap-3 pt-3 pb-3 mb-1" style={{ borderBottom: `1px dashed ${C.grayLine}` }}>
                    <label className="flex items-center gap-2 text-xs font-semibold" style={{ color: C.blueDeep }}>
                      Theme name:
                      <input className={inputCls} style={{ ...inputStyle, maxWidth: 180 }}
                        value={c.name}
                        onChange={(e) => setCycle(c.id, (cy) => ({ ...cy, name: e.target.value }))} />
                    </label>
                    <label className="flex items-center gap-2 text-xs font-semibold" style={{ color: C.blueDeep }}>
                      Period:
                      <input className={inputCls} style={{ ...inputStyle, maxWidth: 120 }}
                        value={c.span}
                        onChange={(e) => setCycle(c.id, (cy) => ({ ...cy, span: e.target.value }))} />
                    </label>
                  </div>
                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
                  {/* Tips & articles */}
                  <div className="pt-3">
                    <h4 className="text-xs font-bold tracking-wide mb-2" style={{ color: C.blueDeep }}>TIPS & ARTICLES TO POST</h4>
                    <ul className="space-y-1.5 mb-2">
                      {c.tips.map((t) => (
                        <li key={t.id} className="flex items-start gap-2 text-sm">
                          <input type="checkbox" checked={t.posted} className="mt-0.5"
                            onChange={() => setCycle(c.id, (cy) => ({
                              ...cy, tips: cy.tips.map((x) => x.id === t.id ? { ...x, posted: !x.posted } : x),
                            }))} />
                          <span style={{
                            color: t.posted ? "#8A958F" : C.ink,
                            textDecoration: t.posted ? "line-through" : "none",
                          }}>{t.text}</span>
                          {!t.posted && <Badge fg={C.amber} bg={C.amberBg}>to post</Badge>}
                          <button className="ml-auto text-xs" style={{ color: "#8A958F" }}
                            onClick={() => setCycle(c.id, (cy) => ({ ...cy, tips: cy.tips.filter((x) => x.id !== t.id) }))}
                            aria-label="Delete tip">✕</button>
                        </li>
                      ))}
                      {c.tips.length === 0 && <li className="text-xs" style={{ color: "#8A958F" }}>Nothing queued yet.</li>}
                    </ul>
                    <div className="flex gap-2">
                      <input className={inputCls} style={inputStyle} placeholder="Add a tip or article…"
                        value={draft(c.id, "tip")} onChange={(e) => setDraft(c.id, "tip", e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter" && draft(c.id, "tip").trim()) {
                            setCycle(c.id, (cy) => ({ ...cy, tips: [...cy.tips, { id: uid(), text: draft(c.id, "tip").trim(), posted: false }] }));
                            setDraft(c.id, "tip", "");
                          }
                        }} />
                      <Btn kind="ghost" onClick={() => {
                        if (draft(c.id, "tip").trim()) {
                          setCycle(c.id, (cy) => ({ ...cy, tips: [...cy.tips, { id: uid(), text: draft(c.id, "tip").trim(), posted: false }] }));
                          setDraft(c.id, "tip", "");
                        }
                      }}>Add</Btn>
                    </div>
                  </div>

                  {/* VP PR notes */}
                  <div className="pt-3">
                    <h4 className="text-xs font-bold tracking-wide mb-2" style={{ color: C.blueDeep }}>VP PR COLLABORATION</h4>
                    <textarea rows={6} className={inputCls} style={inputStyle}
                      placeholder="Joint posts, social media plans, promo ideas…"
                      value={c.prNotes}
                      onChange={(e) => setCycle(c.id, (cy) => ({ ...cy, prNotes: e.target.value }))} />
                  </div>

                  {/* Actions */}
                  <div className="pt-3">
                    <h4 className="text-xs font-bold tracking-wide mb-2" style={{ color: C.blueDeep }}>MEETING ACTION ITEMS</h4>
                    <ul className="space-y-1.5 mb-2">
                      {c.actions.map((a) => (
                        <li key={a.id} className="flex items-start gap-2 text-sm">
                          <input type="checkbox" checked={a.done} className="mt-0.5"
                            onChange={() => setCycle(c.id, (cy) => ({
                              ...cy, actions: cy.actions.map((x) => x.id === a.id ? { ...x, done: !x.done } : x),
                            }))} />
                          <span style={{
                            color: a.done ? "#8A958F" : C.ink,
                            textDecoration: a.done ? "line-through" : "none",
                          }}>{a.text}</span>
                          <button className="ml-auto text-xs" style={{ color: "#8A958F" }}
                            onClick={() => setCycle(c.id, (cy) => ({ ...cy, actions: cy.actions.filter((x) => x.id !== a.id) }))}
                            aria-label="Delete action">✕</button>
                        </li>
                      ))}
                      {c.actions.length === 0 && <li className="text-xs" style={{ color: "#8A958F" }}>No actions yet.</li>}
                    </ul>
                    <div className="flex gap-2">
                      <input className={inputCls} style={inputStyle} placeholder="Add an action item…"
                        value={draft(c.id, "action")} onChange={(e) => setDraft(c.id, "action", e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter" && draft(c.id, "action").trim()) {
                            setCycle(c.id, (cy) => ({ ...cy, actions: [...cy.actions, { id: uid(), text: draft(c.id, "action").trim(), done: false }] }));
                            setDraft(c.id, "action", "");
                          }
                        }} />
                      <Btn kind="ghost" onClick={() => {
                        if (draft(c.id, "action").trim()) {
                          setCycle(c.id, (cy) => ({ ...cy, actions: [...cy.actions, { id: uid(), text: draft(c.id, "action").trim(), done: false }] }));
                          setDraft(c.id, "action", "");
                        }
                      }}>Add</Btn>
                    </div>
                  </div>
                  </div>
                </div>
              )}
            </Card>
          );
        })}
      </div>
    </div>
  );
}

function RecognitionView({ data, setData }) {
  const [form, setForm] = useState({ type: RECOGNITION_TYPES[0], member: "", detail: "" });
  const [guestName, setGuestName] = useState("");
  const unposted = data.recognitions.filter((r) => !r.posted);

  const effectiveMember = form.member === "__guest__" ? guestName.trim() : form.member.trim();

  const add = () => {
    if (!effectiveMember) return;
    setData((d) => ({
      ...d,
      recognitions: [...d.recognitions, { id: uid(), ...form, member: effectiveMember, posted: false }],
    }));
    setForm({ type: RECOGNITION_TYPES[0], member: "", detail: "" });
    setGuestName("");
  };

  const togglePosted = (id) =>
    setData((d) => ({ ...d, recognitions: d.recognitions.map((r) => r.id === id ? { ...r, posted: !r.posted } : r) }));

  const remove = (id) =>
    setData((d) => ({ ...d, recognitions: d.recognitions.filter((r) => r.id !== id) }));

  const clearAll = () => {
    if (window.confirm("Clear all recognitions for the next meeting? This can't be undone (export JSON first if you want a record).")) {
      setData((d) => ({ ...d, recognitions: [] }));
    }
  };

  const copyLLMPrompt = () => {
    const list = data.recognitions.filter((r) => !r.posted);
    if (!list.length) return;
    const items = list.map((r, i) => {
      const detail = r.detail ? `: ${r.detail}` : "";
      return `${i + 1}. ${r.member} – ${r.type}${detail}`;
    }).join("\n");
    const prompt = `## Task\nWrite Game Changers recognition posts for a Toastmasters club meeting.\n\n## Instructions\n- Write one short, enthusiastic post per achievement (2–3 sentences each)\n- Address each member by first name\n- Keep the tone warm, celebratory, and encouraging — suitable for a club group chat\n- Use plain text only (no markdown, no bullet points, no headers in the output)\n- Separate each post with a blank line\n- Label each post with the member's name in bold at the top\n\n## Achievements to recognise\n${items}\n\n## Output format\nReturn only the finished posts, ready to copy-paste into the group chat. Do not include any explanation or preamble.`;
    navigator.clipboard.writeText(prompt).then(() => alert("Prompt copied — paste it into Microsoft Copilot to generate your recognition posts!"));
  };

  const needsDetail = form.type === "First time in a role" || form.type === "Pathways level completion";

  return (
    <div>
      <SectionTitle sub="Capture wins at each meeting, post them to Game Changers, then clear the slate.">
        Meeting Recognition
      </SectionTitle>

      {unposted.length > 0 && (
        <div className="mb-4 px-4 py-3 rounded-lg flex flex-wrap items-center gap-3"
          style={{ backgroundColor: C.amberBg, border: `1px solid ${C.amber}` }}>
          <span className="text-xl">⚠️</span>
          <span className="text-sm font-semibold" style={{ color: C.amber }}>
            {unposted.length} recognition{unposted.length > 1 ? "s" : ""} not yet posted to Game Changers.
          </span>
          <Btn kind="ghost" className="ml-auto" onClick={copyLLMPrompt}
            title="Copy a prompt you can paste into Microsoft Copilot to generate recognition posts">
            ✨ Copy Copilot prompt
          </Btn>
        </div>
      )}

      <Card className="p-4 mb-5" accent={C.maroon}>
        <h3 className="text-sm font-bold mb-3" style={{ color: C.blueDeep }}>Add recognition</h3>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
          <Field label="Type">
            <select className={inputCls} style={inputStyle} value={form.type}
              onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}>
              {RECOGNITION_TYPES.map((t) => <option key={t}>{t}</option>)}
            </select>
          </Field>
          <Field label="Member">
            {data.members.length > 0 ? (
              <select className={inputCls} style={inputStyle} value={form.member}
                onChange={(e) => setForm((f) => ({ ...f, member: e.target.value }))}>
                <option value="">Choose member…</option>
                {data.members.map((m) => <option key={m.id} value={m.name}>{m.name}</option>)}
                <option value="__guest__">Guest / other (type below)</option>
              </select>
            ) : (
              <input className={inputCls} style={inputStyle} value={form.member}
                onChange={(e) => setForm((f) => ({ ...f, member: e.target.value }))} placeholder="Member name" />
            )}
          </Field>
        </div>
        {form.member === "__guest__" && (
          <div className="mb-3">
            <input className={`${inputCls} max-w-xs`} style={inputStyle} placeholder="Guest / other name"
              value={guestName}
              onChange={(e) => setGuestName(e.target.value)} />
          </div>
        )}
        <Field label={needsDetail ? (form.type === "First time in a role" ? "Which role?" : "Which level / path?") : "Detail (optional)"}>
          <textarea className={inputCls} style={{ ...inputStyle, width: "100%", minHeight: 72, resize: "vertical" }}
            value={form.detail}
            onChange={(e) => setForm((f) => ({ ...f, detail: e.target.value }))}
            placeholder={form.type === "First time in a role" ? "e.g. General Evaluator" :form.type === "Contest winner" ? "e.g. International Speech Contest — 1st place" : "Add any extra detail…"} />
        </Field>
        <div className="mt-3"><Btn kind="maroon" onClick={add} disabled={!effectiveMember}>Add recognition</Btn></div>
      </Card>

      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-bold" style={{ color: C.blueDeep }}>
          This meeting cycle ({data.recognitions.length})
        </h3>
        <Btn kind="danger" onClick={clearAll} disabled={data.recognitions.length === 0}>
          Clear all for next meeting
        </Btn>
      </div>

      {data.recognitions.length === 0 ? (
        <Card className="p-8 text-center">
          <p className="text-sm" style={{ color: "#5B6B73" }}>A clean slate. Add recognitions as the meeting happens.</p>
        </Card>
      ) : (
        <div className="space-y-2">
          {data.recognitions.map((r) => (
            <Card key={r.id} className="px-4 py-3 flex flex-wrap items-center gap-3"
              accent={r.posted ? C.green : C.amber}>
              <div className="min-w-0">
                <div className="text-sm font-bold" style={{ color: C.blueDeep }}>
                  {r.member} <span className="font-normal" style={{ color: "#5B6B73" }}>— {r.type}</span>
                </div>
                {r.detail && <div className="text-xs" style={{ color: "#5B6B73" }}>{r.detail}</div>}
              </div>
              <div className="ml-auto flex items-center gap-3">
                {r.posted
                  ? <Badge fg={C.green} bg={C.greenBg}>✓ Posted</Badge>
                  : <Badge fg={C.amber} bg={C.amberBg}>Not posted</Badge>}
                <label className="flex items-center gap-1.5 text-xs font-semibold" style={{ color: C.blueDeep }}>
                  <input type="checkbox" checked={r.posted} onChange={() => togglePosted(r.id)} />
                  Game Changers
                </label>
                <button className="text-xs" style={{ color: "#8A958F" }} onClick={() => remove(r.id)} aria-label="Delete">✕</button>
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Education Goals ----------
function DTMView({ data, setData }) {
  const currentYear = new Date().getFullYear();
  const [newLevel, setNewLevel] = useState("");
  const [newTarget, setNewTarget] = useState("");

  const yearGoals = (data.educationGoals || [])
    .filter((g) => g.year === currentYear)
    .sort((a, b) => a.level - b.level);

  const setCompleted = (id, n) =>
    setData((d) => ({
      ...d,
      educationGoals: (d.educationGoals || []).map((g) =>
        g.id === id ? { ...g, completed: Math.max(0, n) } : g
      ),
    }));

  const saveGoal = () => {
    const lvl = parseInt(newLevel);
    const tgt = parseInt(newTarget);
    if (!lvl || !tgt || tgt < 1) return;
    setData((d) => {
      const existing = (d.educationGoals || []).find((g) => g.year === currentYear && g.level === lvl);
      if (existing) {
        return {
          ...d,
          educationGoals: (d.educationGoals || []).map((g) =>
            g.id === existing.id ? { ...g, target: tgt } : g
          ),
        };
      }
      return {
        ...d,
        educationGoals: [
          ...(d.educationGoals || []),
          { id: uid(), year: currentYear, level: lvl, target: tgt, completed: 0, done: false, doneDate: "" },
        ],
      };
    });
    setNewLevel("");
    setNewTarget("");
  };

  const toggleDone = (id, currentDone) => {
    setData((d) => ({
      ...d,
      educationGoals: (d.educationGoals || []).map((g) =>
        g.id === id
          ? { ...g, done: !currentDone, doneDate: !currentDone ? new Date().toISOString().slice(0, 10) : "" }
          : g
      ),
    }));
  };

  const removeGoal = (id) =>
    setData((d) => ({ ...d, educationGoals: (d.educationGoals || []).filter((g) => g.id !== id) }));

  return (
    <div>
      <SectionTitle sub="Track DCP progress and see who is on each Pathways level.">
        Education Goals
      </SectionTitle>

      {/* DCP Progress link */}
      <div className="mb-5">
        <a
          href="https://dashboards.toastmasters.org/ClubReport.aspx?id=05101315"
          target="_blank"
          rel="noopener noreferrer"
          className="inline-flex items-center gap-2 px-4 py-2 rounded-md text-sm font-semibold"
          style={{ backgroundColor: C.blueDeep, color: "white" }}>
          View DCP Progress on Toastmasters.org ↗
        </a>
      </div>

      {/* Level goals */}
      <Card className="p-4 mb-6">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: C.blueDeep }}>
          LEVEL GOALS — {currentYear}
        </div>

        {yearGoals.length === 0 && (
          <p className="text-sm mb-4" style={{ color: MUTED }}>No level goals set for this year yet.</p>
        )}

        {yearGoals.length > 0 && (
          <div className="mb-4 space-y-3">
            {yearGoals.map((g) => {
              const completed = g.completed || 0;
              const hit = completed >= g.target;
              const pct = g.target > 0 ? Math.min(100, Math.round((completed / g.target) * 100)) : 0;
              return (
                <div key={g.id} className="rounded-lg p-3 flex flex-wrap items-center gap-3"
                  style={{
                    backgroundColor: g.done ? "#F0FAF3" : hit ? "#F5FAF0" : C.paper,
                    border: `1px solid ${g.done || hit ? C.green : C.grayLine}`,
                  }}>
                  <div className="flex-1 min-w-40">
                    <div className="text-sm font-bold mb-1" style={{ color: C.blueDeep }}>
                      Level {g.level}
                    </div>
                    <div className="flex items-center gap-2 mb-1">
                      <label className="text-xs" style={{ color: MUTED }}>Completed:</label>
                      <input
                        type="number" min="0"
                        className={inputCls}
                        style={{ ...inputStyle, width: 60, padding: "2px 6px", fontSize: 13 }}
                        value={completed}
                        onChange={(e) => setCompleted(g.id, parseInt(e.target.value) || 0)}
                      />
                      <span className="text-xs" style={{ color: MUTED }}>/ {g.target} target</span>
                      {hit && !g.done && (
                        <span className="text-xs font-semibold" style={{ color: C.green }}>Target reached!</span>
                      )}
                    </div>
                    <div className="h-1.5 rounded-full overflow-hidden" style={{ backgroundColor: C.grayLine }}>
                      <div className="h-full rounded-full transition-all"
                        style={{
                          width: `${pct}%`,
                          backgroundColor: g.done || hit ? C.green : C.blue,
                        }} />
                    </div>
                  </div>
                  <div className="flex items-center gap-3">
                    {g.done ? (
                      <div className="flex items-center gap-2">
                        <span className="text-xs font-bold px-2 py-1 rounded-full"
                          style={{ backgroundColor: C.green, color: "white" }}>
                          ✓ Complete{g.doneDate ? ` · ${fmtDate(g.doneDate)}` : ""}
                        </span>
                        <button className="text-xs underline" style={{ color: MUTED }}
                          onClick={() => toggleDone(g.id, true)}>Undo</button>
                      </div>
                    ) : (
                      <Btn kind="primary" onClick={() => toggleDone(g.id, false)}>Mark complete</Btn>
                    )}
                    <button className="text-xs px-2 py-1 rounded"
                      style={{ color: C.red, border: `1px solid ${C.grayLine}` }}
                      onClick={() => removeGoal(g.id)}>Remove</button>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* Set / update goal */}
        <div className="flex flex-wrap items-end gap-3 pt-3" style={{ borderTop: `1px dashed ${C.grayLine}` }}>
          <Field label="Level">
            <select className={inputCls} style={inputStyle} value={newLevel}
              onChange={(e) => {
                const lvl = parseInt(e.target.value);
                setNewLevel(e.target.value);
                const existing = yearGoals.find((g) => g.level === lvl);
                setNewTarget(existing ? String(existing.target) : "");
              }}>
              <option value="">Pick level…</option>
              <option value="1">Level 1</option>
              <option value="2">Level 2</option>
              <option value="3">Level 3</option>
              <option value="4">Level 4</option>
              <option value="5">Level 5</option>
            </select>
          </Field>
          <Field label="Target (members)">
            <input type="number" min="1" className={inputCls} style={{ ...inputStyle, width: 100 }}
              value={newTarget} onChange={(e) => setNewTarget(e.target.value)}
              placeholder="e.g. 3" />
          </Field>
          <Btn kind="maroon" onClick={saveGoal} disabled={!newLevel || !newTarget}>
            {yearGoals.some((g) => g.level === parseInt(newLevel)) ? "Update goal" : "Set goal"}
          </Btn>
        </div>
      </Card>

      {/* Members by level */}
      <Card className="p-4 mb-6">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: C.blueDeep }}>MEMBERS BY CURRENT LEVEL</div>
        <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
          {[1, 2, 3, 4, 5].map((lvl) => {
            const atLevel = data.members.filter((m) => m.level === lvl);
            return (
              <div key={lvl} className="rounded-lg p-3" style={{ backgroundColor: C.amberBg, border: `1px solid ${C.grayLine}` }}>
                <div className="text-lg font-bold" style={{ color: C.blueDeep, fontFamily: SERIF }}>{atLevel.length}</div>
                <div className="text-xs font-semibold mb-2" style={{ color: C.blueDeep }}>Level {lvl}</div>
                {atLevel.length > 0 ? (
                  <div className="text-xs space-y-0.5">
                    {atLevel.map((m) => (
                      <div key={m.id} className="truncate" style={{ color: C.ink }}>{m.name}</div>
                    ))}
                  </div>
                ) : (
                  <div className="text-xs" style={{ color: "#8A958F" }}>—</div>
                )}
              </div>
            );
          })}
        </div>
      </Card>
    </div>
  );
}

// ---------- Meeting Rota ----------
function RotaView({ data, setData }) {
  const today = new Date().toISOString().slice(0, 10);
  const rota = (data.meetingRota || []).slice().sort((a, b) => a.date.localeCompare(b.date));
  const upcoming = rota.filter((r) => r.date >= today);
  const next = upcoming[0] || null;
  const memberNames = data.members.map((m) => m.name).sort();

  const [bulkStart, setBulkStart] = useState("");
  const [bulkText, setBulkText] = useState("");
  const [bulkPreview, setBulkPreview] = useState(null);

  const setToastmaster = (id, value) =>
    setData((d) => ({
      ...d,
      meetingRota: (d.meetingRota || []).map((r) =>
        r.id === id ? { ...r, roles: { ...(r.roles || {}), toastmaster: value } } : r
      ),
    }));

  const setMeetingDate = (id, date) =>
    setData((d) => ({
      ...d,
      meetingRota: (d.meetingRota || []).map((r) => r.id === id ? { ...r, date } : r),
    }));

  const removeMeeting = (id) =>
    setData((d) => ({ ...d, meetingRota: (d.meetingRota || []).filter((r) => r.id !== id) }));

  const genMeetingDates = (startDate, count) => {
    const dates = [];
    const d = new Date(startDate + "T00:00:00");
    for (let i = 0; i < count; i++) {
      dates.push(d.toISOString().slice(0, 10));
      const day = d.getDay();
      // Mon(1)→next Tue: +8 days; Tue(2)→next Mon: +6 days; fallback +7
      d.setDate(d.getDate() + (day === 1 ? 8 : day === 2 ? 6 : 7));
    }
    return dates;
  };

  const buildBulkPreview = (startDate, text) => {
    if (!startDate || !text.trim()) return null;
    const names = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    if (names.length === 0) return null;
    return genMeetingDates(startDate, names.length).map((date, i) => ({ date, name: names[i] }));
  };

  const applyBulk = () => {
    if (!bulkPreview || bulkPreview.length === 0) return;
    setData((d) => {
      let rota = [...(d.meetingRota || [])];
      bulkPreview.forEach(({ date, name }) => {
        if (rota.some((r) => r.date === date)) {
          rota = rota.map((r) => r.date === date ? { ...r, roles: { ...(r.roles || {}), toastmaster: name } } : r);
        } else {
          rota = [...rota, { id: uid(), date, theme: "", roles: { toastmaster: name } }];
        }
      });
      return { ...d, meetingRota: rota };
    });
    setBulkText("");
    setBulkStart("");
    setBulkPreview(null);
  };

  return (
    <div>
      <datalist id="rota-member-names">
        {memberNames.map((n) => <option key={n} value={n} />)}
      </datalist>

      <SectionTitle sub="Who is running the next meeting?">Meeting Rota</SectionTitle>

      {/* Next meeting highlight */}
      {next && (
        <Card className="p-5 mb-6" accent={C.gold} style={{ backgroundColor: "#FBF5DA" }}>
          <div className="text-xs font-bold tracking-wide mb-1" style={{ color: C.blueDeep }}>NEXT MEETING</div>
          <div className="text-2xl font-bold mb-2" style={{ color: C.maroon, fontFamily: SERIF }}>
            {fmtDate(next.date)}
          </div>
          <div className="flex items-center gap-2">
            <span className="text-sm font-semibold" style={{ color: C.blueDeep }}>Toastmaster of the Day:</span>
            <span className="text-sm font-bold" style={{ color: (next.roles || {}).toastmaster ? C.ink : MUTED }}>
              {(next.roles || {}).toastmaster || "TBC"}
            </span>
          </div>
        </Card>
      )}

      {/* Bulk schedule Toastmasters — primary workflow */}
      <Card className="p-4 mb-6" accent={C.blue}>
        <div className="text-xs font-bold tracking-wide mb-1" style={{ color: C.blueDeep }}>
          BULK SCHEDULE TOASTMASTERS
        </div>
        <p className="text-xs mb-3" style={{ color: MUTED }}>
          Pick the first meeting date, paste one name per line. Meetings are automatically created on the Mon / Tue alternating schedule.
        </p>
        <div className="mb-3">
          <Field label="First meeting date">
            <input type="date" className={inputCls} style={{ ...inputStyle, maxWidth: 180 }}
              value={bulkStart}
              onChange={(e) => {
                setBulkStart(e.target.value);
                setBulkPreview(buildBulkPreview(e.target.value, bulkText));
              }} />
          </Field>
        </div>
        <div className="flex flex-wrap gap-4 items-start">
          <div className="flex-1 min-w-48">
            <textarea
              className={inputCls}
              style={{ ...inputStyle, width: "100%", minHeight: 120, fontFamily: "monospace", resize: "vertical" }}
              value={bulkText}
              onChange={(e) => {
                setBulkText(e.target.value);
                setBulkPreview(buildBulkPreview(bulkStart, e.target.value));
              }}
              placeholder={"Alice\nBob\nCarol\n…"} />
            <div className="mt-2 flex gap-2">
              <Btn kind="primary" onClick={applyBulk} disabled={!bulkPreview || bulkPreview.length === 0}>
                Apply assignments
              </Btn>
              {bulkPreview && (
                <Btn kind="ghost" onClick={() => { setBulkText(""); setBulkStart(""); setBulkPreview(null); }}>Clear</Btn>
              )}
            </div>
          </div>
          {bulkPreview && bulkPreview.length > 0 && (
            <div className="flex-1 min-w-48">
              <div className="text-xs font-bold mb-2" style={{ color: C.blueDeep }}>Preview</div>
              <div className="space-y-1">
                {bulkPreview.map((p) => (
                  <div key={p.date} className="flex items-center gap-2 text-xs py-1"
                    style={{ borderBottom: `1px solid ${C.grayLine}` }}>
                    <span className="w-28 shrink-0 font-semibold" style={{ color: C.ink }}>{fmtDate(p.date)}</span>
                    <span style={{ color: C.blueDeep }}>{p.name}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </Card>

      {/* Upcoming meetings — editable */}
      {upcoming.length > 0 && (
        <Card className="p-0 mb-6 overflow-hidden">
          <div className="px-4 py-3" style={{ backgroundColor: C.blueDeep }}>
            <span className="text-xs font-bold tracking-wide" style={{ color: "white" }}>
              UPCOMING MEETINGS — {upcoming.length} scheduled
            </span>
          </div>
          <div>
            {upcoming.map((mtg, i) => (
              <div key={mtg.id}
                className="flex flex-wrap items-center gap-3 px-4 py-3"
                style={{
                  borderBottom: i < upcoming.length - 1 ? `1px solid ${C.grayLine}` : "none",
                  backgroundColor: i % 2 === 0 ? "white" : "#FAFAF8",
                }}>
                <span className="text-xs font-bold w-5 shrink-0 text-center"
                  style={{ color: MUTED }}>
                  {i + 1}
                </span>
                <input
                  type="date"
                  className={inputCls}
                  style={{ ...inputStyle, width: 150 }}
                  value={mtg.date}
                  onChange={(e) => setMeetingDate(mtg.id, e.target.value)} />
                <span className="text-xs font-semibold shrink-0 w-8" style={{ color: MUTED }}>
                  {mtg.date ? new Date(mtg.date + "T00:00:00").toLocaleDateString("en-US", { weekday: "short" }) : ""}
                </span>
                <input
                  className={inputCls}
                  list="rota-member-names"
                  style={{ ...inputStyle, flex: 1, minWidth: 140 }}
                  value={(mtg.roles || {}).toastmaster || ""}
                  onChange={(e) => setToastmaster(mtg.id, e.target.value)}
                  placeholder="Toastmaster of the Day…" />
                <button className="text-xs shrink-0 px-2 py-1 rounded"
                  style={{ color: C.red, border: `1px solid ${C.grayLine}` }}
                  onClick={() => removeMeeting(mtg.id)}>Remove</button>
              </div>
            ))}
          </div>
        </Card>
      )}
    </div>
  );
}
// ---------- Events admin view ----------
function EventsView({ data, setData }) {
  const today = new Date().toISOString().slice(0, 10);
  const events = (data.events || []).slice().sort((a, b) => a.date.localeCompare(b.date));
  const upcoming = events.filter((e) => e.date >= today);
  const past = events.filter((e) => e.date < today);
  const [showPastEvt, setShowPastEvt] = useState(false);
  const [form, setForm] = useState({ title: "", detail: "", date: today });

  const addEvent = () => {
    const { title, detail, date } = form;
    if (!title.trim() || !date) return;
    setData((d) => ({
      ...d,
      events: [...(d.events || []), { id: uid(), title: title.trim(), detail: detail.trim(), date, dismissed: [] }],
    }));
    setForm({ title: "", detail: "", date: today });
  };

  const removeEvent = (id) => {
    if (window.confirm("Remove this event notification?"))
      setData((d) => ({ ...d, events: (d.events || []).filter((e) => e.id !== id) }));
  };

  const EventCard = ({ ev, faded }) => (
    <Card className="p-4" accent={faded ? C.grayLine : C.gold}
      style={{ opacity: faded ? 0.6 : 1 }}>
      <div className="flex items-start justify-between gap-2">
        <div>
          <div className="font-bold text-sm" style={{ color: C.blueDeep }}>{ev.title}</div>
          <div className="text-xs mt-0.5" style={{ color: MUTED }}>{fmtDate(ev.date)}</div>
          {ev.detail && <p className="text-sm mt-1" style={{ color: C.ink }}>{ev.detail}</p>}
        </div>
        <button className="text-xs shrink-0 px-2 py-0.5 rounded"
          style={{ color: C.red, border: `1px solid ${C.grayLine}` }}
          onClick={() => removeEvent(ev.id)}>Remove</button>
      </div>
    </Card>
  );

  return (
    <div>
      <SectionTitle sub="Announcements and events shown as pop-up notifications to members in the member portal.">
        Events
      </SectionTitle>

      <Card className="p-4 mb-6" accent={C.gold}>
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: C.blueDeep }}>ADD EVENT</div>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <Field label="Title">
            <input className={inputCls} style={inputStyle} value={form.title}
              onChange={(e) => setForm((f) => ({ ...f, title: e.target.value }))}
              placeholder="Event title…" />
          </Field>
          <Field label="Date">
            <input type="date" className={inputCls} style={inputStyle} value={form.date}
              onChange={(e) => setForm((f) => ({ ...f, date: e.target.value }))} />
          </Field>
          <Field label="Details (optional)">
            <div className="flex gap-2">
              <input className={`${inputCls} flex-1`} style={inputStyle} value={form.detail}
                onChange={(e) => setForm((f) => ({ ...f, detail: e.target.value }))}
                onKeyDown={(e) => { if (e.key === "Enter") addEvent(); }}
                placeholder="Extra info…" />
              <Btn kind="maroon" onClick={addEvent} disabled={!form.title.trim() || !form.date}>Add</Btn>
            </div>
          </Field>
        </div>
      </Card>

      {upcoming.length === 0 && past.length === 0 && (
        <Card className="p-8 text-center">
          <p className="text-sm" style={{ color: MUTED }}>No events yet — add one above.</p>
        </Card>
      )}

      {upcoming.length > 0 && (
        <div className="space-y-3 mb-6">
          <div className="text-sm font-semibold" style={{ color: C.blueDeep }}>Upcoming ({upcoming.length})</div>
          {upcoming.map((ev) => <EventCard key={ev.id} ev={ev} />)}
        </div>
      )}

      {past.length > 0 && (
        <div>
          <button className="flex items-center gap-2 text-sm font-semibold mb-3"
            style={{ color: MUTED }} onClick={() => setShowPastEvt((v) => !v)}>
            <span>{showPastEvt ? "▲" : "▼"}</span>
            Past events ({past.length})
          </button>
          {showPastEvt && (
            <div className="space-y-3">
              {past.map((ev) => <EventCard key={ev.id} ev={ev} faded />)}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ---------- Main app ----------
const NAV = [
  { key: "home", label: "Home", icon: "⌂" },
  { key: "members", label: "Members", icon: "👥" },
  { key: "onboarding", label: "100-Day", icon: "🗓" },
  { key: "cycles", label: "Skill Cycles", short: "Cycles", icon: "🔄" },
  { key: "recognition", label: "Recognition", short: "Awards", icon: "🏆" },
  { key: "dtm", label: "Education Goals", short: "Goals", icon: "🎖" },
  { key: "rota", label: "Meeting Rota", short: "Rota", icon: "📋" },
  { key: "events", label: "Events", icon: "📣" },
];

const LS_KEY = "vpe-dashboard-data";

const loadFromStorage = () => {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    const base = emptyData();
    return {
      ...base,
      ...parsed,
      members: Array.isArray(parsed.members) ? parsed.members : [],
      cycles: Array.isArray(parsed.cycles) && parsed.cycles.length === 6 ? parsed.cycles : base.cycles,
      recognitions: Array.isArray(parsed.recognitions) ? parsed.recognitions : [],
      weeks: Array.isArray(parsed.weeks) ? parsed.weeks : [],
      educationGoals: Array.isArray(parsed.educationGoals) ? parsed.educationGoals : [],
      buddyGroups: Array.isArray(parsed.buddyGroups) ? parsed.buddyGroups : [],
      meetingRota: Array.isArray(parsed.meetingRota) ? parsed.meetingRota : [],
      events: Array.isArray(parsed.events) ? parsed.events : [],
    };
  } catch {
    return null;
  }
};

const WRITE_TOKEN = process.env.REACT_APP_WRITE_TOKEN || "";

// Only attempt cloud sync when deployed to Vercel (not local dev)
const IS_CLOUD = (() => {
  try {
    const h = window.location.hostname;
    return h !== "localhost" && h !== "127.0.0.1" && !h.startsWith("192.168.");
  } catch { return false; }
})();

async function cloudLoad() {
  if (!IS_CLOUD) return null;
  try {
    const res = await fetch("/api/data");
    if (!res.ok) return null;
    return await res.json();
  } catch {
    return null;
  }
}

async function cloudSave(data) {
  if (!IS_CLOUD) return true; // silently succeed in local dev
  try {
    const res = await fetch("/api/data", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-write-token": WRITE_TOKEN },
      body: JSON.stringify(data),
    });
    return res.ok;
  } catch {
    return false;
  }
}

export default function VPEDashboard() {
  const [data, setData] = useState(() => loadFromStorage());
  const [appMode, setAppMode] = useState("admin"); // "admin" | "member"
  const [adminUnlocked, setAdminUnlocked] = useState(
    () => !!sessionStorage.getItem(SS_UNLOCKED_KEY)
  );
  const [view, setView] = useState("home");
  const [memberFilter, setMemberFilter] = useState("all");
  const [toast, setToast] = useState("");
  const [syncStatus, setSyncStatus] = useState("idle"); // "loading"|"ok"|"error"|"idle"
  const fileRef = useRef(null);
  const syncTimer = useRef(null);

  // Load from cloud on first mount
  useEffect(() => {
    setSyncStatus("loading");
    cloudLoad().then((cloudData) => {
      if (cloudData) {
        setData(cloudData);
        try { localStorage.setItem(LS_KEY, JSON.stringify(cloudData)); } catch {}
      }
      setSyncStatus(cloudData ? "ok" : "idle");
    });
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Save to localStorage immediately; debounce cloud sync by 3 s
  useEffect(() => {
    if (!data) return;
    try { localStorage.setItem(LS_KEY, JSON.stringify(data)); } catch {}
    clearTimeout(syncTimer.current);
    syncTimer.current = setTimeout(() => {
      setSyncStatus("loading");
      cloudSave(data).then((ok) => setSyncStatus(ok ? "ok" : "error"));
    }, 3000);
  }, [data]);

  const notify = (msg) => { setToast(msg); setTimeout(() => setToast(""), 2600); };

  const go = (v, filter) => {
    if (filter) setMemberFilter(filter); else setMemberFilter("all");
    setView(v);
  };

  // --- file IO (all local) ---
  const loadFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        const base = emptyData();
        setData({
          ...base,
          ...parsed,
          members: Array.isArray(parsed.members) ? parsed.members : [],
          cycles: Array.isArray(parsed.cycles) && parsed.cycles.length === 6 ? parsed.cycles : base.cycles,
          recognitions: Array.isArray(parsed.recognitions) ? parsed.recognitions : [],
          weeks: Array.isArray(parsed.weeks) ? parsed.weeks : [],
          educationGoals: Array.isArray(parsed.educationGoals) ? parsed.educationGoals : [],
          buddyGroups: Array.isArray(parsed.buddyGroups) ? parsed.buddyGroups : [],
          meetingRota: Array.isArray(parsed.meetingRota) ? parsed.meetingRota : [],
        });
        notify("Data loaded from file.");
      } catch {
        window.alert("That file couldn't be read as valid JSON. Please check it and try again.");
      }
    };
    reader.readAsText(file);
  };

  const exportJSON = () => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `vpe-dashboard-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    notify("JSON exported — keep it somewhere safe.");
  };

  const exportExcel = () => {
    const rows = data.members.map((m) => ({
      Name: m.name,
      Path: memberPaths(m).join(", "),
      Level: m.level,
      "Last attended": m.lastAttended || "",
      "Days since attended": daysSince(m.lastAttended) ?? "",
      Status: memberStatus(m).key,
      "Onboarding start": m.onboardingStart || "",
      "Onboarding day": m.onboardingStart ? (onboardingDay(m.onboardingStart) ?? "") : "",
      Notes: m.notes,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Members");
    XLSX.writeFile(wb, `club-members-${new Date().toISOString().slice(0, 10)}.xlsx`);
    notify("Excel file downloaded.");
  };

  if (appMode === "member") {
    return (
      <MemberPortalView
        data={data}
        setData={setData}
        onSwitchToAdmin={() => setAppMode("admin")}
      />
    );
  }

  // Admin auth gate
  if (!adminUnlocked) {
    return (
      <AdminLogin
        onUnlock={() => setAdminUnlocked(true)}
        onMemberMode={() => setAppMode("member")}
      />
    );
  }

  if (!data) {
    return (
      <Welcome
        onStartFresh={() => setData(emptyData())}
        onLoadFile={loadFile}
        onMemberMode={() => setAppMode("member")}
      />
    );
  }

  const unpostedCount = data.recognitions.filter((r) => !r.posted).length;
  const dormantCount = data.members.filter((m) => memberStatus(m).key === "dormant").length;

  return (
    <div className="min-h-screen flex flex-col md:flex-row" style={{ backgroundColor: C.paper, color: C.ink }}>
      {/* Sidebar (desktop) */}
      <aside className="hidden md:flex md:flex-col w-56 shrink-0 p-4 sticky top-0 h-screen"
        style={{ backgroundColor: C.blue }}>
        <div className="mb-6">
          <div className="text-xs tracking-widest font-bold" style={{ color: C.gold }}>VP EDUCATION</div>
          <div className="text-xl text-white" style={{ fontFamily: SERIF }}>Dashboard</div>
        </div>
        <nav className="flex flex-col gap-1">
          {NAV.map((n) => (
            <button key={n.key} onClick={() => go(n.key)}
              className="flex items-center gap-2 px-3 py-2 rounded-md text-sm font-semibold text-left"
              style={{
                backgroundColor: view === n.key ? "rgba(255,255,255,0.14)" : "transparent",
                color: view === n.key ? C.gold : "rgba(255,255,255,0.85)",
                borderLeft: view === n.key ? `3px solid ${C.gold}` : "3px solid transparent",
              }}>
              <span aria-hidden>{n.icon}</span>{n.label}
              {n.key === "recognition" && unpostedCount > 0 && (
                <span className="ml-auto px-1.5 rounded-full text-xs font-bold"
                  style={{ backgroundColor: C.gold, color: C.blueDeep }}>{unpostedCount}</span>
              )}
              {n.key === "members" && dormantCount > 0 && (
                <span className="ml-auto px-1.5 rounded-full text-xs font-bold"
                  style={{ backgroundColor: C.red, color: "white" }}>{dormantCount}</span>
              )}
            </button>
          ))}
          <div className="mt-2 pt-2" style={{ borderTop: "1px solid rgba(255,255,255,0.12)" }}>
            <button onClick={() => setAppMode("member")}
              className="flex items-center gap-2 px-3 py-2 rounded-md text-sm font-semibold text-left w-full"
              style={{ color: "rgba(255,255,255,0.75)", borderLeft: "3px solid transparent" }}>
              <span aria-hidden>👤</span> Member Portal
            </button>
          </div>
        </nav>
        <div className="mt-auto flex flex-col gap-2">
          <a href={CLUB_LINK} target="_blank" rel="noopener noreferrer"
            className="flex items-center gap-2 px-3 py-2 rounded-md text-sm font-semibold"
            style={{ color: C.gold, border: `1px solid rgba(242,223,116,0.4)` }}>
            ↗ toastmasterclub.org
          </a>
          <div className="flex items-center gap-1.5 text-xs mb-1"
            style={{ color: syncStatus === "ok" ? "#7EE0A0" : syncStatus === "error" ? "#F4A19A" : "rgba(255,255,255,0.55)" }}>
            <span>{syncStatus === "loading" ? "⟳" : syncStatus === "ok" ? "✓" : syncStatus === "error" ? "⚠" : "○"}</span>
            <span>{syncStatus === "loading" ? "Syncing…" : syncStatus === "ok" ? "Saved to cloud" : syncStatus === "error" ? "Cloud sync failed" : "Not connected"}</span>
          </div>
          <Btn kind="maroon" onClick={exportJSON}>Export JSON</Btn>
          <Btn kind="ghost" onClick={exportExcel}>Export Excel</Btn>
          <button onClick={() => fileRef.current?.click()}
            className="text-xs underline mt-1" style={{ color: "rgba(255,255,255,0.7)" }}>
            Load a different JSON file
          </button>
          <button
            className="text-xs underline mt-0.5"
            style={{ color: "rgba(255,255,255,0.55)" }}
            onClick={() => {
              sessionStorage.removeItem(SS_UNLOCKED_KEY);
              setAdminUnlocked(false);
            }}>
            Lock dashboard
          </button>
          <button
            className="text-xs underline mt-0.5"
            style={{ color: "rgba(255,255,255,0.35)" }}
            onClick={() => {
              if (window.confirm("Clear all saved data and start fresh?")) {
                localStorage.removeItem(LS_KEY);
                setData(null);
              }
            }}>
            Clear saved data
          </button>
          <p className="text-xs mt-2" style={{ color: "rgba(255,255,255,0.45)" }}>
            Changes sync to Vercel cloud and are saved locally as backup.
          </p>
        </div>
      </aside>

      {/* Mobile top bar */}
      <header className="md:hidden flex items-center justify-between px-4 py-3 sticky top-0 z-40"
        style={{ backgroundColor: C.blue }}>
        <div>
          <div className="text-xs tracking-widest font-bold" style={{ color: C.gold }}>VP EDUCATION</div>
          <div className="text-base text-white" style={{ fontFamily: SERIF }}>Dashboard</div>
        </div>
        <div className="flex gap-2">
          <Btn kind="maroon" onClick={exportJSON}>JSON</Btn>
          <Btn kind="ghost" onClick={exportExcel}>Excel</Btn>
        </div>
      </header>

      {/* Main */}
      <main className="flex-1 p-4 sm:p-6 lg:p-8 pb-24 md:pb-8 max-w-6xl">
        {view === "home" && <HomeView data={data} go={go} />}
        {view === "members" && <MembersView key={memberFilter} data={data} setData={setData} initialFilter={memberFilter} />}
        {view === "onboarding" && <OnboardingView data={data} setData={setData} />}
        {view === "cycles" && <CyclesView data={data} setData={setData} />}
        {view === "recognition" && <RecognitionView data={data} setData={setData} />}
        {view === "dtm" && <DTMView data={data} setData={setData} />}
        {view === "rota" && <RotaView data={data} setData={setData} />}
        {view === "events" && <EventsView data={data} setData={setData} />}
      </main>

      {/* Mobile bottom nav */}
      <nav className="md:hidden fixed bottom-0 inset-x-0 z-40 flex"
        style={{ backgroundColor: C.blue, borderTop: `2px solid ${C.maroon}` }}>
        {NAV.map((n) => (
          <button key={n.key} onClick={() => go(n.key)}
            className="flex-1 py-2 flex flex-col items-center gap-0.5 text-xs font-semibold relative"
            style={{ color: view === n.key ? C.gold : "rgba(255,255,255,0.8)" }}>
            <span aria-hidden>{n.icon}</span>{n.short || n.label}
            {n.key === "recognition" && unpostedCount > 0 && (
              <span className="absolute top-1 right-1/4 w-2 h-2 rounded-full" style={{ backgroundColor: C.gold }} />
            )}
            {n.key === "members" && dormantCount > 0 && (
              <span className="absolute top-1 right-1/4 w-2 h-2 rounded-full" style={{ backgroundColor: C.red }} />
            )}
          </button>
        ))}
      </nav>

      {/* Hidden file input for re-loading */}
      <input ref={fileRef} type="file" accept=".json,application/json" className="hidden"
        onChange={(e) => { loadFile(e.target.files?.[0]); e.target.value = ""; }} />

      {/* Toast */}
      {toast && (
        <div className="fixed bottom-20 md:bottom-6 left-1/2 -translate-x-1/2 px-4 py-2 rounded-full text-sm font-semibold shadow-lg z-50"
          style={{ backgroundColor: C.blueDeep, color: "white" }}>
          {toast}
        </div>
      )}
    </div>
  );
}
