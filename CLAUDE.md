# VPE Dashboard — Claude context

## What this is
A single-file React app for a Toastmasters VP Education to manage members, track Pathways progress, run education goals, manage meeting cycles, and handle recognition. Data is persisted to Vercel Blob storage via a serverless API and can also be exported/imported as JSON. Nothing is cached in browser storage — every load reads from the blob and every change debounce-saves back to it.

## Stack
- **React 19** (CRA / react-scripts 5)
- **Tailwind CSS** (via PostCSS)
- **xlsx** for Excel export
- **@vercel/blob** for cloud persistence, via `api/data.js` (Vercel serverless function)
- Single source file: `src/vpe-dashboard.jsx`

## Architecture — everything is in one file
| Section | Location |
|---|---|
| Brand tokens (`C`), constants | Top of file |
| `emptyData()` — data shape | Near top |
| Helper functions (`uid`, `daysSince`, `memberStatus`, `memberPaths`, …) | After constants |
| Small UI atoms (`Badge`, `Card`, `Btn`, `Field`, `SectionTitle`) | After helpers |
| Welcome screen | `Welcome` component |
| Member form modal | `MemberModal` component |
| Onboarding bar + progress | `OnboardingBar` component |
| 100-day plan editor | `PlanModal` component |
| Views (one per nav tab) | `HomeView`, `MembersView`, `OnboardingView`, `CyclesView`, `WeeklyView`, `RecognitionView`, `DTMView` |
| Nav definition | `NAV` array (keys: home, members, onboarding, cycles, weekly, recognition, dtm) |
| Root component + file I/O | `VPEDashboard` (default export) |

## Data shape (`data` state object)
```js
{
  version: 2,
  members: [Member],
  cycles: [Cycle],        // 6 fixed theme cycles per Toastmasters year
  recognitions: [Recognition],
  weeks: [Week],
  educationGoals: [EducationGoal],
}
```

### Member fields
```js
{
  id, name,
  paths: string[],        // Pathways paths (multi-select). Legacy: path: string
  level: 1-5,             // Current Pathways level (working on)
  levelDates: { "1": "YYYY-MM-DD", ... },  // Date each level was completed
  currentProject: string,
  lastAttended: "YYYY-MM-DD",
  totalMeetings: number,
  roles: string[],        // Roles ever completed
  roleLog: [{ id, role, date }],  // Dated role history (upsert by role name)
  isNew: boolean,
  onboardingStart: "YYYY-MM-DD",
  customPlan: [{ from, to, label }] | undefined,
  stagesDone: boolean[] | undefined,   // Manual checkbox overrides for onboarding stages
  devFeeling: "thriving"|"good"|"unsure"|"struggling"|"",
  devNextStep: string,
  notes: string,
  dtm: boolean[] | undefined,  // DTM requirement checklist (7 items = DTM_REQUIREMENTS)
}
```

### EducationGoal fields
```js
{ id, year: number, level: 1-5, target: number }
```

### Cycle fields
```js
{ id, name, span, months: number[], tips: [{ id, text, posted }], prNotes, actions: [{ id, text, done }] }
```

### Recognition fields
```js
{ id, type, member, detail, posted }
```

### Week fields
```js
{ id, label, created, tasks: [{ id, text, done }] }
```

## Key design decisions
- **No browser storage** — no `localStorage`/`sessionStorage` caching of app data. `data` state starts `null`; on mount the app fetches `GET /api/data` (Vercel Blob) and only that response populates it. Every change debounce-saves via `POST /api/data`. JSON export/import remains available as a manual backup, but is not part of the load/save path. `sessionStorage` is still used for the admin-unlock flag (`SS_UNLOCKED_KEY`), which is session auth state, not app data.
- **Local dev has no persistence** — `IS_CLOUD` gates cloud calls off on `localhost`/`127.0.0.1`/`192.168.*`, so `npm start` reloads always start from an empty/Welcome state. This is expected: there is no local fallback store.
- **`cloudLoad()`/`cloudSave()`** live near the bottom of `src/vpe-dashboard.jsx`, alongside `api/data.js` (the serverless GET/POST handler backed by `@vercel/blob`). `cloudLoad()` returns `{ ok, data }` so "blob is empty" and "storage errored" are distinguishable — a failed load renders an explicit "Can't reach cloud storage" screen instead of silently falling through to the file-upload prompt.
- **Smart quotes ban** — JSX attribute values must use straight ASCII `"` (U+0022). Curly/smart quotes (U+201C/D) break the Babel JSX parser. Use module-level string constants (e.g. `CLS_DEV_ROW`, `MUTED`) to avoid JSX attribute string literals when in doubt.
- **`memberStatus(m)`** returns `{ key: "dormant"|"nudge"|"ok", label, fg, bg }`. Members with no `lastAttended` are treated as **dormant** (red).
- **`memberPaths(m)`** normalises the legacy `path: string` field to `paths: string[]`.
- **`DEV_FEELING_MAP`** is a module-level constant (not inline object) to avoid parser issues.
- **Education Goals tab** (`DTMView` function) is the primary tab; individual DTM tracker is collapsed inside it.

## Nav tabs
| Key | Label | View function |
|---|---|---|
| home | Home | HomeView |
| members | Members | MembersView |
| onboarding | 100-Day | OnboardingView |
| cycles | Cycles | CyclesView |
| weekly | Weekly | WeeklyView |
| recognition | Recognition | RecognitionView |
| dtm | Education Goals | DTMView |

## Member form sections (in order)
1. Name
2. Pathways paths (toggle pills, multi-select)
3. Current level (working on)
4. Current project
5. Last attended meeting
6. Total meetings attended
7. **Level completion dates** (grid of 5 date pickers, one per level — feeds education goals)
8. Roles completed (toggle pills + custom)
9. **Role history log** (dated entries, upsert by role name)
10. Development check-in (feeling + next step)
11. Notes
12. New member checkbox + onboarding start date

## Run / build
```bash
npm start    # dev server on :3000
npm run build  # production build
```
