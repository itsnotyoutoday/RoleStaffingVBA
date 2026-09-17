# Role Staffing for MS Project (VBA)

Six components, imported into your Project **Global.MPT**, so they're available in every project
file you open.

Role Staffing answers one question: **how long will this actually take, given who's available to do what?**

It does that by:

- **Suggesting** who should fill each open slot, from each person's authority and preferences
- **Enforcing your rules** — e.g. a drafter may not check their own work — across linked tasks
- **Working out what a check covers** from your predecessor links, with an override for exceptions
- **Levelling** against people's real availability, so the finish date reflects reality
- **Reporting** role load, drift, rule breaches, and what's still unstaffed

## Files

| File | What it is |
|---|---|
| `RoleStaffing.bas` | **Core.** Config, role catalogue, rules engine, field access, relations, weekly load, helpers. Everything leans on this - **import it first.** |
| `RS_TaskData.bas` | Task data that maintains itself: `SameWorkAs`, hammock tasks, `SpanCalc`, backfill, pre-flight |
| `RS_Staffing.bas` | Staffing and scheduling: suggest, apply, unstaff, level, level-of-effort, forward scheduler |
| `RS_UI.bas` | Menu, setup check and its form, control panel, ribbon, reports |
| `clsRS_AppEvents.cls` | The manual-edit guard. Must be a **Class Module**. |
| `clsRS_RowBtn.cls` | Click handler for the setup-check form's per-row Fix buttons. Must be a **Class Module**. |

Everything you paste is also in `paste/`, with the `Attribute`/`VERSION` headers already stripped -
paste those into a new module rather than the `.bas` files themselves, which carry headers VBA's
editor won't accept from the clipboard.

The CONFIG block at the top of `RoleStaffing.bas` is the only part you'd normally edit.

They share one namespace, so nothing needs importing between them - but import `RoleStaffing.bas`
first, since it holds the state the others use.

<details>
<summary>Why four files rather than one</summary>

It was one file until it reached 6,000 lines and 240 KB, which is past what a VBA module should
carry and awkward to work in. The split is by what the code does, not by size: core state, task
data, staffing, interface.

Every module-level variable the others needed became Public in the core module. Nothing else
changed - VBA modules share a namespace, so the split is filing, not restructuring.
</details>

## Installing into Global.MPT

**Global.MPT** is Project's own global template. Anything in it is available in *every* project file you open, which is what you want — you install this once, not per schedule.

### 1. Allow macros

File > Options > Trust Center > Trust Center Settings > Macro Settings. If your IT department controls this, you may need them to allow macros in Global.MPT specifically.

### 2. Open the VBA editor and pick the right project

Alt+F11. In the **Project Explorer** on the left you'll see two or more top-level nodes:

```
VBAProject (ProjectGlobal)      <-- Global.MPT. THIS ONE.
VBAProject (YourSchedule.mpp)   <-- just this one file
```

**Click ProjectGlobal before you do anything else.** This is the single most common mistake: `Insert > Module` and `File > Import File` both drop the code into **whichever node is selected**, so if your schedule is highlighted, the code goes into that `.mpp` and works only while that file is open. If the macros vanish when you open a different project, this is why.

Not sure which you're on? Click a node and look at the Properties window (F4) — the `Name` will say `ProjectGlobal`.

### 3. Get the code in

**File > Import File** on both files, one at a time. This is the better route: the file headers set the module name and the module type for you.

No way to get files onto the machine? See **Pasting** below.

### 4. Check it compiles

**Debug > Compile ProjectGlobal.** This is the editor's built-in syntax check, not a separate compiler. It should complete silently. If it stops on a line, see "If the editor complains" near the end of this README.

### 5. Make it stick

Close the VBA editor, then **close Project entirely**. Global.MPT is written when Project shuts down cleanly — say **Yes** if it asks whether to save the global file.

This matters: if Project crashes, or you kill it from Task Manager, **everything you just imported is lost**. Restart Project and check before you rely on it.

### 6. Check it took

Open **any** project file — ideally a different one from the one you had open — and press **Alt+F8**. You should see `RS_AutoStaff`, `RS_Menu`, `RS_Setup` and the rest in the macro list. If they only appear in one file, go back to step 2.

### 7. Add a ribbon button

File > Options > Customize Ribbon. In the tab list on the right, click **New Group** (right-click > Rename it "Role Staffing"), set "Choose commands from" on the left to **Macros**, pick `RS_Panel`, click **Add**, then **Rename** for an icon and a shorter label.

One button gets you everything. `RS_Panel` opens the control panel; `RS_Menu` is the no-form fallback. Any individual `RS_` macro can be its own button the same way — `RS_AutoStaff` and `RS_Schedule` are the two worth having directly.

### Backing it up

Global.MPT lives here:

```
%APPDATA%\Microsoft\MS Project\<version>\1033\Global.MPT
```

Paste that into Explorer's address bar. Copy the file somewhere safe **before** a big change — if it gets corrupted, Project silently rebuilds an empty one and your macros, views and tables go with it.

`RS_ExportAll` is the other half of this: it writes the current code back out to a folder, so your repo is the real backup.

### Why there's a `paste/` folder as well

**If you can get the files onto the machine, import `RoleStaffing.bas` and `clsRS_AppEvents.cls` and ignore `paste/` entirely.** It exists for one situation only: when the only way to move code is the clipboard. A `.cls` can't be pasted as-is, because its first lines are importer metadata rather than VBA — so `paste/` holds the same code with those lines already stripped. Same source, different transport.

### Pasting instead of importing? Use the `paste/` folder

**Don't copy the `.bas` and `.cls` files.** Their first lines (`VERSION 1.0 CLASS`, `BEGIN`, `Attribute VB_Name = ...`) are metadata for Project's importer, **not VBA** — paste them and the editor turns them red. You don't need them: a module's name comes from the Properties window, not from that line.

The **`paste/`** folder has both files with those headers already stripped. Select all, paste, done:

| Copy this (Raw view) | Into |
|---|---|
| `paste/RoleStaffing.txt` | Insert > **Module**, Name (F4) = `RoleStaffing` |
| `paste/RS_TaskData.txt` | Insert > **Module**, Name = `RS_TaskData` |
| `paste/RS_Staffing.txt` | Insert > **Module**, Name = `RS_Staffing` |
| `paste/RS_UI.txt` | Insert > **Module**, Name = `RS_UI` |
| `paste/clsRS_AppEvents.txt` | Insert > **Class Module**, Name = `clsRS_AppEvents` |

**Select ProjectGlobal in the Project Explorer first.** `Insert` adds to whichever node is highlighted, so with your schedule selected the code lands in that `.mpp` and disappears when you open a different file.

Two things that have to be right:

- **Class Module, not a plain Module**, for `clsRS_AppEvents` — `WithEvents` won't compile in a standard module.
- **The name must be exactly `clsRS_AppEvents`**, because `RoleStaffing.bas` does `New clsRS_AppEvents`.

Set the name in the **Properties window** (F4 > Name), not by typing anything into the code. Copy from GitHub's **Raw** view, not the rendered page, or you'll pick up stray formatting.

> `paste/` is generated from the `.bas` and `.cls` — those two stay the source of truth, and are what `RS_ExportAll` writes.

## The control panel

There's a proper dialog — a status line plus a button per action — but it can't be shipped as a file. **A VBA UserForm keeps every control in a binary `.frx`**, so there's nothing readable to put in a repo or paste into an editor. Instead the code builds the form on your machine:

1. Run **`RS_BuildPanel`** (or `RS_Menu` > 11). Once, ever.
2. Close Project and say **Yes** when it offers to save Global.MPT — otherwise the form is gone next time.
3. From then on, **`RS_Panel`** opens it.

It needs **Trust access to the VBA project object model** ticked in File > Options > Trust Center > Trust Center Settings > Macro Settings. If IT won't allow that, `RS_BuildPanel` says so and `RS_Menu` does the same job without it.

### Do it once, then commit the form

You only have to go through that on **one** machine. Once the form exists, export it and it becomes an ordinary repo file:

1. Build it (`RS_BuildPanel`) or draw it by hand — whichever works on that machine
2. Run **`RS_ExportAll`** (menu 12, or the button on the panel). Point it at your local clone.
3. Out come `RoleStaffing.bas`, `clsRS_AppEvents.cls`, `frmRoleStaffing.frm` **and** `frmRoleStaffing.frx` — commit all four
4. Every machine after that just imports four files: no trust setting, no building, no drawing

`RS_ExportAll` is also the easy way to get any change you make at work back into the repo — it dumps the current state of every Role Staffing component in one go, rather than right-clicking each one. (Without trust access, right-click each component in the Project Explorer > Export File instead.)

That's the way to get the binary: VBA is the only thing that writes a correct `.frx`, so let it, once. (The format is documented as [MS-OFORMS] if you ever want to generate one, but it's an OLE-storage serialization where a wrong byte gives you a silently corrupt form rather than an error — not worth it for a dialog this size.)

<details>
<summary>Or draw it by hand (5 minutes, no trust setting needed)</summary>

In the VBA editor with **ProjectGlobal** selected: Insert > UserForm. Set its **Name** to `frmRoleStaffing` and Caption to `Role Staffing`. Drop on a Label named `lblStatus` and ten CommandButtons named `cmdSetup`, `cmdLoe`, `cmdSuggest`, `cmdApply`, `cmdClear`, `cmdLevel`, `cmdValidate`, `cmdLoad`, `cmdLinks`, `cmdGuard`, plus `cmdClose`. Then double-click the form and paste:

```vb
Private Sub UserForm_Initialize()
    lblStatus.Caption = RS_PanelStatus()
End Sub
Private Sub cmdSetup_Click():    Unload Me: RS_Setup:             End Sub
Private Sub cmdLoe_Click():      Unload Me: RS_AddLevelOfEffort:  End Sub
Private Sub cmdSuggest_Click():  Unload Me: RS_Suggest:           End Sub
Private Sub cmdApply_Click():    Unload Me: RS_Apply:             End Sub
Private Sub cmdClear_Click():    Unload Me: RS_ClearSuggestions:  End Sub
Private Sub cmdLevel_Click():    Unload Me: RS_Level:             End Sub
Private Sub cmdValidate_Click(): Unload Me: RS_Validate:          End Sub
Private Sub cmdLoad_Click():     Unload Me: RS_RoleLoad:          End Sub
Private Sub cmdLinks_Click():    Unload Me: RS_RuleLinks:         End Sub
Private Sub cmdClose_Click():    Unload Me:                       End Sub
Private Sub cmdGuard_Click()
    If gRS_GuardOn Then RS_StopGuard Else RS_StartGuard
    lblStatus.Caption = RS_PanelStatus()
End Sub
```
</details>

The guard starts automatically via `Auto_Open`. If you hit Reset in the VBA editor it stops; run `RS_StartGuard` or restart Project.

---

# The model

Three ideas, and the whole thing follows from them.

### 1. Max Units is the only hard constraint

A person's **Max Units** is their real availability to this project. Nothing may exceed it. This is also exactly what Project's own leveller and overallocation indicators understand — which is why levelling is native here rather than reinvented, and why the red overallocation markers you already know still mean what they've always meant.

### 2. RoleAvail is authority *and* preference

One column, `RoleAvail`, does two jobs:

```
Jane        Check:40*, Approve:40*, Engineer:10, Draft:10
Bob     Draft:100
Liaison     Check:80*, Draft:20
```

- **Listed at all = they're allowed to do it.** Bob has no `Check`, so he can *never* be suggested as a checker. He won't even appear as a candidate. That's an authority boundary, not a low score.
- **The number = preference** among the roles they *are* allowed to do — where their time goes first.
- **`*` = this is their job.** A bump that beats raw availability, for when someone owns a role regardless of who has more spare time.
- **`Check:0` is valid** — allowed, but last resort.

### 3. The percentages are soft

They're **priority levers, not walls.** If there's no Approve work this month, Jane's approve time lends to checking. The numbers don't have to add up to 100 — they're weights, not a partition.

So "actual" drifting from "target" in the Role Load report is normal and expected, not a fault. The only thing that's genuinely a problem is exceeding Max Units, and that's what levelling fixes.

### Who does what

| Concern | Handled by |
|---|---|
| Who's allowed to do what | RoleAvail — hard boundary |
| Who gets picked | Our scoring: preference, `*`, current load, headroom |
| When work happens | **Project's native leveller** — Max Units is the real constraint |
| Overallocation alerts | **Project's native indicators** — correct, live, no setup |
| Rule breaches | Our guard (as you type) and RS_Validate (on demand) |
| Preference drift | Our Role Load report — information, not an error |

---

# Fields

Five custom fields. `RS_Setup` renames them for you.

**In the normal flow you type exactly one of them: `RoleAvail`.**

| You type | Macros write | You type, occasionally |
|---|---|---|
| Resource: **Name, Max Units, RoleAvail** | Resource: `IsRole` (on the buckets it creates) | Task row: `Role`, when staffing by hand |
| Task: **Name, Work, links, task type** | Task: `Role` on assignment rows, `Suggested`, Notes | Task: `CoversUIDs`, when links lie |
| Task: **assign a bucket** | | |

Everything else is built-in Project fields you already use.

## Resource sheet

| Field | Name | Example | Filled by |
|---|---|---|---|
| Text1 | `RoleAvail` | `Check:40*,Approve:40*,Engineer:10,Draft:10` | **You** |
| Text2 | `IsRole` | `Check` | You, on bucket rows |

Plus built-in **Max Units** — real availability.

Two kinds of row:

| Type | Example | Max Units | RoleAvail | IsRole | Assigned to tasks? |
|---|---|---|---|---|---|
| **Person** | Jane | 40% | `Check:40*,...` | — | Yes, by staffing |
| **Bucket** | Checker | 999% | — | `Check` | Yes, to mark a slot open |

A **bucket** is unstaffed demand. Name it whatever you like — `Checker`, `Redline Support`, `Approver`. `IsRole` is what carries the meaning, not the name.

## Task sheet

| Field | Name | Purpose |
|---|---|---|
| Text2 | `Role` | On a **task row**: what kind of work this is. On an **assignment row**: the role that person fills. |
| Text4 | `CoversUIDs` | Override: which upstream work this task's rules reach. `none` = covers nothing. |
| Text5 | `Suggested` | Suggested person, for review before Apply |

`Role` being one column with two meanings is deliberate — read it down the outline and the indent tells you which:

```
ID  Name                 Role     Suggested    Work
45  Draft wing rib       Draft                  40h
      Bob            Draft                  40h
      Jane               Check                   8h
```

Assignment **Notes** holds why a person was suggested.

## Roles are discovered, never hardcoded

There's no fixed list of roles anywhere. They're whatever appears in `RoleAvail` and `IsRole` across the resource sheet. **Add a role by typing it in somebody's RoleAvail** and re-running `RS_Setup` to get a bucket for it.

Typo containment: `RS_Validate` lists every role it found with how many people have it. A typo shows up as a role with exactly one person.

---

# Rules

Set in the CONFIG block at the top of `RoleStaffing.bas`:

```vb
Public Const RS_RULES As String = "Draft&!Check"
```

Read as: *whoever Drafts may not Check* — **the same work**.

| Written | Means |
|---|---|
| `Draft&!Check` | a drafter may not check their own work |
| `Draft&!Check;Check&!Approve` | ...and a checker may not approve it |
| `Draft&!Check&!Approve` | a drafter may neither check nor approve it |

Role names are arbitrary. `Design&!Verify` works the same way.

### "The same work" spans two tasks

The point of this is the two-UID case:

```
Task 112  Draft wing rib    ← Bob drafts
    ↓ predecessor link
Task 118  Check wing rib    ← Bob is now blocked from checking
```

From any task carrying a role a rule mentions, it walks back through predecessor links to find the related work. While walking it:

- passes through milestones and any task whose roles **no rule mentions** (a Redline rework task doesn't hide the drafting behind it)
- looks inside a summary task if a link points at one
- stops at other rule-relevant work
- collects all of them if there are several

Walk depth is `RS_MAX_LINK_DEPTH` (8 links).

**Rules are symmetric.** It doesn't matter whether you assign the drafter or the checker first — both directions are blocked.

**`CoversUIDs` overrides the links.** Where links don't tell the true story, type the related tasks' **Unique IDs** into `CoversUIDs`, e.g. `112, 118`. Enter `none` for a check that covers no drafting. Use Unique IDs, not IDs — row IDs shift when tasks move, Unique IDs never do.

---

# Day-to-day

**Once per project file:**

1. **`RS_Setup`** (menu 1) — renames fields, creates a bucket per role, builds the `RS Roles` and `RS Staffing` views, a group and some filters.
2. Open **RS Roles** and fill in each person's **Max Units** and **RoleAvail**.
3. Re-run **`RS_Setup`** if you added roles it hadn't seen — it creates the missing buckets.

> ### Suggested and assignment-level Role are invisible in a Gantt Chart
>
> `RS_Suggest` writes the name onto the **assignment row** (`Assignment.Text5`). A Gantt Chart has
> one row per task and **no assignment rows**, so that column reads blank there however well Suggest
> worked. Same for the assignment-level `Role`.
>
> Use a **Task Usage** view (View > Task Usage, or the `RS Staffing` view) and expand a task:
>
> ```
> 25  Transcribe cockpit lighting    Draft
>       Role: Draft                  Draft      Bob    <- the suggestion is here
> ```
>
> This is the most common reason it looks like nothing happened.

## Just staff it

**`RS_AutoStaff`** (menu 0, or the top button on the panel) does the lot in one go: picks the best
available person for every open slot, puts them into Resource Names, and levels the schedule against
everyone's real availability. One Undo puts it all back.

`RS_Suggest` / `RS_Apply` separately are for when you want to look the picks over first.

If it comes back with **"Nobody was assigned"**, the Staffing Suggestions report it just opened says
why slot by slot - read the **Excluded** lines under each task. The two usual causes are that nobody
has that role in their RoleAvail, or a rule excluded everyone who does.

**What each macro actually changes** — the usual confusion is expecting `RS_Level` to assign people:

| Macro | Changes |
|---|---|
| `RS_MarkRoles` | Marks role rows, writes `Role`. Resource Names unchanged. |
| `RS_Suggest` | Writes the **`Suggested`** column only. Resource Names unchanged, on purpose - so you can review first. |
| **`RS_Apply`** | **Replaces the role name with the person in Resource Names.** This is the one that staffs the schedule. |
| `RS_Level` | Moves dates to fit availability. Assigns nobody. |

If `Suggested` comes back empty, it's almost always that nobody has **RoleAvail** filled in - a person
is only a candidate for a role listed there. The Suggest report tells you which:

| Report says | Meaning |
|---|---|
| `Nothing to staff` | `IsRole` isn't set on the role rows - run `RS_MarkRoles` |
| `Nobody qualified and eligible for Check` | Nobody has `Check` in RoleAvail - fill in the resource sheet |
| `-> Suggested: Jane` | Working. Run `RS_Apply`. |

**Each planning cycle:**

4. **Build and link the schedule** as you do now.
5. **Assign buckets** to the tasks that need somebody. ← **This is the trigger.** A slot is open because a bucket sits on it; nothing auto-detects unstaffed work.
6. **`RS_AddLevelOfEffort`** (menu 2) for ongoing support — select summary tasks, pick a role and a %.
7. **`RS_Suggest`** (menu 3) — scores everyone, fills `Suggested`, opens a report showing every candidate, their score, and who was excluded and why.
8. **Review** in RS Staffing. Type a different name to override, clear it to leave the slot open.
9. **`RS_Apply`** (menu 4) — swaps each bucket for its person, re-checks the rules. One Undo step.
10. **`RS_Level`** (menu 6) — levels against real availability and reports the finish date before and after.
11. **`RS_Validate`** (7) and **`RS_RoleLoad`** (8) before issuing.

---

# Reports

| Macro | Tells you |
|---|---|
| `RS_Suggest` | every candidate per slot, scores, exclusions and why |
| `RS_Level` | finish date before/after, which tasks were delayed |
| `RS_Validate` | rule breaches, unknown roles, roles nobody can do, real overallocation |
| `RS_RoleLoad` | per person: target vs actual per role, weeks past availability, unstaffed demand |
| `RS_RuleLinks` | what each task covers, and how it worked that out |

`RS_RoleLoad` also totals **unstaffed demand per role** — useful for "do I even have enough checkers on this programme?"

---

# Level-of-effort

A bucket on a **summary task** spans first to last task underneath, so support follows development through release, and stretches as the schedule moves.

The % applies to whoever ends up with the assignment. To commit a share of a whole *group* rather than one person, make a resource like `Redline Group` with Max Units equal to the group size (300% for three people). Remember 15% of that group is 45% units.

**Project's leveller won't move summary-task assignments.** That load is real but it isn't levelled — check `RS_RoleLoad` to see what it's doing to the people carrying it. `RS_Level` warns you when the file has any.

---

# The manual-edit guard

With the guard on, typing a name over a bucket in Task Usage:

- **Blocks** it if it breaks a rule, naming the task where the conflict is
- **Asks** if the person has no authority for that role (it's not in their RoleAvail)
- **Records the role** on the assignment row so the rules keep working afterwards

Bulk paste may not raise events at all, so `RS_Validate` is the real backstop.

---

# If the editor complains

**"Expected: end of statement" on `VERSION 1.0 CLASS`** means the file reached Project with **Unix (LF) line endings**. The importer can't parse the class header without CRLF, so it treats the header as code. The files in this repo are stored CRLF and `.gitattributes` stops git normalising them — but if you ever regenerate or re-save them on a Mac or Linux box, convert back to CRLF before importing.


Written against the documented Project object model but not run here, so the first Debug > Compile may find something. Likely spots:

- **Event signatures** in `clsRS_AppEvents`. Delete the `Sub` line, pick **App** in the left dropdown and the event in the right, so the editor writes the exact signature, then paste the body back.
- **View, filter, group and levelling methods** vary between Project versions — `GroupEditEx` doesn't exist in all of them. These are called **late-bound through a `PA()` helper** on purpose: an early-bound call to a missing method, or one with a renamed argument, is a *compile* error that stops the whole module, and no `On Error` can catch it. Late-bound, it's a runtime error that `RS_Setup` reports and carries on from. **Don't "tidy" `PA.` away.**

  If a view, filter or group fails, build it by hand — everything else still works. The group is pure convenience: in Resource Usage, Project > Group by > Customize, group **assignments** by `Text2` to see each person's load split by role.
- **Units.** Versions differ on whether 100% is `1` or `100`. The code handles both. If a level-of-effort % comes out 100× off, look at `SetUnits`.

# Checking the code before it reaches Project

VBA reports compile errors one at a time, in front of you, after a paste. Two checks run here
instead, and between them they catch most of what used to be found the hard way.

### `python3 tools/check_vba.py`

Project-specific things no general tool can know:

- **CRLF line endings** — LF makes the importer treat a class header as code
- **Reserved words as identifiers** — `Text`, `Val`, `Line`, `Open`, `Lib` and friends, in `Dim`,
  parameters, `For` loops, and module-level `Public`/`Private`/`Const`
- **Declarations after a procedure**, unbalanced `If`/`With`/`For` blocks, too many line continuations
- **`pj*` constants that don't exist** — checked against all 4,907 real ones. `pjStandard` was
  invented and stopped a module compiling
- **Assignment to a function call**, `Private` called across modules, a module named after a procedure
- **Form markers** — the generated forms are found by a procedure their code calls; when that
  procedure was renamed, the check form stopped recognising itself and built a new copy every run
- **`paste/` drifting out of step** with the source files
- **The probe macro naming no methods** — `CallByName` with `VbGet` *runs* a zero-argument method

### `vbalidator . --host project`

A real VBA parser, with a type model for MS Project. Needs the submodule and Python 3.10+:

```bash
git submodule update --init
pip install -e vendor/VBAlidator
vbalidator . --host project
```

It knows scope and types, which regexes can't. On its first run it found `For Each x` with no
`Dim x` under `Option Explicit` — a compile error that would have stopped all of `RS_UI.bas`
compiling, and so every menu item in it.

The MS Project type model is in the submodule: 103 classes, 2,925 members and all 4,907 `pj*`
constants, generated from Microsoft's published VBA reference. It covers the *documented* API — a
type library also carries hidden members, which is why `Project.NewTasksAreManual` is carried as an
explicit exception. See `tools/README.md`.

# Test these early

- **A rule across two tasks.** Draft on 112, Check on 118, linked. Try to put the drafter on the check and confirm it's blocked.
- **Pass-through.** Put a Redline task between them and confirm the block still fires.
- **Summary-task assignment.** Redline at 15% on a summary; move a child later; confirm it stretches.
- **Suggest then Apply** on one task — confirm the person gets the work the bucket had.
- **Level** on a deliverable you already know the answer to, before trusting it on the whole programme.

# Known limitations

- **Summary-task level-of-effort isn't levelled.** Project won't move those assignments.
- **Weekly capacity uses the project's hours per week.** Resource calendars, vacations and holidays aren't considered.
- **Editing a Suggested name doesn't re-score the other rows.** Re-run Suggest.
- **Roles don't constrain the schedule, only the choice of person.** If Jane ends up 80% drafting against a 10% preference, that's reported as drift, not prevented. Preference is soft by design — that's the point of lending time between roles.
- **Speed.** Reports read the whole project. Fine for hundreds of tasks, slower for very large files.

---

# What each command does

The panel shows the eight you'd actually use. **More...** opens the full list for the one-off things.

### On the panel

| Command | Does | When |
|---|---|---|
| **STAFF AND SCHEDULE** | Picks people, writes them into Resource Names, levels the schedule. One Undo. | The main one. Every time you re-plan. |
| **Check & prep file** | Finds manually scheduled tasks and wrong task types, offers to fix them. Reports missing Work and links. | Before a run, and after anyone edits the schedule by hand |
| **Validate** | Rule breaches, unknown roles, roles nobody can do, real overallocation. Changes nothing. | Every time, before staffing. It's the cheap pre-flight. |
| **What-if: project the finish date** | Schedules forward from availability and reports a finish date and the tightest role. **Writes nothing.** | Change somebody's Max Units, run this, see what it buys |
| **Role load** | Per person: target vs actual hours per role, weeks past availability, unstaffed demand | Checking whether the plan is humane |
| **Rule links** | What each task covers, and how it worked that out | When a rule fires and you don't see why |
| **Unstaff everyone** | Puts people back on role rows, clears levelling. Ready to run again. | Starting over |
| **More...** | The full menu | Setup and migration |

### Behind More...

| Command | Does | When |
|---|---|---|
| **Set up this file** | Renames the custom fields, builds the views and filters | Once per project file |
| **Mark role rows** | Marks Engineer/Checker/etc. as roles and copies the role onto every assignment | Once, migrating a schedule built before this existed |
| **Re-record roles** | Same copy step on its own | After hand-assigning or pasting |
| **Suggest** / **Apply** | The two halves of STAFF AND SCHEDULE, separately | When you want to review the picks before committing |
| **Clear suggestions** | Empties the Suggested column | Abandoning a review |
| **Level** | Levels only - assigns nobody | Re-levelling after hand edits |
| **Level-of-effort** | Puts a role on a summary task at a percentage | Ongoing support spanning a deliverable |
| **Guard on/off** | The warning when you type a name over a role row by hand | Rarely |
| **Build the control panel** | Creates the panel itself | Once, ever |
| **Export all files** | Writes the code out to a folder | Getting changes back into git |

---

# Before you run it

**`RS_PrepFile` (menu 17, or the second button on the panel) does most of this for you.** It checks
the file, tells you what's wrong, and offers to fix what's fixable.

It **fixes**: manually scheduled tasks (levelling ignores those entirely), task type (must be Fixed
Work, or reducing somebody's rate changes the WORK instead of the dates and nothing stretches), new
tasks defaulting to manual, and the levelling option that otherwise pins your finish date.

It **reports but won't guess at**: tasks with no Work in hours, slots with no predecessor links,
level-of-effort on summaries. Those need a decision.

What it can't see is the resource sheet, so check that yourself:

| | Resource sheet |
|---|---|
| ☐ | Every person has **Max Units** - their real availability to *this* project |
| ☐ | Every person has **RoleAvail** - not listed means not allowed |
| ☐ | One role row per role, **IsRole** set, Max Units **1000%** |
| ☐ | No row has both IsRole and RoleAvail |

Then, in order:

1. **`RS_Setup`** - once per file
2. **`RS_PrepFile`** (17) - fixes the task side
3. **`RS_Validate`** (7) - should be **0 must-fix**. This is the real pre-flight.
4. **`RS_Schedule`** (15) - writes nothing, projects a finish date. If it looks insane, something
   above is wrong and finding out now is free.
5. **`RS_AutoStaff`** (0) - the real run

Steps 2-4 are all read-mostly and quick. Step 3 is the one worth doing every single time.

---

### Editing someone's roles

Select a person on the resource sheet, then **Menu > P** (or the control panel's **Edit RoleAvail**
button). It shows their current roles, the roles this file knows about, and checks the percentages
add to 100 - offering to scale them proportionally if not.

There's no right-click item by default: **Project 365 keeps its shortcut menus in the ribbon
system**, not in CommandBars. The collection VBA can see holds only legacy toolbars and Office task
panes - on a 365 build it contains exactly one shortcut menu, and it's empty. Nothing to attach to.

The CommandBars approach was tried and removed for that reason. What remains is the route that can
work: **Menu > X > 58** writes a `<contextMenus>` block into the customUI part of the ribbon file,
and **59** takes it back out. That needs a Project restart to take effect, because the ribbon file
is read at startup.

# Which view for what

Project has several views over the same data, and they show different things. Neither Gantt nor
Task Usage is "the right one" - they answer different questions.

| View | One row per | Use it for |
|---|---|---|
| **Gantt Chart** | task | Building the schedule: names, Work, links, task type. Assigning role rows. Task-level `Role` and `CoversUIDs`. |
| **Task Usage** (`RS Staffing`) | task **+ each assignment** | Reviewing staffing: `Suggested`, who is on a task and at what units, assignment-level `Role`. |
| **Resource Sheet** (`RS Roles`) | person or role row | `Max Units`, `RoleAvail`, `IsRole`. Where red overallocation shows up. |
| **Resource Usage** | person, by week | Who is overloaded, in which weeks, on what. |

The thing that catches people: **assignment-level fields do not exist in a Gantt Chart.** A Gantt has
one row per task and no assignment rows, so `Suggested` reads blank there however well `RS_Suggest`
worked. Same for the per-person `Role` and units.

So, in practice:

1. **Gantt** - build the schedule, link it, put role rows in Resource Names
2. **Task Usage** - run Suggest, check the picks, run Apply
3. **Resource Sheet / Resource Usage** - see who is overloaded and by how much

You can assign resources from either Gantt or Task Usage; only *reviewing* needs Task Usage.

---

# Migrating a schedule you've already built

If your Gantt already has role names in **Resource Names** — `Engineer`, `Checker`, `Draft` — that
work isn't wasted. Those rows *are* the unstaffed slots; nothing in the file just says so yet.

Run **`RS_MarkRoles`** (menu 13). It asks which resources are roles rather than real people,
pre-filled with every resource name so you only delete the actual people:

```
Engineer, Checker, Draft, Redline Support, Bob Smith, Jane Doe
        ^ delete the people, leave the roles
```

If a row's name isn't the role name, say so with `=`:

```
Checker=Check, Redline Support=Redline
```

It then, in one Undo step:

1. **Marks those rows as roles** (sets `IsRole`) so the macros treat them as unstaffed demand
2. **Raises their Max Units** so open slots don't read as an overallocated person
3. **Copies the role onto every assignment** that uses them — this is the part that makes your
   existing Gantt work with the rules and the reports
4. Sets the task-row `Role` too, wherever a task carries only one role

Without step 1, `RS_Suggest` would report "nothing to staff" on a completely populated
schedule — it identifies a slot by `IsRole`, and your old rows don't have it.

**Then:** fill in **Max Units** and **RoleAvail** for your real people (nobody is a candidate for
a role that isn't in their RoleAvail), and run `RS_Suggest` > `RS_Apply` > `RS_Level`.

`RS_BackfillRoles` (menu 14) re-runs just the role-recording part — useful after hand-assigning or
pasting. Neither macro overwrites a `Role` that's already filled in.

---

# Smoke test: seeing the scheduling work

A small file, rigged so availability is obviously the binding constraint. Ten minutes.

## Build it

New blank project. Two people:

| Name | Max Units | RoleAvail |
|---|---|---|
| Bob | `50%` | `Draft:100` |
| Liaison | `50%` | `Check:80*,Draft:20` |

At 50% units each that's **20 h/week** apiece. Run **`RS_Setup`** — its message should list the `Draft` and `Check` buckets it created. If it says no roles found, RoleAvail didn't take.

Four tasks. Two things must be right or nothing will move:

- **Auto Scheduled**, not Manually Scheduled — manual tasks are invisible to the leveller
- **Task type = Fixed Work** (Task Information > Advanced)

| ID | Name | Work | Predecessor |
|---|---|---|---|
| 1 | Draft A | `40h` | — |
| 2 | Draft B | `40h` | — |
| 3 | Draft C | `40h` | — |
| 4 | Check A | `8h` | **1** |

Enter **Work**, not Duration. Leave 1-3 unlinked so they all want to start at once.

Assign the **buckets** to the tasks - `Draft` on 1-3, `Check` on 4. Not people: a bucket is what says the slot is open.

## 1. Does it spread the load?

Run **`RS_Suggest`** and read the report:

- **Task 1 > Bob** — score ~100 (`Draft:100`) against Liaison's ~20 (`Draft:20`)
- **Task 2 > Liaison** — the interesting one. Bob is already full those weeks, so the
  over-availability penalty drags him below her. If both go to Bob, load tracking is broken.
- **Task 4 > Liaison**, with Bob under **Excluded**:
  `Bob: was Draft on task 1 (Draft A), which this task covers`

That last line is `Draft&!Check` firing across two task UIDs through the predecessor link. It's
the most important line in the test.

Then **`RS_Apply`**.

## 2. Does Project see the overallocation?

Before levelling, the **Resource Sheet** should show Bob in **red** with the overallocation
icon, because tasks 1 and 3 overlap. That's the native alerting proving the model: role
percentages are soft, Max Units is the wall, and Project understands the wall.

Note the current **Project Finish** (Project > Project Information).

## 3. Does the date move?

Run **`RS_Level`**. Bob has 80 h of drafting at 20 h/week = 4 weeks, so tasks 1 and 3
serialise and the finish pushes out. The red should clear.

That's the whole point of this tool: a finish date that reflects who is actually available.

Undo with Resource > Level > **Clear Leveling** and re-run as often as you like.

## 4. Sanity-check the reports

- **`RS_RoleLoad`** — Liaison's Draft *actual* should exceed her 20% *target*. That's correct,
  not a fault: she's lending time from checking, which is the soft-percentage model working.
- **`RS_Validate`** — 0 must-fix.
- **`RS_RuleLinks`** — task 4 shows `covers task 1`.

## If nothing moves

| Symptom | Cause |
|---|---|
| Levelling changes nothing | Tasks are **Manually Scheduled**, or already sequential |
| Durations look wrong | Task type isn't **Fixed Work** |
| "Nothing to staff" | You assigned people instead of buckets |
| Rule didn't fire | No predecessor link 4 > 1. Check `RS_RuleLinks`. |
| Everything went to one person | Load tracking problem - keep the Suggest report |
