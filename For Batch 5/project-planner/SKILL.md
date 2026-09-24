---
name: project-planner
description: Creates a project planner Excel workbook (tasks, milestones, timeline) from user-provided project details, meeting notes, briefs, or task lists in any format, following a fixed template. Use whenever the user wants to start, build, or set up a project plan/planner/tracker as an Excel file.
license: Internal use.
---

# Project Planner Skill

You create project planner Excel workbooks by filling in the bundled
template file (`assets/project_planner_template.xlsx`) — never build a
workbook from scratch, and never change its tab names, table names, or
column headers.

## Template structure

Reference `assets/project_planner_template.xlsx` directly when writing code:

- **Sheet "Dashboard"**: cell `C4` = project name, `C5` = project manager.
  All other cells are formulas — never overwrite them.
- **Sheet "Tasks"**: Excel Table `TasksTable`, columns A-I = ID, Task Name,
  Owner, Start Date, Due Date, Status, % Complete, Notes, Priority.
- **Sheet "Milestones"**: Excel Table `MilestonesTable`, columns A-E = ID,
  Milestone, Due Date, Status, Notes.
- **Sheet "Gantt Timeline"**: cell `C3` = project start date. All other
  cells are formulas or mirrored from Tasks — never overwrite them.
- Status must be one of: `Not Started`, `In Progress`, `Complete`, `Blocked`.
- Priority must be one of: `Low`, `Medium`, `High`.
- % Complete is stored as a fraction (0.25, not 25) formatted as a percentage.
- Dates must be resolved to explicit calendar dates (convert relative
  phrases like "next Friday" or "two weeks from today" using the current
  date) before writing them.

## Conversation flow

1. Determine whether the user is starting a new project or providing
   source material (a brief, meeting notes, a task list, a spreadsheet)
   to build one from. Either is fine — extract what you can from any
   source material provided, and ask for the rest.
2. Collect, asking if not already given: project name, project manager,
   project start date.
3. Collect tasks: task name, owner, start date, due date, status (default
   "Not Started"), % complete (default 0), notes (optional), priority
   (default "Medium"). Multiple tasks can be collected or extracted at once.
4. Collect milestones: milestone name, due date, status (default "Not
   Started"), notes (optional).
5. **Never guess a required field silently — ask.** This applies even when
   working from source material: if a task's owner, dates, or name aren't
   present in the source, ask rather than inventing them. It's fine to ask
   for several missing fields in one message.
6. Once you have at least the project basics and one task, generate the
   file: open `assets/project_planner_template.xlsx` with code, write the
   project name/PM into Dashboard, the start date into Gantt Timeline,
   append one row per task to `TasksTable` and one row per milestone to
   `MilestonesTable` (do not touch any other cells), then save the result
   as `"<Project Name> - Project Planner.xlsx"` and share it with the user.
7. Confirm in plain language what was added. Don't describe cell
   addresses, code, or table internals to the user.
8. If the user wants to keep adding tasks/milestones after the file is
   generated, collect the new details and regenerate the file with the
   additional rows appended, keeping everything already in it.
9. Remind the user that the Dashboard and Gantt Timeline sheets
   recalculate automatically when they open the file in Excel — those
   numbers may show as 0 or blank in a quick preview until then.

## Reference: what the generation code should do

```python
import openpyxl

wb = openpyxl.load_workbook("assets/project_planner_template.xlsx")

# Dashboard
ws = wb["Dashboard"]
ws["C4"] = project_name
ws["C5"] = project_manager

# Gantt Timeline
wb["Gantt Timeline"]["C3"] = start_date

# Tasks — append rows into the existing table, don't overwrite existing ones
tasks_ws = wb["Tasks"]
table = tasks_ws.tables["TasksTable"]
# find the next empty row within the table range (or next ID number),
# and write task fields into columns A-I of that row

# Milestones — same pattern against MilestonesTable

wb.save(f"{project_name} - Project Planner.xlsx")
```

Verify before delivering a file:
- It loaded the existing template rather than building a new workbook
  (this preserves formulas, formatting, and table definitions).
- It appended into the Table ranges rather than overwriting row 5 every time.
- It never wrote into Dashboard or Gantt Timeline cells other than the
  two input cells named above.
