---
name: project-status-dashboard
description: Use this skill in Excel when the user asks to build, create, or update a project status dashboard from a task list or tracker table.
license: Internal use.
---


# Project Status Dashboard Builder

Build a clean project status dashboard from the task data in this workbook.

## Steps
1. Find the first table in the workbook that contains task-level data
   (look for columns like Task, Owner, Due Date, Status, or % Complete).
2. Group tasks by workstream or category (use a "Workstream" or "Category"
   column if present; otherwise group by Owner).
3. For each group, calculate:
   - Total tasks
   - % complete (based on the Status or % Complete column)
   - Count of overdue tasks (Due Date is in the past and Status is not "Done")
4. Create a new sheet named "Dashboard."
5. On the Dashboard sheet, build a summary table with one row per
   workstream/group, showing: Workstream, Total Tasks, % Complete,
   Overdue Count.
6. Apply conditional formatting:
   - % Complete below 50% → yellow fill
   - Overdue Count greater than 0 → red fill
7. Add a title at the top of the sheet: "Project Status Dashboard —
   [today's date]."

## Workbook output
Produce a dashboard sheet with the following characteristics:
- Named "Dashboard"
- One summary row per workstream
- Conditional formatting applied as described above
- No raw task data duplicated — summary only

## Common pitfalls to avoid
- Do not create the dashboard if no task table can be found — ask the
  user to point to the correct sheet or table instead.
- Do not treat header rows as data.
- Do not overwrite an existing "Dashboard" sheet without asking first.