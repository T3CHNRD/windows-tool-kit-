# Contributing

## Task Script Rules

- Task scripts live in `Scripts/Tasks`.
- Exit with `0` when the task succeeds or safely skips work.
- Exit with `1` or throw only when the task truly failed.
- Write clear `Completed`, `Skipped`, and `Failed` summary lines for maintenance, update, security, and misc tools.
- Avoid destructive actions unless the script clearly warns the user and requires confirmation.
- Prefer structured PowerShell cmdlets over parsing command-line text when Windows provides them.

<!-- FIX: LOW-03 - document a consistent task exit-code convention. -->
