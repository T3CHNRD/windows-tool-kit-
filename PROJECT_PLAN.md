# Windows Maintenance Toolkit Project Plan

## Phase 1 - Intake and Inventory
1. Extract imported scripts from ZIP without unpacking unrelated large content.
2. Inventory each script and classify by purpose, risk, and dependencies.
3. Preserve untouched originals under `LegacyScripts/`.

## Phase 2 - Core Architecture
1. Build reusable module (`Modules/MaintenanceToolkit`) with:
   - task catalog
   - logging
   - admin checks
   - shared command runner
2. Build WinForms launcher with:
   - task list
   - status text
   - per-task progress bar
   - background worker execution
   - success/failure messaging

## Phase 3 - Maintenance Functions
1. Add network diagnostics and network reset tasks.
2. Add cleanup, free-space, and disk monitoring tasks.
3. Add app update and Windows repair tasks.
4. Add de-bloat inventory/export helper.

## Phase 4 - External Integrations
1. Add integration wrapper for `mallockey/Install-Microsoft365`.
2. Add Microsoft official Media Creation Tool workflow helper.

## Phase 5 - GitHub Readiness
1. Add README and setup instructions.
2. Add script adaptation notes.
3. Add logs/screenshots placeholders.
4. Validate launcher/module syntax.
