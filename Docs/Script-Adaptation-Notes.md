# ZIP Script Adaptation Notes

The following scripts were extracted from the ZIP and preserved in `LegacyScripts/` with original names.

## Adaptation Strategy
- Keep originals untouched for traceability.
- Expose each script in the launcher as a `Legacy Imported Scripts` task.
- Run legacy scripts in background jobs so UI remains responsive.
- Add centralized logging and progress around legacy task execution.
- Flag OneNote- and network-share-specific scripts as environment-dependent.

## Script-by-Script Notes
1. `bootcertchcker.ps1`
- Purpose: quick secure boot and cert presence check.
- Adaptation: runnable legacy task; future improvement is to fold logic into core repair module.

2. `Check-SecureBootCert.ps1`
- Purpose: local/remote secure boot certificate auditing.
- Adaptation: runnable legacy task; candidate for future first-class “Security Audit” tool with parameter UI.

3. `Fix_OneNote_Duplicates.ps1`
- Purpose: OneNote duplicate page cleanup with logs/progress.
- Adaptation: runnable legacy task; retained for OneNote environments using mapped `V:` paths.

4. `how_to_guide.ps1`
- Purpose: creates a OneNote “START HERE” page.
- Adaptation: runnable legacy task; should eventually become a OneNote helper sub-tool with editable templates.

5. `Invoke-MassDuplicateCleanup.ps1`
- Purpose: hardened duplicate and untitled page cleanup for OneNote sections.
- Adaptation: runnable legacy task; overlaps with other duplicate cleanup scripts and can be consolidated later.

6. `Invoke-ps2exe.ps1`
- Purpose: compiles a script into EXE and creates desktop shortcut.
- Adaptation: runnable legacy task; paths are user-specific and should be parameterized before broad use.

7. `Launch_OneNote.ps1`
- Purpose: consolidated OneNote import and cleanup workflow.
- Adaptation: runnable legacy task; remains available for existing OneNote documentation pipeline.

8. `Master Audit & Smart-Skip.ps1`
- Purpose: OneNote ingest/audit flow with duplicate-aware behavior.
- Adaptation: runnable legacy task; file name contains spaces/special chars and is handled safely by full path invocation.

9. `master_onenote_importer.ps1`
- Purpose: large OneNote rebuild tool with Excel/Word COM dependencies.
- Adaptation: runnable legacy task; kept as advanced/import heavy operation.

10. `move_to_onenote.ps1`
- Purpose: OneNote import from local staging path.
- Adaptation: runnable legacy task; staging path currently hardcoded.

11. `move_to_onenote_2.ps1`
- Purpose: retries failed imports based on log parsing.
- Adaptation: runnable legacy task; depends on presence of prior log output.

12. `move_to_onenote_selfcheck.ps1`
- Purpose: OneNote import with staging purge and self-check logic.
- Adaptation: runnable legacy task; suitable for scheduled execution after parameterization.

13. `sortdoc.ps1`
- Purpose: document categorization/copy into staging folders.
- Adaptation: runnable legacy task; strong candidate for future first-class “Document Prep” task.

14. `The Network Access Script.ps1`
- Purpose: prompt for remote computer and open UNC admin share.
- Adaptation: runnable legacy task; overlaps with new toolkit network category and can be merged later.

## Environment Dependencies to Document
- OneNote COM automation (`OneNote.Application`)
- Office COM automation (Excel/Word) for importer script
- Path assumptions (`V:\`, `P:\`, and custom staging folders)
- Admin rights for selected operations
