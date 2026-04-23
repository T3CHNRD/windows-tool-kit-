# Script Review

Review scope: `C:\Users\tvedders\Documents\windows_script_toolkit`

Review type:
- Static code review completed for launcher, module, task scripts, and imported legacy scripts
- Runtime smoke testing completed for launcher parsing and EXE build
- Destructive or environment-specific scripts were not executed end-to-end to avoid unwanted system changes

## Findings

### Current task scripts
| Script | Status | Input needed | Review notes |
|---|---|---|---|
| `Invoke-ClearTempJunk.ps1` | UI-safe | No | Works as a background cleanup task. Should stay admin-only. |
| `Invoke-DebloatInventory.ps1` | UI-safe after launcher update | Yes | Previously had no in-app selection UI. Launcher now provides selected-app review input and passes a selection file. |
| `Invoke-DiskSpaceMonitor.ps1` | UI-safe | No | Safe read-only task. |
| `Invoke-FreeCDriveSpace.ps1` | UI-safe | No | Safe as admin task. May trigger long-running cleanup and should remain progress-tracked. |
| `Invoke-NetworkMaintenance.ps1` | UI-safe | No | Safe read-only diagnostics. |
| `Invoke-ResetNetworkStack.ps1` | UI-safe | No | Admin task with disruptive effect; UI warning/admin gate is appropriate. |
| `Invoke-UpdateAllApps.ps1` | UI-safe | No | Admin task. Long-running task already routed through progress/status UI. |
| `Invoke-WindowsRepairChecks.ps1` | UI-safe | No | Admin task. Long-running task already routed through progress/status UI. |

### Imported legacy scripts
| Script | Status | Input needed | Review notes |
|---|---|---|---|
| `bootcertchcker.ps1` | Review-safe | No | Good candidate for first-class repair/security tab task. |
| `Check-SecureBootCert.ps1` | Review-safe | Optional | Has remote/computer parameters; should eventually get proper UI input fields instead of raw legacy launch. |
| `Fix_OneNote_Duplicates.ps1` | Not UI-ready | No direct prompt | Hardcoded OneNote and path assumptions (`V:`). Needs configuration UI before reliable use. |
| `how_to_guide.ps1` | Not UI-ready | No direct prompt | OneNote-specific behavior with hardcoded notebook/section names. |
| `Invoke-MassDuplicateCleanup.ps1` | Not UI-ready | No direct prompt | OneNote environment-specific. |
| `Invoke-ps2exe.ps1` | Not UI-ready | No direct prompt | User-specific file/icon paths; should not be exposed as a production runtime task. |
| `Launch_OneNote.ps1` | Not UI-ready | No direct prompt | Heavy environment dependency on OneNote plus mapped drives. |
| `Master Audit & Smart-Skip.ps1` | Not UI-ready | No direct prompt | Hardcoded network and OneNote assumptions. |
| `master_onenote_importer.ps1` | Not UI-ready | No direct prompt | Complex rebuild script with Excel/Word/OneNote dependencies and mapped-drive assumptions. |
| `move_to_onenote.ps1` | Not UI-ready | No direct prompt | Requires staged content and OneNote environment. |
| `move_to_onenote_2.ps1` | Not UI-ready | No direct prompt | Depends on prior log files and staging layout. |
| `move_to_onenote_selfcheck.ps1` | Not UI-ready | No direct prompt | Environment-specific cleanup/import workflow. |
| `sortdoc.ps1` | Not UI-ready | No direct prompt | Hardcoded remote/local paths. Good future candidate for parameterized document-prep UI. |
| `The Network Access Script.ps1` | Not UI-ready | Yes | Uses `Read-Host`; requires a dedicated textbox in the launcher before it should be a supported UI task. |

## UI input coverage
- Covered now: Debloat helper has a launcher-based selection UI.
- Still missing: legacy scripts that require either explicit parameters or environment configuration.
- Recommendation: do not treat the OneNote legacy scripts as end-user-ready until they are refactored into parameterized tasks.
