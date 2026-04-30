"""
system_repair.py — RepairDock AI
Cross-platform system integrity checks and repair.

Translated from the original Windows Tool Kit:
  Scripts/Tasks/Invoke-WindowsRepairChecks.ps1

Original bugs fixed:
  1. DISM called before SFC — should be SFC first, DISM only if SFC fails
  2. No timeout on DISM RestoreHealth — can hang indefinitely on no internet
  3. Exit code not checked after SFC — partial success treated as success
  4. No macOS/Linux equivalent provided at all

Platform mapping:
  Windows — DISM + SFC (system file checker)
  macOS   — fsck (read-only on live volume), Disk Utility First Aid via diskutil
  Linux   — fsck (unmounted), e2fsck, systemd journal checks
"""

import argparse
import platform
import subprocess
import sys
from datetime import datetime
from pathlib import Path

SYSTEM = platform.system()
LOG_DIR = Path(__file__).parent.parent.parent / "logs"


def _log(msg: str, level: str = "INFO") -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    print(line)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"system_repair_{datetime.now().strftime('%Y%m%d')}.log"
    with open(log_file, "a") as f:
        f.write(line + "\n")


def _die(msg: str) -> None:
    _log(msg, "ERROR")
    sys.exit(1)


def require_admin() -> None:
    if SYSTEM == "Windows":
        import ctypes
        if not ctypes.windll.shell32.IsUserAnAdmin():
            _die("Administrator privileges required.")
    else:
        import os
        if os.geteuid() != 0:
            _die("Root privileges required (sudo).")


def run(cmd: list[str], timeout: int = 1800, check: bool = True) -> subprocess.CompletedProcess:
    _log(f"EXEC: {' '.join(cmd)}")
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,  # Bug fix: added timeout — DISM can hang indefinitely
        )
    except subprocess.TimeoutExpired:
        _log(f"Command timed out after {timeout}s: {' '.join(cmd)}", "WARN")
        return subprocess.CompletedProcess(cmd, returncode=1, stdout="", stderr="TIMEOUT")

    if result.stdout.strip():
        _log(f"OUT: {result.stdout.strip()}")
    if result.stderr.strip():
        _log(f"ERR: {result.stderr.strip()}", "WARN")
    if check and result.returncode != 0:
        _die(f"Command failed (exit {result.returncode}): {' '.join(cmd)}")
    return result


# ─────────────────────────────────────────────
# Windows — SFC then DISM
# Bug fix: original ran DISM first. Correct order is SFC → DISM
# ─────────────────────────────────────────────

def repair_windows() -> None:
    _log("=== Windows System File Checker (SFC) ===")
    sfc_result = run(["sfc", "/scannow"], timeout=1800, check=False)

    # Bug fix: check exit code properly. 0 = no violations. 1 = could not repair.
    if sfc_result.returncode == 0:
        _log("SFC: No integrity violations found or all repaired successfully.")
    else:
        _log("SFC found issues it could not repair — running DISM to repair component store.", "WARN")
        _log("=== DISM — Restore Health ===")
        # Bug fix: added 30-minute timeout. DISM on no-internet can hang forever.
        dism_result = run(
            ["DISM", "/Online", "/Cleanup-Image", "/RestoreHealth"],
            timeout=1800,
            check=False,
        )
        if dism_result.returncode == 0:
            _log("DISM repair succeeded. Running SFC again to verify...")
            run(["sfc", "/scannow"], timeout=1800, check=False)
        else:
            _log(
                "DISM could not restore health. Component store may be corrupt, "
                "or internet access is required for source files.",
                "WARN",
            )

    _log("Windows repair checks complete.")


# ─────────────────────────────────────────────
# macOS — diskutil First Aid on all volumes
# ─────────────────────────────────────────────

def repair_macos() -> None:
    _log("=== macOS Disk First Aid ===")

    result = run(["diskutil", "list", "-plist"], check=False)
    # Run First Aid on all internal disks
    disks_result = run(["diskutil", "list"], check=False)
    disks = [
        line.strip().split()[0]
        for line in disks_result.stdout.splitlines()
        if line.strip().startswith("/dev/disk")
    ]

    for disk in disks:
        _log(f"Running First Aid on {disk}...")
        run(["diskutil", "repairVolume", disk], check=False, timeout=600)

    # Check system log for kernel panics and disk errors
    _log("=== Checking system logs for disk errors ===")
    run(["log", "show", "--predicate", "eventMessage contains 'disk error'",
         "--last", "24h"], check=False, timeout=30)

    _log("macOS repair checks complete.")


# ─────────────────────────────────────────────
# Linux — journal checks + recommend fsck on unmount
# ─────────────────────────────────────────────

def repair_linux() -> None:
    _log("=== Linux System Integrity Checks ===")

    # Check systemd journal for critical errors
    _log("Checking systemd journal for errors...")
    run(["journalctl", "-p", "err", "-b", "--no-pager"], check=False, timeout=30)

    # List filesystems and their fsck status
    _log("Filesystem status:")
    run(["df", "-h"], check=False)
    run(["lsblk", "-o", "NAME,FSTYPE,MOUNTPOINT,SIZE,STATE"], check=False)

    # Check dmesg for disk errors
    _log("Checking kernel log for disk I/O errors...")
    result = run(["dmesg", "--level=err,crit,alert,emerg"], check=False, timeout=10)
    if "error" in result.stdout.lower() or "failed" in result.stdout.lower():
        _log("Disk-related errors found in kernel log — recommend running fsck offline.", "WARN")
    else:
        _log("No critical disk errors in kernel log.")

    _log("NOTE: fsck on mounted filesystems is not safe. To deep-scan:")
    _log("  Boot from RepairDock USB → run: fsck -f /dev/sdX")

    _log("Linux repair checks complete.")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="RepairDock AI — Cross-platform system repair checks"
    )
    parser.add_argument(
        "--platform-override",
        choices=["windows", "macos", "linux"],
        help="Force a specific platform (for testing)",
    )
    args = parser.parse_args()

    effective_platform = args.platform_override or SYSTEM
    _log(f"RepairDock AI — System Repair — platform: {effective_platform}")
    require_admin()

    platform_key = effective_platform.lower()
    if platform_key == "windows":
        repair_windows()
    elif platform_key in ("darwin", "macos"):
        repair_macos()
    elif platform_key == "linux":
        repair_linux()
    else:
        _die(f"Unsupported platform: {effective_platform}")


if __name__ == "__main__":
    main()
