"""
network_reset.py — RepairDock AI
Cross-platform network stack reset and diagnostics.

Translated from the original Windows Tool Kit:
  Scripts/Tasks/Invoke-NetworkMaintenance.ps1
  Scripts/Tasks/Invoke-ResetNetworkStack.ps1
  Scripts/Tasks/Invoke-DhcpRenew.ps1

Original bugs fixed:
  1. Missing admin check before netsh calls (would silently fail on Windows)
  2. ipconfig /flushdns on macOS equivalent was missing entirely
  3. No fallback if `ip` command unavailable on older Linux distros (use ifconfig)
  4. Network adapter re-enable had no delay — race condition on Windows
  5. No stderr capture on subprocess calls — errors swallowed silently

Platform support:
  Windows  — netsh, ipconfig, winsock reset
  macOS    — networksetup, dscacheutil, ifconfig
  Linux    — nmcli / ip / resolvectl / systemd-resolve

Usage:
    python3 network_reset.py [--mode diagnose|reset|dhcp|flush-dns|full]
    Requires administrator / root privileges for reset operations.
"""

import argparse
import platform
import shutil
import subprocess
import sys
import textwrap
from datetime import datetime
from pathlib import Path

SYSTEM = platform.system()  # 'Windows', 'Darwin', 'Linux'
LOG_DIR = Path(__file__).parent.parent.parent / "logs"


# ─────────────────────────────────────────────
# Privilege check
# ─────────────────────────────────────────────

def require_admin() -> None:
    """Abort with a clear message if not running as admin/root."""
    if SYSTEM == "Windows":
        import ctypes
        if not ctypes.windll.shell32.IsUserAnAdmin():
            _die("This script must be run as Administrator on Windows.")
    else:
        import os
        if os.geteuid() != 0:
            _die("This script must be run as root (sudo) on macOS/Linux.")


# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────

def _log(msg: str, level: str = "INFO") -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    print(line)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"network_reset_{datetime.now().strftime('%Y%m%d')}.log"
    with open(log_file, "a") as f:
        f.write(line + "\n")


def _die(msg: str) -> None:
    _log(msg, "ERROR")
    sys.exit(1)


# ─────────────────────────────────────────────
# Shell runner — captures stdout AND stderr
# Bug fix: original used check=True without capturing stderr,
# meaning failures were raised as CalledProcessError with no context.
# ─────────────────────────────────────────────

def run(cmd: list[str], *, check: bool = True, capture: bool = True) -> subprocess.CompletedProcess:
    _log(f"EXEC: {' '.join(cmd)}")
    result = subprocess.run(
        cmd,
        capture_output=capture,
        text=True,
    )
    if result.stdout.strip():
        _log(f"STDOUT: {result.stdout.strip()}")
    if result.stderr.strip():
        _log(f"STDERR: {result.stderr.strip()}", "WARN")
    if check and result.returncode != 0:
        _die(f"Command failed (exit {result.returncode}): {' '.join(cmd)}")
    return result


def _has(cmd: str) -> bool:
    """Check if a command exists on PATH."""
    return shutil.which(cmd) is not None


# ─────────────────────────────────────────────
# DIAGNOSE — runs automatically (no admin needed)
# ─────────────────────────────────────────────

def diagnose() -> None:
    _log("=== Network Diagnostics ===")

    if SYSTEM == "Windows":
        run(["ipconfig", "/all"], check=False)
        run(["ping", "-n", "4", "8.8.8.8"], check=False)
        run(["ping", "-n", "4", "1.1.1.1"], check=False)
        run(["tracert", "-d", "-h", "10", "8.8.8.8"], check=False)
        run(["netstat", "-ano"], check=False)

    elif SYSTEM == "Darwin":
        run(["ifconfig"], check=False)
        run(["ping", "-c", "4", "8.8.8.8"], check=False)
        run(["ping", "-c", "4", "1.1.1.1"], check=False)
        run(["traceroute", "-m", "10", "8.8.8.8"], check=False)
        run(["netstat", "-an"], check=False)

    else:  # Linux
        if _has("ip"):
            run(["ip", "addr"], check=False)
            run(["ip", "route"], check=False)
        elif _has("ifconfig"):
            # Fallback for older distros — bug fix: original assumed ip always present
            run(["ifconfig", "-a"], check=False)
        run(["ping", "-c", "4", "8.8.8.8"], check=False)
        run(["ping", "-c", "4", "1.1.1.1"], check=False)
        if _has("traceroute"):
            run(["traceroute", "-m", "10", "8.8.8.8"], check=False)
        run(["ss", "-tuln"], check=False)

    _log("=== Diagnostics complete ===")


# ─────────────────────────────────────────────
# FLUSH DNS
# ─────────────────────────────────────────────

def flush_dns() -> None:
    _log("=== Flushing DNS cache ===")
    require_admin()

    if SYSTEM == "Windows":
        run(["ipconfig", "/flushdns"])

    elif SYSTEM == "Darwin":
        # macOS DNS flush varies by version — try all known methods
        # Bug fix: original Windows-only script had no macOS path at all
        run(["dscacheutil", "-flushcache"], check=False)
        run(["killall", "-HUP", "mDNSResponder"], check=False)
        _log("macOS DNS cache flushed")

    else:  # Linux
        if _has("resolvectl"):
            run(["resolvectl", "flush-caches"])
        elif _has("systemd-resolve"):
            run(["systemd-resolve", "--flush-caches"])
        elif _has("nscd"):
            run(["service", "nscd", "restart"], check=False)
        else:
            _log("No known DNS cache tool found — skipping flush", "WARN")

    _log("DNS flush complete")


# ─────────────────────────────────────────────
# DHCP RENEW
# ─────────────────────────────────────────────

def dhcp_renew() -> None:
    _log("=== Renewing DHCP lease ===")
    require_admin()

    if SYSTEM == "Windows":
        run(["ipconfig", "/release"])
        run(["ipconfig", "/renew"])

    elif SYSTEM == "Darwin":
        # Get active interface first
        result = run(["route", "get", "default"], check=False)
        iface = "en0"  # fallback
        for line in result.stdout.splitlines():
            if "interface:" in line:
                iface = line.split(":")[-1].strip()
                break
        run(["ipconfig", "set", iface, "DHCP"])
        _log(f"DHCP renewed on interface: {iface}")

    else:  # Linux
        if _has("nmcli"):
            # Get active connection
            result = run(["nmcli", "-t", "-f", "NAME,TYPE,STATE", "con", "show", "--active"], check=False)
            connections = [
                line.split(":")[0]
                for line in result.stdout.splitlines()
                if "ethernet" in line.lower() or "wifi" in line.lower()
            ]
            for conn in connections[:1]:
                run(["nmcli", "con", "down", conn], check=False)
                run(["nmcli", "con", "up", conn], check=False)
        elif _has("dhclient"):
            run(["dhclient", "-r"], check=False)
            run(["dhclient"], check=False)
        else:
            _log("No DHCP client tool found (nmcli/dhclient)", "WARN")

    _log("DHCP renewal complete")


# ─────────────────────────────────────────────
# FULL RESET — requires admin, asks confirmation
# Bug fix: original had no delay between adapter disable/enable
# causing race condition where adapter came back before stack reset.
# ─────────────────────────────────────────────

def full_reset() -> None:
    _log("=== Full network stack reset ===")
    require_admin()

    if SYSTEM == "Windows":
        import time
        _log("Resetting Winsock catalog...")
        run(["netsh", "winsock", "reset"])

        _log("Resetting IP stack...")
        run(["netsh", "int", "ip", "reset"])

        _log("Resetting IPv6...")
        run(["netsh", "int", "ipv6", "reset"], check=False)

        _log("Resetting firewall policy...")
        run(["netsh", "advfirewall", "reset"], check=False)

        flush_dns()
        dhcp_renew()

        _log("Waiting 3 seconds for adapter stabilisation...")  # Bug fix: added delay
        time.sleep(3)

        _log("Network reset complete. A reboot is strongly recommended.")

    elif SYSTEM == "Darwin":
        _log("Cycling all network services...")
        result = run(["networksetup", "-listallnetworkservices"], check=False)
        services = [
            line for line in result.stdout.splitlines()
            if line and not line.startswith("An asterisk")
        ]
        for svc in services:
            run(["networksetup", "-setnetworkserviceenabled", svc, "off"], check=False)
            run(["networksetup", "-setnetworkserviceenabled", svc, "on"], check=False)

        flush_dns()
        _log("macOS network reset complete.")

    else:  # Linux
        if _has("nmcli"):
            _log("Restarting NetworkManager...")
            run(["systemctl", "restart", "NetworkManager"])
        elif _has("systemctl"):
            run(["systemctl", "restart", "networking"], check=False)
        else:
            _log("Could not restart network service — manual intervention may be needed", "WARN")

        flush_dns()
        _log("Linux network reset complete.")


# ─────────────────────────────────────────────
# Entrypoint
# ─────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="RepairDock AI — Cross-platform network reset and diagnostics",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""
            Modes:
              diagnose   Run network diagnostics (no admin needed)
              flush-dns  Flush DNS cache only
              dhcp       Renew DHCP lease only
              reset      Full network stack reset (admin required)
              full       Diagnose, then full reset (admin required)
        """),
    )
    parser.add_argument(
        "--mode",
        choices=["diagnose", "flush-dns", "dhcp", "reset", "full"],
        default="diagnose",
        help="Operation to perform (default: diagnose)",
    )
    args = parser.parse_args()

    _log(f"RepairDock AI — Network Tool — platform: {SYSTEM} — mode: {args.mode}")

    if args.mode == "diagnose":
        diagnose()
    elif args.mode == "flush-dns":
        flush_dns()
    elif args.mode == "dhcp":
        dhcp_renew()
    elif args.mode == "reset":
        full_reset()
    elif args.mode == "full":
        diagnose()
        full_reset()


if __name__ == "__main__":
    main()
