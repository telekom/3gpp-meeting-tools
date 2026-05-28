import platform
import re
import socket
import subprocess
import json
import psutil
from typing import Tuple, List, Dict, Any


def _check_windows_vpn() -> Tuple[bool, List[Dict[str, str]]]:
    """
    Detects active VPNs on Windows using PowerShell.
    Implements strict data validation, encoding safety, and structural filtering.
    """
    vpn_keywords: List[str] = [
        'cisco', 'anyconnect', 'vpn', 'tap', 'virtual',
        'wireguard', 'fortinet', 'globalprotect'
    ]

    cmd: List[str] = [
        "powershell", "-NoProfile", "-Command",
        "Get-NetAdapter | Select-Object Name, InterfaceDescription, Status | ConvertTo-Json"
    ]

    try:
        # Added encoding='utf-8' and errors='ignore' to prevent UnicodeDecodeError
        # if localized Windows environments return special characters (e.g., ä, ö, é)
        result = subprocess.run(
            cmd, capture_output=True, text=True, check=True,
            encoding='utf-8', errors='ignore'
        )

        output = result.stdout.strip()

        if not output:
            return False, []

        raw_data = json.loads(output)

        # Ensure we are always iterating over a list
        if isinstance(raw_data, dict):
            raw_data = [raw_data]

        # --- PHASE 1: Data Validation & Filtering ---
        valid_adapters: List[Dict[str, str]] = []

        for item in raw_data:
            if not isinstance(item, dict):
                continue

            name: Any = item.get('Name')
            desc: Any = item.get('InterfaceDescription')
            status: Any = item.get('Status')

            # Reject the entry entirely if any critical data is missing (null)
            if name is None or desc is None or status is None:
                continue

            valid_adapters.append({
                "Name": str(name).strip(),
                "Description": str(desc).strip(),
                "Status": str(status).strip().lower()
            })

        # --- PHASE 2: Logic Processing ---
        active_vpns: List[Dict[str, str]] = []

        for adapter in valid_adapters:
            status = adapter["Status"]
            desc_lower = adapter["Description"].lower()

            if status == 'up' and any(keyword in desc_lower for keyword in vpn_keywords):
                active_vpns.append({
                    "Name": adapter["Name"],
                    "Description": adapter["Description"]
                })

        return len(active_vpns) > 0, active_vpns

    except FileNotFoundError:
        print("Error: PowerShell executable not found in system PATH.")
        return False, []
    except json.JSONDecodeError:
        print("Error: Windows returned malformed JSON data.")
        return False, []
    except Exception as e:
        print(f"Windows VPN check encountered an unexpected error: {e}")
        return False, []


def _check_unix_vpn() -> Tuple[bool, List[Dict[str, str]]]:
    """
    Detects active VPNs on Linux/macOS using psutil.
    Includes try/except blocks for restricted environments (like Docker or strict VMs).
    """
    vpn_indicators: List[str] = [
        'tun', 'tap', 'ppp', 'wg', 'utun', 'ipsec', 'cscotun'
    ]

    active_vpns: List[Dict[str, str]] = []

    try:
        interfaces = psutil.net_if_addrs().keys()

        for interface in interfaces:
            interface_lower: str = interface.lower()

            if any(indicator in interface_lower for indicator in vpn_indicators):
                stats = psutil.net_if_stats().get(interface)

                if stats is not None and stats.isup:
                    active_vpns.append({
                        "Name": interface,
                        "Description": "Virtual Network Tunnel"
                    })

        return len(active_vpns) > 0, active_vpns

    except PermissionError:
        print("Error: Insufficient permissions to read network interfaces.")
        return False, []
    except Exception as e:
        print(f"Unix VPN check encountered an unexpected error: {e}")
        return False, []


def is_vpn_active() -> Tuple[bool, List[Dict[str, str]]]:
    """
    Main Entry Point: Detects if a VPN is currently active, supporting Windows, Linux, and macOS.
    """
    current_os: str = platform.system()

    if current_os == "Windows":

        vpn_status = _check_windows_vpn()
        # print(f'Checking Windows: {vpn_status}')
        return vpn_status
    elif current_os in ["Linux", "Darwin"]:
        vpn_status = _check_unix_vpn()
        # print(f'Checking Linux: {vpn_status}')
        return vpn_status
    else:
        print(f"Unsupported Operating System: {current_os}")
        return False, []


# --- Example Usage ---
if __name__ == "__main__":
    is_connected, vpn_adapters = is_vpn_active()

    print(f"System Detected: {platform.system()}\n" + "-" * 30)

    if is_connected:
        print("🔒 VPN is ACTIVE.")
        for adapter in vpn_adapters:
            print(f" -> Interface: {adapter['Name']}")
            print(f" -> Hardware:  {adapter['Description']}")
    else:
        print("🌐 No active VPN detected. You are on your standard connection.")


def we_are_in_meeting_network():
    # Before, 10.10.10.10 used only FTP, so we had to differentiate between files and folders. Now,
    # we can always just use HTTP (albeit no HTTPs in 10.10.10.10)
    ip_addresses = [i[4][0] for i in socket.getaddrinfo(socket.gethostname(), None)]
    matches = [re.match(r'10.10.(\d)+.(\d)+', ip_address) for ip_address in ip_addresses]
    matches = [match for match in matches if match is not None]
    ip_is_meeting_ip = (len(matches) != 0)
    return ip_is_meeting_ip
