import platform
import subprocess
import sys
import json
import csv
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, scrolledtext
import webbrowser

# ------------------ HELPERS ------------------

def run_ps(cmd):
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", cmd],
            capture_output=True,
            text=True
        )
        return result.stdout.strip()
    except:
        return ""

def ps_json(cmd):
    out = run_ps(cmd + " | ConvertTo-Json -Depth 6")
    try:
        return json.loads(out) if out else None
    except:
        return None

def is_admin():
    try:
        return subprocess.run(
            ["net", "session"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        ).returncode == 0
    except:
        return False

# ------------------ DATA COLLECTION ------------------

def cpu():
    return ps_json("Get-CimInstance Win32_Processor | Select Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed")

def ram():
    return ps_json("Get-CimInstance Win32_PhysicalMemory | Select Capacity,Speed")

def gpu():
    return ps_json("Get-CimInstance Win32_VideoController | Select Name")

def disks():
    return ps_json("Get-PhysicalDisk | Select FriendlyName,BusType,Size,FirmwareVersion")

def nvme_controllers():
    return ps_json("""
    Get-CimInstance -Namespace root\\Microsoft\\Windows\\Storage MSFT_NvmeController |
    Select FriendlyName,FirmwareVersion,PcieLinkWidth,PcieLinkSpeed
    """)

def nvme_wear():
    return ps_json("""
    Get-CimInstance -Namespace root\\Microsoft\\Windows\\Storage MSFT_NvmeSmartHealthInformation |
    Select PercentageUsed
    """)

def smart_health():
    return ps_json("""
    Get-CimInstance -Namespace root\\wmi MSStorageDriver_FailurePredictStatus |
    Select InstanceName,PredictFailure
    """)

def bitlocker():
    return ps_json("Get-BitLockerVolume | Select MountPoint,VolumeStatus,EncryptionPercentage")

def bios():
    return ps_json("Get-CimInstance Win32_BIOS | Select Manufacturer,SMBIOSBIOSVersion,ReleaseDate")

def secure_boot():
    try:
        return run_ps("Confirm-SecureBootUEFI")
    except:
        return "Unsupported"

def tpm():
    t = ps_json("Get-Tpm")
    if not t:
        return None
    return {
        "present": t.get("TpmPresent"),
        "ready": t.get("TpmReady"),
        "version": "2.0" if str(t.get("ManufacturerVersion","")).startswith("2") else "1.2"
    }

def boot_mode():
    return "UEFI" if "UEFI" in run_ps("bcdedit") else "Legacy"

# ------------------ AGGREGATOR ------------------

def collect_all_data(progress, log):
    steps = [
        ("CPU", cpu),
        ("RAM", ram),
        ("GPU", gpu),
        ("Disks", disks),
        ("NVMe Controllers", nvme_controllers),
        ("NVMe Wear", nvme_wear),
        ("SMART Health", smart_health),
        ("BitLocker", bitlocker),
        ("BIOS", bios),
        ("Secure Boot", secure_boot),
        ("TPM", tpm),
        ("Boot Mode", boot_mode)
    ]

    data = {
        "timestamp": datetime.now().isoformat(),
        "hostname": platform.node(),
        "admin": is_admin(),
        "os": {
            "name": platform.system(),
            "release": platform.release(),
            "version": platform.version(),
            "arch": platform.architecture()[0]
        }
    }

    for i, (name, func) in enumerate(steps, start=1):
        log(f"Collecting {name}...\n")
        data[name.lower().replace(" ", "_")] = func()
        progress["value"] = (i / len(steps)) * 100
        progress.update()

    return data

# ------------------ EXPORT ------------------

def export_reports(data, base):
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    host = platform.node()

    html = base / f"audit_{host}_{stamp}.html"
    js = base / f"audit_{host}_{stamp}.json"
    csvf = base / f"audit_{host}_{stamp}.csv"

    # HTML
    with open(html, "w", encoding="utf-8") as f:
        f.write(f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Hardware Audit</title>
<style>
body {{
    background:#121212;
    color:#e0e0e0;
    font-family:Segoe UI,Arial;
}}
h1,h2 {{ color:#90caf9; }}
pre {{
    background:#1e1e1e;
    padding:15px;
    border-radius:8px;
}}
</style>
</head>
<body>
<h1>Hardware & Security Audit</h1>
<h2>{data['hostname']} — {data['timestamp']}</h2>
<pre>{json.dumps(data, indent=2)}</pre>
</body>
</html>
""")

    with open(js, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

    with open(csvf, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Section", "Key", "Value"])
        for section, content in data.items():
            if isinstance(content, dict):
                for k, v in content.items():
                    writer.writerow([section, k, v])
            else:
                writer.writerow([section, "", content])

    return html, js, csvf

# ------------------ GUI ------------------

def gui():
    root = tk.Tk()
    root.title("Windows Hardware & Security Audit")
    root.geometry("800x500")
    root.configure(bg="#121212")

    style = ttk.Style()
    style.theme_use("default")
    style.configure(
        "TProgressbar",
        background="#90caf9",
        troughcolor="#1e1e1e",
        thickness=20
    )

    output = scrolledtext.ScrolledText(
        root,
        bg="#1e1e1e",
        fg="#e0e0e0",
        insertbackground="white"
    )
    output.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    progress = ttk.Progressbar(root, length=700, mode="determinate")
    progress.pack(pady=5)

    def log(msg):
        output.insert(tk.END, msg)
        output.see(tk.END)
        output.update()

    def run_audit():
        try:
            output.delete("1.0", tk.END)
            progress["value"] = 0

            log(f"Admin privileges: {'YES' if is_admin() else 'NO'}\n\n")

            data = collect_all_data(progress, log)

            base = Path(sys.executable if getattr(sys, "frozen", False) else __file__).parent
            html, js, csvf = export_reports(data, base)

            log("\nAudit complete.\n")
            log(f"HTML: {html}\nJSON: {js}\nCSV: {csvf}\n")

            webbrowser.open(html)

        except Exception as e:
            log(f"\nERROR:\n{e}\n")

    btn = tk.Button(
        root,
        text="Run Audit",
        command=run_audit,
        bg="#1e1e1e",
        fg="#90caf9",
        relief="flat",
        padx=20,
        pady=8
    )
    btn.pack(pady=10)

    root.mainloop()

# ------------------ ENTRY ------------------

if __name__ == "__main__":
    if platform.system() != "Windows":
        print("Windows only.")
        sys.exit(1)
    gui()
