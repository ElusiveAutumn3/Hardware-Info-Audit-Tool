# Hardware-Info-Audit-Tool
Lightweight GUI tool to scan CPU/GPU/RAM/disks/NVMe health, BitLocker, TPM, Secure Boot + export HTML/JSON/CSV reports.
Standalone Windows tool that performs a quick yet thorough audit of your system's hardware and key security features. It collects detailed info on CPU, RAM, GPU, physical disks, NVMe controllers & wear levels, SMART health, BitLocker status, BIOS version, Secure Boot, TPM presence/readiness, and boot mode — all via native PowerShell queries.

Features include:

• Comprehensive hardware overview (CPU cores/speed, RAM modules/speed, GPU name, disk types/sizes/firmware, NVMe PCIe link & % used)
• Security posture snapshot (BitLocker encryption, Secure Boot state, TPM 1.2/2.0 readiness)
• SMART failure prediction & general disk health
• UEFI vs Legacy boot detection
• Dark-mode GUI with scrolled log and progress tracking
• Exports professional-looking reports (HTML with JSON-style pre for easy reading, plus machine-readable JSON & CSV)
No bloat, no telemetry, no internet needed — built for Windows admins, technicians, and power users who want reliable info without running heavy suites like HWiNFO or Speccy every time.
