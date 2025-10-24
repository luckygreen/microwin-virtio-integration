# MicroWin VirtIO Integration

Automated VirtIO driver integration for MicroWin Windows 11 ISOs. This PowerShell script post-processes MicroWin ISOs to create fully "first-boot-ready" virtual machine images with complete VirtIO support.

**Works with any VirtIO-capable virtualization platform:** Proxmox VE, QEMU, KVM, oVirt, and more.

## 🎯 What This Does

This script eliminates all manual driver installation steps by:

- **Injecting all 10 VirtIO drivers** into both `install.wim` (installed system) and `boot.wim` (Windows installer/WinPE)
- **Bundling virtio-win-guest-tools.exe** and configuring automatic silent installation after OOBE
- **Generating bootable ISOs** with UEFI/BIOS dual-boot support
- **Auto-detecting source files** in your working directory
- **Smart ISO naming** with descriptive volume labels (under 32 characters)

### The Problem This Solves

Deploying Windows VMs with VirtIO hardware typically involves a frustrating multi-step process with numerous manual interventions:

#### Problem #1: The Installer Can't See Your Disk
The Windows installer immediately greets you with **"no drives found"** because `boot.wim` (the Windows PE environment) doesn't include VirtIO storage drivers. You're forced to either:
- Load drivers manually during setup (requires mounting the VirtIO ISO, browsing folders, multiple clicks)
- Use IDE/SATA emulation as a workaround, then reconfigure to VirtIO later
- Create a custom answer file to inject drivers (complex and error-prone)

#### Problem #2: Missing Drivers After Installation
Even after getting Windows installed, your work isn't done. Opening Device Manager reveals a sea of yellow warning triangles:
- Network adapter not working (no connectivity)
- Unknown devices for balloon driver, serial ports, RNG device
- Missing drivers for input devices and other VirtIO hardware
- Each driver requires manual installation, system reboots, and verification

#### Problem #3: Guest Tools Installation
To get full VM integration (QEMU Guest Agent for proper shutdown, snapshots, and monitoring), you need to:
- Mount the VirtIO ISO again in the running VM
- Manually run virtio-win-guest-tools.exe
- Wait for installation and reboot
- Verify services are running correctly

#### Problem #4: Repeat for Every VM
Every single Windows VM deployment means repeating this entire process. For organizations deploying multiple VMs, this becomes a significant time sink and potential source of errors.

### This Script's Solution

**This script eliminates ALL of these problems completely.**

Your output ISO:
- ✅ **Boots and installs immediately** with VirtIO storage visible from the start
- ✅ **Has zero missing drivers** - Device Manager is completely clean after first boot
- ✅ **Installs guest tools automatically** - QEMU Guest Agent, SPICE components, and all VirtIO services start automatically after OOBE
- ✅ **Requires zero manual intervention** - True "hands-off" deployment from ISO boot to fully functional VM

The result: **What used to take 30-60 minutes of manual work per VM now takes zero.** Just boot the ISO, let Windows install, and you have a fully functional VirtIO-enabled VM with all drivers and tools pre-configured.

## ✨ Features

- ✅ **Complete automation** - Zero manual driver installation required
- ✅ **Boot.wim injection** - Installer can see VirtIO storage immediately (no IDE workaround!)
- ✅ **Install.wim injection** - All drivers available in installed system
- ✅ **Guest tools auto-install** - QEMU Guest Agent and SPICE components install silently after OOBE
- ✅ **Smart file detection** - Automatically finds MicroWin ISO, VirtIO ISO, and guest tools
- ✅ **Proper bootable ISOs** - Full UEFI and BIOS support via oscdimg
- ✅ **Clean volume labels** - Descriptive names under 32 characters (ISO9660 compatible)

## 🚀 Quick Start

### Prerequisites

- Windows with PowerShell 5.1+ (run as Administrator)
- [MicroWin](https://github.com/ChrisTitusTech/winutil) processed Windows 11 ISO
- [VirtIO drivers ISO](https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/stable-virtio/) (e.g., virtio-win-0.1.285.iso)
- [virtio-win-guest-tools.exe](https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/stable-virtio/) (optional but recommended)

### Usage

1. **Place all files in the same directory:**
   ```
   your-folder/
   ├── Integrate-VirtIO-MicroWin.ps1
   ├── MicroWin11_25H2_English_x64.iso
   ├── virtio-win-0.1.285.iso
   └── virtio-win-guest-tools.exe (optional)
   ```

2. **Run the script:**
   ```powershell
   .\Integrate-VirtIO-MicroWin.ps1
   ```

3. **Wait for completion** (~15-20 minutes)

4. **Use the output ISO** with your VirtIO-enabled VM!

### Manual File Selection

If you prefer to specify files explicitly:

```powershell
.\Integrate-VirtIO-MicroWin.ps1 `
    -MicroWinISO ".\MicroWin11_25H2_English_x64.iso" `
    -VirtIOISO ".\virtio-win-0.1.285.iso" `
    -GuestToolsExe ".\virtio-win-guest-tools.exe"
```

## 📋 VirtIO Drivers Included

The script injects all essential VirtIO drivers for Windows 11:

| Driver | Purpose |
|--------|---------|
| **NetKVM** | Virtual network adapter |
| **viostor** | VirtIO SCSI block storage |
| **vioscsi** | VirtIO SCSI controller |
| **Balloon** | Dynamic memory management |
| **viorng** | Random number generator |
| **vioserial** | Virtual serial port |
| **qemupciserial** | PCI-based serial |
| **vioinput** | Keyboard and mouse input |
| **pvpanic** | Guest crash notification |
| **viofs** | VirtIO-FS shared filesystem |

## 🎓 How It Works

1. **Mounts source ISOs** - MicroWin and VirtIO ISOs
2. **Extracts ISO contents** - Copies MicroWin ISO to working directory
3. **Verifies boot files** - Ensures BIOS and UEFI boot files exist
4. **Injects drivers into install.wim** - Adds VirtIO drivers to the installed Windows image
5. **Injects drivers into boot.wim** - Adds VirtIO drivers to Windows installer/WinPE environment
6. **Adds guest tools** - Copies virtio-win-guest-tools.exe and creates SetupComplete.cmd for auto-install
7. **Creates bootable ISO** - Uses oscdimg to generate properly bootable ISO with dual-boot support
8. **Cleans up** - Removes temporary files

## 🖥️ Tested Platforms

- ✅ Proxmox VE (primary target)
- ✅ QEMU/KVM
- ✅ oVirt
- ✅ Any other VirtIO-capable virtualization platform

## 📝 Example Output

```
========================================
Success!
========================================
New ISO: MicroWin11_25H2_Eng_x64_VIO285.iso

Features integrated:
  ✓ VirtIO drivers injected into install.wim (post-install)
  ✓ VirtIO drivers injected into boot.wim (installer phase)
  ✓ Guest tools auto-install via SetupComplete.cmd
  ✓ Full hands-off VirtIO deployment - no IDE workaround needed!
  ✓ Volume label: MicroWin11_25H2_Eng_x64_VIO285
```

## 🤝 Contributing

This script bridges a gap that we believe should be filled directly by MicroWin. We strongly encourage and hope the [MicroWin development team](https://github.com/ChrisTitusTech/winutil) will consider integrating this functionality into their excellent tool.

Until then, contributions, bug reports, and feature requests are welcome!

## 📜 License

This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or distribute this software, either in source code form or as a compiled binary, for any purpose, commercial or non-commercial, and by any means.

See the [LICENSE](LICENSE) file for details, or visit [unlicense.org](https://unlicense.org/).

## 👥 Authors

- Claude (AI Assistant)
- Lucky Green <<shamrock@cypherpunks.to>>

## 🙏 Acknowledgments

- [Chris Titus Tech](https://github.com/ChrisTitusTech) for the excellent MicroWin tool
- [Red Hat](https://www.redhat.com/) for the VirtIO drivers
- The Proxmox and QEMU communities

## 📞 Support

For issues, questions, or suggestions, please [open an issue](https://github.com/luckygreen/microwin-virtio-integration/issues) on GitHub.

---

**Note:** This script requires Administrator privileges and processes large files (Windows installation images). Ensure you have adequate disk space (~15-20 GB free) and time (~15-20 minutes) for the operation to complete.
