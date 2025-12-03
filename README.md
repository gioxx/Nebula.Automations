# Nebula.Automations (formerly Nebula.Tools)

**Nebula.Automations** provides reusable PowerShell functions for scripting, automation and cloud integration.

![PowerShell Gallery](https://img.shields.io/powershellgallery/v/Nebula.Automations?label=PowerShell%20Gallery)
![Downloads](https://img.shields.io/powershellgallery/dt/Nebula.Automations?color=blue)

> [!NOTE]  
> **Why the name change?**  
> I want this module to remain specifically designed and developed to be an integral part of other scripts and automations that have little to do instead with everyday use tools. Nebula.Tools will return in another form, with different functions related to everyday use via PowerShell.  
> I apologize if this is confusing, I realize I could have thought of this long before!  
> The old GUID was invalidated and I removed the old versions from the GitHub releases. It's starting from scratch. I recommend you do the same with any installations on your machine!

---

## 📦 Installation

Install from PowerShell Gallery:

```powershell
Install-Module -Name Nebula.Automations -Scope CurrentUser
```

---

## 🚀 Usage

All documentation for using the module is available at **[kb.gioxx.org/docs/Nebula/Automations](https://kb.gioxx.org/docs/Nebula/Automations/intro)**.

---

## 🧽 How to clean up old module versions (optional)

When updating from previous versions, old files (such as unused `.psm1`, `.yml`, or `LICENSE` files) are not automatically deleted.  
If you want a completely clean setup, you can remove all previous versions manually:

```powershell
# Remove all installed versions of the module
Uninstall-Module -Name Nebula.Automations -AllVersions -Force

# Reinstall the latest clean version
Install-Module -Name Nebula.Automations -Scope CurrentUser -Force
```

ℹ️ This is entirely optional — PowerShell always uses the most recent version installed.

---

## 📄 License

All scripts in this repository are licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## 🔧 Development

This module is part of the [Nebula](https://github.com/gioxx?tab=repositories&q=Nebula) PowerShell tools family.

Feel free to fork, improve and submit pull requests.

---

## 📬 Feedback and Contributions

Feedback, suggestions, and pull requests are welcome!  
Feel free to [open an issue](https://github.com/gioxx/Nebula.Automations/issues) or contribute directly.