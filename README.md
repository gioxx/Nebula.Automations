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

## ✨ Included Functions

- `Send-Mail`  
  Send emails via SMTP with support for:
  - Attachments
  - CC / BCC
  - Custom SMTP server and port

- `CheckMGGraphConnection`  
  Connect to Microsoft Graph using application credentials:
  - Automatically handles module install
  - Authenticates with client ID/secret
  - Logs connection status

---

## 📦 Installation

Install from PowerShell Gallery:

```powershell
Install-Module -Name Nebula.Automations -Scope CurrentUser
```

---

## 🚀 Usage

Example to send an email:

```powershell
Send-Mail -SMTPServer "smtp.example.com" -From "me@example.com" -To "you@example.com" -Subject "Hello" -Body "Test message"
```

Example to connect to Microsoft Graph:

```powershell
$Graph = CheckMGGraphConnection -tenantId "<tenant>" -clientId "<client>" -clientSecret "<secret>"
```

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

## 🔧 Development

This module is part of the [Nebula](https://github.com/gioxx?tab=repositories&q=Nebula) PowerShell tools family.

Feel free to fork, improve and submit pull requests.

---

## 📄 License

Licensed under the [MIT License](https://opensource.org/licenses/MIT).
