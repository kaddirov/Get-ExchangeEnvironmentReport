# 📧 Exchange Environment Report 3.1

![PowerShell](https://img.shields.io/badge/PowerShell-5.1-blue.svg) ![Exchange](https://img.shields.io/badge/Exchange-2016%20%7C%202019%20%7C%20SE-orange.svg)


A high-performance PowerShell script to generate **modern HTML dashboards** for your Microsoft Exchange Server infrastructure.

## ⭐ Credits & Acknowledgements
This project is a modernized **3.1** evolution by **B.O** based on the original work by **Steve Goodman** and **Thomas Stensitzki**.
*   Original Script: [Get-ExchangeEnvironmentReport](https://github.com/Apoc70/Get-ExchangeEnvironmentReport/blob/master/Get-ExchangeEnvironmentReport.ps1)

---

## 🚀 Features
*   **⚡ Ultra-Fast:** Uses "Bulk Collection" logic (single query per server) to scan massive infrastructures in seconds.
*   **📊 KPI Dashboard:** Immediate overview (Active/Archive Mailboxes, Database Whitespace, Server Health).
*   **🎨 100% Customizable:** Adapt the report's logo, title, and colors to your corporate branding via simple parameters.
*   **🔋 Visual Gauges:** Instant monitoring of disk space (Database & Logs) with color codes.
*   **🛡️ Secure by Design:** No complex external calls. Easily hostable on IIS.
*   **🔍 Live Filtering:** Interactive sorting and filtering directly in your browser.
*   **📧 Email Support:** Automated delivery of HTML reports to your mailbox.
*   **📜 Certificate Badges:** Visual breakdown of IIS, SMTP, POP, IMAP, and AUTH certificates.

## 📋 Prerequisites
*   **OS:** Windows Server 2012 R2 or higher.
*   **Exchange:** Management Shell (EMS) installed.
*   **Permissions:** Member of `Organization Management` or `View-Only Organization Management` group.
*   **Network:** WinRM (5985) and RPC/DCOM open to target servers.

## 🛠️ Installation & Usage

1.  Download the `Get-ExchangeEnvironmentReport.ps1` file.
2.  Run the script from an **Exchange Management Shell** (Administrator).

### Standard Command
```powershell
.\Get-ExchangeEnvironmentReport.ps1 `
    -HTMLReport "C:\inetpub\wwwroot\Report\Exchange_Status.html"
```

### Customization (White Label)

You can adapt the report to your visual identity:

```powershell
.\Get-ExchangeEnvironmentReport.ps1 `
    -HTMLReport "C:\Reports\Dashboard.html" `
    -CompanyLogo "MYCORP" `
    -ReportTitle "MESSAGING OPS" `
    -ThemeColor "#0078D4"
```

| Parameter | Description | Default |
|-----------|-------------|--------|
| `-CompanyLogo` | The "Logo" text top left (e.g., Company Name). | `EXCHANGE` |
| `-ReportTitle` | The main title of the report. | `REPORTING` |
| `-ThemeColor` | HEX color code for titles, borders, and theme. | `#F27A00` |

### Sending via Email
```powershell
.\Get-ExchangeEnvironmentReport.ps1 `
    -HTMLReport "C:\Reports\Daily.html" `
    -SendMail $true `
    -MailFrom "report@mycorp.com" `
    -MailTo "admins@mycorp.com" `
    -MailServer "smtp.mycorp.com"
```

## 🛡️ Security (IIS)
To restrict access to the HTML report (e.g., Internal only), it is recommended to configure **IP Address and Domain Restrictions** directly on your IIS web server. The script does not handle authentication to remain lightweight and compatible.

## 🤝 Contribution
This script is a "Community Standard". Feel free to open Pull Requests to add DAG metrics, Transport queues, or optimizations.



---

## 📝 Changelog

### [V3.1] - 2026-05-04
*   **Added**: Restored **Email Delivery** functionality (`Send-MailMessage`).
*   **Added**: **Client-side filtering** and sorting in the HTML report.
*   **Added**: **Advanced Certificate Analysis** with service badges (IIS, SMTP, etc.).
*   **Added**: Visible versioning in console and HTML footer/header.
*   **Added**: New **Standard Edition** in a separate folder for generic usage.

**License:** MIT
**Author:** B.O
