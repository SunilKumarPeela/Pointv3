# Point 

[![C++20](https://img.shields.io/badge/language-C%2B%2B20-blue.svg)](https://en.wikipedia.org/wiki/C%2B%2B20)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011%20(x64)-0078D6.svg)](https://www.microsoft.com/windows)
[![Security Standard](https://img.shields.io/badge/compliance-NIST%20CSF%202.0%20%7C%20PCI%20DSS%204.0.1%20%7C%20GDPR-success.svg)](COMPLIANCE.md)
[![License](https://img.shields.io/badge/license-Proprietary-lightgrey.svg)](LICENSE.txt)

> **Zero-Trust, Local-First Analytics Workspace for Security, Identity, and Operational Data.**

Point is a native, dependency-free Windows data workspace engineered to correlate, cross-reference, and analyze multi-source CSV and Excel datasets without cloud exposure, database overhead, or spreadsheet formula fragility[cite: 4, 8].

---

## ⚡ Overview & Architectural Highlights

Point operates as an air-gapped desktop workspace where datasets are parsed and evaluated strictly in local memory[cite: 4, 8]. It enables security analysts, compliance officers, and IT operations teams to join disparate identity and operational data at scale[cite: 4, 8].

* **Local-First Isolation**: Core binary (`Point.exe`) executes entirely offline with zero outbound network calls, telemetry, scripting hooks, or shell execution[cite: 4, 8].
* **2-Million-Row Virtual Grid**: Sparse Win32 control architecture renders up to 2,000,000 rows and 256 columns with viewport-only updates and sub-millisecond scrolling[cite: 8, 18].
* **Automated Identity Mapping**: Discovers shared keys (`Employee ID`, `Username`, `Device ID`, `Asset Tag`, etc.) across disparate reports via a 64×64 Field Synonym Manager[cite: 8, 13, 18].
* **Read-Only Ingestion**: Ingests `.csv`, `.xlsx`, `.xls`, and `.xlsm` files safely, treating all original documents as immutable[cite: 8, 13, 15].
* **Network Boundary Segregation**: Outbound HTTPS report fetching is strictly isolated into an independent companion utility (`PointFetcher.exe`)[cite: 4, 8].

---

## 🔍 Multi-Mode Analytical Intelligence

Point includes specialized operational modes designed to analyze data without complex joins or formulas:

| Mode | Functionality |
|---|---|
| **Universal** | Type or paste known IDs down a single column to resolve cross-report profiles across all indexed datasets[cite: 8, 13]. |
| **Narrow** | Perform multi-field exact `AND` filtering with safe cross-report relationship joins[cite: 8, 13]. |
| **Count** | Compute grouped occurrence distributions; automatically resolves affected object IDs for counts $\le 50$[cite: 8, 13, 18]. |
| **Compare** | Compare two identity profiles across multiple evidence fields, producing Jaccard similarity scores and risk-flagged differences[cite: 8, 13]. |
| **Analyze** | Inspect key integrity across datasets for exact duplicates, diverging details, and missing keys[cite: 8, 13]. |
| **Insights** | Run offline, deterministic data quality profiling (missing-value measurement, cardinality checks, 1.5×IQR outlier detection)[cite: 8, 13]. |
| **Chart** | Interactive GDI-rendered Bar, Column, Pie, and 100%-Stacked visualizations with direct record drill-down[cite: 8, 18]. |
| **Change** | Track record-level changes (`Added`, `Removed`, `Modified`) against an in-memory baseline[cite: 8, 13]. |

---

## 🛡️ Enterprise Compliance & Security Controls

Point is built to align with **NIST CSF 2.0**, **PCI DSS 4.0.1**, and **GDPR** technical safeguard requirements[cite: 4, 8]:

* **Fail-Closed Access Control**: Enforces native Windows security token memberships (`Point Users`, `Point Administrators`, `Point Exporters`)[cite: 4, 11].
* **DPAPI Encryption**: Saved workspaces and synonym configurations are encrypted at rest using Windows DPAPI bound to the current user[cite: 4, 11, 18].
* **Data Loss Prevention (DLP)**: Built-in Luhn validation blocks exports containing valid payment card numbers (PCI DSS avoidance posture), and sensitive credentials/national IDs are masked[cite: 4, 13].
* **Tamper-Evident Audit Trail**: Generates an append-only UTC audit log secured via rolling BCrypt SHA-256 hash chains[cite: 4, 13].
* **Formula Injection Defense**: Escapes spreadsheet formula prefixes (`=`, `+`, `-`, `@`) during export[cite: 4, 13].
* **Isolated Fetcher Security**: Outbound companion (`PointFetcher.exe`) enforces HTTPS, applies a 2 GiB ceiling, validates file magic numbers (ZIP/OLE), and stores authentication secrets in Windows Credential Manager[cite: 4, 8, 17].

---

## 📁 Repository Structure

```text
├── build.bat                   # Release build script (MSVC C++20, /O2, /MT, /guard:cf)
├── build_debug.bat             # Debug build script (/Od, /Zi, /RTC1)
├── build_installer.bat         # Inno Setup packaging script with code-signing hooks
├── configure_compliance.bat    # Windows local security group setup script
├── test_core.bat               # Schema mapping and core regression test runner
├── point-security.conf         # Fail-closed security policy configuration
├── COMPLIANCE.md               # Compliance control matrix & deployment obligations
├── LICENSE.txt                 # Proprietary evaluation license
├── src/
│   ├── point.ico               # Application vector branding icon
│   ├── point.rc                # Win32 resource definitions & VERSIONINFO metadata
│   ├── point_compliance.h/.cpp # Group auth, DPAPI encryption, DACL hardening, retention
│   ├── point_core.h/.cpp       # Parsing, query engine, hash indexers, SHA-256 audit logger
│   ├── point_excel_import.h/.cpp # Excel COM automation bridge and worksheet cache
│   ├── point_fetcher.cpp       # Standalone HTTPS downloader & Credential Manager client
│   ├── point_win32.cpp         # Native Win32 UI, virtualized grid, GDI charts, event loop
│   └── resource.h              # Resource header definitions
└── tests/
    └── schema_mapping_test.cpp # Regression tests for core schema indexing and joins
# Point Windows Access Configuration

Point uses Windows security groups to control access.

| Group | Permissions |
|---|---|
| `Point Users` | Open, import, search, view |
| `Point Exporters` | Export results |
| `Point Administrators` | Full access + export |

## Setup

**1. Get your account name**
```powershell
whoami
```

**2. Open PowerShell as Administrator** → Start menu → search PowerShell → Run as administrator

**3. Create the groups**
```powershell
New-LocalGroup -Name "Point Users" -Description "Users authorized to run Point" -ErrorAction SilentlyContinue
New-LocalGroup -Name "Point Exporters" -Description "Users authorized to export Point results" -ErrorAction SilentlyContinue
New-LocalGroup -Name "Point Administrators" -Description "Administrators authorized for full Point access" -ErrorAction SilentlyContinue
```

**4. Grant access** (replace `DESKTOP-1234\john` with your `whoami` output)

Standard access only:
```powershell
$PointUser = "DESKTOP-1234\john"
Add-LocalGroupMember -Group "Point Users" -Member $PointUser
```

Access + export:
```powershell
$PointUser = "DESKTOP-1234\john"
Add-LocalGroupMember -Group "Point Users" -Member $PointUser
Add-LocalGroupMember -Group "Point Exporters" -Member $PointUser
```

Full admin:
```powershell
$PointUser = "DESKTOP-1234\john"
Add-LocalGroupMember -Group "Point Administrators" -Member $PointUser
```

**5. Verify**
```powershell
Get-LocalGroupMember -Group "Point Users"
Get-LocalGroupMember -Group "Point Exporters"
Get-LocalGroupMember -Group "Point Administrators"
```

**6. Apply** — Close Point → Sign out → Sign back in → Reopen Point

## Remove access
```powershell
$PointUser = "DESKTOP-1234\john"
Remove-LocalGroupMember -Group "Point Users" -Member $PointUser
Remove-LocalGroupMember -Group "Point Exporters" -Member $PointUser
Remove-LocalGroupMember -Group "Point Administrators" -Member $PointUser
```
Sign out/in afterward.

## Required config (`point-security.conf`, next to `Point.exe`)
```ini
enforce_windows_groups=true
allowed_windows_groups=Point Users;Point Administrators
export_windows_groups=Point Exporters;Point Administrators
export_retention_days=30
workspace_retention_days=30
log_retention_days=365
```

## Troubleshooting
- **Access Denied** → reopen PowerShell as Administrator
- **User still unauthorized** → verify group membership, then sign out/in
- **Group already exists** → safe to ignore, `-ErrorAction SilentlyContinue` handles it
- **Company-managed devices** → don't change groups yourself; contact IT/security
```
##Copyright
```
Copyright (c) 2026 Sunil Kumar Peela.
All rights reserved.

This prototype source and associated documentation files (the "Software")
are provided strictly for the owner's private evaluation, research, and
development.

No permission is granted to any person or organization to copy, modify,
merge, publish, distribute, sublicense, sell, or deploy the Software in
production environments without the explicit, prior written authorization
of the copyright holder.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE, TITLE, AND NON-INFRINGEMENT. IN NO EVENT
SHALL THE AUTHOR OR COPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES, OR
OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT, OR OTHERWISE,
ARISING FROM, OUT OF, OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.
```[cite: 6]
