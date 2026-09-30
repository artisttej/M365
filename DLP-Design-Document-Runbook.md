# Purview DLP Design Document Export – Runbook

| | |
|---|---|
| **Script** | `Export-DlpDesignDocument.ps1` (v1.6.0) |
| **Purpose** | Export Microsoft Purview DLP policies and rules from a tenant and generate one Excel design document per policy |
| **Access required** | Read-only: `View-Only DLP Compliance Management` (or Compliance Administrator / Compliance Data Administrator) |
| **Changes made to tenant** | **None** – the script only uses `Get-*` cmdlets |

---

## 1. Overview

The process runs in **two stages** so that client machines only ever need the Exchange Online module:

```
 CLIENT MACHINE (locked down)                     OUR MACHINE (build)
 ─────────────────────────────                    ──────────────────────────
 ExchangeOnlineManagement only                    ImportExcel only
 Connect-IPPSSession                              No tenant connection
        │                                                  ▲
        ▼                                                  │
 Export-DlpDesignDocument.ps1 -ExportOnly  ──copy folder──►  Export-DlpDesignDocument.ps1 -FromExport
        │                                                  │
        ▼                                                  ▼
 DLP-Export-<timestamp>\Raw\*.xml                  <OutputFolder>\Policies\<nn> - <policy>.xlsx
 (snapshot of policies, rules, SITs)               (one design document per policy)
```

| Stage | Where | Modules needed | Connects to Purview? |
|---|---|---|---|
| 1. Export | Client machine | `ExchangeOnlineManagement` (already approved) | Yes – read only |
| 2. Build | Our machine | `ImportExcel` (manually copied) | **No** |

---

## 2. Prerequisites

### Client machine (export)
- Windows PowerShell 5.1 or PowerShell 7.x
- `ExchangeOnlineManagement` module (standard client-approved module)
- Account with a DLP read role (see table above)
- A copy of `Export-DlpDesignDocument.ps1`

### Our machine (build)
- Windows PowerShell 5.1 or PowerShell 7.x
- `ImportExcel` module – installed manually (Section 3). **Microsoft Excel does not need to be installed.**
- A copy of `Export-DlpDesignDocument.ps1`

### Script execution policy
If the script is blocked when run, unblock it or allow it for the current window only:

```powershell
Unblock-File .\Export-DlpDesignDocument.ps1
# or, for this PowerShell window only:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

---

## 3. Download and save ImportExcel (build machine only)

> ImportExcel is **only** needed on the machine that builds the design documents. Do **not** copy it to client machines.

### Option A – `Save-Module` (machine with internet access)

```powershell
New-Item -ItemType Directory C:\temp -Force | Out-Null
Save-Module -Name ImportExcel -Path C:\temp
```

This creates `C:\temp\ImportExcel\<version>\...`.

### Option B – manual download from the PowerShell Gallery

1. Browse to **https://www.powershellgallery.com/packages/ImportExcel** → **Manual Download** tab → **Download the raw nupkg file**.
2. Rename `importexcel.<version>.nupkg` to `.zip` and extract it.
3. Delete the packaging files: `_rels`, `package`, `[Content_Types].xml`, `ImportExcel.nuspec`.
4. Rename the extracted folder to `ImportExcel` and place it at `C:\temp\ImportExcel` (so `ImportExcel.psd1` is directly inside it).

### Unblock and import

Run in the PowerShell window you will build from:

```powershell
Get-ChildItem C:\temp\ImportExcel -Recurse | Unblock-File

# Option B layout:
Import-Module C:\temp\ImportExcel\ImportExcel.psd1
# Option A layout (versioned subfolder):
Import-Module (Get-ChildItem C:\temp\ImportExcel -Recurse -Filter ImportExcel.psd1 | Select-Object -First 1).FullName

Get-Module ImportExcel    # should list the module
```

> The script uses a module that is **already imported** in the session and will not search for or install it. Alternatively pass `-ModulePath C:\temp` to let the script find it.

---

## 4. Stage 1 – Export on the client machine

### 4.1 Connect to Security & Compliance PowerShell

```powershell
Connect-IPPSSession -UserPrincipalName admin@client.com

# Verify the session - this must return a command:
Get-Command Get-DlpCompliancePolicy
```

> **Run the script in the same PowerShell window** you connected in. A new window, "Run with PowerShell", or switching between PowerShell 5.1 and 7 starts a fresh, unconnected session.

When an IPPS session is already connected the script detects it automatically: **it does not import, search for, or install any module and does not sign in again.**

### 4.2 Choose what to export

All commands below use `-SkipConnect -ExportOnly -SkipModuleInstall`:
- `-SkipConnect` – use the existing IPPS session
- `-ExportOnly` – save the snapshot and stop (no Excel, ImportExcel not needed)
- `-SkipModuleInstall` – never attempt to download anything

**Single policy (exact name or GUID)**
```powershell
.\Export-DlpDesignDocument.ps1 -SkipConnect -ExportOnly -SkipModuleInstall `
    -PolicyName 'SEG | D-USB | Block removable media'
```

**Multiple policies (exact names)**
```powershell
.\Export-DlpDesignDocument.ps1 -SkipConnect -ExportOnly -SkipModuleInstall `
    -PolicyName 'SEG | D-USB | Block removable media','SEG | D-PRT | Printing controls'
```

**Policies whose names contain text** (case-insensitive, literal match – `|` and spaces are matched exactly)
```powershell
# Any policy containing "SEG | D-USB |" OR "| D-PRT |"
.\Export-DlpDesignDocument.ps1 -SkipConnect -ExportOnly -SkipModuleInstall `
    -PolicyNameContains 'SEG | D-USB |','| D-PRT |'
```

**All policies**
```powershell
.\Export-DlpDesignDocument.ps1 -SkipConnect -ExportOnly -SkipModuleInstall
```

Optionally add `-OutputFolder 'C:\Exports\<Client>'` to control where the snapshot is written.

### 4.3 Load on Purview

| Export type | Calls made to Purview |
|---|---|
| `-PolicyName` (exact names) | 1 per named policy + 1 rules call per policy + 1 per distinct SIT referenced |
| `-PolicyNameContains` | 1 policy list call (filtered locally) + 1 rules call per **matched** policy + 1 per distinct SIT referenced |
| All policies | 3 bulk calls: all policies, all rules, SIT catalogue |

All calls are read-only. There is no server-side name filter for DLP policies, so `-PolicyNameContains` reads the policy list once and filters it locally.

**Reference timing (full export):** 397 policies, 2,335 rules and 778 SITs were retrieved in a few minutes. After retrieval the script processes the data locally for a few more minutes before writing the snapshot – **do not cancel during this phase**, as the snapshot files are written at the end.

> Avoid `-IncludeDistributionDetail` unless required – it adds a slow extra call per policy.

### 4.4 Output

On completion the script prints the snapshot folder path, e.g.:

```
Snapshot folder : C:\...\DLP-Export-20260930-150300\Raw
Policies: 397  Rules: 2335  SITs saved: 778
```

| File | Contents |
|---|---|
| `Raw\DlpPolicies.xml` | Policies (full fidelity – used for the build) |
| `Raw\DlpRules.xml` | Rules (full fidelity – used for the build) |
| `Raw\DlpSits.xml` | Sensitive information types referenced (all of them for a full export) |
| `Raw\ExportInfo.xml` | Tenant, exported by, export date |
| `Raw\DlpPolicies.json`, `Raw\DlpRules.json` | Human-readable copies (reference only) |

---

## 5. Transfer the snapshot

Copy the **whole** `DLP-Export-<timestamp>` folder from the client machine to the build machine using the approved transfer method.

> ⚠️ The snapshot contains the client's full DLP configuration (policy scoping, users/groups, SIT definitions). Handle it as **client confidential** and store it only in the approved project location.

---

## 6. Stage 2 – Build the design documents (our machine)

Import ImportExcel first (Section 3), then run from the folder containing the script.

**All policies in the snapshot**
```powershell
.\Export-DlpDesignDocument.ps1 -FromExport 'C:\Exports\DLP-Export-20260930-150300' `
    -OutputFolder 'C:\DesignDocs\<Client>' -CRRef 'CHG0012345' -SkipModuleInstall
```

**A subset of the snapshot** (same filters as export)
```powershell
.\Export-DlpDesignDocument.ps1 -FromExport 'C:\Exports\DLP-Export-20260930-150300' `
    -OutputFolder 'C:\DesignDocs\<Client>' -PolicyNameContains '| D-PRT |' -SkipModuleInstall

.\Export-DlpDesignDocument.ps1 -FromExport 'C:\Exports\DLP-Export-20260930-150300' `
    -PolicyName 'SEG | D-USB | Block removable media' -SkipModuleInstall
```

- `-FromExport` accepts the export folder **or** its `Raw` subfolder.
- `-CRRef` is optional; without it the CR Ref field is left blank.
- No connection to Purview is made. The build can be re-run any number of times, with different filters, from the same snapshot.
- **Timing:** roughly 1–2 seconds per policy; allow 10–20 minutes for ~400 policies. Test with a small subset first.

### Output

```
C:\DesignDocs\<Client>\
├── Policies\
│   ├── 01 - SEG - D-USB - Block removable media.xlsx
│   ├── 02 - SEG - D-PRT - Printing controls.xlsx
│   └── ...
└── Raw\            (copy of the snapshot used)
```

File names are `<priority position> - <policy name>.xlsx`. Characters not allowed in Windows file names are replaced (`|` becomes `-`).

### Design document contents (per policy)

| Sheet | Contents |
|---|---|
| **Policy_Config** | CR Ref, priority position, name & description, created / last modified (by and date, UTC), mode, admin units, locations per workload (included / excluded), rule list |
| **Rule_Config** | Per rule: details incl. created / modified; numbered condition blocks (1.1, 1.2 …) with *Is Group*, *Condition* (Purview predicate name), *Condition Properties & Values*, AND/OR operators between blocks, exceptions; actions; user notifications; user overrides; incident reports / alerts; additional options |
| **SIT_Config** | Per rule: one block per SIT / trainable classifier / label – confidence level, instance count, group logic, built-in vs custom |
| **Additional Information** | Export details, distribution status, SIT summary, every raw policy and rule property |
| **Lists** | Drop-down values used in the workbook and the Purview predicate → PowerShell parameter mapping |

---

## 7. Parameter reference

| Parameter | Stage | Description |
|---|---|---|
| `-SkipConnect` | Export | Use the existing `Connect-IPPSSession` session. (An existing session is also auto-detected.) |
| `-ExportOnly` | Export | Save the snapshot and stop. Does not load ImportExcel. |
| `-FromExport <path>` | Build | Build design documents from a snapshot folder. Does not connect to Purview or load the Exchange module. |
| `-PolicyName <names>` | Both | One or more exact policy names or GUIDs. |
| `-PolicyNameContains <text>` | Both | One or more text fragments; policies whose name contains **any** of them are selected (case-insensitive, literal). Can be combined with `-PolicyName`. |
| `-OutputFolder <path>` | Both | Output location. Default: `.\DLP-Export-<yyyyMMdd-HHmmss>` in the current directory. |
| `-CRRef <text>` | Build | Change reference written into every design document. |
| `-SkipModuleInstall` | Both | Never attempt to download or install modules; fail with a clear message instead. |
| `-ModulePath <path>` | Both | Folder holding manually copied modules (e.g. `C:\temp`). Added to the module path for the current session only. |
| `-IncludeDistributionDetail` | Export | Also capture per-policy distribution results (slow – one extra call per policy). |
| `-IncludeCombinedWorkbook` | Build | Also produce a single tabular workbook covering all selected policies. |
| `-UserPrincipalName` / `-AppId`, `-CertificateThumbprint`, `-Organization` | Export | Let the script connect itself (interactive or certificate auth) instead of using an existing session. |

---

## 8. Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `...no Security & Compliance session was found in this PowerShell window` | Script run in a different session from the one that ran `Connect-IPPSSession` | Connect and run the script in the **same** window. Check with `Get-Command Get-DlpCompliancePolicy`. |
| `Module 'ExchangeOnlineManagement' ... was not loaded` | Old script version, or not connected | Confirm the first output line shows `v1.6.0`; connect IPPS first. |
| `Module 'ImportExcel' is not loaded or installed` | ImportExcel not imported in the build window | `Import-Module C:\temp\ImportExcel\ImportExcel.psd1`, or add `-ModulePath C:\temp`. |
| Import of ImportExcel fails / is blocked | Files marked as downloaded from the internet, or application control (AppLocker/WDAC) | `Get-ChildItem C:\temp\ImportExcel -Recurse \| Unblock-File`. If still blocked, raise with the machine owner – ImportExcel includes the EPPlus .NET DLL. |
| `Snapshot file 'DlpPolicies.xml' not found` | `-FromExport` points to the wrong folder | Point it at the `DLP-Export-<timestamp>` folder or its `Raw` subfolder. |
| `Policy not found or not accessible` | Exact name typo or insufficient role | Check the name in the Purview portal, or use `-PolicyNameContains`. |
| Script file blocked by execution policy | Downloaded script | `Unblock-File .\Export-DlpDesignDocument.ps1` or `Set-ExecutionPolicy -Scope Process Bypass`. |
| Cells show "Not recorded" for created / modified by | Value is empty in Purview (older / template / service-created policies) | Expected – no action. |

---

## 9. Jira ticket

**Summary:** Produce Purview DLP design documents for `<Client>` – `<policy scope, e.g. SEG | D-USB | and | D-PRT | policies>`

**Description:**
Export the in-scope Microsoft Purview DLP policies and rules from the `<Client>` tenant using `Export-DlpDesignDocument.ps1` (v1.6.0) and generate one Excel design document per policy, following the attached runbook. Export is read-only; no changes are made to the tenant. ImportExcel must not be installed on client machines – the build is performed offline on our machine from the exported snapshot.

**Scope:**
- Policies: `<-PolicyName list / -PolicyNameContains fragments / All>`
- CR reference: `<CHGxxxxxxx>`

**Sub-tasks:**
- [ ] Confirm read-only DLP role for the export account
- [ ] Prepare build machine: download and save ImportExcel (Runbook §3)
- [ ] Connect to IPPS on client machine and verify `Get-Command Get-DlpCompliancePolicy` (§4.1)
- [ ] Run export with agreed scope (§4.2); record policy / rule / SIT counts from the output
- [ ] Transfer snapshot folder via approved method; store as client confidential (§5)
- [ ] Build design documents from the snapshot (§6)
- [ ] Spot-check at least one design document against the Purview portal (scope, conditions, actions)
- [ ] Upload design documents and snapshot to the project location; attach / link to this ticket

**Acceptance criteria:**
- [ ] One design document exists for every in-scope policy
- [ ] Policy / rule counts in the export output match the agreed scope
- [ ] Spot-checked document(s) match the Purview portal configuration
- [ ] No modules other than ExchangeOnlineManagement were installed or imported on the client machine
- [ ] Snapshot and documents stored in the approved location only
