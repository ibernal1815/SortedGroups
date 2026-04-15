<#
.SYNOPSIS
    Builds a hash table from Server2016_Upgrade_ServerList.xlsx grouped by Group name.

.DESCRIPTION
    Reads the server list Excel file and organizes servers into a hash table keyed
    by Group. Each entry contains a deduplicated contact list and server details,
    making it easy for the lead cloud administrator to send bulk emails per group.

.PARAMETER XlsxPath
    Path to the Excel file. Defaults to the same directory as the script.

.PARAMETER Group
    Show a single group only (e.g. -Group "IT-DA").

.PARAMETER Email
    Print ready-to-paste email targets (To: lines) per group.

.PARAMETER Json
    Dump the full hash table as JSON to stdout.

.EXAMPLE
    .\group_server_map.ps1
    .\group_server_map.ps1 -Email
    .\group_server_map.ps1 -Group "IT-DA"
    .\group_server_map.ps1 -Json
#>

[CmdletBinding()]
param (
    [string]$XlsxPath = (Join-Path $PSScriptRoot "Server2016_Upgrade_ServerList.xlsx"),
    [string]$Group    = "",
    [switch]$Email,
    [switch]$Json
)

# Excel COM can reject calls if it's still spinning up, so retry on 0x80010001
function Invoke-COM {
    param (
        [scriptblock]$Action,
        [int]$Retries = 5,
        [int]$DelayMs = 300
    )
    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            return & $Action
        } catch {
            if ($_.Exception.HResult -eq [int]0x80010001 -and $i -lt ($Retries - 1)) {
                Start-Sleep -Milliseconds $DelayMs
            } else {
                throw
            }
        }
    }
}

function Import-ExcelData {
    param ([string]$Path)

    if (-not (Test-Path $Path)) {
        Write-Error "File not found: $Path"
        exit 1
    }

    $excel    = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false
        $excel.Interactive   = $false

        # COM needs a moment before it'll take calls without throwing a fit
        Start-Sleep -Milliseconds 500

        $workbook = Invoke-COM { $excel.Workbooks.Open((Resolve-Path $Path).Path) }
        Start-Sleep -Milliseconds 300

        $sheet   = Invoke-COM { $workbook.Sheets.Item(1) }
        $lastRow = Invoke-COM { $sheet.UsedRange.Rows.Count }
        $lastCol = Invoke-COM { $sheet.UsedRange.Columns.Count }

        # Grab headers first so we can reference columns by name instead of index
        $headers = @{}
        for ($col = 1; $col -le $lastCol; $col++) {
            $val = Invoke-COM { $sheet.Cells.Item(1, $col).Text }
            $headers[$val.Trim()] = $col
        }

        $rows = @()
        for ($row = 2; $row -le $lastRow; $row++) {
            $entry = @{}
            foreach ($h in $headers.Keys) {
                $val       = Invoke-COM { $sheet.Cells.Item($row, $headers[$h]).Text }
                $entry[$h] = $val.Trim()
            }
            $rows += $entry
        }

        return $rows
    }
    finally {
        # Clean up COM refs in reverse order or Excel won't actually close
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
            Start-Sleep -Milliseconds 500
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Build-GroupTable {
    param ([array]$Rows)

    $table = @{}

    foreach ($row in $Rows) {
        $groupName = $row["Group"]
        if (-not $groupName) { continue }

        if (-not $table.ContainsKey($groupName)) {
            $table[$groupName] = @{
                Contacts   = [System.Collections.Generic.List[string]]::new()
                ContactSet = [System.Collections.Generic.HashSet[string]]::new()
                Servers    = [System.Collections.Generic.List[hashtable]]::new()
            }
        }

        $entry = $table[$groupName]

        # Some rows share the same owner/manager, HashSet keeps dupes out
        foreach ($col in @("AppOwner", "AppOwner_2", "Manager")) {
            $email = $row[$col]
            if ($email -and $entry.ContactSet.Add($email)) {
                $entry.Contacts.Add($email)
            }
        }

        $entry.Servers.Add(@{
            Server              = $row["Server"]
            "Provisioned Space" = $row["Provisioned Space"]
            "Used Space"        = $row["Used Space"]
            "Host Mem"          = $row["Host Mem"]
            Site                = $row["Site"]
            Status              = $row["Status"]
        })
    }

    # ContactSet was just a helper, don't need it in the final output
    foreach ($key in $table.Keys) {
        $table[$key].Remove("ContactSet")
    }

    return $table
}

function Show-Summary {
    param ([hashtable]$Table, [string]$FilterGroup = "")

    $groups = if ($FilterGroup) { @($FilterGroup) } else { $Table.Keys | Sort-Object }

    foreach ($g in $groups) {
        if (-not $Table.ContainsKey($g)) {
            Write-Warning "Group '$g' not found."
            continue
        }

        $entry = $Table[$g]
        $bar   = "=" * 60

        Write-Host ""
        Write-Host $bar -ForegroundColor Cyan
        Write-Host ("  GROUP : {0}  ({1} server(s))" -f $g, $entry.Servers.Count) -ForegroundColor Yellow
        Write-Host $bar -ForegroundColor Cyan

        Write-Host ("  Contacts ({0}):" -f $entry.Contacts.Count) -ForegroundColor Green
        foreach ($c in $entry.Contacts) {
            Write-Host "    - $c"
        }

        Write-Host ""
        Write-Host "  Servers:" -ForegroundColor Green
        foreach ($srv in $entry.Servers) {
            Write-Host ("    [{0}] {1,-26} Provisioned: {2,-12} Used: {3}" -f `
                $srv.Site, $srv.Server, $srv."Provisioned Space", $srv."Used Space")
        }
    }
}

function Show-EmailTargets {
    param ([hashtable]$Table)

    Write-Host ""
    Write-Host "── Email Targets by Group ──────────────────────────────────" -ForegroundColor Cyan

    foreach ($g in ($Table.Keys | Sort-Object)) {
        $entry   = $Table[$g]
        $toLine  = $entry.Contacts -join "; "
        $servers = ($entry.Servers | ForEach-Object { $_.Server }) -join ", "

        Write-Host ""
        Write-Host "Group   : $g" -ForegroundColor Yellow
        Write-Host "To      : $toLine" -ForegroundColor Green
        Write-Host ("Servers ({0}): {1}" -f $entry.Servers.Count, $servers)
    }
}

function Show-Json {
    param ([hashtable]$Table)

    # convert to a serializable structure
    $output = @{}
    foreach ($g in $Table.Keys) {
        $output[$g] = @{
            contacts     = @($Table[$g].Contacts)
            server_count = $Table[$g].Servers.Count
            servers      = @($Table[$g].Servers)
        }
    }

    $output | ConvertTo-Json -Depth 5
}

Write-Host "Loading $XlsxPath ..." -ForegroundColor DarkGray

$rows  = Import-ExcelData -Path $XlsxPath
$table = Build-GroupTable -Rows $rows

if ($Json) {
    Show-Json  -Table $table
} elseif ($Email) {
    Show-EmailTargets -Table $table
} else {
    Show-Summary -Table $table -FilterGroup $Group
}
