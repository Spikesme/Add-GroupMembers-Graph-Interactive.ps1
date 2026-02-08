<# 
Add-GroupMembers-Graph-Interactive.ps1

Interaktives Skript: Mehrere User (Copy/Paste aus Ticket) zu Entra ID (Azure AD) Sicherheitsgruppen hinzufügen – via Microsoft Graph.

Fixes / Features:
- Stabiler Graph-Login: Interactive, bei MSAL-Konflikten automatisch Device-Code Fallback
- Importiert nur benötigte Graph-Submodule (verhindert FunctionOverflow 4096)
- Gruppensuche + Auswahlmenü (Anzeige-Fix)
- Userliste: mehrere Zeilen ODER kommagetrennt/semicolon-getrennt in EINER Zeile
- Parsing: "Vorname Nachname" oder "Nachname, Vorname"
- Auto-Mode: genau 1 User-Treffer -> automatisch adden, bei mehreren Treffern Auswahlmenü
#>

#region Helpers
function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Modul '$Name' fehlt – installiere..." -ForegroundColor Yellow
        Install-Module $Name -Scope CurrentUser -Force -ErrorAction Stop
    }
}

function Ensure-GraphConnection {
    param([string[]]$Scopes = @("User.Read.All","Group.Read.All","GroupMember.ReadWrite.All"))

    # Wenn schon verbunden und Scopes passen -> fertig
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account) {
            $missing = @($Scopes | Where-Object { $_ -notin $ctx.Scopes })
            if ($missing.Count -eq 0) {
                Write-Host "Bereits mit Microsoft Graph angemeldet als $($ctx.Account)." -ForegroundColor Green
                return
            }
            Write-Host "Scopes fehlen: $($missing -join ', ') -> Re-Login" -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
        }
    } catch {}

    Write-Host "Verbinde zu Microsoft Graph..." -ForegroundColor Cyan

    # Versuch 1: Interactive
    try {
        Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null
        $ctx = Get-MgContext
        Write-Host "Verbunden (Interactive) als $($ctx.Account) (Tenant: $($ctx.TenantId))" -ForegroundColor Green
        return
    } catch {
        $msg = $_.Exception.Message
        Write-Host "Interactive Login fehlgeschlagen: $msg" -ForegroundColor Yellow

        # Fallback: Device Code (stabil bei MSAL/Assembly-Konflikten)
        Write-Host "Wechsle auf Device Code Login..." -ForegroundColor Cyan
        Connect-MgGraph -UseDeviceAuthentication -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null
        $ctx = Get-MgContext
        Write-Host "Verbunden (Device Code) als $($ctx.Account) (Tenant: $($ctx.TenantId))" -ForegroundColor Green
    }
}

function Select-FromList {
    param(
        [Parameter(Mandatory)] $Items,
        [Parameter(Mandatory)] [string] $Title,
        [Parameter(Mandatory)] [scriptblock] $Label
    )

    if (-not $Items -or $Items.Count -eq 0) { return $null }
    if ($Items.Count -eq 1) { return $Items[0] }

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan

    for ($i = 0; $i -lt $Items.Count; $i++) {
        $text = & $Label $Items[$i]
        if ($null -eq $text) { $text = "" }
        Write-Host ("[{0}] {1}" -f ($i + 1), $text)
    }

    while ($true) {
        $raw = Read-Host "Nummer wählen (1-$($Items.Count))"

        if ($raw -notmatch '^\s*\d+\s*$') {
            Write-Host "Bitte nur eine Zahl eingeben (z.B. 7)." -ForegroundColor Yellow
            continue
        }

        $idx = [int]$raw
        if ($idx -lt 1 -or $idx -gt $Items.Count) {
            Write-Host "Bitte Zahl zwischen 1 und $($Items.Count) eingeben." -ForegroundColor Yellow
            continue
        }

        return $Items[$idx - 1]
    }
}

function Split-UserInput {
    param([Parameter(Mandatory)][string[]]$Lines)

    # Unterstützt: mehrere Zeilen UND/ODER "Name, Name, Name" in einer Zeile (auch ; als Trenner)
    $allText = ($Lines -join "`n")
    $tokens = $allText -split '[,\n;]'
    $tokens = $tokens | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    return $tokens
}

function Parse-Name {
    param([Parameter(Mandatory)][string]$Text)

    $t = $Text.Trim().TrimEnd(',')

    # "Nachname, Vorname"
    if ($t -match '^\s*([^,]+)\s*,\s*(.+?)\s*$') {
        $last  = $matches[1].Trim()
        $first = $matches[2].Trim()
        return [pscustomobject]@{ First=$first; Last=$last; Raw=$Text }
    }

    # "Vorname Nachname" (alles bis letztes Wort = Vorname, letztes Wort = Nachname)
    $parts = $t -split '\s+'
    if ($parts.Count -lt 2) { return $null }

    $last  = $parts[-1]
    $first = ($parts[0..($parts.Count-2)] -join ' ')
    return [pscustomobject]@{ First=$first; Last=$last; Raw=$Text }
}
#endregion Helpers

#region Main
try {
    # Minimal-Import gegen FunctionOverflow
    Ensure-Module -Name "Microsoft.Graph.Authentication"
    Ensure-Module -Name "Microsoft.Graph.Users"
    Ensure-Module -Name "Microsoft.Graph.Groups"

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop

    Ensure-GraphConnection

    $grpSearch = Read-Host "Gruppenname (Teilstring)"
    if ([string]::IsNullOrWhiteSpace($grpSearch)) { throw "Kein Gruppen-Suchstring angegeben." }

    $groups = Get-MgGroup -Search "displayName:$grpSearch" -ConsistencyLevel eventual -All -ErrorAction Stop |
        Where-Object { $_.SecurityEnabled -eq $true } |
        Sort-Object DisplayName

    if (-not $groups -or $groups.Count -eq 0) { throw "Keine Sicherheitsgruppe gefunden zu '$grpSearch'." }

    $group = Select-FromList -Items $groups -Title "Gefundene Gruppen:" -Label { param($g) $g.DisplayName }
    if (-not $group) { throw "Keine Gruppe ausgewählt." }

    Write-Host ""
    Write-Host "Zielgruppe: $($group.DisplayName)" -ForegroundColor Cyan
    Write-Host "Userliste einfügen (mehrere Zeilen ODER kommagetrennt/semicolon-getrennt). Ende: leere Zeile." -ForegroundColor Gray

    $lines = New-Object System.Collections.Generic.List[string]
    while ($true) {
        $l = Read-Host
        if ([string]::IsNullOrWhiteSpace($l)) { break }
        $lines.Add($l) | Out-Null
    }

    $tokens = Split-UserInput -Lines $lines
    if (-not $tokens -or $tokens.Count -eq 0) {
        Write-Host "Keine Namen erkannt. Ende." -ForegroundColor Yellow
        return
    }

    $countOK = 0; $countSkip = 0; $countFail = 0

    foreach ($token in $tokens) {
        $n = Parse-Name -Text $token
        if (-not $n) {
            Write-Host "Ungültig (konnte nicht parsen): $token" -ForegroundColor Yellow
            $countFail++
            continue
        }

        Write-Host ""
        Write-Host "Suche: $($n.First) $($n.Last)" -ForegroundColor Gray

        $users = Get-MgUser -Search "displayName:$($n.First) $($n.Last)" -ConsistencyLevel eventual -All -ErrorAction Stop |
            Select-Object Id,DisplayName,UserPrincipalName,GivenName,Surname,AccountEnabled

        if (-not $users -or $users.Count -eq 0) {
            Write-Host "FAIL: Kein Treffer für '$($n.First) $($n.Last)' (Input: $token)" -ForegroundColor Red
            $countFail++
            continue
        }

        # Auto-Mode: 1 Treffer -> direkt, sonst Auswahl
        $u = if ($users.Count -eq 1) { 
            $users[0] 
        } else {
            Select-FromList -Items $users -Title "Mehrere Treffer:" -Label { param($x) "$($x.DisplayName) | $($x.UserPrincipalName)" }
        }

        if (-not $u) {
            Write-Host "SKIP: Keine Auswahl getroffen für '$($n.First) $($n.Last)'." -ForegroundColor DarkYellow
            $countSkip++
            continue
        }

        try {
            New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($u.Id)"
            } -ErrorAction Stop | Out-Null

            Write-Host "OK: $($u.DisplayName) <$($u.UserPrincipalName)>" -ForegroundColor Green
            $countOK++
        } catch {
            $msg = $_.Exception.Message
            if ($msg -match "added object references already exist") {
                Write-Host "SKIP (schon drin): $($u.DisplayName) <$($u.UserPrincipalName)>" -ForegroundColor DarkYellow
                $countSkip++
            } else {
                Write-Host "FAIL: $($u.DisplayName) <$($u.UserPrincipalName)> :: $msg" -ForegroundColor Red
                $countFail++
            }
        }
    }

    Write-Host ""
    Write-Host "===== Zusammenfassung =====" -ForegroundColor Cyan
    Write-Host ("OK:    {0}" -f $countOK)
    Write-Host ("SKIP:  {0}" -f $countSkip)
    Write-Host ("FAIL:  {0}" -f $countFail)
    Write-Host "===========================" -ForegroundColor Cyan

} catch {
    Write-Host "Fataler Fehler: $($_.Exception.Message)" -ForegroundColor Red
    throw
}
#endregion Main
