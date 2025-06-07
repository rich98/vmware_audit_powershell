<#
    VMware Workstation Auditor
    Version: 1.1.6

    Copyright 2025 Richard Wadsworth

    Licensed under the Apache License, Version 2.0 (the "License");
    you may not use this file except in compliance with the License.
    You may obtain a copy of the License at

        http://www.apache.org/licenses/LICENSE-2.0

    Unless required by applicable law or agreed to in writing, software
    distributed under the License is distributed on an "AS IS" BASIS,
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    See the License for the specific language governing permissions and
    limitations under the License.

    CHANGE LOG:
    - v1.1.6 (2025-06-07)
      * Fixed remote version check using improved newline stripping.
      * Updated GUI title with version display.
      * Improved layout for version check dialog.
      * Added compatibility for EXE wrapping via BAT launcher.
      * General refinements to update logic and error handling.
      * Alternating row colors (pale blue and white) for readability.
      * Fixed null handling for ListView subitems to prevent crash.
      * Color alternation is now based on VM identity, not line count.
#>

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
} catch {
    Write-Error "Failed to load required assemblies: $_"
    exit
}

$script:vmDirectory = "$env:USERPROFILE\Documents\Virtual Machines"
$script:vmAuditData = @()
$script:vmwareVersion = ""
$logFile = "$env:TEMP\VMwareAudit_ErrorLog.txt"
$script:currentVersion = "1.1.6"

function Write-ErrorLog {
    param([string]$message)
    Add-Content -Path $logFile -Value "$(Get-Date -Format u): $message"
}

function Check-ForUpdate {
    try {
        $url = "https://raw.githubusercontent.com/rich98/vmware_audit_powershell/main/VERSION.TXT"
        $remoteVersion = (Invoke-RestMethod -Uri $url -UseBasicParsing) -replace '[\r\n]+', ''
        if ($remoteVersion -ne $script:currentVersion) {
            [System.Windows.Forms.MessageBox]::Show("A newer version ($remoteVersion) is available.", "Update Available")
        }
    } catch {
        Write-ErrorLog -message "Update check failed: $_"
    }
}

function Show-GUI {
    Check-ForUpdate
    $form = New-Object Windows.Forms.Form
    $form.Text = "VMware Workstation Auditor v$script:currentVersion (Table View)"
    $form.Size = New-Object Drawing.Size(1000, 660)
    $form.StartPosition = "CenterScreen"

    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = "VM Directory:"
    $lbl.Location = '10,15'
    $lbl.Size = '90,20'
    $form.Controls.Add($lbl)

    $txtPath = New-Object Windows.Forms.TextBox
    $txtPath.Text = $script:vmDirectory
    $txtPath.Location = '100,12'
    $txtPath.Size = '640,22'
    $form.Controls.Add($txtPath)

    $btnBrowse = New-Object Windows.Forms.Button
    $btnBrowse.Text = "Browse..."
    $btnBrowse.Location = '750,10'
    $btnBrowse.Size = '80,24'
    $btnBrowse.Add_Click({
        $dlg = New-Object Windows.Forms.FolderBrowserDialog
        if ($dlg.ShowDialog() -eq "OK") {
            $txtPath.Text = $dlg.SelectedPath
            $script:vmDirectory = $dlg.SelectedPath
        }
    })
    $form.Controls.Add($btnBrowse)

    $btnAudit = New-Object Windows.Forms.Button
    $btnAudit.Text = "Run Audit"
    $btnAudit.Location = '840,10'
    $btnAudit.Size = '90,24'
    $form.Controls.Add($btnAudit)

    $listView = New-Object Windows.Forms.ListView
    $listView.View = 'Details'
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.Location = '10,45'
    $listView.Size = '960,530'
    $listView.Columns.Add("VM Name", 180)
    $listView.Columns.Add("Key", 280)
    $listView.Columns.Add("Value", 280)
    $listView.Columns.Add("HW Version", 90)
    $listView.Columns.Add("Snapshot Info", 120)
    $form.Controls.Add($listView)

    $btnSave = New-Object Windows.Forms.Button
    $btnSave.Text = "Save to CSV"
    $btnSave.Location = '10,585'
    $btnSave.Size = '100,26'
    $btnSave.Add_Click({ Save-ToCSV })
    $form.Controls.Add($btnSave)

    $btnHelp = New-Object Windows.Forms.Button
    $btnHelp.Text = "Help"
    $btnHelp.Location = '120,585'
    $btnHelp.Size = '80,26'
    $btnHelp.Add_Click({ Show-Help })
    $form.Controls.Add($btnHelp)

    $lblCount = New-Object Windows.Forms.Label
    $lblCount.Text = "VMs Detected: 0"
    $lblCount.Location = New-Object Drawing.Point(220, 588)
    $lblCount.Size = New-Object Drawing.Size(200, 20)
    $form.Controls.Add($lblCount)

    $lblVersion = New-Object Windows.Forms.Label
    $lblVersion.Text = "Detected version of VMware Workstation is: (not yet scanned)"
    $lblVersion.Location = New-Object Drawing.Point(($form.ClientSize.Width - 460), 588)
    $lblVersion.Size = New-Object Drawing.Size(450, 20)
    $lblVersion.Anchor = "Right,Bottom"
    $lblVersion.TextAlign = 'MiddleRight'
    $form.Controls.Add($lblVersion)

    $btnAudit.Add_Click({
        $script:vmDirectory = $txtPath.Text
        if (-not (Test-Path $script:vmDirectory)) {
            [System.Windows.Forms.MessageBox]::Show("Directory not found.", "Error")
            return
        }

        $listView.Items.Clear()
        Invoke-Audit

        $colorMap = @{}
        $vmColorIndex = 0

        foreach ($item in $script:vmAuditData) {
            $vmName      = if ($item.VMName) { $item.VMName } else { "" }
            $key         = if ($item.Key) { $item.Key } else { "" }
            $value       = if ($item.Value) { $item.Value } else { "" }
            $vmVersion   = if ($item.VMVersion) { $item.VMVersion } else { "" }
            $snapshotUID = if ($item.SnapshotUID) { $item.SnapshotUID } else { "" }

            if (-not $colorMap.ContainsKey($vmName)) {
                $colorMap[$vmName] = $vmColorIndex % 2
                $vmColorIndex++
            }

            $row = New-Object Windows.Forms.ListViewItem($vmName)
            [void]$row.SubItems.Add($key)
            [void]$row.SubItems.Add($value)
            [void]$row.SubItems.Add($vmVersion)
            [void]$row.SubItems.Add($snapshotUID)

            if ($colorMap[$vmName] -eq 0) {
                $row.BackColor = [System.Drawing.Color]::White
            } else {
                $row.BackColor = [System.Drawing.Color]::FromArgb(240, 248, 255)
            }

            $listView.Items.Add($row)
        }

        $listView.AutoResizeColumns("HeaderSize")
        $lblVersion.Text = "Detected version of VMware Workstation is: $($script:vmwareVersion)"

        $uniqueVMs = ($script:vmAuditData | Select-Object -ExpandProperty VMName -Unique).Count
        $lblCount.Text = "VMs Detected: $uniqueVMs"

        Check-ForUpdate
    })

    [void]$form.ShowDialog()
}

Show-GUI
