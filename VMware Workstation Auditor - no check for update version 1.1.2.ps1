<#
    VMware Workstation Auditor
    Version: 1.1.2 without update check

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

function Write-ErrorLog {
    param([string]$message)
    Add-Content -Path $logFile -Value "$(Get-Date -Format u): $message"
}

function Read-VMXFlat {
    param (
        [string]$filePath,
        [string]$vmName
    )
    $result = @()
    $vmxData = @{}
    try {
        $lines = Get-Content -Path $filePath -ErrorAction Stop
        foreach ($line in $lines) {
            if ($line -match '^[^#]*?([^=\s]+?)\s*=\s*"?(.+?)"?$') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $vmxData[$key] = $value
                $result += [PSCustomObject]@{
                    VMName      = $vmName
                    Key         = $key
                    Value       = $value
                    VMVersion   = $vmxData['virtualHW.version']
                    SnapshotUID = ""
                }
            }
        }
    } catch {
        Write-ErrorLog -message $_.Exception.Message
        $result += [PSCustomObject]@{
            VMName      = $vmName
            Key         = "Error"
            Value       = $_.Exception.Message
            VMVersion   = ""
            SnapshotUID = ""
        }
    }
    return $result
}

function Read-SnapshotMeta {
    param (
        [string]$vmDir,
        [string]$vmName
    )
    $vmsd = Join-Path $vmDir "$vmName.vmsd"
    $results = @()

    if (Test-Path $vmsd) {
        try {
            $content = Get-Content $vmsd | Where-Object { $_ -match "snapshot\\." }
            foreach ($line in $content) {
                $uid = ""
                $desc = $line.Trim()
                if ($line -match 'snapshot\\.([^.]+)\\.displayName\s*=\s*"(.+?)"') {
                    $uid = $matches[1]
                    $desc = $matches[2]
                } elseif ($line -match 'snapshot\\.([^.]+)') {
                    $uid = $matches[1]
                }
                $results += [PSCustomObject]@{
                    VMName      = $vmName
                    Key         = "SnapshotMeta"
                    Value       = $desc
                    VMVersion   = ""
                    SnapshotUID = $uid
                }
            }
        } catch {
            Write-ErrorLog -message $_.Exception.Message
            $results += [PSCustomObject]@{
                VMName      = $vmName
                Key         = "SnapshotMetaError"
                Value       = $_.Exception.Message
                VMVersion   = ""
                SnapshotUID = ""
            }
        }
    }

    return $results
}

function Invoke-Audit {
    $script:vmAuditData = @()

    $vmwareReg = "HKLM:\\SOFTWARE\\WOW6432Node\\VMware, Inc.\\VMware Workstation"
    if (Test-Path $vmwareReg) {
        $script:vmwareVersion = (Get-ItemProperty -Path $vmwareReg).ProductVersion
    } else {
        $script:vmwareVersion = "Not Detected"
    }

    $vmxFiles = Get-ChildItem -Path $script:vmDirectory -Recurse -Filter *.vmx -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($vmx in $vmxFiles) {
        $vmName = [System.IO.Path]::GetFileNameWithoutExtension($vmx.FullName)
        $vmDir = Split-Path $vmx.FullName
        $script:vmAuditData += Read-VMXFlat -filePath $vmx.FullName -vmName $vmName
        $script:vmAuditData += Read-SnapshotMeta -vmDir $vmDir -vmName $vmName
    }
}

function Save-ToCSV {
    $dlg = New-Object Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV files (*.csv)|*.csv"
    $dlg.Title = "Save Audit Report"
    $dlg.FileName = "VMwareAuditReport.csv"
    if ($dlg.ShowDialog() -eq "OK") {
        try {
            $script:vmAuditData | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Report saved to:`n$($dlg.FileName)", "Export Successful")
        } catch {
            Write-ErrorLog -message $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show("Failed to save report.`nError: $_", "Error")
        }
    }
}

function Show-Help {
    [System.Windows.Forms.MessageBox]::Show("This tool audits VMware Workstation VMs, extracting VMX data, hardware versions, and snapshot metadata. Results can be exported to CSV.", "Help")
}

function Show-GUI {
    $form = New-Object Windows.Forms.Form
    $form.Text = "VMware Workstation Auditor (Table View)"
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

        foreach ($item in $script:vmAuditData) {
            $vmName      = if ($item.VMName) { $item.VMName } else { "" }
            $key         = if ($item.Key) { $item.Key } else { "" }
            $value       = if ($item.Value) { $item.Value } else { "" }
            $vmVersion   = if ($item.VMVersion) { $item.VMVersion } else { "" }
            $snapshotUID = if ($item.SnapshotUID) { $item.SnapshotUID } else { "" }

            $row = New-Object Windows.Forms.ListViewItem($vmName)
            [void]$row.SubItems.Add($key)
            [void]$row.SubItems.Add($value)
            [void]$row.SubItems.Add($vmVersion)
            [void]$row.SubItems.Add($snapshotUID)
            $listView.Items.Add($row)
        }

        $listView.AutoResizeColumns("HeaderSize")
        $lblVersion.Text = "Detected version of VMware Workstation is: $($script:vmwareVersion)"

        $uniqueVMs = ($script:vmAuditData | Select-Object -ExpandProperty VMName -Unique).Count
        $lblCount.Text = "VMs Detected: $uniqueVMs"
    })

    [void]$form.ShowDialog()
}

Show-GUI
