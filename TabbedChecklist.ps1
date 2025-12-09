# Variables

## Gets working directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$config = Get-Content "$scriptDir\config\config.json" | ConvertFrom-Json
$softwareDir = "$scriptDir\software"

# Functions and Switch

function Invoke-ExeInstall {
    param($item, $listBox)

    $processArgs = @{
        FilePath = $item.path
    }

    if ($item.args) {
        $processArgs.ArgumentList = $item.args
    }

    if ($item.wait -eq $true) {
        $processArgs.Wait = $true
    }

    if ($item.runAs -eq $true) {
        $processArgs.Verb = "RunAs"
    }

    Start-Process @processArgs
}

function Invoke-MsiInstall {
    param($item, $listBox)

    $processArgs = @{
        FilePath = $item.path
    }

    if ($item.args) {
        $processArgs.ArgumentList = $item.args
    }

    if ($item.wait -eq $true) {
        $processArgs.Wait = $true
    }

    if ($item.runAs -eq $true) {
        $processArgs.Verb = "RunAs"
    }

    Start-Process @processArgs
}

function Invoke-PowerShellScript {
    param($item, $listBox)

    $processArgs = @{
        FilePath = "powershell"
        ArgumentList = "-command & '$($item.path)'"
    }

    if ($item.runAs -eq $true) {
        $processArgs.Verb = "RunAs"
    }

    Start-Process @processArgs
}

function Invoke-RegistryChange {
    param($item, $listBox)

    $command = "reg add '$($item.key)' /v $($item.valueName) /t $($item.valueType) /d $($item.valueData) /f"

    if ($item.runAs -eq $true) {
        Start-Process powershell -Verb RunAs -ArgumentList "-command $command" -Wait
    } else {
        Invoke-Expression $command
    }
}

function Invoke-FileCopy {
    param($item, $listBox)

    $copyArgs = @{
        Path = $item.source
        Destination = $item.destination
        Recurse = $true
    }

    if ($item.force -eq $true) {
        $copyArgs.Force = $true
    }

    Copy-Item @copyArgs
}

function Invoke-FileMove {
    param($item, $listBox)

    $moveArgs = @{
        Path = $item.source
        Destination = $item.destination
    }

    if ($item.force -eq $true) {
        $moveArgs.Force = $true
    }

    Move-Item @moveArgs
}

function Invoke-ShortcutInstall {
    param($item, $listBox)

    if ($item.icon) {
        Copy-Item -Path $item.icon.source -Destination $item.icon.destination -Force
    }

    if ($item.shortcut.tempDest) {
        Copy-Item -Path $item.shortcut.source -Destination $item.shortcut.tempDest -Force
        Move-Item -Path $item.shortcut.tempDest -Destination $item.shortcut.destination -Force
    } else {
        Copy-Item -Path $item.shortcut.source -Destination $item.shortcut.destination -Force
    }
}

function Invoke-Command {
    param($item, $listBox)

    if ($item.runAs -eq $true) {
        Start-Process powershell -Verb RunAs -ArgumentList "-command $($item.command)" -Wait
    } else {
        Invoke-Expression $item.command
    }
}

## allows for wait time between operations
function Invoke-Sleep {
    param($item, $listBox)

    Start-Sleep -Seconds $item.seconds
}

## allows for cleaner json 
function Invoke-MultiStep {
    param($item, $listBox)

    foreach ($step in $item.steps) {
        Execute-Item $step $listBox
    }
}

## switch function
function Execute-Item {
    param($item, $listBox)

    switch ($item.type) {
        "exe" { Invoke-ExeInstall $item $listBox }
        "msi" { Invoke-MsiInstall $item $listBox }
        "powershell" { Invoke-PowerShellScript $item $listBox }
        "registry" { Invoke-RegistryChange $item $listBox }
        "copy" { Invoke-FileCopy $item $listBox }
        "move" { Invoke-FileMove $item $listBox }
        "shortcut" { Invoke-ShortcutInstall $item $listBox }
        "command" { Invoke-Command $item $listBox }
        "sleep" { Invoke-Sleep $item $listBox }
        "multi" { Invoke-MultiStep $item $listBox }
        "none" { }  # Placeholder items
        default {
            $listBox.Items.Add("ERROR: Unknown type '$($item.type)' for item '$($item.name)'")
        }
    }
}

# creates the form

function GenerateForm {
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null


    ## fonts matter
    $fontFamily = $config.fonts.fontFamily
    $checkBoxFont = New-Object System.Drawing.Font($fontFamily, $config.fonts.checkBoxSize, [System.Drawing.FontStyle]::Regular)
    $formFont = New-Object System.Drawing.Font($fontFamily, $config.fonts.formSize, [System.Drawing.FontStyle]::Regular)
    $tabFont = New-Object System.Drawing.Font($fontFamily, $config.fonts.tabSize, [System.Drawing.FontStyle]::Bold)
    $buttonFont = New-Object System.Drawing.Font($fontFamily, $config.fonts.buttonSize, [System.Drawing.FontStyle]::Bold)
    $buttonFontRegular = New-Object System.Drawing.Font($fontFamily, $config.fonts.statusLabelSize, [System.Drawing.FontStyle]::Regular)
    $statusLabelFont = New-Object System.Drawing.Font($fontFamily, $config.fonts.statusLabelSize, [System.Drawing.FontStyle]::Italic)


    ## Creates the looks and feel
    $backgroundColor = $config.backgroundColor
    $listviewColor = "#f2eded"

    $form1 = New-Object System.Windows.Forms.Form
    $form1.Text = "Software Installation Checklist"
    $form1.ClientSize = New-Object System.Drawing.Size(800, 610)
    $form1.Font = $formFont
    $form1.BackColor = $backgroundColor

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(35, 590)
    $statusLabel.Size = New-Object System.Drawing.Size(350, 20)
    $statusLabel.ForeColor = [System.Drawing.Color]::White
    $statusLabel.Font = $statusLabelFont
    $statusLabel.Text = "Select items and click 'Complete Selected' to install"
    $form1.Controls.Add($statusLabel)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(370, 530)
    $tabControl.Font = $tabFont
    $form1.Controls.Add($tabControl)

    $listBox1 = New-Object System.Windows.Forms.ListBox
    $listBox1.FormattingEnabled = $True
    $listBox1.Size = New-Object System.Drawing.Size(400, 600)
    $listBox1.Location = New-Object System.Drawing.Point(390, 10)
    $listBox1.BackColor = $listviewColor
    $form1.Controls.Add($listBox1)

    $button1 = New-Object System.Windows.Forms.Button
    $button1.Text = "Complete Selected"
    $button1.Size = New-Object System.Drawing.Size(150, 30)
    $button1.Location = New-Object System.Drawing.Point(115, 550)
    $button1.ForeColor = [System.Drawing.Color]::White
    $button1.Font = $buttonFont
    $form1.Controls.Add($button1)

    $jsonFiles = Get-ChildItem -Path $softwareDir -Filter "*.json" | Sort-Object Name

    foreach ($jsonFile in $jsonFiles) {
        $jsonData = Get-Content $jsonFile.FullName | ConvertFrom-Json

        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = $jsonData.tabName
        $tab.BackColor = $backgroundColor
        $tabControl.Controls.Add($tab)

        $yPosition = 15
        foreach ($item in $jsonData.items) {
            if ($item.displayAs -eq "button") {
                $button = New-Object System.Windows.Forms.Button
                $button.Text = $item.name
                $button.Location = New-Object System.Drawing.Point(15, $yPosition)
                $button.Size = New-Object System.Drawing.Size(330, 30)
                $button.ForeColor = [System.Drawing.Color]::White
                $button.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
                $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $button.Font = $checkBoxFont

                $button.Tag = $item

                $button.add_Click({
                    $clickedItem = $this.Tag

                    if ($clickedItem.logMessage) {
                        $messages = $clickedItem.logMessage -split "`n"
                        foreach ($msg in $messages) {
                            $listBox1.Items.Add($msg)
                        }
                    }

                    try {
                        Execute-Item $clickedItem $listBox1
                        $listBox1.Items.Add("Action completed.")
                    }
                    catch {
                        $listBox1.Items.Add("ERROR: $($_.Exception.Message)")
                    }
                })

                $tab.Controls.Add($button)
                $yPosition += 35 
            }
            else {
                $checkbox = New-Object System.Windows.Forms.CheckBox
                $checkbox.Text = $item.name
                $checkbox.Location = New-Object System.Drawing.Point(15, $yPosition)
                $checkbox.Size = New-Object System.Drawing.Size(330, 25)
                $checkbox.ForeColor = [System.Drawing.Color]::White
                $checkbox.Font = $checkBoxFont

                $checkbox.Tag = $item

                $tab.Controls.Add($checkbox)
                $yPosition += 30
            }
        }
    }

    $handler_button1_Click = {

        foreach ($tab in $tabControl.Controls) {
            foreach ($control in $tab.Controls) {
                if ($control -is [System.Windows.Forms.CheckBox] -and $control.Checked) {
                    $item = $control.Tag
                    if ($item.logMessage) {
                        $messages = $item.logMessage -split "`n"
                        foreach ($msg in $messages) {
                            $listBox1.Items.Add($msg)
                        }
                    }

                    try {
                        Execute-Item $item $listBox1
                    }
                    catch {
                        $listBox1.Items.Add("ERROR: $($_.Exception.Message)")
                    }

                    $control.Checked = $false
                }
            }
        }
    }

    $button1.add_Click($handler_button1_Click)

    # Show the Form
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    $OnLoadForm_StateCorrection = {
        $form1.WindowState = $InitialFormWindowState
    }
    $form1.add_Load($OnLoadForm_StateCorrection)
    $form1.ShowDialog() | Out-Null
}

GenerateForm
