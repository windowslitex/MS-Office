@echo off
powershell -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content '%~f0' | Select-Object -Skip 3 | Out-String)"
exit /b

# -----------------------------------------------------------------------------
# PowerShell Windows Config Studio - Worldwide Full-Name Language Engine
# -----------------------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Xml

$currentDir = [Environment]::CurrentDirectory

# ---------- Dynamic Accent Theme Extraction Engine ----------
$accentColor = [System.Drawing.Color]::FromArgb(0, 120, 212) 
try {
    $regPath = "HKCU:\Software\Microsoft\Windows\DWM"
    if (Test-Path $regPath) {
        $colorValue = Get-ItemProperty -Path $regPath -Name ColorizationColor -ErrorAction SilentlyContinue
        if ($colorValue) {
            $bytes = [BitConverter]::GetBytes([uint32]$colorValue.ColorizationColor)
            $accentColor = [System.Drawing.Color]::FromArgb(255, $bytes[2], $bytes[1], $bytes[0])
        }
    }
} catch {}

$glassPanelBg  = [System.Drawing.Color]::FromArgb(245, 247, 250)
$glassCanvasBg = [System.Drawing.Color]::FromArgb(220, 226, 235)
$textPrimary   = [System.Drawing.Color]::FromArgb(20, 20, 20)
$textMuted     = [System.Drawing.Color]::FromArgb(90, 95, 105)

$fontSection = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5)
$fontRegular = New-Object System.Drawing.Font("Segoe UI", 9)
$fontConsole = New-Object System.Drawing.Font("Consolas", 9.5)

# ---------- Base Form Container ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Office Deployment Utility"
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.Size = New-Object System.Drawing.Size(540, 520)
$form.BackColor = $glassCanvasBg
$form.ForeColor = $textPrimary

# Top Accent Strip
$aeroRibbon = New-Object System.Windows.Forms.Panel
$aeroRibbon.Dock = "Top"
$aeroRibbon.Height = 4
$aeroRibbon.BackColor = $accentColor
$form.Controls.Add($aeroRibbon)

# Main Flow Layout Panel
$flowContainer = New-Object System.Windows.Forms.FlowLayoutPanel
$flowContainer.Dock = "Fill"
$flowContainer.FlowDirection = "TopDown"
$flowContainer.WrapContents = $false
$flowContainer.AutoScroll = $true
$flowContainer.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)
$form.Controls.Add($flowContainer)

function Create-FlowLabel {
    param($text, $font = $fontRegular, $color = $textPrimary)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Font = $font
    $lbl.ForeColor = $color
    $lbl.AutoSize = $true
    $lbl.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 4)
    return $lbl
}

# Data Arrays
$apps = @("Word", "Excel", "Outlook", "Teams", "PowerPoint", "Access", "Bing", "Groove", "Lync", "OneDrive", "OneNote", "Publisher")
$coreApps = @("Word", "Excel", "Outlook", "Teams", "PowerPoint")
$years    = @("O365", "2016", "2019", "2021", "2024")
$editions = @("Retail", "Volume")
$products = @("ProPlus", "Professional", "Standard", "HomeBusiness", "Business")

# Comprehensive Global Language Dictionary Matrix (Ordered: US -> Indian Reg -> World Famous)
$langMap = [ordered]@{
    "English (United States)"  = "en-us"
    "English (United Kingdom)" = "en-gb"
    "English (India)"          = "en-in"
    "Hindi (India)"            = "hi-in"
    "Bengali (India)"          = "bn-in"
    "Telugu (India)"           = "te-in"
    "Tamil (India)"            = "ta-in"
    "Marathi (India)"          = "mr-in"
    "Gujarati (India)"         = "gu-in"
    "Kannada (India)"          = "kn-in"
    "Malayalam (India)"        = "ml-in"
    "Punjabi (India)"          = "pa-in"
    "Urdu (India)"             = "ur-in"
    "Spanish (Spain)"          = "es-es"
    "Spanish (Mexico)"         = "es-mx"
    "French (France)"          = "fr-fr"
    "French (Canada)"          = "fr-ca"
    "German (Germany)"         = "de-de"
    "Chinese (Simplified)"     = "zh-cn"
    "Chinese (Traditional)"    = "zh-tw"
    "Japanese (Japan)"         = "ja-jp"
    "Korean (South Korea)"     = "ko-kr"
    "Russian (Russia)"         = "ru-ru"
    "Arabic (Saudi Arabia)"    = "ar-sa"
    "Portuguese (Brazil)"      = "pt-br"
    "Portuguese (Portugal)"    = "pt-pt"
    "Italian (Italy)"          = "it-it"
    "Dutch (Netherlands)"      = "nl-nl"
    "Turkish (Turkey)"         = "tr-tr"
    "Polish (Poland)"          = "pl-pl"
    "Swedish (Sweden)"         = "sv-se"
    "Vietnamese (Vietnam)"     = "vi-vn"
    "Indonesian (Indonesia)"   = "id-id"
    "Thai (Thailand)"          = "th-th"
}

$displayLanguages = [System.Collections.ArrayList]::new()
foreach ($key in $langMap.Keys) { $displayLanguages.Add($key) | Out-Null }

# --- UI Build Execution Stack ---

# Section 1: Edition Mapping
$flowContainer.Controls.Add((Create-FlowLabel "1. Deployment Edition Mapping" $fontSection $accentColor))

$panelEdition = New-Object System.Windows.Forms.TableLayoutPanel
$panelEdition.Size = New-Object System.Drawing.Size(480, 40)
$panelEdition.ColumnCount = 3
$panelEdition.RowCount = 1
$panelEdition.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$panelEdition.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$panelEdition.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$panelEdition.BackColor = $glassPanelBg
$panelEdition.BorderStyle = "FixedSingle"
$panelEdition.Padding = New-Object System.Windows.Forms.Padding(4)
$panelEdition.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)

$comboYear = New-Object System.Windows.Forms.ComboBox -Property @{DropDownStyle="DropDownList"; Font=$fontRegular; BackColor=$glassPanelBg; Dock="Fill"}
$comboYear.Items.AddRange($years); $comboYear.SelectedIndex = 0
$comboEdition = New-Object System.Windows.Forms.ComboBox -Property @{DropDownStyle="DropDownList"; Font=$fontRegular; BackColor=$glassPanelBg; Dock="Fill"}
$comboEdition.Items.AddRange($editions); $comboEdition.SelectedIndex = 0
$comboProduct = New-Object System.Windows.Forms.ComboBox -Property @{DropDownStyle="DropDownList"; Font=$fontRegular; BackColor=$glassPanelBg; Dock="Fill"}
$comboProduct.Items.AddRange($products); $comboProduct.SelectedIndex = 0

$panelEdition.Controls.Add($comboYear, 0, 0)
$panelEdition.Controls.Add($comboEdition, 1, 0)
$panelEdition.Controls.Add($comboProduct, 2, 0)
$flowContainer.Controls.Add($panelEdition)

# Section 2: App Exclusions
$flowContainer.Controls.Add((Create-FlowLabel "2. App Exclusion Matrix (Checked = Blocked)" $fontSection $accentColor))

$panelExclusions = New-Object System.Windows.Forms.Panel
$panelExclusions.Size = New-Object System.Drawing.Size(480, 90)
$panelExclusions.BackColor = $glassPanelBg
$panelExclusions.BorderStyle = "FixedSingle"
$panelExclusions.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)

$flowApps = New-Object System.Windows.Forms.FlowLayoutPanel
$flowApps.Dock = "Fill"
$flowApps.AutoScroll = $true
$flowApps.Padding = New-Object System.Windows.Forms.Padding(6)

$checkboxes = @{}
foreach ($app in $apps) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $app
    $cb.Width = 105
    $cb.Height = 22
    $cb.Font = $fontRegular
    if ($coreApps -notcontains $app) { $cb.Checked = $true }
    $checkboxes[$app] = $cb
    $flowApps.Controls.Add($cb)
}
$panelExclusions.Controls.Add($flowApps)
$flowContainer.Controls.Add($panelExclusions)

# Section 3: Localization Setup
$flowContainer.Controls.Add((Create-FlowLabel "3. Localized Packages" $fontSection $accentColor))

$panelLang = New-Object System.Windows.Forms.TableLayoutPanel
$panelLang.Size = New-Object System.Drawing.Size(480, 140)
$panelLang.ColumnCount = 2
$panelLang.RowCount = 1
$panelLang.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$panelLang.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$panelLang.BackColor = $glassPanelBg
$panelLang.BorderStyle = "FixedSingle"
$panelLang.Padding = New-Object System.Windows.Forms.Padding(4)
$panelLang.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)

$p1 = New-Object System.Windows.Forms.Panel -Property @{Dock="Fill"}
$l1 = Create-FlowLabel "Primary System UI" $fontRegular $textMuted
$l1.Dock = "Top"; $l1.Margin = New-Object System.Windows.Forms.Padding(0)
$clbMain = New-Object System.Windows.Forms.CheckedListBox -Property @{Dock="Fill"; BorderStyle="None"; CheckOnClick=$true; Font=$fontRegular; BackColor=$glassPanelBg}
$clbMain.Items.AddRange($displayLanguages)
$idx = $clbMain.Items.IndexOf("English (United States)"); if ($idx -ge 0) { $clbMain.SetItemChecked($idx, $true) }
$p1.Controls.AddRange(@($clbMain, $l1))

$p2 = New-Object System.Windows.Forms.Panel -Property @{Dock="Fill"}
$l2 = Create-FlowLabel "Optional Proofing" $fontRegular $textMuted
$l2.Dock = "Top"; $l2.Margin = New-Object System.Windows.Forms.Padding(0)
$clbProof = New-Object System.Windows.Forms.CheckedListBox -Property @{Dock="Fill"; BorderStyle="None"; CheckOnClick=$true; Font=$fontRegular; BackColor=$glassPanelBg}
$clbProof.Items.AddRange($displayLanguages)
$idx = $clbProof.Items.IndexOf("English (United States)"); if ($idx -ge 0) { $clbProof.SetItemChecked($idx, $true) }
$p2.Controls.AddRange(@($clbProof, $l2))

$panelLang.Controls.Add($p1, 0, 0)
$panelLang.Controls.Add($p2, 1, 0)
$flowContainer.Controls.Add($panelLang)

# Section 4: Advanced Configurations
$flowContainer.Controls.Add((Create-FlowLabel "4. Advanced Virtualization Config" $fontSection $accentColor))

$panelOpt = New-Object System.Windows.Forms.Panel
$panelOpt.Size = New-Object System.Drawing.Size(480, 36)
$panelOpt.BackColor = $glassPanelBg
$panelOpt.BorderStyle = "FixedSingle"
$panelOpt.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 15)

$chkShared = New-Object System.Windows.Forms.CheckBox
$chkShared.Text = "Enable Shared Computer Licensing System Layout"
$chkShared.Font = $fontRegular
$chkShared.Location = New-Object System.Drawing.Point(8, 7)
$chkShared.Size = New-Object System.Drawing.Size(450, 20)
$panelOpt.Controls.Add($chkShared)
$flowContainer.Controls.Add($panelOpt)

# Action Buttons Container Setup
$panelAction = New-Object System.Windows.Forms.TableLayoutPanel
$panelAction.Size = New-Object System.Drawing.Size(480, 42)
$panelAction.ColumnCount = 2
$panelAction.RowCount = 1
$panelAction.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$panelAction.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$panelAction.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 5)

$btnPreview = New-Object System.Windows.Forms.Button -Property @{Text="Preview XML Script"; FlatStyle="Flat"; BackColor=$accentColor; ForeColor=[System.Drawing.Color]::White; Font=$fontSection; Dock="Fill"}
$btnPreview.FlatAppearance.BorderColor = $accentColor
$btnPreview.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)

$btnCancel = New-Object System.Windows.Forms.Button -Property @{Text="Close Engine"; FlatStyle="Flat"; BackColor=[System.Drawing.Color]::FromArgb(120, 130, 140); ForeColor=[System.Drawing.Color]::White; Font=$fontSection; Dock="Fill"}
$btnCancel.FlatAppearance.BorderSize = 0
$btnCancel.Margin = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$btnCancel.Add_Click({ $form.Close() })

$panelAction.Controls.Add($btnPreview, 0, 0)
$panelAction.Controls.Add($btnCancel, 1, 0)
$flowContainer.Controls.Add($panelAction)

$statusStrip = New-Object System.Windows.Forms.StatusStrip -Property @{BackColor=$glassPanelBg; SizingGrip=$false}
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel -Property @{Text=" Ready."; Font=$fontRegular}
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

# ---------- Data Processing Logic Matrix ----------
function Build-XmlString {
    $excluded = @()
    foreach ($key in $checkboxes.Keys) {
        if ($checkboxes[$key].Checked) { $excluded += $key }
    }

    $year = $comboYear.SelectedItem
    $editionType = $comboEdition.SelectedItem
    $product = $comboProduct.SelectedItem
    
    $productId = if ($year -eq "O365") { "O365" + $product + $editionType } else { $product + $year + $editionType }

    # Parse display text mapping back to standard ISO identifiers
    $mainLangs = @()
    for ($i = 0; $i -lt $clbMain.Items.Count; $i++) {
        if ($clbMain.GetItemChecked($i)) { 
            $displayName = $clbMain.Items[$i]
            $mainLangs += $langMap[$displayName]
        }
    }
    
    $proofLangs = @()
    for ($i = 0; $i -lt $clbProof.Items.Count; $i++) {
        if ($clbProof.GetItemChecked($i)) { 
            $displayName = $clbProof.Items[$i]
            $proofLangs += $langMap[$displayName]
        }
    }

    $xml = [System.Xml.XmlDocument]::new()
    $config = $xml.CreateElement("Configuration")
    $xml.AppendChild($config) | Out-Null
    
    $add = $xml.CreateElement("Add")
    $add.SetAttribute("OfficeClientEdition", "64")
    $add.SetAttribute("Channel", "Current")
    $config.AppendChild($add) | Out-Null
    
    $productElement = $xml.CreateElement("Product")
    $productElement.SetAttribute("ID", $productId)
    $add.AppendChild($productElement) | Out-Null
    
    foreach ($lang in $mainLangs) {
        $langElement = $xml.CreateElement("Language")
        $langElement.SetAttribute("ID", $lang)
        $productElement.AppendChild($langElement) | Out-Null
    }
    
    foreach ($app in $excluded) {
        $excludeElement = $xml.CreateElement("ExcludeApp")
        $excludeElement.SetAttribute("ID", $app)
        $productElement.AppendChild($excludeElement) | Out-Null
    }
    
    if ($proofLangs.Count -gt 0) {
        $proofProduct = $xml.CreateElement("Product")
        $proofProduct.SetAttribute("ID", "ProofingTools")
        $add.AppendChild($proofProduct) | Out-Null
        foreach ($lang in $proofLangs) {
            $langElement = $xml.CreateElement("Language")
            $langElement.SetAttribute("ID", $lang)
            $proofProduct.AppendChild($langElement) | Out-Null
        }
    }
    
    $updates = $xml.CreateElement("Updates")
    $updates.SetAttribute("Enabled", "TRUE")
    $updates.SetAttribute("Channel", "Current")
    $config.AppendChild($updates) | Out-Null
    
    $display = $xml.CreateElement("Display")
    $display.SetAttribute("Level", "Full")
    $display.SetAttribute("AcceptEULA", "TRUE")
    $config.AppendChild($display) | Out-Null
    
    if ($chkShared.Checked) {
        $property = $xml.CreateElement("Property")
        $property.SetAttribute("Name", "SharedComputerLicensing")
        $property.SetAttribute("Value", "1")
        $config.AppendChild($property) | Out-Null
    }
    
    $property2 = $xml.CreateElement("Property")
    $property2.SetAttribute("Name", "FORCEAPPSHUTDOWN")
    $property2.SetAttribute("Value", "TRUE")
    $config.AppendChild($property2) | Out-Null
    
    $stringWriter = New-Object System.IO.StringWriter
    $xmlWriter = New-Object System.Xml.XmlTextWriter($stringWriter)
    $xmlWriter.Formatting = "Indented"
    $xmlWriter.Indentation = 4
    $xml.WriteTo($xmlWriter)
    $xmlWriter.Flush()
    return $stringWriter.ToString()
}

# ---------- Workspace Popup Window Engine ----------
function Show-XmlPopup {
    param($xmlContent)
    
    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = "Configuration Script Preview Workspace"
    $popup.Size = New-Object System.Drawing.Size(560, 480)
    $popup.StartPosition = "CenterParent"
    $popup.FormBorderStyle = "FixedDialog"
    $popup.MaximizeBox = $false
    $popup.MinimizeBox = $false
    $popup.BackColor = $glassCanvasBg
    
    $popupRibbon = New-Object System.Windows.Forms.Panel
    $popupRibbon.Dock = "Top"
    $popupRibbon.Height = 4
    $popupRibbon.BackColor = $accentColor
    $popup.Controls.Add($popupRibbon)

    $layout = New-Object System.Windows.Forms.FlowLayoutPanel
    $layout.Dock = "Fill"
    $layout.FlowDirection = "TopDown"
    $layout.WrapContents = $false
    $layout.Padding = New-Object System.Windows.Forms.Padding(15)
    $popup.Controls.Add($layout)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "REAL-TIME XML CONFIGURATION DOCUMENT"
    $lbl.Font = $fontSection
    $lbl.ForeColor = $accentColor
    $lbl.AutoSize = $true
    $lbl.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
    $layout.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Size = New-Object System.Drawing.Size(515, 300)
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Both"
    $txt.BackColor = $glassPanelBg
    $txt.ForeColor = $textPrimary
    $txt.Font = $fontConsole
    $txt.BorderStyle = "FixedSingle"
    $txt.Text = $xmlContent
    $layout.Controls.Add($txt)

    $btnPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $btnPanel.Size = New-Object System.Drawing.Size(515, 42)
    $btnPanel.ColumnCount = 2
    $btnPanel.RowCount = 1
    $btnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $btnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $btnPanel.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 0)

    $btnSave = New-Object System.Windows.Forms.Button -Property @{Text="Save Configuration.xml"; FlatStyle="Flat"; BackColor=$accentColor; ForeColor=[System.Drawing.Color]::White; Font=$fontSection; Dock="Fill"}
    $btnSave.FlatAppearance.BorderColor = $accentColor
    $btnSave.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
    
    $btnSave.Add_Click({
        $configPath = Join-Path $currentDir "configuration.xml"
        try {
            $xmlContent | Out-File -FilePath $configPath -Encoding UTF8
            $statusLabel.Text = " Configuration script successfully saved."
            [System.Windows.Forms.MessageBox]::Show("Configuration saved successfully!`nPath: $configPath", "Status Engine Complete", "OK", "Information")
            $popup.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving configuration document.", "Error", "OK", "Error")
        }
    })

    $btnClosePopup = New-Object System.Windows.Forms.Button -Property @{Text="Return to Designer"; FlatStyle="Flat"; BackColor=[System.Drawing.Color]::FromArgb(120, 130, 140); ForeColor=[System.Drawing.Color]::White; Font=$fontSection; Dock="Fill"}
    $btnClosePopup.FlatAppearance.BorderSize = 0
    $btnClosePopup.Margin = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
    $btnClosePopup.Add_Click({ $popup.Close() })

    $btnPanel.Controls.Add($btnSave, 0, 0)
    $btnPanel.Controls.Add($btnClosePopup, 1, 0)
    $layout.Controls.Add($btnPanel)

    $popup.ShowDialog($form) | Out-Null
}

# ---------- Route Interaction Wire Hooks ----------
$btnPreview.Add_Click({
    $currentXmlData = Build-XmlString
    Show-XmlPopup -xmlContent $currentXmlData
})

# Display Window Pipeline Loop Initialization
$form.Add_Shown({ $form.Activate() })
$form.ShowDialog() | Out-Null