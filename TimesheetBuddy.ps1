Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook

function Get-TimesheetData {
    param([datetime]$targetDate)

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
    $items = $calendar.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")

    $startDate = $targetDate.Date
    $endDate = $startDate.AddDays(1)
    $restriction = "[Start] >= '$($startDate.ToString("g"))' AND [Start] < '$($endDate.ToString("g"))'"
    $filteredItems = $items.Restrict($restriction)

    $output = @()
    foreach ($item in $filteredItems) {
        $start = $item.Start.ToString("h:mm tt")
        $end = $item.End.ToString("h:mm tt")
        $subject = $item.Subject
        $location = $item.Location
        $output += "$start - ${end}: $subject" + ($(if ($location) { " @ $location" } else { "" }))
    }

    return $output -join "`r`n"
}

function Save-TimesheetFile {
    param($text, $targetDate)
    $filename = "timesheet_$($targetDate.ToString('yyyy-MM-dd')).txt"
    $filepath = Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath $filename
    $text | Set-Content -Path $filepath
    [System.Windows.Forms.MessageBox]::Show("Saved to:`n$filepath", "Timesheet Saved")
}

function Copy-ToClipboard {
    param($text)
    [System.Windows.Forms.Clipboard]::SetText($text)
    [System.Windows.Forms.MessageBox]::Show("Timesheet copied to clipboard.", "Copied")
}

# --- UI Styling ---
$font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Regular)
$btnFont = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$bgColor = [System.Drawing.Color]::FromArgb(40, 44, 52)
$fgColor = [System.Drawing.Color]::WhiteSmoke
$btnColor = [System.Drawing.Color]::FromArgb(60, 63, 65)

# --- Main Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Timesheet Buddy"
$form.Size = New-Object System.Drawing.Size(260, 430)
$form.FormBorderStyle = "FixedToolWindow"
$form.TopMost = $true
$form.StartPosition = "Manual"
$form.Location = New-Object System.Drawing.Point(0, 100)
$form.BackColor = $bgColor
$form.Font = $font

# --- Label ---
$label = New-Object System.Windows.Forms.Label
$label.Text = "Select a date:"
$label.Location = New-Object System.Drawing.Point(20, 20)
$label.Size = New-Object System.Drawing.Size(200, 30)
$label.ForeColor = $fgColor
$form.Controls.Add($label)

# --- Date Picker ---
$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$datePicker.Location = New-Object System.Drawing.Point(20, 55)
$datePicker.Width = 200
$form.Controls.Add($datePicker)

# --- Save Button ---
$saveBtn = New-Object System.Windows.Forms.Button
$saveBtn.Text = "📄 Capture Timesheet"
$saveBtn.Location = New-Object System.Drawing.Point(20, 95)
$saveBtn.Size = New-Object System.Drawing.Size(200, 40)
$saveBtn.BackColor = $btnColor
$saveBtn.ForeColor = $fgColor
$saveBtn.Font = $btnFont
$saveBtn.FlatStyle = "Flat"
$saveBtn.Add_Click({
    $text = Get-TimesheetData -targetDate $datePicker.Value
    if (-not $text) {
        [System.Windows.Forms.MessageBox]::Show("No appointments found for selected date.", "Nothing Found")
        return
    }
    Save-TimesheetFile -text $text -targetDate $datePicker.Value
})
$form.Controls.Add($saveBtn)

# --- Copy Button ---
$copyBtn = New-Object System.Windows.Forms.Button
$copyBtn.Text = "📋 Copy to Clipboard"
$copyBtn.Location = New-Object System.Drawing.Point(20, 145)
$copyBtn.Size = New-Object System.Drawing.Size(200, 40)
$copyBtn.BackColor = $btnColor
$copyBtn.ForeColor = $fgColor
$copyBtn.Font = $btnFont
$copyBtn.FlatStyle = "Flat"
$copyBtn.Add_Click({
    $text = Get-TimesheetData -targetDate $datePicker.Value
    if (-not $text) {
        [System.Windows.Forms.MessageBox]::Show("No appointments found for selected date.", "Nothing Found")
        return
    }
    Copy-ToClipboard -text $text
})
$form.Controls.Add($copyBtn)

# --- Exit Button ---
$exitBtn = New-Object System.Windows.Forms.Button
$exitBtn.Text = "❌ Close"
$exitBtn.Location = New-Object System.Drawing.Point(20, 195)
$exitBtn.Size = New-Object System.Drawing.Size(200, 35)
$exitBtn.BackColor = $btnColor
$exitBtn.ForeColor = $fgColor
$exitBtn.Font = $btnFont
$exitBtn.FlatStyle = "Flat"
$exitBtn.Add_Click({ $form.Close() })
$form.Controls.Add($exitBtn)

# --- Banner Panel ---
$bannerPanel = New-Object System.Windows.Forms.Panel
$bannerPanel.Location = New-Object System.Drawing.Point(20, 240)
$bannerPanel.Size = New-Object System.Drawing.Size(200, 130)
$bannerPanel.BackColor = [System.Drawing.Color]::Black
$form.Controls.Add($bannerPanel)

# --- Scrolling Label ---
$bannerLabel = New-Object System.Windows.Forms.Label
$bannerLabel.Text = " ✨ doing the needful... ✨ "
$bannerLabel.ForeColor = [System.Drawing.Color]::White
$bannerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$bannerLabel.AutoSize = $true
$bannerLabel.Location = New-Object System.Drawing.Point($bannerPanel.Width, 50)
$bannerPanel.Controls.Add($bannerLabel)

# --- Timer for Scrolling ---
$scrollTimer = New-Object System.Windows.Forms.Timer
$scrollTimer.Interval = 30
$scrollTimer.Add_Tick({
    $bannerLabel.Left -= 2
    if ($bannerLabel.Right -lt 0) {
        $bannerLabel.Left = $bannerPanel.Width
    }
})
$scrollTimer.Start()

# --- Show Form ---
[void]$form.ShowDialog()
