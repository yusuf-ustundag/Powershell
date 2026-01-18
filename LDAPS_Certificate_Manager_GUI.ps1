#Requires -RunAsAdministrator
<#
.SYNOPSIS
    LDAPS Certificate Manager - Professional Light Theme GUI
.NOTES
    Version: 4.1 - Bug fixes + Issuer info + Enhanced logging
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Light Theme Color Palette
$script:Colors = @{
    # Backgrounds
    BgMain       = [System.Drawing.Color]::FromArgb(248, 249, 250)
    BgWhite      = [System.Drawing.Color]::White
    BgCard       = [System.Drawing.Color]::White
    BgHeader     = [System.Drawing.Color]::FromArgb(37, 99, 235)
    BgInput      = [System.Drawing.Color]::FromArgb(241, 243, 245)
    
    # Borders
    Border       = [System.Drawing.Color]::FromArgb(222, 226, 230)
    BorderDark   = [System.Drawing.Color]::FromArgb(206, 212, 218)
    
    # Text
    TextDark     = [System.Drawing.Color]::FromArgb(33, 37, 41)
    TextMuted    = [System.Drawing.Color]::FromArgb(108, 117, 125)
    TextLight    = [System.Drawing.Color]::White
    
    # Accent Colors
    Blue         = [System.Drawing.Color]::FromArgb(37, 99, 235)
    BlueHover    = [System.Drawing.Color]::FromArgb(29, 78, 216)
    BlueLight    = [System.Drawing.Color]::FromArgb(219, 234, 254)
    
    Green        = [System.Drawing.Color]::FromArgb(34, 197, 94)
    GreenHover   = [System.Drawing.Color]::FromArgb(22, 163, 74)
    GreenLight   = [System.Drawing.Color]::FromArgb(220, 252, 231)
    GreenText    = [System.Drawing.Color]::FromArgb(21, 128, 61)
    
    Orange       = [System.Drawing.Color]::FromArgb(249, 115, 22)
    OrangeHover  = [System.Drawing.Color]::FromArgb(234, 88, 12)
    OrangeLight  = [System.Drawing.Color]::FromArgb(255, 237, 213)
    OrangeText   = [System.Drawing.Color]::FromArgb(194, 65, 12)
    
    Red          = [System.Drawing.Color]::FromArgb(239, 68, 68)
    RedHover     = [System.Drawing.Color]::FromArgb(220, 38, 38)
    RedLight     = [System.Drawing.Color]::FromArgb(254, 226, 226)
    RedText      = [System.Drawing.Color]::FromArgb(185, 28, 28)
    
    Gray         = [System.Drawing.Color]::FromArgb(107, 114, 128)
    GrayHover    = [System.Drawing.Color]::FromArgb(75, 85, 99)
    GrayLight    = [System.Drawing.Color]::FromArgb(243, 244, 246)
}

$script:ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { Get-Location }
$script:LogDirectory = Join-Path $script:ScriptPath "LDAPS_Manager_Logs"
if (-not (Test-Path $script:LogDirectory)) { New-Item -ItemType Directory -Path $script:LogDirectory -Force | Out-Null }
#endregion

#region Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "LDAPS Certificate Manager"
$form.Size = New-Object System.Drawing.Size(1000, 835)
$form.StartPosition = "CenterScreen"
$form.BackColor = $script:Colors.BgMain
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
#endregion

#region Header
$header = New-Object System.Windows.Forms.Panel
$header.Location = New-Object System.Drawing.Point(0, 0)
$header.Size = New-Object System.Drawing.Size(1000, 90)
$header.BackColor = $script:Colors.BgHeader

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "LDAPS Certificate Manager"
$lblTitle.Location = New-Object System.Drawing.Point(24, 18)
$lblTitle.AutoSize = $true
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $script:Colors.TextLight
$header.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Professional Certificate Lifecycle Management"
$lblSub.Location = New-Object System.Drawing.Point(26, 52)
$lblSub.AutoSize = $true
$lblSub.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSub.ForeColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)
$header.Controls.Add($lblSub)

$lblServer = New-Object System.Windows.Forms.Label
$lblServer.Text = $env:COMPUTERNAME
$lblServer.Location = New-Object System.Drawing.Point(800, 20)
$lblServer.Size = New-Object System.Drawing.Size(170, 24)
$lblServer.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$lblServer.Font = New-Object System.Drawing.Font("Consolas", 14, [System.Drawing.FontStyle]::Bold)
$lblServer.ForeColor = $script:Colors.TextLight
$header.Controls.Add($lblServer)

$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = (Get-Date -Format "yyyy-MM-dd HH:mm")
$lblDate.Location = New-Object System.Drawing.Point(800, 48)
$lblDate.Size = New-Object System.Drawing.Size(170, 20)
$lblDate.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$lblDate.ForeColor = [System.Drawing.Color]::FromArgb(180, 255, 255, 255)
$header.Controls.Add($lblDate)

$form.Controls.Add($header)
#endregion

#region Certificate Cards Container
$cardsPanel = New-Object System.Windows.Forms.Panel
$cardsPanel.Location = New-Object System.Drawing.Point(24, 106)
$cardsPanel.Size = New-Object System.Drawing.Size(950, 200)
$cardsPanel.BackColor = $script:Colors.BgMain
$form.Controls.Add($cardsPanel)

# Current Certificate Card - Using GroupBox for border
$cardCurrent = New-Object System.Windows.Forms.GroupBox
$cardCurrent.Location = New-Object System.Drawing.Point(0, 0)
$cardCurrent.Size = New-Object System.Drawing.Size(460, 200)
$cardCurrent.BackColor = $script:Colors.BgWhite
$cardCurrent.Text = "Current LDAPS Certificate"
$cardCurrent.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
$cardCurrent.ForeColor = $script:Colors.TextDark
$cardsPanel.Controls.Add($cardCurrent)

# Thumbprint
$lblCurThumbT = New-Object System.Windows.Forms.Label
$lblCurThumbT.Text = "Thumbprint:"
$lblCurThumbT.Location = New-Object System.Drawing.Point(12, 30)
$lblCurThumbT.Size = New-Object System.Drawing.Size(75, 18)
$lblCurThumbT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCurThumbT.ForeColor = $script:Colors.TextMuted
$cardCurrent.Controls.Add($lblCurThumbT)

$txtCurThumb = New-Object System.Windows.Forms.Label
$txtCurThumb.Text = "Loading..."
$txtCurThumb.Location = New-Object System.Drawing.Point(90, 30)
$txtCurThumb.Size = New-Object System.Drawing.Size(355, 18)
$txtCurThumb.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$txtCurThumb.ForeColor = $script:Colors.TextDark
$cardCurrent.Controls.Add($txtCurThumb)

# Subject
$lblCurSubjT = New-Object System.Windows.Forms.Label
$lblCurSubjT.Text = "Subject:"
$lblCurSubjT.Location = New-Object System.Drawing.Point(12, 52)
$lblCurSubjT.Size = New-Object System.Drawing.Size(75, 18)
$lblCurSubjT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCurSubjT.ForeColor = $script:Colors.TextMuted
$cardCurrent.Controls.Add($lblCurSubjT)

$txtCurSubj = New-Object System.Windows.Forms.Label
$txtCurSubj.Text = "-"
$txtCurSubj.Location = New-Object System.Drawing.Point(90, 52)
$txtCurSubj.Size = New-Object System.Drawing.Size(355, 18)
$txtCurSubj.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtCurSubj.ForeColor = $script:Colors.TextDark
$cardCurrent.Controls.Add($txtCurSubj)

# Issuer (NEW)
$lblCurIssuerT = New-Object System.Windows.Forms.Label
$lblCurIssuerT.Text = "Issuer:"
$lblCurIssuerT.Location = New-Object System.Drawing.Point(12, 74)
$lblCurIssuerT.Size = New-Object System.Drawing.Size(75, 18)
$lblCurIssuerT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCurIssuerT.ForeColor = $script:Colors.TextMuted
$cardCurrent.Controls.Add($lblCurIssuerT)

$txtCurIssuer = New-Object System.Windows.Forms.Label
$txtCurIssuer.Text = "-"
$txtCurIssuer.Location = New-Object System.Drawing.Point(90, 74)
$txtCurIssuer.Size = New-Object System.Drawing.Size(355, 18)
$txtCurIssuer.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtCurIssuer.ForeColor = $script:Colors.Blue
$cardCurrent.Controls.Add($txtCurIssuer)

# Expires
$lblCurExpT = New-Object System.Windows.Forms.Label
$lblCurExpT.Text = "Expires:"
$lblCurExpT.Location = New-Object System.Drawing.Point(12, 96)
$lblCurExpT.Size = New-Object System.Drawing.Size(75, 18)
$lblCurExpT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCurExpT.ForeColor = $script:Colors.TextMuted
$cardCurrent.Controls.Add($lblCurExpT)

$txtCurExp = New-Object System.Windows.Forms.Label
$txtCurExp.Text = "-"
$txtCurExp.Location = New-Object System.Drawing.Point(90, 96)
$txtCurExp.Size = New-Object System.Drawing.Size(150, 18)
$txtCurExp.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtCurExp.ForeColor = $script:Colors.TextDark
$cardCurrent.Controls.Add($txtCurExp)

# Status Badge
$badgeCur = New-Object System.Windows.Forms.Label
$badgeCur.Location = New-Object System.Drawing.Point(320, 92)
$badgeCur.Size = New-Object System.Drawing.Size(90, 26)
$badgeCur.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$badgeCur.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$badgeCur.Text = "UNKNOWN"
$badgeCur.BackColor = $script:Colors.GrayLight
$badgeCur.ForeColor = $script:Colors.Gray
$cardCurrent.Controls.Add($badgeCur)

# Days Left
$lblDaysT = New-Object System.Windows.Forms.Label
$lblDaysT.Text = "Days Left:"
$lblDaysT.Location = New-Object System.Drawing.Point(12, 135)
$lblDaysT.Size = New-Object System.Drawing.Size(75, 18)
$lblDaysT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblDaysT.ForeColor = $script:Colors.TextMuted
$cardCurrent.Controls.Add($lblDaysT)

$progDays = New-Object System.Windows.Forms.ProgressBar
$progDays.Location = New-Object System.Drawing.Point(90, 135)
$progDays.Size = New-Object System.Drawing.Size(260, 22)
$progDays.Maximum = 365
$progDays.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$cardCurrent.Controls.Add($progDays)

$lblDaysVal = New-Object System.Windows.Forms.Label
$lblDaysVal.Text = "- days"
$lblDaysVal.Location = New-Object System.Drawing.Point(360, 135)
$lblDaysVal.Size = New-Object System.Drawing.Size(85, 22)
$lblDaysVal.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$lblDaysVal.ForeColor = $script:Colors.TextDark
$cardCurrent.Controls.Add($lblDaysVal)

# EKU Info
$lblCurEKU = New-Object System.Windows.Forms.Label
$lblCurEKU.Text = ""
$lblCurEKU.Location = New-Object System.Drawing.Point(90, 165)
$lblCurEKU.Size = New-Object System.Drawing.Size(355, 18)
$lblCurEKU.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblCurEKU.ForeColor = $script:Colors.TextMuted
$cardCurrent.Controls.Add($lblCurEKU)

# Recommended Certificate Card
$cardRec = New-Object System.Windows.Forms.GroupBox
$cardRec.Location = New-Object System.Drawing.Point(490, 0)
$cardRec.Size = New-Object System.Drawing.Size(460, 200)
$cardRec.BackColor = $script:Colors.BgWhite
$cardRec.Text = "Recommended Certificate"
$cardRec.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
$cardRec.ForeColor = $script:Colors.GreenText
$cardsPanel.Controls.Add($cardRec)

# Thumbprint
$lblRecThumbT = New-Object System.Windows.Forms.Label
$lblRecThumbT.Text = "Thumbprint:"
$lblRecThumbT.Location = New-Object System.Drawing.Point(12, 30)
$lblRecThumbT.Size = New-Object System.Drawing.Size(75, 18)
$lblRecThumbT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblRecThumbT.ForeColor = $script:Colors.TextMuted
$cardRec.Controls.Add($lblRecThumbT)

$txtRecThumb = New-Object System.Windows.Forms.Label
$txtRecThumb.Text = "Scanning..."
$txtRecThumb.Location = New-Object System.Drawing.Point(90, 30)
$txtRecThumb.Size = New-Object System.Drawing.Size(355, 18)
$txtRecThumb.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$txtRecThumb.ForeColor = $script:Colors.GreenText
$cardRec.Controls.Add($txtRecThumb)

# Subject
$lblRecSubjT = New-Object System.Windows.Forms.Label
$lblRecSubjT.Text = "Subject:"
$lblRecSubjT.Location = New-Object System.Drawing.Point(12, 52)
$lblRecSubjT.Size = New-Object System.Drawing.Size(75, 18)
$lblRecSubjT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblRecSubjT.ForeColor = $script:Colors.TextMuted
$cardRec.Controls.Add($lblRecSubjT)

$txtRecSubj = New-Object System.Windows.Forms.Label
$txtRecSubj.Text = "-"
$txtRecSubj.Location = New-Object System.Drawing.Point(90, 52)
$txtRecSubj.Size = New-Object System.Drawing.Size(355, 18)
$txtRecSubj.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtRecSubj.ForeColor = $script:Colors.TextDark
$cardRec.Controls.Add($txtRecSubj)

# Issuer (NEW)
$lblRecIssuerT = New-Object System.Windows.Forms.Label
$lblRecIssuerT.Text = "Issuer:"
$lblRecIssuerT.Location = New-Object System.Drawing.Point(12, 74)
$lblRecIssuerT.Size = New-Object System.Drawing.Size(75, 18)
$lblRecIssuerT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblRecIssuerT.ForeColor = $script:Colors.TextMuted
$cardRec.Controls.Add($lblRecIssuerT)

$txtRecIssuer = New-Object System.Windows.Forms.Label
$txtRecIssuer.Text = "-"
$txtRecIssuer.Location = New-Object System.Drawing.Point(90, 74)
$txtRecIssuer.Size = New-Object System.Drawing.Size(355, 18)
$txtRecIssuer.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtRecIssuer.ForeColor = $script:Colors.Blue
$cardRec.Controls.Add($txtRecIssuer)

# Expires
$lblRecExpT = New-Object System.Windows.Forms.Label
$lblRecExpT.Text = "Expires:"
$lblRecExpT.Location = New-Object System.Drawing.Point(12, 96)
$lblRecExpT.Size = New-Object System.Drawing.Size(75, 18)
$lblRecExpT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblRecExpT.ForeColor = $script:Colors.TextMuted
$cardRec.Controls.Add($lblRecExpT)

$txtRecExp = New-Object System.Windows.Forms.Label
$txtRecExp.Text = "-"
$txtRecExp.Location = New-Object System.Drawing.Point(90, 96)
$txtRecExp.Size = New-Object System.Drawing.Size(150, 18)
$txtRecExp.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtRecExp.ForeColor = $script:Colors.GreenText
$cardRec.Controls.Add($txtRecExp)

# Status Badge
$badgeRec = New-Object System.Windows.Forms.Label
$badgeRec.Location = New-Object System.Drawing.Point(320, 92)
$badgeRec.Size = New-Object System.Drawing.Size(90, 26)
$badgeRec.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$badgeRec.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$badgeRec.Text = "VALID"
$badgeRec.BackColor = $script:Colors.GreenLight
$badgeRec.ForeColor = $script:Colors.GreenText
$cardRec.Controls.Add($badgeRec)

# Valid Days
$lblRecDaysT = New-Object System.Windows.Forms.Label
$lblRecDaysT.Text = "Valid for:"
$lblRecDaysT.Location = New-Object System.Drawing.Point(12, 130)
$lblRecDaysT.Size = New-Object System.Drawing.Size(75, 20)
$lblRecDaysT.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblRecDaysT.ForeColor = $script:Colors.TextMuted
$cardRec.Controls.Add($lblRecDaysT)

$lblRecDaysVal = New-Object System.Windows.Forms.Label
$lblRecDaysVal.Text = "- days"
$lblRecDaysVal.Location = New-Object System.Drawing.Point(90, 125)
$lblRecDaysVal.Size = New-Object System.Drawing.Size(200, 40)
$lblRecDaysVal.Font = New-Object System.Drawing.Font("Segoe UI", 22, [System.Drawing.FontStyle]::Bold)
$lblRecDaysVal.ForeColor = $script:Colors.Green
$cardRec.Controls.Add($lblRecDaysVal)

# EKU Info
$lblRecEKU = New-Object System.Windows.Forms.Label
$lblRecEKU.Text = ""
$lblRecEKU.Location = New-Object System.Drawing.Point(90, 170)
$lblRecEKU.Size = New-Object System.Drawing.Size(355, 18)
$lblRecEKU.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblRecEKU.ForeColor = $script:Colors.TextMuted
$cardRec.Controls.Add($lblRecEKU)
#endregion

#region Certificate Store Grid
$lblGrid = New-Object System.Windows.Forms.Label
$lblGrid.Text = "Certificate Store - LDAPS Compatible Only"
$lblGrid.Location = New-Object System.Drawing.Point(24, 321)
$lblGrid.AutoSize = $true
$lblGrid.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
$lblGrid.ForeColor = $script:Colors.TextDark
$form.Controls.Add($lblGrid)

$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(24, 348)
$dgv.Size = New-Object System.Drawing.Size(950, 180)
$dgv.BackgroundColor = $script:Colors.BgWhite
$dgv.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dgv.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$dgv.GridColor = $script:Colors.Border
$dgv.EnableHeadersVisualStyles = $false
$dgv.ColumnHeadersDefaultCellStyle.BackColor = $script:Colors.BgInput
$dgv.ColumnHeadersDefaultCellStyle.ForeColor = $script:Colors.TextDark
$dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$dgv.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8)
$dgv.ColumnHeadersHeight = 36
$dgv.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$dgv.DefaultCellStyle.BackColor = $script:Colors.BgWhite
$dgv.DefaultCellStyle.ForeColor = $script:Colors.TextDark
$dgv.DefaultCellStyle.SelectionBackColor = $script:Colors.BlueLight
$dgv.DefaultCellStyle.SelectionForeColor = $script:Colors.Blue
$dgv.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dgv.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4)
$dgv.RowTemplate.Height = 34
$dgv.RowHeadersVisible = $false
$dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgv.MultiSelect = $false
$dgv.AllowUserToAddRows = $false
$dgv.AllowUserToDeleteRows = $false
$dgv.ReadOnly = $true
$dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

# Added Issuer column
@("Status", "Thumbprint", "Subject", "Issuer", "Expires", "Days", "Score", "Type") | ForEach-Object { $dgv.Columns.Add($_, $_) | Out-Null }
$dgv.Columns["Status"].FillWeight = 8
$dgv.Columns["Thumbprint"].FillWeight = 20
$dgv.Columns["Subject"].FillWeight = 18
$dgv.Columns["Issuer"].FillWeight = 16
$dgv.Columns["Expires"].FillWeight = 10
$dgv.Columns["Days"].FillWeight = 7
$dgv.Columns["Score"].FillWeight = 7
$dgv.Columns["Type"].FillWeight = 16

$form.Controls.Add($dgv)
#endregion

#region Options Panel
$optPanel = New-Object System.Windows.Forms.Panel
$optPanel.Location = New-Object System.Drawing.Point(24, 544)
$optPanel.Size = New-Object System.Drawing.Size(600, 36)
$optPanel.BackColor = $script:Colors.BgMain
$form.Controls.Add($optPanel)

$chkRemove = New-Object System.Windows.Forms.CheckBox
$chkRemove.Text = "Also remove old certificate from store"
$chkRemove.Location = New-Object System.Drawing.Point(0, 8)
$chkRemove.AutoSize = $true
$chkRemove.ForeColor = $script:Colors.TextDark
$optPanel.Controls.Add($chkRemove)

$chkAutoLog = New-Object System.Windows.Forms.CheckBox
$chkAutoLog.Text = "Auto-export log after operation"
$chkAutoLog.Location = New-Object System.Drawing.Point(280, 8)
$chkAutoLog.AutoSize = $true
$chkAutoLog.ForeColor = $script:Colors.TextDark
$chkAutoLog.Checked = $true
$optPanel.Controls.Add($chkAutoLog)
#endregion

#region Operation Log
$lblLogTitle = New-Object System.Windows.Forms.Label
$lblLogTitle.Text = "Operation Log"
$lblLogTitle.Location = New-Object System.Drawing.Point(24, 591)
$lblLogTitle.AutoSize = $true
$lblLogTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$lblLogTitle.ForeColor = $script:Colors.TextDark
$form.Controls.Add($lblLogTitle)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(24, 616)
$txtLog.Size = New-Object System.Drawing.Size(950, 110)
$txtLog.BackColor = $script:Colors.BgInput
$txtLog.ForeColor = $script:Colors.TextDark
$txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)
#endregion

#region Buttons
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh"
$btnRefresh.Location = New-Object System.Drawing.Point(540, 741)
$btnRefresh.Size = New-Object System.Drawing.Size(100, 36)
$btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefresh.FlatAppearance.BorderColor = $script:Colors.Border
$btnRefresh.BackColor = $script:Colors.BgWhite
$btnRefresh.ForeColor = $script:Colors.TextDark
$btnRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnRefresh.Add_MouseEnter({ $this.BackColor = $script:Colors.GrayLight })
$btnRefresh.Add_MouseLeave({ $this.BackColor = $script:Colors.BgWhite })
$form.Controls.Add($btnRefresh)

$btnSelect = New-Object System.Windows.Forms.Button
$btnSelect.Text = "Use Selected"
$btnSelect.Location = New-Object System.Drawing.Point(650, 741)
$btnSelect.Size = New-Object System.Drawing.Size(110, 36)
$btnSelect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSelect.FlatAppearance.BorderSize = 0
$btnSelect.BackColor = $script:Colors.Orange
$btnSelect.ForeColor = $script:Colors.TextLight
$btnSelect.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSelect.Enabled = $false
$btnSelect.Add_MouseEnter({ $this.BackColor = $script:Colors.OrangeHover })
$btnSelect.Add_MouseLeave({ $this.BackColor = $script:Colors.Orange })
$form.Controls.Add($btnSelect)

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text = "Apply Fix"
$btnApply.Location = New-Object System.Drawing.Point(770, 741)
$btnApply.Size = New-Object System.Drawing.Size(110, 36)
$btnApply.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnApply.FlatAppearance.BorderSize = 0
$btnApply.BackColor = $script:Colors.Green
$btnApply.ForeColor = $script:Colors.TextLight
$btnApply.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$btnApply.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnApply.Enabled = $false
$btnApply.Add_MouseEnter({ $this.BackColor = $script:Colors.GreenHover })
$btnApply.Add_MouseLeave({ $this.BackColor = $script:Colors.Green })
$form.Controls.Add($btnApply)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(890, 741)
$btnClose.Size = New-Object System.Drawing.Size(84, 36)
$btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClose.FlatAppearance.BorderColor = $script:Colors.Border
$btnClose.BackColor = $script:Colors.BgWhite
$btnClose.ForeColor = $script:Colors.TextDark
$btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClose.Add_MouseEnter({ $this.BackColor = $script:Colors.RedLight; $this.ForeColor = $script:Colors.RedText })
$btnClose.Add_MouseLeave({ $this.BackColor = $script:Colors.BgWhite; $this.ForeColor = $script:Colors.TextDark })
$form.Controls.Add($btnClose)
#endregion

#region Global Variables
$script:CurCert = $null
$script:BestCert = $null
$script:AllCerts = @()
$script:Scores = @{}
$script:NeedsChange = $false
$script:LogContent = New-Object System.Text.StringBuilder
#endregion

#region Helper Functions
function Get-IssuerCN {
    param([string]$Issuer)
    if ($Issuer -match 'CN=([^,]+)') { return $matches[1] }
    return $Issuer
}

function Get-EKUList {
    param($Cert)
    $ekus = @()
    if ($Cert.EnhancedKeyUsageList.FriendlyName -contains "KDC Authentication") { $ekus += "KDC Auth" }
    if ($Cert.EnhancedKeyUsageList.FriendlyName -contains "Smart Card Logon") { $ekus += "Smart Card" }
    if ($Cert.EnhancedKeyUsageList.FriendlyName -contains "Server Authentication") { $ekus += "Server Auth" }
    if ($Cert.EnhancedKeyUsageList.FriendlyName -contains "Client Authentication") { $ekus += "Client Auth" }
    return ($ekus -join ", ")
}

function Update-Badge {
    param($Label, [string]$Status)
    switch ($Status) {
        "VALID"   { $Label.Text = "VALID"; $Label.BackColor = $script:Colors.GreenLight; $Label.ForeColor = $script:Colors.GreenText }
        "WARNING" { $Label.Text = "EXPIRING"; $Label.BackColor = $script:Colors.OrangeLight; $Label.ForeColor = $script:Colors.OrangeText }
        "EXPIRED" { $Label.Text = "EXPIRED"; $Label.BackColor = $script:Colors.RedLight; $Label.ForeColor = $script:Colors.RedText }
        default   { $Label.Text = "UNKNOWN"; $Label.BackColor = $script:Colors.GrayLight; $Label.ForeColor = $script:Colors.Gray }
    }
}

function Get-CertStatus {
    param([datetime]$Expire, [int]$Warn = 30)
    $days = [math]::Round(($Expire - (Get-Date)).TotalDays)
    if ($days -lt 0) { return @{ Status = "EXPIRED"; Days = $days; Color = $script:Colors.Red } }
    elseif ($days -le $Warn) { return @{ Status = "WARNING"; Days = $days; Color = $script:Colors.Orange } }
    else { return @{ Status = "VALID"; Days = $days; Color = $script:Colors.Green } }
}

function Write-Log {
    param([string]$Msg, [string]$Type = "INFO")
    $ts = Get-Date -Format "HH:mm:ss"
    $line = "[$ts] $Msg"
    $script:LogContent.AppendLine($line) | Out-Null
    
    $color = switch ($Type) {
        "OK"   { $script:Colors.GreenText }
        "WARN" { $script:Colors.OrangeText }
        "ERR"  { $script:Colors.RedText }
        default { $script:Colors.TextMuted }
    }
    
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.SelectionColor = $color
    $txtLog.AppendText("$line`r`n")
    $txtLog.ScrollToCaret()
}

function Save-Log {
    param([string]$Prefix = "LDAPS")
    $file = Join-Path $script:LogDirectory "${Prefix}_${env:COMPUTERNAME}_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $content = @"
================================================================================
                    LDAPS Certificate Manager - Operation Log
================================================================================
Server:     $env:COMPUTERNAME
Date:       $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
User:       $env:USERNAME
================================================================================

$($script:LogContent.ToString())

================================================================================
                              End of Log
================================================================================
"@
    $content | Out-File $file -Encoding UTF8
    return $file
}

function Refresh-Data {
    $txtLog.Clear()
    $script:LogContent.Clear() | Out-Null
    $dgv.Rows.Clear()
    $btnApply.Enabled = $false
    $btnSelect.Enabled = $false
    $script:NeedsChange = $false
    
    Write-Log "============================================" "INFO"
    Write-Log "Starting LDAPS Certificate Scan" "INFO"
    Write-Log "============================================" "INFO"
    Write-Log "Server: $env:COMPUTERNAME" "INFO"
    Write-Log "User: $env:USERNAME" "INFO"
    Write-Log "" "INFO"
    
    # Get Current LDAPS Certificate
    try {
        Write-Log "[Step 1] Connecting to LDAPS port 636..." "INFO"
        $tcp = New-Object System.Net.Sockets.TcpClient("localhost", 636)
        $ssl = New-Object System.Net.Security.SslStream($tcp.GetStream(), $false, { $true })
        $ssl.AuthenticateAsClient($env:COMPUTERNAME)
        $script:CurCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
        $ssl.Close(); $tcp.Close()
        
        $st = Get-CertStatus -Expire $script:CurCert.NotAfter
        $issuerCN = Get-IssuerCN -Issuer $script:CurCert.Issuer
        $ekuList = Get-EKUList -Cert $script:CurCert
        
        $txtCurThumb.Text = $script:CurCert.Thumbprint
        $txtCurSubj.Text = $script:CurCert.Subject
        $txtCurIssuer.Text = $issuerCN
        $txtCurExp.Text = $script:CurCert.NotAfter.ToString("yyyy-MM-dd HH:mm")
        $txtCurExp.ForeColor = $st.Color
        $lblDaysVal.Text = "$($st.Days) days"
        $lblDaysVal.ForeColor = $st.Color
        $progDays.Value = [Math]::Max(0, [Math]::Min(365, $st.Days))
        $lblCurEKU.Text = "EKU: $ekuList"
        Update-Badge -Label $badgeCur -Status $st.Status
        
        Write-Log "[OK] Current LDAPS certificate found" "OK"
        Write-Log "     Thumbprint: $($script:CurCert.Thumbprint)" "INFO"
        Write-Log "     Subject:    $($script:CurCert.Subject)" "INFO"
        Write-Log "     Issuer:     $issuerCN" "INFO"
        Write-Log "     Expires:    $($script:CurCert.NotAfter.ToString('yyyy-MM-dd HH:mm'))" "INFO"
        Write-Log "     Status:     $($st.Status) ($($st.Days) days left)" $(if($st.Status -eq "VALID"){"OK"}elseif($st.Status -eq "WARNING"){"WARN"}else{"ERR"})
        Write-Log "     EKU:        $ekuList" "INFO"
        Write-Log "" "INFO"
    }
    catch {
        Write-Log "[ERROR] LDAPS connection failed: $_" "ERR"
        $txtCurThumb.Text = "Connection Failed"
        Update-Badge -Label $badgeCur -Status "UNKNOWN"
    }
    
    # Scan Certificate Store
    Write-Log "[Step 2] Scanning certificate store..." "INFO"
    $exclude = @("Cloudbase", "WinRM", "IIS", "Web Server", "Code Signing", "Root CA")
    
    $script:AllCerts = Get-ChildItem Cert:\LocalMachine\My | Where-Object {
        $c = $_
        $ex = $false; foreach ($p in $exclude) { if ($c.Subject -match $p -or $c.Issuer -match $p) { $ex = $true; break } }
        $kdc = $c.EnhancedKeyUsageList.FriendlyName -contains "KDC Authentication"
        $dc = $c.Subject -match $env:COMPUTERNAME -or $c.DnsNameList.Unicode -match $env:COMPUTERNAME
        $srv = $c.EnhancedKeyUsageList.FriendlyName -contains "Server Authentication"
        $c.HasPrivateKey -and (-not $ex) -and ($kdc -or ($dc -and $srv))
    }
    
    Write-Log "[OK] Found $($script:AllCerts.Count) LDAPS-compatible certificate(s)" "OK"
    
    # Score certificates
    $script:Scores = @{}
    foreach ($c in $script:AllCerts) {
        $s = 0
        if ($c.EnhancedKeyUsageList.FriendlyName -contains "KDC Authentication") { $s += 100 }
        if ($c.EnhancedKeyUsageList.FriendlyName -contains "Smart Card Logon") { $s += 50 }
        if ($c.Subject -match $env:COMPUTERNAME) { $s += 30 }
        if ($c.Issuer -match "Trendyol|Enterprise|Corp|Issuing") { $s += 20 }
        $script:Scores[$c.Thumbprint] = $s
    }
    
    $script:AllCerts = $script:AllCerts | Where-Object { $_.NotAfter -gt (Get-Date) } | Sort-Object { $script:Scores[$_.Thumbprint] } -Descending
    $script:BestCert = $script:AllCerts | Select-Object -First 1
    
    Write-Log "" "INFO"
    Write-Log "[Step 3] Certificate analysis:" "INFO"
    
    # Populate Grid
    $idx = 1
    foreach ($c in $script:AllCerts) {
        $st = Get-CertStatus -Expire $c.NotAfter
        $cur = $script:CurCert -and ($c.Thumbprint -eq $script:CurCert.Thumbprint)
        $best = $script:BestCert -and ($c.Thumbprint -eq $script:BestCert.Thumbprint)
        $score = $script:Scores[$c.Thumbprint]
        $issuerCN = Get-IssuerCN -Issuer $c.Issuer
        
        $types = @()
        if ($c.EnhancedKeyUsageList.FriendlyName -contains "KDC Authentication") { $types += "KDC" }
        if ($c.EnhancedKeyUsageList.FriendlyName -contains "Smart Card Logon") { $types += "SC" }
        $typeStr = $types -join "+"
        
        $type = ""
        if ($cur -and $best) { $type = "CURRENT+BEST" }
        elseif ($cur) { $type = "CURRENT" }
        elseif ($best) { $type = "RECOMMENDED" }
        if ($typeStr) { $type = if ($type) { "$type [$typeStr]" } else { "[$typeStr]" } }
        
        $ri = $dgv.Rows.Add($st.Status, $c.Thumbprint, $c.Subject, $issuerCN, $c.NotAfter.ToString("yyyy-MM-dd"), $st.Days, $score, $type)
        $dgv.Rows[$ri].Cells["Status"].Style.ForeColor = $st.Color
        $dgv.Rows[$ri].Cells["Days"].Style.ForeColor = $st.Color
        $dgv.Rows[$ri].Cells["Score"].Style.ForeColor = $script:Colors.Blue
        if ($best) { $dgv.Rows[$ri].Cells["Type"].Style.ForeColor = $script:Colors.GreenText }
        if ($cur) { $dgv.Rows[$ri].DefaultCellStyle.BackColor = $script:Colors.BlueLight }
        
        Write-Log "     [$idx] $($c.Thumbprint.Substring(0,16))... | Score: $score | $($st.Status) | $issuerCN" $(if($cur -and $best){"OK"}elseif($best){"WARN"}else{"INFO"})
        $idx++
    }
    
    Write-Log "" "INFO"
    
    # Update Recommended panel
    if ($script:BestCert) {
        $bst = Get-CertStatus -Expire $script:BestCert.NotAfter
        $issuerCN = Get-IssuerCN -Issuer $script:BestCert.Issuer
        $ekuList = Get-EKUList -Cert $script:BestCert
        
        $txtRecThumb.Text = $script:BestCert.Thumbprint
        $txtRecSubj.Text = $script:BestCert.Subject
        $txtRecIssuer.Text = $issuerCN
        $txtRecExp.Text = $script:BestCert.NotAfter.ToString("yyyy-MM-dd HH:mm")
        $lblRecDaysVal.Text = "$($bst.Days) days"
        $lblRecEKU.Text = "EKU: $ekuList"
        Update-Badge -Label $badgeRec -Status $bst.Status
        
        if ($script:CurCert -and $script:CurCert.Thumbprint -ne $script:BestCert.Thumbprint) {
            $script:NeedsChange = $true
            $btnApply.Enabled = $true
            Write-Log "[!] ACTION REQUIRED: Certificate change recommended" "WARN"
            Write-Log "    Current cert expires in $((Get-CertStatus -Expire $script:CurCert.NotAfter).Days) days" "WARN"
            Write-Log "    Recommended cert valid for $($bst.Days) days" "OK"
        } else {
            Write-Log "[OK] Status OK: Already using best certificate" "OK"
        }
    }
    
    Write-Log "" "INFO"
    Write-Log "============================================" "INFO"
    Write-Log "Scan completed at $(Get-Date -Format 'HH:mm:ss')" "INFO"
    Write-Log "============================================" "INFO"
}

function Apply-Fix {
    if (-not $script:NeedsChange -or -not $script:BestCert) { return }
    
    $r = [System.Windows.Forms.MessageBox]::Show(
        "This will:`n`n1. Remove old certificate from NTDS Registry`n2. Restart NTDS service (brief interruption)`n3. Activate new certificate for LDAPS`n`nContinue?",
        "Confirm Certificate Change", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    
    if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { Write-Log "[X] Operation cancelled by user." "WARN"; return }
    
    $btnApply.Enabled = $false
    $btnRefresh.Enabled = $false
    $old = $script:CurCert.Thumbprint
    
    Write-Log "" "INFO"
    Write-Log "========================================================" "WARN"
    Write-Log "     STARTING CERTIFICATE ROTATION" "WARN"
    Write-Log "========================================================" "INFO"
    Write-Log "Old Certificate: $old" "INFO"
    Write-Log "New Certificate: $($script:BestCert.Thumbprint)" "INFO"
    Write-Log "" "INFO"
    
    try {
        # Clean NTDS Registry
        Write-Log "[Step 1/4] Cleaning NTDS Registry..." "INFO"
        $path = "HKLM:\SOFTWARE\Microsoft\Cryptography\Services\NTDS\SystemCertificates\My\Certificates"
        if (Test-Path $path) {
            $removed = 0
            Get-ChildItem $path | Where-Object { $_.PSChildName -ne $script:BestCert.Thumbprint } | ForEach-Object {
                Remove-Item $_.PSPath -Recurse -Force
                Write-Log "   Removed: $($_.PSChildName)" "OK"
                $removed++
            }
            Write-Log "[OK] Removed $removed certificate(s) from NTDS Registry" "OK"
        }
        
        # Restart NTDS
        Write-Log "" "INFO"
        Write-Log "[Step 2/4] Restarting NTDS service..." "WARN"
        Write-Log "   Stopping NTDS..." "INFO"
        Stop-Service NTDS -Force
        Start-Sleep 3
        Write-Log "   Starting NTDS..." "INFO"
        Start-Service NTDS
        Start-Sleep 5
        Write-Log "[OK] NTDS restarted successfully" "OK"
        
        # Remove old cert from store if requested
        if ($chkRemove.Checked -and $old -and $old -ne $script:BestCert.Thumbprint) {
            Write-Log "" "INFO"
            Write-Log "[Step 3/4] Removing old certificate from store..." "INFO"
            $op = "Cert:\LocalMachine\My\$old"
            if (Test-Path $op) { 
                Remove-Item $op -Force
                Write-Log "[OK] Old certificate removed from Certificate Store" "OK"
            }
        } else {
            Write-Log "[Step 3/4] Skipped (old cert removal not selected)" "INFO"
        }
        
        # Verify new certificate
        Write-Log "" "INFO"
        Write-Log "[Step 4/4] Verifying new certificate..." "INFO"
        Start-Sleep 2
        $tcp2 = New-Object System.Net.Sockets.TcpClient("localhost", 636)
        $ssl2 = New-Object System.Net.Security.SslStream($tcp2.GetStream(), $false, { $true })
        $ssl2.AuthenticateAsClient($env:COMPUTERNAME)
        $new = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl2.RemoteCertificate)
        $ssl2.Close(); $tcp2.Close()
        
        if ($new.Thumbprint -eq $script:BestCert.Thumbprint) {
            Write-Log "[OK] Verification successful!" "OK"
            Write-Log "    New Thumbprint: $($new.Thumbprint)" "OK"
            Write-Log "" "INFO"
            Write-Log "========================================================" "OK"
            Write-Log "     OPERATION COMPLETED SUCCESSFULLY" "OK"
            Write-Log "========================================================" "OK"
            
            if ($chkAutoLog.Checked) { 
                $f = Save-Log -Prefix "LDAPS_Success"
                Write-Log "" "INFO"
                Write-Log "Log exported to: $f" "INFO"
            }
            [System.Windows.Forms.MessageBox]::Show("Certificate changed successfully!`n`nNew Thumbprint:`n$($new.Thumbprint)", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            Write-Log "[!] WARNING: Verification failed - unexpected certificate" "ERR"
            Write-Log "    Expected: $($script:BestCert.Thumbprint)" "ERR"
            Write-Log "    Got:      $($new.Thumbprint)" "ERR"
        }
        Refresh-Data
    }
    catch {
        Write-Log "" "ERR"
        Write-Log "[ERROR] Operation failed: $($_.Exception.Message)" "ERR"
        if ($chkAutoLog.Checked) { Save-Log -Prefix "LDAPS_Error" }
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally { $btnRefresh.Enabled = $true }
}
#endregion

#region Event Handlers
$btnRefresh.Add_Click({ Refresh-Data })
$btnApply.Add_Click({ Apply-Fix })
$btnClose.Add_Click({ $form.Close() })

$dgv.Add_SelectionChanged({
    if ($dgv.SelectedRows.Count -gt 0) {
        $sel = $dgv.SelectedRows[0].Cells["Thumbprint"].Value
        $btnSelect.Enabled = $script:CurCert -and $sel -ne $script:CurCert.Thumbprint
    }
})

$btnSelect.Add_Click({
    if ($dgv.SelectedRows.Count -gt 0) {
        $sel = $dgv.SelectedRows[0].Cells["Thumbprint"].Value
        $script:BestCert = $script:AllCerts | Where-Object { $_.Thumbprint -eq $sel }
        if ($script:BestCert) {
            $bst = Get-CertStatus -Expire $script:BestCert.NotAfter
            $issuerCN = Get-IssuerCN -Issuer $script:BestCert.Issuer
            $ekuList = Get-EKUList -Cert $script:BestCert
            
            $txtRecThumb.Text = $script:BestCert.Thumbprint
            $txtRecSubj.Text = $script:BestCert.Subject
            $txtRecIssuer.Text = $issuerCN
            $txtRecExp.Text = $script:BestCert.NotAfter.ToString("yyyy-MM-dd HH:mm")
            $lblRecDaysVal.Text = "$($bst.Days) days"
            $lblRecEKU.Text = "EKU: $ekuList"
            
            if ($script:CurCert -and $script:CurCert.Thumbprint -ne $script:BestCert.Thumbprint) {
                $script:NeedsChange = $true
                $btnApply.Enabled = $true
                Write-Log "[i] User selected certificate: $($sel.Substring(0,20))..." "INFO"
            }
            $btnSelect.Enabled = $false
        }
    }
})

$form.Add_Shown({ Refresh-Data })
#endregion

[System.Windows.Forms.Application]::EnableVisualStyles()
$form.ShowDialog() | Out-Null
