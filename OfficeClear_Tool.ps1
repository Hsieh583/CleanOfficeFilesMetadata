# OfficeClear_Tool.ps1
# Version 2.0
# Office File Metadata Cleaner Tool

# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Constants for Office COM operations
$script:RemoveAllDocumentInfo = 1
$script:AlertsOff = 0

# Global variables
$script:CancelRequested = $false
$script:ProcessedFiles = @()
$script:FailedFiles = @()
$script:WordApp = $null
$script:ExcelApp = $null
$script:PowerPointApp = $null

# Function to write log with color coding
function Write-Log {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $symbol = switch ($Type) {
        "Success" { "[PASS]" }
        "Error" { "[FAIL]" }
        "Info" { "[INFO]" }
        default { "[INFO]" }
    }
    
    $logMessage = "$timestamp $symbol $Message"
    $logTextBox.AppendText("$logMessage`r`n")
    $logTextBox.SelectionStart = $logTextBox.Text.Length
    $logTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Function to initialize Office application
function Initialize-OfficeApp {
    param(
        [string]$AppType
    )
    
    try {
        switch ($AppType) {
            "Word" {
                if ($null -eq $script:WordApp) {
                    $script:WordApp = New-Object -ComObject Word.Application
                    $script:WordApp.Visible = $false
                    $script:WordApp.DisplayAlerts = $script:AlertsOff
                    $script:WordApp.ScreenUpdating = $false
                }
                return $script:WordApp
            }
            "Excel" {
                if ($null -eq $script:ExcelApp) {
                    $script:ExcelApp = New-Object -ComObject Excel.Application
                    $script:ExcelApp.Visible = $false
                    $script:ExcelApp.DisplayAlerts = $false
                    $script:ExcelApp.ScreenUpdating = $false
                }
                return $script:ExcelApp
            }
            "PowerPoint" {
                if ($null -eq $script:PowerPointApp) {
                    $script:PowerPointApp = New-Object -ComObject PowerPoint.Application
                }
                return $script:PowerPointApp
            }
        }
    }
    catch {
        Write-Log "Failed to initialize $AppType application: $($_.Exception.Message)" "Error"
        return $null
    }
}

# Function to clean metadata from Word document
function Clean-WordMetadata {
    param(
        [string]$FilePath,
        [array]$MetadataItems
    )
    
    $app = Initialize-OfficeApp -AppType "Word"
    if ($null -eq $app) {
        throw "Word application not available"
    }
    
    $doc = $app.Documents.Open($FilePath)
    
    try {
        foreach ($item in $MetadataItems) {
            try {
                switch ($item) {
                    "Title" { $doc.BuiltInDocumentProperties.Item("Title").Value = "" }
                    "Subject" { $doc.BuiltInDocumentProperties.Item("Subject").Value = "" }
                    "Author" { $doc.BuiltInDocumentProperties.Item("Author").Value = "" }
                    "Manager" { $doc.BuiltInDocumentProperties.Item("Manager").Value = "" }
                    "Company" { $doc.BuiltInDocumentProperties.Item("Company").Value = "" }
                    "Category" { $doc.BuiltInDocumentProperties.Item("Category").Value = "" }
                    "Keywords" { $doc.BuiltInDocumentProperties.Item("Keywords").Value = "" }
                    "Comments" { $doc.BuiltInDocumentProperties.Item("Comments").Value = "" }
                }
            }
            catch {
                # Some properties might not exist or be read-only, continue
            }
        }
        
        $doc.RemoveDocumentInformation($script:RemoveAllDocumentInfo)
        $doc.Save()
    }
    finally {
        $doc.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
}

# Function to clean metadata from Excel workbook
function Clean-ExcelMetadata {
    param(
        [string]$FilePath,
        [array]$MetadataItems
    )
    
    $app = Initialize-OfficeApp -AppType "Excel"
    if ($null -eq $app) {
        throw "Excel application not available"
    }
    
    $workbook = $app.Workbooks.Open($FilePath)
    
    try {
        foreach ($item in $MetadataItems) {
            try {
                switch ($item) {
                    "Title" { $workbook.BuiltInDocumentProperties.Item("Title").Value = "" }
                    "Subject" { $workbook.BuiltInDocumentProperties.Item("Subject").Value = "" }
                    "Author" { $workbook.BuiltInDocumentProperties.Item("Author").Value = "" }
                    "Manager" { $workbook.BuiltInDocumentProperties.Item("Manager").Value = "" }
                    "Company" { $workbook.BuiltInDocumentProperties.Item("Company").Value = "" }
                    "Category" { $workbook.BuiltInDocumentProperties.Item("Category").Value = "" }
                    "Keywords" { $workbook.BuiltInDocumentProperties.Item("Keywords").Value = "" }
                    "Comments" { $workbook.BuiltInDocumentProperties.Item("Comments").Value = "" }
                }
            }
            catch {
            }
        }
        
        $workbook.RemoveDocumentInformation($script:RemoveAllDocumentInfo)
        $workbook.Save()
    }
    finally {
        $workbook.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
}

# Function to clean metadata from PowerPoint presentation
function Clean-PowerPointMetadata {
    param(
        [string]$FilePath,
        [array]$MetadataItems
    )
    
    $app = Initialize-OfficeApp -AppType "PowerPoint"
    if ($null -eq $app) {
        throw "PowerPoint application not available"
    }
    
    $ReadOnly = $false
    $Untitled = $false  
    $WithWindow = $false
    
    $presentation = $app.Presentations.Open($FilePath, $ReadOnly, $Untitled, $WithWindow)
    
    try {
        foreach ($item in $MetadataItems) {
            try {
                switch ($item) {
                    "Title" { $presentation.BuiltInDocumentProperties.Item("Title").Value = "" }
                    "Subject" { $presentation.BuiltInDocumentProperties.Item("Subject").Value = "" }
                    "Author" { $presentation.BuiltInDocumentProperties.Item("Author").Value = "" }
                    "Manager" { $presentation.BuiltInDocumentProperties.Item("Manager").Value = "" }
                    "Company" { $presentation.BuiltInDocumentProperties.Item("Company").Value = "" }
                    "Category" { $presentation.BuiltInDocumentProperties.Item("Category").Value = "" }
                    "Keywords" { $presentation.BuiltInDocumentProperties.Item("Keywords").Value = "" }
                    "Comments" { $presentation.BuiltInDocumentProperties.Item("Comments").Value = "" }
                }
            }
            catch {
            }
        }
        
        $presentation.RemoveDocumentInformation($script:RemoveAllDocumentInfo)
        $presentation.Save()
    }
    finally {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
}

# Function to process a single file
function Process-File {
    param(
        [string]$FilePath,
        [array]$MetadataItems
    )
    
    try {
        $fileStream = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
        $fileStream.Close()
        
        $fileInfo = Get-Item $FilePath
        $originalCreationTime = $fileInfo.CreationTime
        $originalLastWriteTime = $fileInfo.LastWriteTime
        
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        
        switch ($extension) {
            {$_ -in ".doc", ".docx"} {
                Clean-WordMetadata -FilePath $FilePath -MetadataItems $MetadataItems
            }
            {$_ -in ".xls", ".xlsx"} {
                Clean-ExcelMetadata -FilePath $FilePath -MetadataItems $MetadataItems
            }
            {$_ -in ".ppt", ".pptx"} {
                Clean-PowerPointMetadata -FilePath $FilePath -MetadataItems $MetadataItems
            }
        }
        
        $fileInfo = Get-Item $FilePath
        $fileInfo.CreationTime = $originalCreationTime
        $fileInfo.LastWriteTime = $originalLastWriteTime
        
        return @{ Success = $true; Error = "" }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

# Function to cleanup Office applications
function Cleanup-OfficeApps {
    try {
        if ($null -ne $script:WordApp) {
            $script:WordApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:WordApp) | Out-Null
            $script:WordApp = $null
        }
        
        if ($null -ne $script:ExcelApp) {
            $script:ExcelApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:ExcelApp) | Out-Null
            $script:ExcelApp = $null
        }
        
        if ($null -ne $script:PowerPointApp) {
            $script:PowerPointApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:PowerPointApp) | Out-Null
            $script:PowerPointApp = $null
        }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    catch {
        Write-Log "Error cleaning up Office applications: $($_.Exception.Message)" "Error"
    }
}

# Function to generate CSV report
function Generate-Report {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path (Get-Location) "Cleanup_Report_$timestamp.csv"
    
    try {
        $report = @()
        
        foreach ($file in $script:ProcessedFiles) {
            $report += [PSCustomObject]@{
                "File Path" = $file
                "Status" = "Success"
                "Error Message" = ""
            }
        }
        
        foreach ($file in $script:FailedFiles) {
            $report += [PSCustomObject]@{
                "File Path" = $file.Path
                "Status" = "Failed"
                "Error Message" = $file.Error
            }
        }
        
        $report | Export-Csv -Path $reportPath -Encoding UTF8 -NoTypeInformation
        
        return $reportPath
    }
    catch {
        Write-Log "Failed to generate report: $($_.Exception.Message)" "Error"
        return $null
    }
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Office Metadata Cleaner v2.0"
$form.Size = New-Object System.Drawing.Size(800, 650)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Create tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(765, 540)
$form.Controls.Add($tabControl)

# ===== Tab 1: Settings =====
$tabSettings = New-Object System.Windows.Forms.TabPage
$tabSettings.Text = "Settings"
$tabControl.Controls.Add($tabSettings)

# Folder selection
$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Location = New-Object System.Drawing.Point(10, 15)
$lblFolder.Size = New-Object System.Drawing.Size(100, 20)
$lblFolder.Text = "Target Folder:"
$tabSettings.Controls.Add($lblFolder)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(110, 12)
$txtFolder.Size = New-Object System.Drawing.Size(530, 20)
$tabSettings.Controls.Add($txtFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(650, 10)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 25)
$btnBrowse.Text = "Browse..."
$btnBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select folder to process"
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolder.Text = $folderBrowser.SelectedPath
    }
})
$tabSettings.Controls.Add($btnBrowse)

# File types group
$grpFileTypes = New-Object System.Windows.Forms.GroupBox
$grpFileTypes.Location = New-Object System.Drawing.Point(10, 50)
$grpFileTypes.Size = New-Object System.Drawing.Size(360, 120)
$grpFileTypes.Text = "File Types"
$tabSettings.Controls.Add($grpFileTypes)

$chkWord = New-Object System.Windows.Forms.CheckBox
$chkWord.Location = New-Object System.Drawing.Point(20, 25)
$chkWord.Size = New-Object System.Drawing.Size(300, 20)
$chkWord.Text = "Word (.doc, .docx)"
$chkWord.Checked = $true
$grpFileTypes.Controls.Add($chkWord)

$chkExcel = New-Object System.Windows.Forms.CheckBox
$chkExcel.Location = New-Object System.Drawing.Point(20, 55)
$chkExcel.Size = New-Object System.Drawing.Size(300, 20)
$chkExcel.Text = "Excel (.xls, .xlsx)"
$chkExcel.Checked = $true
$grpFileTypes.Controls.Add($chkExcel)

$chkPowerPoint = New-Object System.Windows.Forms.CheckBox
$chkPowerPoint.Location = New-Object System.Drawing.Point(20, 85)
$chkPowerPoint.Size = New-Object System.Drawing.Size(300, 20)
$chkPowerPoint.Text = "PowerPoint (.ppt, .pptx)"
$chkPowerPoint.Checked = $true
$grpFileTypes.Controls.Add($chkPowerPoint)

# Metadata items group
$grpMetadata = New-Object System.Windows.Forms.GroupBox
$grpMetadata.Location = New-Object System.Drawing.Point(10, 180)
$grpMetadata.Size = New-Object System.Drawing.Size(360, 280)
$grpMetadata.Text = "Metadata Items"
$tabSettings.Controls.Add($grpMetadata)

$metadataCheckboxes = @{}

$chkTitle = New-Object System.Windows.Forms.CheckBox
$chkTitle.Location = New-Object System.Drawing.Point(20, 25)
$chkTitle.Size = New-Object System.Drawing.Size(150, 20)
$chkTitle.Text = "Title"
$chkTitle.Checked = $true
$grpMetadata.Controls.Add($chkTitle)
$metadataCheckboxes["Title"] = $chkTitle

$chkSubject = New-Object System.Windows.Forms.CheckBox
$chkSubject.Location = New-Object System.Drawing.Point(20, 55)
$chkSubject.Size = New-Object System.Drawing.Size(150, 20)
$chkSubject.Text = "Subject"
$chkSubject.Checked = $true
$grpMetadata.Controls.Add($chkSubject)
$metadataCheckboxes["Subject"] = $chkSubject

$chkAuthor = New-Object System.Windows.Forms.CheckBox
$chkAuthor.Location = New-Object System.Drawing.Point(20, 85)
$chkAuthor.Size = New-Object System.Drawing.Size(150, 20)
$chkAuthor.Text = "Author"
$chkAuthor.Checked = $true
$grpMetadata.Controls.Add($chkAuthor)
$metadataCheckboxes["Author"] = $chkAuthor

$chkManager = New-Object System.Windows.Forms.CheckBox
$chkManager.Location = New-Object System.Drawing.Point(20, 115)
$chkManager.Size = New-Object System.Drawing.Size(150, 20)
$chkManager.Text = "Manager"
$chkManager.Checked = $true
$grpMetadata.Controls.Add($chkManager)
$metadataCheckboxes["Manager"] = $chkManager

$chkCompany = New-Object System.Windows.Forms.CheckBox
$chkCompany.Location = New-Object System.Drawing.Point(190, 25)
$chkCompany.Size = New-Object System.Drawing.Size(150, 20)
$chkCompany.Text = "Company"
$chkCompany.Checked = $true
$grpMetadata.Controls.Add($chkCompany)
$metadataCheckboxes["Company"] = $chkCompany

$chkCategory = New-Object System.Windows.Forms.CheckBox
$chkCategory.Location = New-Object System.Drawing.Point(190, 55)
$chkCategory.Size = New-Object System.Drawing.Size(150, 20)
$chkCategory.Text = "Category"
$chkCategory.Checked = $true
$grpMetadata.Controls.Add($chkCategory)
$metadataCheckboxes["Category"] = $chkCategory

$chkKeywords = New-Object System.Windows.Forms.CheckBox
$chkKeywords.Location = New-Object System.Drawing.Point(190, 85)
$chkKeywords.Size = New-Object System.Drawing.Size(150, 20)
$chkKeywords.Text = "Keywords"
$chkKeywords.Checked = $true
$grpMetadata.Controls.Add($chkKeywords)
$metadataCheckboxes["Keywords"] = $chkKeywords

$chkComments = New-Object System.Windows.Forms.CheckBox
$chkComments.Location = New-Object System.Drawing.Point(190, 115)
$chkComments.Size = New-Object System.Drawing.Size(150, 20)
$chkComments.Text = "Comments"
$chkComments.Checked = $true
$grpMetadata.Controls.Add($chkComments)
$metadataCheckboxes["Comments"] = $chkComments

# Select all / Deselect all buttons
$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Location = New-Object System.Drawing.Point(20, 150)
$btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
$btnSelectAll.Text = "Select All"
$btnSelectAll.Add_Click({
    foreach ($checkbox in $metadataCheckboxes.Values) {
        $checkbox.Checked = $true
    }
})
$grpMetadata.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Location = New-Object System.Drawing.Point(130, 150)
$btnDeselectAll.Size = New-Object System.Drawing.Size(100, 30)
$btnDeselectAll.Text = "Deselect All"
$btnDeselectAll.Add_Click({
    foreach ($checkbox in $metadataCheckboxes.Values) {
        $checkbox.Checked = $false
    }
})
$grpMetadata.Controls.Add($btnDeselectAll)

# Processing options
$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Location = New-Object System.Drawing.Point(390, 50)
$grpOptions.Size = New-Object System.Drawing.Size(360, 80)
$grpOptions.Text = "Processing Options"
$tabSettings.Controls.Add($grpOptions)

$chkIncludeSubfolders = New-Object System.Windows.Forms.CheckBox
$chkIncludeSubfolders.Location = New-Object System.Drawing.Point(20, 30)
$chkIncludeSubfolders.Size = New-Object System.Drawing.Size(300, 20)
$chkIncludeSubfolders.Text = "Include Subfolders"
$chkIncludeSubfolders.Checked = $true
$grpOptions.Controls.Add($chkIncludeSubfolders)

# ===== Tab 2: Progress =====
$tabProgress = New-Object System.Windows.Forms.TabPage
$tabProgress.Text = "Progress"
$tabControl.Controls.Add($tabProgress)

# Status label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(10, 15)
$lblStatus.Size = New-Object System.Drawing.Size(730, 20)
$lblStatus.Text = "Ready"
$tabProgress.Controls.Add($lblStatus)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 45)
$progressBar.Size = New-Object System.Drawing.Size(730, 25)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$tabProgress.Controls.Add($progressBar)

# Log textbox
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 80)
$logTextBox.Size = New-Object System.Drawing.Size(730, 440)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = "Vertical"
$logTextBox.ReadOnly = $true
$logTextBox.BackColor = [System.Drawing.Color]::Black
$logTextBox.ForeColor = [System.Drawing.Color]::Lime
$logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabProgress.Controls.Add($logTextBox)

# ===== Bottom buttons =====
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Location = New-Object System.Drawing.Point(10, 560)
$btnStart.Size = New-Object System.Drawing.Size(150, 40)
$btnStart.Text = "Start"
$btnStart.BackColor = [System.Drawing.Color]::LightGreen
$btnStart.Font = New-Object System.Drawing.Font($btnStart.Font.FontFamily, 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Location = New-Object System.Drawing.Point(170, 560)
$btnCancel.Size = New-Object System.Drawing.Size(150, 40)
$btnCancel.Text = "Cancel"
$btnCancel.BackColor = [System.Drawing.Color]::LightCoral
$btnCancel.Font = New-Object System.Drawing.Font($btnCancel.Font.FontFamily, 10, [System.Drawing.FontStyle]::Bold)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$btnOpenReport = New-Object System.Windows.Forms.Button
$btnOpenReport.Location = New-Object System.Drawing.Point(330, 560)
$btnOpenReport.Size = New-Object System.Drawing.Size(150, 40)
$btnOpenReport.Text = "Open Report"
$btnOpenReport.BackColor = [System.Drawing.Color]::LightBlue
$btnOpenReport.Font = New-Object System.Drawing.Font($btnOpenReport.Font.FontFamily, 10, [System.Drawing.FontStyle]::Bold)
$btnOpenReport.Enabled = $false
$form.Controls.Add($btnOpenReport)

# Global variable for report path
$script:ReportPath = $null

# Start button click event
$btnStart.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtFolder.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a target folder", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not (Test-Path $txtFolder.Text)) {
        [System.Windows.Forms.MessageBox]::Show("The specified path does not exist", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not ($chkWord.Checked -or $chkExcel.Checked -or $chkPowerPoint.Checked)) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one file type", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $hasSelectedMetadata = $false
    foreach ($checkbox in $metadataCheckboxes.Values) {
        if ($checkbox.Checked) {
            $hasSelectedMetadata = $true
            break
        }
    }
    
    if (-not $hasSelectedMetadata) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one metadata item", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $btnStart.Enabled = $false
    $btnCancel.Enabled = $true
    $btnOpenReport.Enabled = $false
    $tabSettings.Enabled = $false
    $tabControl.SelectedTab = $tabProgress
    
    $script:CancelRequested = $false
    $script:ProcessedFiles = @()
    $script:FailedFiles = @()
    $logTextBox.Clear()
    $progressBar.Value = 0
    
    $fileExtensions = @()
    if ($chkWord.Checked) { $fileExtensions += "*.doc", "*.docx" }
    if ($chkExcel.Checked) { $fileExtensions += "*.xls", "*.xlsx" }
    if ($chkPowerPoint.Checked) { $fileExtensions += "*.ppt", "*.pptx" }
    
    $selectedMetadata = @()
    foreach ($key in $metadataCheckboxes.Keys) {
        if ($metadataCheckboxes[$key].Checked) {
            $selectedMetadata += $key
        }
    }
    
    Write-Log "Starting file scan..." "Info"
    
    $files = @()
    foreach ($extension in $fileExtensions) {
        if ($chkIncludeSubfolders.Checked) {
            $files += Get-ChildItem -Path $txtFolder.Text -Filter $extension -Recurse -File -ErrorAction SilentlyContinue
        } else {
            $files += Get-ChildItem -Path $txtFolder.Text -Filter $extension -File -ErrorAction SilentlyContinue
        }
    }
    
    if ($files.Count -eq 0) {
        Write-Log "No matching files found" "Error"
        [System.Windows.Forms.MessageBox]::Show("No matching files found", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $btnStart.Enabled = $true
        $btnCancel.Enabled = $false
        $tabSettings.Enabled = $true
        return
    }
    
    Write-Log "Found $($files.Count) files" "Info"
    Write-Log "Starting metadata cleanup..." "Info"
    
    $totalFiles = $files.Count
    $currentFile = 0
    
    foreach ($file in $files) {
        if ($script:CancelRequested) {
            Write-Log "User cancelled operation" "Info"
            break
        }
        
        $currentFile++
        $lblStatus.Text = "Processing: $($file.Name) ($currentFile/$totalFiles)"
        $progressBar.Value = [int](($currentFile / $totalFiles) * 100)
        [System.Windows.Forms.Application]::DoEvents()
        
        Write-Log "Processing file: $($file.FullName)" "Info"
        
        $result = Process-File -FilePath $file.FullName -MetadataItems $selectedMetadata
        
        if ($result.Success) {
            Write-Log "Successfully cleaned: $($file.Name)" "Success"
            $script:ProcessedFiles += $file.FullName
        } else {
            Write-Log "Failed: $($file.Name) - $($result.Error)" "Error"
            $script:FailedFiles += @{
                Path = $file.FullName
                Error = $result.Error
            }
        }
    }
    
    Write-Log "Cleaning up Office applications..." "Info"
    Cleanup-OfficeApps
    
    Write-Log "Generating report..." "Info"
    $script:ReportPath = Generate-Report
    
    if ($null -ne $script:ReportPath) {
        Write-Log "Report generated: $script:ReportPath" "Success"
        $btnOpenReport.Enabled = $true
    }
    
    Write-Log "====== Processing Complete ======" "Info"
    Write-Log "Successful: $($script:ProcessedFiles.Count) files" "Success"
    Write-Log "Failed: $($script:FailedFiles.Count) files" "Error"
    
    $lblStatus.Text = "Processing Complete"
    $progressBar.Value = 100
    
    $btnStart.Enabled = $true
    $btnCancel.Enabled = $false
    $tabSettings.Enabled = $true
    
    [System.Windows.Forms.MessageBox]::Show(
        "Processing Complete`nSuccessful: $($script:ProcessedFiles.Count)`nFailed: $($script:FailedFiles.Count)",
        "Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

# Cancel button click event
$btnCancel.Add_Click({
    $script:CancelRequested = $true
    $btnCancel.Enabled = $false
    Write-Log "Cancelling operation..." "Info"
})

# Open report button click event
$btnOpenReport.Add_Click({
    if ($null -ne $script:ReportPath -and (Test-Path $script:ReportPath)) {
        Start-Process $script:ReportPath
    }
})

# Form closing event - cleanup
$form.Add_FormClosing({
    Cleanup-OfficeApps
})

# Show the form
[void]$form.ShowDialog()
