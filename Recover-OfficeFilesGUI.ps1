# Recover-OfficeFilesGUI.ps1
# PowerShell script with Windows Forms GUI to find and recover unsaved Microsoft Office files
# Includes all Office file types, .wbk, new locations, Recover with date-appended filename and extension matching, File Type column, and dropdown

# Load Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-OfficeRecoveryFiles {
    param (
        [string]$UserProfile = $env:USERPROFILE,
        [bool]$ExcludeZeroByte = $false
    )

    try {
        # Define Office recovery and temp file paths with associated file types
        $pathFileTypes = @(
            @{ Path = "$UserProfile\AppData\Roaming\Microsoft\Word\"; FileType = "Word" },
            @{ Path = "$UserProfile\AppData\Local\Microsoft\Word\"; FileType = "Word" },
            @{ Path = "$UserProfile\AppData\Roaming\Microsoft\Excel\"; FileType = "Excel" },
            @{ Path = "$UserProfile\AppData\Roaming\Microsoft\PowerPoint\"; FileType = "PowerPoint" },
            @{ Path = "$UserProfile\AppData\Roaming\Microsoft\Access\"; FileType = "Access" },
            @{ Path = "$UserProfile\AppData\Roaming\Microsoft\Publisher\"; FileType = "Publisher" },
            @{ Path = "$UserProfile\AppData\Roaming\Microsoft\Visio\"; FileType = "Visio" },
            @{ Path = "$UserProfile\AppData\Local\Microsoft\Office\UnsavedFiles\"; FileType = "Unknown" },
            @{ Path = "$UserProfile\AppData\Local\Temp\"; FileType = "Unknown" },
            @{ Path = "shell:RecycleBinFolder"; FileType = "Unknown" }
        )

        # Define file extensions with associated file types
        $extensionFileTypes = @{
            ".asd" = "Word"; ".docx" = "Word"; ".doc" = "Word"; ".wbk" = "Word";
            ".xar" = "Excel"; ".xlsx" = "Excel"; ".xls" = "Excel";
            ".pptx" = "PowerPoint"; ".ppt" = "PowerPoint";
            ".accdb" = "Access"; ".mdb" = "Access";
            ".pub" = "Publisher";
            ".vsdx" = "Visio"; ".vsd" = "Visio";
            ".tmp" = "Unknown"
        }

        # Define extensions for search
        $extensions = $extensionFileTypes.Keys

        $files = @()
        foreach ($pathInfo in $pathFileTypes) {
            $path = $pathInfo.Path
            $defaultFileType = $pathInfo.FileType

            if ($path -eq "shell:RecycleBinFolder") {
                # Handle Recycle Bin via shell namespace
                $shell = New-Object -ComObject Shell.Application
                $recycleBin = $shell.NameSpace(0x0a) # Recycle Bin
                foreach ($ext in $extensions) {
                    $items = $recycleBin.Items() | Where-Object { $_.Name -like "*$ext" }
                    foreach ($item in $items) {
                        $filePath = $item.Path
                        $fileSize = [math]::Round($item.Size / 1KB, 2)
                        $fileModified = $item.ModifyDate
                        $fileExt = [System.IO.Path]::GetExtension($item.Name).ToLower()
                        $fileType = if ($extensionFileTypes[$fileExt] -eq $null) { "Unknown" } else { $extensionFileTypes[$fileExt] }
                        $files += [PSCustomObject]@{
                            Name = $item.Name
                            FullName = $filePath
                            SizeKB = $fileSize
                            LastWriteTime = $fileModified
                            FileType = $fileType
                        }
                    }
                }
            } elseif (Test-Path $path) {
                foreach ($ext in $extensions) {
                    $items = Get-ChildItem -Path $path -Filter "*$ext" -Recurse -ErrorAction SilentlyContinue
                    foreach ($item in $items) {
                        $fileType = if ($defaultFileType -eq "Unknown") {
                            if ($extensionFileTypes[$item.Extension.ToLower()] -eq $null) { "Unknown" } else { $extensionFileTypes[$item.Extension.ToLower()] }
                        } else { $defaultFileType }
                        $files += [PSCustomObject]@{
                            Name = $item.Name
                            FullName = $item.FullName
                            SizeKB = [math]::Round($item.Length / 1KB, 2)
                            LastWriteTime = $item.LastWriteTime
                            FileType = $fileType
                        }
                    }
                }
            }
        }

        # Filter out zero-byte files if requested
        if ($ExcludeZeroByte) {
            $files = $files | Where-Object { $_.SizeKB -gt 0 }
        }

        return $files | Sort-Object LastWriteTime -Descending
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error searching for files: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
}

function Initialize-GUI {
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Office File Recovery"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"

    # Create DataGridView to display files
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object System.Drawing.Point(10, 10)
    $dataGridView.Size = New-Object System.Drawing.Size(760, 400)
    $dataGridView.SelectionMode = "FullRowSelect"
    $dataGridView.MultiSelect = $false
    $dataGridView.ReadOnly = $true
    $dataGridView.ColumnCount = 5
    $dataGridView.Columns[0].Name = "File Name"
    $dataGridView.Columns[1].Name = "Path"
    $dataGridView.Columns[2].Name = "Size (KB)"
    $dataGridView.Columns[3].Name = "Last Modified"
    $dataGridView.Columns[4].Name = "File Type"
    $dataGridView.AutoSizeColumnsMode = "Fill"
    $form.Controls.Add($dataGridView)

    # Create buttons
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(10, 420)
    $refreshButton.Size = New-Object System.Drawing.Size(100, 30)
    $refreshButton.Text = "Refresh"
    $form.Controls.Add($refreshButton)

    $openButton = New-Object System.Windows.Forms.Button
    $openButton.Location = New-Object System.Drawing.Point(120, 420)
    $openButton.Size = New-Object System.Drawing.Size(100, 30)
    $openButton.Text = "Open File"
    $form.Controls.Add($openButton)

    $recoverButton = New-Object System.Windows.Forms.Button
    $recoverButton.Location = New-Object System.Drawing.Point(240, 420)
    $recoverButton.Size = New-Object System.Drawing.Size(120, 30)
    $recoverButton.Text = "Recover"
    $form.Controls.Add($recoverButton)

    # Create file type dropdown
    $fileTypeComboBox = New-Object System.Windows.Forms.ComboBox
    $fileTypeComboBox.Location = New-Object System.Drawing.Point(520, 420)
    $fileTypeComboBox.Size = New-Object System.Drawing.Size(100, 30)
    $fileTypeComboBox.Items.AddRange(@("Auto", "Word", "Excel", "PowerPoint", "Access", "Publisher", "Visio"))
    $fileTypeComboBox.SelectedIndex = 0 # Default to "Auto"
    $form.Controls.Add($fileTypeComboBox)

    # Create checkbox for excluding zero-byte files
    $excludeZeroCheckBox = New-Object System.Windows.Forms.CheckBox
    $excludeZeroCheckBox.Location = New-Object System.Drawing.Point(630, 420)
    $excludeZeroCheckBox.Size = New-Object System.Drawing.Size(150, 30)
    $excludeZeroCheckBox.Text = "Exclude Zero-Byte Files"
    $excludeZeroCheckBox.Checked = $true
    $form.Controls.Add($excludeZeroCheckBox)

    # Create status bar
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(10, 460)
    $statusLabel.Size = New-Object System.Drawing.Size(760, 20)
    $statusLabel.Text = "Ready"
    $form.Controls.Add($statusLabel)

    # Extension mapping for recovery
    $recoveryExtensions = @{
        "Word" = ".docx"
        "Excel" = ".xlsx"
        "PowerPoint" = ".pptx"
        "Access" = ".accdb"
        "Publisher" = ".pub"
        "Visio" = ".vsdx"
    }

    # Refresh file list function
    function Update-FileList {
        $dataGridView.Rows.Clear()
        $files = Get-OfficeRecoveryFiles -ExcludeZeroByte $excludeZeroCheckBox.Checked
        foreach ($file in $files) {
            $dataGridView.Rows.Add($file.Name, $file.FullName, $file.SizeKB, $file.LastWriteTime, $file.FileType)
        }
        $statusLabel.Text = "Found $($files.Count) potential recovery files."
    }

    # Button click events
    $refreshButton.Add_Click({
        $statusLabel.Text = "Refreshing file list..."
        Update-FileList
    })

    $openButton.Add_Click({
        if ($dataGridView.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a file.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $filePath = $dataGridView.SelectedRows[0].Cells[1].Value
        try {
            Start-Process -FilePath $filePath
            $statusLabel.Text = "Opened $filePath"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error opening file: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error opening file."
        }
    })

    $recoverButton.Add_Click({
        if ($dataGridView.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a file.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $filePath = $dataGridView.SelectedRows[0].Cells[1].Value
        $fileName = $dataGridView.SelectedRows[0].Cells[0].Value
        $inferredFileType = $dataGridView.SelectedRows[0].Cells[4].Value
        $selectedFileType = $fileTypeComboBox.SelectedItem.ToString()

        # Determine the file type for recovery
        $recoveryFileType = if ($selectedFileType -eq "Auto") { $inferredFileType } else { $selectedFileType }
        $newExtension = if ($recoveryExtensions.ContainsKey($recoveryFileType)) { $recoveryExtensions[$recoveryFileType] } else { [System.IO.Path]::GetExtension($fileName) }

        # Create new filename with recovery date
        $recoveryDate = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $newFileName = "${baseName}_${recoveryDate}${newExtension}"
        $desktopPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), $newFileName)

        try {
            Copy-Item -Path $filePath -Destination $desktopPath -Force
            $statusLabel.Text = "Recovered $newFileName to Desktop."
            [System.Windows.Forms.MessageBox]::Show("File recovered to Desktop as $newFileName.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error recovering file: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error recovering file."
        }
    })

    # Checkbox change event
    $excludeZeroCheckBox.Add_CheckedChanged({
        Update-FileList
    })

    # Initialize file list
    Update-FileList

    # Show form
    [System.Windows.Forms.Application]::Run($form)
}

# Execute GUI
try {
    Initialize-GUI
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error starting application: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

