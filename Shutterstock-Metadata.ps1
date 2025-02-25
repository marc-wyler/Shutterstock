# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
try {
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel -ErrorAction Stop
} catch {
    Write-Host "Excel interop assembly could not be loaded. Please install Microsoft Office."
}

# Function to load image from file without locking it
function Get-ImageFromFile ($path) {
    if (-Not (Test-Path $path)) { return $null }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $ms = New-Object System.IO.MemoryStream(, $bytes)
        return [System.Drawing.Image]::FromStream($ms)
    } catch {
        return $null
    }
}

# Function to show a temporary confirmation message (timeout in ms)
function Show-TemporaryMessage($message, $timeout, $backgroundColor) {
    $tempForm = New-Object System.Windows.Forms.Form
    $tempForm.FormBorderStyle = 'None'
    $tempForm.StartPosition = 'CenterScreen'
    $tempForm.Size = New-Object System.Drawing.Size(400,100)
    
    # Default color if none provided
    if (-not $backgroundColor) {
        $backgroundColor = [System.Drawing.Color]::FromArgb(147,112,219)
    }
    $tempForm.BackColor = $backgroundColor
    $tempForm.TopMost = $true
    $tempForm.Cursor = 'Hand'  # Show hand cursor to indicate clickable

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $message
    $label.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
    $label.ForeColor = [System.Drawing.Color]::White
    $label.TextAlign = 'MiddleCenter'
    $label.Dock = 'Fill'
    $label.Cursor = 'Hand'  # Show hand cursor on label too
    $tempForm.Controls.Add($label)

    # Add click handlers to close the form
    $tempForm.Add_Click({ $tempForm.Close() })
    $label.Add_Click({ $tempForm.Close() })

    # Reduce all timeouts to 1000ms (1 second)
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = [Math]::Min($timeout, 1000)  # Use 1 second or less
    $timer.Add_Tick({ 
        $timer.Stop()
        $tempForm.Close()
    })

    $timer.Start()
    $tempForm.ShowDialog()
}

# Function to update the file list (global variable $global:filePaths)
function Update-FileList {
    param($newFiles)
    $global:filePaths += $newFiles
    foreach ($path in $newFiles) {
        $listBox.Items.Add([System.IO.Path]::GetFileName($path))
    }
}

# Function to create the main form
function Create-MetadataEditorForm {
    param($filePaths)

    # Store file paths globally so we can update them later
    $global:filePaths = $filePaths

    # Add this at the start of the script with other global variables
    $global:currentKeywords = ""

    # Create the main UI form with improved aesthetics
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Image Metadata Editor"
    $form.Size = New-Object System.Drawing.Size(1280,900)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.TopMost = $true
    $form.Focus()

    # Use a bold Segoe UI font for buttons and labels
    $boldFont = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)

    # Change the red title label to be "Metadata Editor" at top left
    $metadataEditorLabel = New-Object System.Windows.Forms.Label
    $metadataEditorLabel.Text = "Metadata Editor"
    $metadataEditorLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $metadataEditorLabel.Location = New-Object System.Drawing.Point(10,8)
    $metadataEditorLabel.ForeColor = [System.Drawing.Color]::Red
    $metadataEditorLabel.AutoSize = $true

    # Base Path Label and TextBox
    $basePathLabel = New-Object System.Windows.Forms.Label
    $basePathLabel.Text = "Base Path:"
    $basePathLabel.Font = $boldFont
    $basePathLabel.Location = New-Object System.Drawing.Point(270,10)  # Changed from 10 to 270
    $basePathLabel.AutoSize = $true

    $basePathTextBox = New-Object System.Windows.Forms.TextBox
    $basePathTextBox.Font = $boldFont
    $basePathTextBox.Location = New-Object System.Drawing.Point(370,8)  # Changed from 80 to 370
    $basePathTextBox.Size = New-Object System.Drawing.Size(550,25)  # Match width with Selected File textbox
    $basePathTextBox.ReadOnly = $true
    $basePathTextBox.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Right"

    # Update Pictures label position
    $picturesLabel = New-Object System.Windows.Forms.Label
    $picturesLabel.Text = "Picture Selection:"
    $picturesLabel.Font = $boldFont
    $picturesLabel.Location = New-Object System.Drawing.Point(10,75)  # Changed from 40 to 75
    $picturesLabel.AutoSize = $true

    # Update listbox position
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Font = $boldFont
    $listBox.Location = New-Object System.Drawing.Point(10,100)  # Changed from 65 to 100 to align with picture box
    $listBox.Size = New-Object System.Drawing.Size(250,545)
    $listBox.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Bottom"
    foreach ($path in $global:filePaths) {
        $listBox.Items.Add([System.IO.Path]::GetFileName($path))
    }

    # Update Add More Pictures button position
    $addPicsButton = New-Object System.Windows.Forms.Button
    $addPicsButton.Text = "Add More Pictures"
    $addPicsButton.Font = $boldFont
    $addPicsButton.Size = New-Object System.Drawing.Size(250,35)
    $addPicsButton.Location = New-Object System.Drawing.Point(10,655)
    $addPicsButton.BackColor = [System.Drawing.Color]::FromArgb(255,165,0)  # Initial color (Orange)
    $addPicsButton.ForeColor = [System.Drawing.Color]::Black
    $addPicsButton.FlatStyle = 'Flat'
    $addPicsButton.FlatAppearance.BorderSize = 0

    # Add this after creating the Add More Pictures button
    $blinkTimer = New-Object System.Windows.Forms.Timer
    $blinkTimer.Interval = 1000  # Blink every second
    $blinkState = $false

    $blinkTimer.Add_Tick({
        if ($blinkState) {
            $addPicsButton.BackColor = [System.Drawing.Color]::FromArgb(255,165,0)  # Orange
        } else {
            $addPicsButton.BackColor = [System.Drawing.Color]::FromArgb(255,215,0)  # Yellow
        }
        $blinkState = -not $blinkState
    })

    # Start the blinking
    $blinkTimer.Start()

    # Add cleanup to form closing
    $form.Add_FormClosing({
        $blinkTimer.Stop()
        $blinkTimer.Dispose()
    })

    # Add the click handler for the Add More Pictures button
    $addPicsButton.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff"
        $dlg.Multiselect = $true
        $dlg.Title = "Select Additional Images"
        
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            foreach ($path in $dlg.FileNames) {
                $fileName = [System.IO.Path]::GetFileName($path)
                if (-not $listBox.Items.Contains($fileName)) {
                    $global:filePaths += $path
                    $listBox.Items.Add($fileName)
                }
            }
            Show-TemporaryMessage "Pictures added successfully" 1000 ([System.Drawing.Color]::FromArgb(51,153,255))
        }
    })

    # Add Bulk Rename button below Add More Pictures
    $bulkRenameButton = New-Object System.Windows.Forms.Button
    $bulkRenameButton.Text = "Bulk Rename"
    $bulkRenameButton.Font = $boldFont
    $bulkRenameButton.Size = New-Object System.Drawing.Size(250,35)
    $bulkRenameButton.Location = New-Object System.Drawing.Point(10,700)  # Positioned below Add More Pictures
    $bulkRenameButton.BackColor = [System.Drawing.Color]::FromArgb(255,165,0)  # Initial color (Orange)
    $bulkRenameButton.ForeColor = [System.Drawing.Color]::Black
    $bulkRenameButton.FlatStyle = 'Flat'
    $bulkRenameButton.FlatAppearance.BorderSize = 0

    # Add Delete Picture button below Bulk Rename
    $deletePictureButton = New-Object System.Windows.Forms.Button
    $deletePictureButton.Text = "Delete Picture"
    $deletePictureButton.Font = $boldFont
    $deletePictureButton.Size = New-Object System.Drawing.Size(250,35)
    $deletePictureButton.Location = New-Object System.Drawing.Point(10,745)  # Positioned below Bulk Rename
    $deletePictureButton.BackColor = [System.Drawing.Color]::FromArgb(255,165,0)  # Initial color (Orange)
    $deletePictureButton.ForeColor = [System.Drawing.Color]::Black
    $deletePictureButton.FlatStyle = 'Flat'
    $deletePictureButton.FlatAppearance.BorderSize = 0

    # Add click handler for Bulk Rename button
    $bulkRenameButton.Add_Click({
        if ($listBox.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No pictures selected for renaming.", "Info")
            return
        }

        if ([string]::IsNullOrWhiteSpace($fileNameTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a base name in 'Edit File Name'.", "Error")
            return
        }

        $baseName = $fileNameTextBox.Text
        $counter = 1

        foreach ($item in $listBox.Items) {
            $originalFileName = $item.ToString()
            $originalFilePath = $global:filePaths | Where-Object { 
                [System.IO.Path]::GetFileName($_) -eq $originalFileName 
            } | Select-Object -First 1

            if ($originalFilePath -and (Test-Path $originalFilePath)) {
                $directory = [System.IO.Path]::GetDirectoryName($originalFilePath)
                $extension = [System.IO.Path]::GetExtension($originalFilePath)

                # Rename all files with the same base name and a unique number
                $newFileName = "$baseName-$counter$extension"
                $newFilePath = Join-Path $directory $newFileName

                try {
                    Rename-Item -Path $originalFilePath -NewName $newFileName
                    $listBox.Items[$listBox.Items.IndexOf($item)] = $newFileName
                    $global:filePaths[$global:filePaths.IndexOf($originalFilePath)] = $newFilePath
                    $counter++
                } catch {
                    # Log the error silently without showing a message box
                    Write-Debug "Error renaming file: $originalFileName"
                }
            }
        }

        Show-TemporaryMessage "Bulk rename completed successfully!" 2000 ([System.Drawing.Color]::FromArgb(51,153,255))
    })

    # Add click handler for Delete Picture button
    $deletePictureButton.Add_Click({
        if ($listBox.SelectedItem) {
            $selectedFileName = $listBox.SelectedItem.ToString()
            $global:filePaths = $global:filePaths | Where-Object { 
                [System.IO.Path]::GetFileName($_) -ne $selectedFileName 
            }
            $listBox.Items.Remove($selectedFileName)
        } else {
            [System.Windows.Forms.MessageBox]::Show("No picture selected to delete.", "Info")
        }
    })

    # Editable File Name TextBox
    $fileNameLabel = New-Object System.Windows.Forms.Label
    $fileNameLabel.Text = "Edit File Name:"
    $fileNameLabel.Font = $boldFont
    $fileNameLabel.Location = New-Object System.Drawing.Point(270,40)
    $fileNameLabel.AutoSize = $true

    $fileNameTextBox = New-Object System.Windows.Forms.TextBox
    $fileNameTextBox.Font = $boldFont
    $fileNameTextBox.Location = New-Object System.Drawing.Point(370,38)
    $fileNameTextBox.Size = New-Object System.Drawing.Size(550,25)  # Made wider to almost touch .JPG
    $fileNameTextBox.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Right"

    # Non-editable textbox to display selected file name
    $displayedFileNameLabel = New-Object System.Windows.Forms.Label
    $displayedFileNameLabel.Text = "Selected File:"
    $displayedFileNameLabel.Font = $boldFont
    $displayedFileNameLabel.Location = New-Object System.Drawing.Point(270,70)
    $displayedFileNameLabel.AutoSize = $true

    $displayedFileNameTextBox = New-Object System.Windows.Forms.TextBox
    $displayedFileNameTextBox.Font = $boldFont
    $displayedFileNameTextBox.Location = New-Object System.Drawing.Point(370,68)
    $displayedFileNameTextBox.Size = New-Object System.Drawing.Size(550,25)  # Match width with other textboxes
    $displayedFileNameTextBox.ReadOnly = $true
    $displayedFileNameTextBox.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Right"

    # Preview Panel and PictureBox for image preview
    $previewPanel = New-Object System.Windows.Forms.Panel
    $previewPanel.Size = New-Object System.Drawing.Size(600,360)
    $previewPanel.Location = New-Object System.Drawing.Point(270,100)
    $previewPanel.BorderStyle = 'FixedSingle'
    $previewPanel.BackColor = [System.Drawing.Color]::White
    $previewPanel.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Right"

    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.SizeMode = 'Zoom'
    $pictureBox.Dock = 'Fill'
    $previewPanel.Controls.Add($pictureBox)

    # New TextBox on right side of preview to display Excel Tags for the current Event/Subject
    $excelTagsTextBox = New-Object System.Windows.Forms.RichTextBox
    $excelTagsTextBox.Multiline = $true
    $excelTagsTextBox.ReadOnly = $true
    $excelTagsTextBox.ScrollBars = "Vertical"
    $excelTagsTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $excelTagsTextBox.Location = New-Object System.Drawing.Point(900,100)
    $excelTagsTextBox.Size = New-Object System.Drawing.Size(360,330)  # Made slightly narrower
    $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
    $excelTagsTextBox.BorderStyle = "FixedSingle"
    $excelTagsTextBox.Text = "Keywords of Metadata.xlsx, Select Folder to change the filepath"

    # Update AI Suggestions label position
    $aiSuggestionsLabel = New-Object System.Windows.Forms.Label
    $aiSuggestionsLabel.Text = "AI Suggestions:"
    $aiSuggestionsLabel.Font = $boldFont
    $aiSuggestionsLabel.Location = New-Object System.Drawing.Point(900,470)
    $aiSuggestionsLabel.AutoSize = $true

    # Update suggestions textbox position
    $suggestedTagsTextBox = New-Object System.Windows.Forms.RichTextBox
    $suggestedTagsTextBox.Multiline = $true
    $suggestedTagsTextBox.ReadOnly = $false
    $suggestedTagsTextBox.ScrollBars = "Vertical"
    $suggestedTagsTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $suggestedTagsTextBox.Location = New-Object System.Drawing.Point(900,495)
    $suggestedTagsTextBox.Size = New-Object System.Drawing.Size(360,235)
    $suggestedTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
    $suggestedTagsTextBox.BorderStyle = "FixedSingle"
    $suggestedTagsTextBox.Text = "Copilot suggestions will appear here..."

    # Image Rotation Buttons
    $rotateLeftButton = New-Object System.Windows.Forms.Button
    $rotateLeftButton.Text = "Rotate Left"
    $rotateLeftButton.Font = $boldFont
    $rotateLeftButton.Size = New-Object System.Drawing.Size(100,30)
    $rotateLeftButton.Location = New-Object System.Drawing.Point(270,470)
    $rotateLeftButton.BackColor = [System.Drawing.Color]::FromArgb(51,153,255)
    $rotateLeftButton.ForeColor = [System.Drawing.Color]::White
    $rotateLeftButton.FlatStyle = 'Flat'
    $rotateLeftButton.FlatAppearance.BorderSize = 0
    $rotateLeftButton.Add_Click({
        if ($pictureBox.Image) {
            $pictureBox.Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone)
            $global:currentRotation = ($global:currentRotation - 90) % 360
            $pictureBox.Invalidate()
        }
    })

    $rotateRightButton = New-Object System.Windows.Forms.Button
    $rotateRightButton.Text = "Rotate Right"
    $rotateRightButton.Font = $boldFont
    $rotateRightButton.Size = New-Object System.Drawing.Size(100,30)
    $rotateRightButton.Location = New-Object System.Drawing.Point(380,470)
    $rotateRightButton.BackColor = [System.Drawing.Color]::FromArgb(51,153,255)
    $rotateRightButton.ForeColor = [System.Drawing.Color]::White
    $rotateRightButton.FlatStyle = 'Flat'
    $rotateRightButton.FlatAppearance.BorderSize = 0
    $rotateRightButton.Add_Click({
        if ($pictureBox.Image) {
            $pictureBox.Image.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone)
            $global:currentRotation = ($global:currentRotation + 90) % 360
            $pictureBox.Invalidate()
        }
    })

    # Metadata Panel (below rotation buttons)
    $metadataPanel = New-Object System.Windows.Forms.Panel
    $metadataPanel.Size = New-Object System.Drawing.Size(600,210)
    $metadataPanel.Location = New-Object System.Drawing.Point(270,520)
    $metadataPanel.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Right,Bottom"

    # Metadata controls inside Metadata Panel
    $selectedFolderPathLabel = New-Object System.Windows.Forms.Label
    $selectedFolderPathLabel.Text = "Save Folder:"
    $selectedFolderPathLabel.Font = $boldFont
    $selectedFolderPathLabel.Location = New-Object System.Drawing.Point(10,10)
    $selectedFolderPathLabel.AutoSize = $true

    $selectedFolderPathTextBox = New-Object System.Windows.Forms.TextBox
    $selectedFolderPathTextBox.Font = $boldFont
    $selectedFolderPathTextBox.Location = New-Object System.Drawing.Point(150,8)
    $selectedFolderPathTextBox.Size = New-Object System.Drawing.Size(320,25)
    $selectedFolderPathTextBox.ReadOnly = $true

    $selectFolderButton = New-Object System.Windows.Forms.Button
    $selectFolderButton.Text = "Select Folder"
    $selectFolderButton.Font = $boldFont
    $selectFolderButton.Size = New-Object System.Drawing.Size(100,25)
    $selectFolderButton.Location = New-Object System.Drawing.Point(480,8)
    $selectFolderButton.BackColor = [System.Drawing.Color]::FromArgb(102,153,0)
    $selectFolderButton.ForeColor = [System.Drawing.Color]::White
    $selectFolderButton.FlatStyle = 'Flat'
    $selectFolderButton.FlatAppearance.BorderSize = 0

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Title:"
    $titleLabel.Font = $boldFont
    $titleLabel.Location = New-Object System.Drawing.Point(10,45)
    $titleLabel.AutoSize = $true

    $titleTextBox = New-Object System.Windows.Forms.TextBox
    $titleTextBox.Font = $boldFont
    $titleTextBox.Location = New-Object System.Drawing.Point(150,43)
    $titleTextBox.Size = New-Object System.Drawing.Size(430,25)

    $subjectLabel = New-Object System.Windows.Forms.Label
    $subjectLabel.Text = "Subject:"
    $subjectLabel.Font = $boldFont
    $subjectLabel.Location = New-Object System.Drawing.Point(10,80)
    $subjectLabel.AutoSize = $true

    $subjectTextBox = New-Object System.Windows.Forms.TextBox
    $subjectTextBox.Font = $boldFont
    $subjectTextBox.Location = New-Object System.Drawing.Point(150,78)
    $subjectTextBox.Size = New-Object System.Drawing.Size(430,25)

    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Text = "Description:"
    $descriptionLabel.Font = $boldFont
    $descriptionLabel.Location = New-Object System.Drawing.Point(10,115)
    $descriptionLabel.AutoSize = $true

    $descriptionTextBox = New-Object System.Windows.Forms.TextBox
    $descriptionTextBox.Font = $boldFont
    $descriptionTextBox.Location = New-Object System.Drawing.Point(150,113)
    $descriptionTextBox.Size = New-Object System.Drawing.Size(400,25)  # Match tags textbox size
    $descriptionTextBox.Anchor = [System.Windows.Forms.AnchorStyles] "Top,Left,Right"

    # Add description info button and character limit
    $descriptionInfoButton = New-Object System.Windows.Forms.Button
    $descriptionInfoButton.Text = "i"
    $descriptionInfoButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $descriptionInfoButton.Size = New-Object System.Drawing.Size(25,25)
    $descriptionInfoButton.Location = New-Object System.Drawing.Point(555,113)  # Match tags info button position
    $descriptionInfoButton.BackColor = [System.Drawing.Color]::FromArgb(51,153,255)
    $descriptionInfoButton.ForeColor = [System.Drawing.Color]::White
    $descriptionInfoButton.FlatStyle = "Flat"
    $descriptionInfoButton.FlatAppearance.BorderSize = 0
    $descriptionInfoButton.Cursor = "Hand"

    # Add tooltip for description info button
    $descriptionToolTip = New-Object System.Windows.Forms.ToolTip
    $descriptionToolTip.InitialDelay = 100
    $descriptionToolTip.ReshowDelay = 100
    $descriptionToolTip.AutoPopDelay = 5000
    $descriptionToolTip.ShowAlways = $true
    $descriptionToolTip.SetToolTip($descriptionInfoButton, "Maximum 200 characters allowed for Shutterstock descriptions")

    # Add character limit handling for description textbox
    $descriptionTextBox.MaxLength = 200
    $descriptionTextBox.Add_TextChanged({
        $remainingChars = 200 - $descriptionTextBox.Text.Length
        if ($remainingChars -le 0) {
            $descriptionTextBox.BackColor = [System.Drawing.Color]::FromArgb(255,200,200)  # Light red
        } else {
            $descriptionTextBox.BackColor = [System.Drawing.Color]::White
        }
    })

    $tagsLabel = New-Object System.Windows.Forms.Label
    $tagsLabel.Text = "Tags:"
    $tagsLabel.Font = $boldFont
    $tagsLabel.Location = New-Object System.Drawing.Point(10,150)
    $tagsLabel.AutoSize = $true

    $tagsTextBox = New-Object System.Windows.Forms.TextBox
    $tagsTextBox.Font = $boldFont
    $tagsTextBox.Location = New-Object System.Drawing.Point(150,148)
    $tagsTextBox.Size = New-Object System.Drawing.Size(400,25)
    $tagsTextBox.Text = ""

    # Add info button for Tags
    $tagsInfoButton = New-Object System.Windows.Forms.Button
    $tagsInfoButton.Text = "i"
    $tagsInfoButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $tagsInfoButton.Size = New-Object System.Drawing.Size(25,25)
    $tagsInfoButton.Location = New-Object System.Drawing.Point(555,148)
    $tagsInfoButton.BackColor = [System.Drawing.Color]::FromArgb(51,153,255)
    $tagsInfoButton.ForeColor = [System.Drawing.Color]::White
    $tagsInfoButton.FlatStyle = "Flat"
    $tagsInfoButton.FlatAppearance.BorderSize = 0
    $tagsInfoButton.Cursor = "Hand"

    # Create and configure tooltip
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.InitialDelay = 100
    $toolTip.ReshowDelay = 100
    $toolTip.AutoPopDelay = 5000
    $toolTip.ShowAlways = $true
    $toolTip.SetToolTip($tagsInfoButton, "Separate multiple tags with commas.`nExample: nature, landscape, mountain")

    $dateLabel = New-Object System.Windows.Forms.Label
    $dateLabel.Text = "Date Taken:"
    $dateLabel.Font = $boldFont
    $dateLabel.Location = New-Object System.Drawing.Point(10,185)
    $dateLabel.AutoSize = $true

    $dateTimePicker = New-Object System.Windows.Forms.DateTimePicker
    $dateTimePicker.Font = $boldFont
    $dateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $dateTimePicker.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    $dateTimePicker.Location = New-Object System.Drawing.Point(150,183)
    $dateTimePicker.Size = New-Object System.Drawing.Size(200,25)

    # Update the controls order in the Metadata Panel
    $metadataPanel.Controls.Clear()
    $metadataPanel.Controls.AddRange(@(
        # Save Folder controls first
        $selectedFolderPathLabel,
        $selectedFolderPathTextBox,
        $selectFolderButton,
        
        # Then the rest of the controls
        $titleLabel,
        $titleTextBox,
        $subjectLabel,
        $subjectTextBox,
        $descriptionLabel,
        $descriptionTextBox,
        $descriptionInfoButton,
        $tagsLabel,
        $tagsTextBox,
        $tagsInfoButton,
        $dateLabel,
        $dateTimePicker
    ))

    # Update the control positions
    $selectedFolderPathLabel.Location = New-Object System.Drawing.Point(10,10)
    $selectedFolderPathTextBox.Location = New-Object System.Drawing.Point(150,8)
    $selectFolderButton.Location = New-Object System.Drawing.Point(480,8)

    $titleLabel.Location = New-Object System.Drawing.Point(10,45)
    $titleTextBox.Location = New-Object System.Drawing.Point(150,43)

    $subjectLabel.Location = New-Object System.Drawing.Point(10,80)
    $subjectTextBox.Location = New-Object System.Drawing.Point(150,78)

    $descriptionLabel.Location = New-Object System.Drawing.Point(10,115)
    $descriptionTextBox.Location = New-Object System.Drawing.Point(150,113)
    $descriptionInfoButton.Location = New-Object System.Drawing.Point(555,113)

    $tagsLabel.Location = New-Object System.Drawing.Point(10,150)
    $tagsTextBox.Location = New-Object System.Drawing.Point(150,148)
    $tagsInfoButton.Location = New-Object System.Drawing.Point(555,148)

    $dateLabel.Location = New-Object System.Drawing.Point(10,185)
    $dateTimePicker.Location = New-Object System.Drawing.Point(150,183)

    # Bottom Action Buttons (centered in a FlowLayoutPanel)
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save Metadata"
    $saveButton.Font = $boldFont
    $saveButton.Size = New-Object System.Drawing.Size(130,35)
    $saveButton.BackColor = [System.Drawing.Color]::FromArgb(51,153,255)
    $saveButton.ForeColor = [System.Drawing.Color]::White
    $saveButton.FlatStyle = 'Flat'
    $saveButton.FlatAppearance.BorderSize = 0

    # Add the tooltip here
    $saveButtonTooltip = New-Object System.Windows.Forms.ToolTip
    $saveButtonTooltip.InitialDelay = 100
    $saveButtonTooltip.ReshowDelay = 100
    $saveButtonTooltip.AutoPopDelay = 5000
    $saveButtonTooltip.ShowAlways = $true
    $saveButtonTooltip.SetToolTip($saveButton, "Save Metadata and Delete Picture from List")

    $saveToExcelButton = New-Object System.Windows.Forms.Button
    $saveToExcelButton.Text = "Save to Excel"
    $saveToExcelButton.Font = $boldFont
    $saveToExcelButton.Size = New-Object System.Drawing.Size(130,35)
    $saveToExcelButton.BackColor = [System.Drawing.Color]::FromArgb(255,204,0)
    $saveToExcelButton.ForeColor = [System.Drawing.Color]::Black
    $saveToExcelButton.FlatStyle = 'Flat'
    $saveToExcelButton.FlatAppearance.BorderSize = 0

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = "Reset"
    $resetButton.Font = $boldFont
    $resetButton.Size = New-Object System.Drawing.Size(80,35)
    $resetButton.BackColor = [System.Drawing.Color]::FromArgb(255,77,77)
    $resetButton.ForeColor = [System.Drawing.Color]::White
    $resetButton.FlatStyle = 'Flat'
    $resetButton.FlatAppearance.BorderSize = 0

    $infoButton = New-Object System.Windows.Forms.Button
    $infoButton.Text = "Info"
    $infoButton.Font = $boldFont
    $infoButton.Size = New-Object System.Drawing.Size(80,35)
    $infoButton.Location = New-Object System.Drawing.Point(1150,8)  # Moved from 1200 to 1150
    $infoButton.BackColor = [System.Drawing.Color]::FromArgb(51,102,255)
    $infoButton.ForeColor = [System.Drawing.Color]::White
    $infoButton.FlatStyle = 'Flat'
    $infoButton.FlatAppearance.BorderSize = 0
    $infoButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

    $findTagsButton = New-Object System.Windows.Forms.Button
    $findTagsButton.Text = "Copilot Suggestions"
    $findTagsButton.Font = $boldFont
    $findTagsButton.Size = New-Object System.Drawing.Size(160,35)
    $findTagsButton.BackColor = [System.Drawing.Color]::FromArgb(255,192,203)  # Pink color
    $findTagsButton.ForeColor = [System.Drawing.Color]::Black
    $findTagsButton.FlatStyle = 'Flat'
    $findTagsButton.FlatAppearance.BorderSize = 0  # Remove border
    $findTagsButton.Margin = New-Object System.Windows.Forms.Padding(20,5,20,5)
    $findTagsButton.Cursor = 'Hand'  # Add hand cursor

    $aiCheckBox = New-Object System.Windows.Forms.CheckBox
    $aiCheckBox.Text = "AI is Open"
    $aiCheckBox.Font = $boldFont
    $aiCheckBox.Size = New-Object System.Drawing.Size(300,35)
    $aiCheckBox.ForeColor = [System.Drawing.Color]::Black
    $aiCheckBox.TextAlign = "MiddleLeft"
    $aiCheckBox.Padding = New-Object System.Windows.Forms.Padding(0,0,0,0)
    $aiCheckBox.Margin = New-Object System.Windows.Forms.Padding(15,5,0,5)  # Match other button margins
    $aiCheckBox.Location = New-Object System.Drawing.Point(270,810)  # Adjusted position to fit within the frame

    # Add Copy Suggestions button next to Copilot Suggestions
    $copySuggestionsButton = New-Object System.Windows.Forms.Button
    $copySuggestionsButton.Text = "Copy Suggestions"
    $copySuggestionsButton.Font = $boldFont
    $copySuggestionsButton.Size = New-Object System.Drawing.Size(160,35)
    $copySuggestionsButton.BackColor = [System.Drawing.Color]::FromArgb(255,192,203)  # Pink color
    $copySuggestionsButton.ForeColor = [System.Drawing.Color]::Black
    $copySuggestionsButton.FlatStyle = 'Flat'
    $copySuggestionsButton.FlatAppearance.BorderSize = 0
    $copySuggestionsButton.Margin = New-Object System.Windows.Forms.Padding(20,5,20,5)
    $copySuggestionsButton.Cursor = 'Hand'

    $actionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $actionPanel.Location = New-Object System.Drawing.Point(270,760)
    $actionPanel.Size = New-Object System.Drawing.Size(1100,50)  # Made wider to fit checkbox
    $actionPanel.FlowDirection = "LeftToRight"
    $actionPanel.WrapContents = $false
    foreach ($btn in @($saveButton, $saveToExcelButton, $resetButton, $findTagsButton)) {
        $btn.Margin = New-Object System.Windows.Forms.Padding(15,5,15,5)  # Reduce margins between buttons
    }
    $actionPanel.Controls.AddRange(@($saveButton, $saveToExcelButton, $resetButton, $findTagsButton, $copySuggestionsButton, $aiCheckBox))

    # Create the extension textbox (keep this one, remove the duplicate later in the code)
    $fileExtensionTextBox = New-Object System.Windows.Forms.TextBox
    $fileExtensionTextBox.Font = $boldFont
    $fileExtensionTextBox.Location = New-Object System.Drawing.Point(920,38)  # Moved closer to filename textbox
    $fileExtensionTextBox.Size = New-Object System.Drawing.Size(50,25)
    $fileExtensionTextBox.Text = ".JPG"
    $fileExtensionTextBox.ReadOnly = $true
    $fileExtensionTextBox.TextAlign = "Center"
    $fileExtensionTextBox.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $fileExtensionTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top

    # Update the filename textbox size and remove its text changed handler
    $fileNameTextBox.Size = New-Object System.Drawing.Size(400,25)  # Make smaller to fit extension box
    $fileNameTextBox.Add_TextChanged({})  # Remove any existing handlers

    # Add this definition for keywordCountTextBox
    $keywordCountTextBox = New-Object System.Windows.Forms.Label
    $keywordCountTextBox.Font = $boldFont
    $keywordCountTextBox.Location = New-Object System.Drawing.Point(900,740)
    $keywordCountTextBox.AutoSize = $true
    $keywordCountTextBox.Text = "0/50 Keywords"

    # Add the keywordCountTextBox to the form controls
    $form.Controls.Add($keywordCountTextBox)

    # Update the form controls - remove any duplicate textbox references
    $form.Controls.AddRange(@(
        $metadataEditorLabel,  # Changed from $titleLabel to $metadataEditorLabel
        $basePathLabel, 
        $basePathTextBox,
        $picturesLabel,
        $listBox, 
        $addPicsButton,
        $bulkRenameButton,  # Add Bulk Rename button
        $deletePictureButton,  # Add Delete Picture button
        $fileNameLabel, 
        $fileNameTextBox,
        $fileExtensionTextBox,
        $displayedFileNameLabel, 
        $displayedFileNameTextBox,
        $previewPanel,
        $excelTagsTextBox,
        $aiSuggestionsLabel,
        $suggestedTagsTextBox,
        $rotateLeftButton, 
        $rotateRightButton,
        $metadataPanel,
        $actionPanel,
        $infoButton
    ))

    # ListBox Selection Changed event
    $listBox.Add_SelectedIndexChanged({
        if ($listBox.SelectedItem) {
            try {
                $selectedFileName = $listBox.SelectedItem.ToString()
                $selectedFilePath = $global:filePaths | Where-Object { 
                    [System.IO.Path]::GetFileName($_).Equals($selectedFileName, [StringComparison]::OrdinalIgnoreCase)
                } | Select-Object -First 1

                if (-not $selectedFilePath -or -not (Test-Path $selectedFilePath)) {
                    return  # Skip processing if the file is not found
                }

                $basePathTextBox.Text = [System.IO.Path]::GetDirectoryName($selectedFilePath)
                $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($selectedFilePath)  # Remove .jpg extension
                $fileNameTextBox.Text = $baseFileName  # Set the file name text box
                $displayedFileNameTextBox.Text = [System.IO.Path]::GetFileName($selectedFilePath)

                # Load image
                if ($pictureBox.Image) {
                    $pictureBox.Image.Dispose()
                    $pictureBox.Image = $null
                }
                $img = Get-ImageFromFile $selectedFilePath
                if ($img) {
                    $pictureBox.Image = $img
                }

                # Check if the file exists in Metadata.xlsx
                $selectedFolderPath = $selectedFolderPathTextBox.Text
                if ($selectedFolderPath -and (Test-Path (Join-Path $selectedFolderPath "Metadata.xlsx"))) {
                    $excelApp = New-Object -ComObject Excel.Application
                    $excelApp.Visible = $false
                    $excelApp.DisplayAlerts = $false

                    $excelPath = Join-Path $selectedFolderPath "Metadata.xlsx"
                    $workbook = $excelApp.Workbooks.Open($excelPath)
                    $worksheet = $workbook.Sheets.Item(1)

                    $usedRange = $worksheet.UsedRange
                    $found = $false

                    for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
                        $excelFileName = $worksheet.Cells.Item($row, 1).Text
                        if ($excelFileName -eq $baseFileName) {
                            $titleTextBox.Text = $worksheet.Cells.Item($row, 1).Text
                            $subjectTextBox.Text = $worksheet.Cells.Item($row, 2).Text
                            $descriptionTextBox.Text = $worksheet.Cells.Item($row, 3).Text
                            $tagsTextBox.Text = $worksheet.Cells.Item($row, 4).Text
                            $found = $true
                            break
                        }
                    }

                    if (-not $found) {
                        $titleTextBox.Text = ""
                        $subjectTextBox.Text = ""
                        $descriptionTextBox.Text = ""
                        $tagsTextBox.Text = ""
                    }

                    # Cleanup
                    $workbook.Close()
                    $excelApp.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }

            } catch {
                Write-Debug "Error processing selected file: $_"
            }
        } else {
            # Clear everything if no selection
            if ($pictureBox.Image) {
                $pictureBox.Image.Dispose()
                $pictureBox.Image = $null
            }
            $basePathTextBox.Text = ""
            $fileNameTextBox.Text = ""
            $displayedFileNameTextBox.Text = ""
            $titleTextBox.Text = ""
            $subjectTextBox.Text = ""
            $descriptionTextBox.Text = ""
            $tagsTextBox.Text = ""
            $dateTimePicker.Value = Get-Date
            $excelTagsTextBox.Text = "Keywords of Metadata.xlsx, Select Folder to change the filepath"
            $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
        }
    })

    # If files were provided, select the first one by default
    if ($global:filePaths.Count -gt 0) {
        $listBox.SelectedIndex = 0
        $basePathTextBox.Text = [System.IO.Path]::GetDirectoryName($global:filePaths[0])
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($global:filePaths[0])
        $fileNameTextBox.Text = $baseFileName
        $displayedFileNameTextBox.Text = [System.IO.Path]::GetFileName($global:filePaths[0])
        $img = Get-ImageFromFile $global:filePaths[0]
        if ($img) { $pictureBox.Image = $img } else { $pictureBox.Image = $null }
    }

    # Add a function to update keywords
    function Update-Keywords {
        param(
            [string]$title,
            [string]$folderPath
        )
        
        if ($folderPath -and $title -and (Test-Path (Join-Path $folderPath "Metadata.xlsx"))) {
            try {
                $excelApp = New-Object -ComObject Excel.Application
                $excelApp.Visible = $false
                $excelApp.DisplayAlerts = $false
                
                $excelPath = Join-Path $folderPath "Metadata.xlsx"
                $workbook = $excelApp.Workbooks.Open($excelPath)
                $worksheet = $workbook.Sheets.Item(1)
                
                $usedRange = $worksheet.UsedRange
                $found = $false
                
                for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
                    if ($worksheet.Cells.Item($row, 1).Text -eq $title) {
                        $global:currentKeywords = $worksheet.Cells.Item($row, 4).Text
                        $found = $true
                        break
                    }
                }
                
                if ($found) {
                    $excelTagsTextBox.Text = "Keywords of Metadata.xlsx, Select Folder to change the filepath`n`n$($global:currentKeywords)"
                    $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(240,255,240)  # Light green
                } else {
                    $global:currentKeywords = ""
                    $excelTagsTextBox.Text = "Keywords of Metadata.xlsx, Select Folder to change the filepath"
                    $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)  # Default
                }
                
                # Cleanup
                $workbook.Close()
                $excelApp.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                
            } catch {
                $excelTagsTextBox.Text = "Error reading Excel file: $_"
                $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(255,240,240)  # Light red
            }
        } else {
            $global:currentKeywords = ""
            $excelTagsTextBox.Text = "Select a folder containing Metadata.xlsx to see keywords"
            $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
        }
    }

    # Update the Select Folder button click handler
    $selectFolderButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFolderPath = $folderBrowser.SelectedPath
            if (-not $selectedFolderPath.EndsWith("\")) {
                $selectedFolderPath += "\"
            }
            $selectedFolderPathTextBox.Text = $selectedFolderPath
            # Update keywords if we have a title
            if (-not [string]::IsNullOrWhiteSpace($titleTextBox.Text)) {
                Update-Keywords -title $titleTextBox.Text -folderPath $selectedFolderPath
            }
        }
    })

    # Function to format keywords display
    function Format-KeywordsDisplay {
        param(
            [string]$title,
            [string[]]$keywords
        )
        
        # Clear existing text and formatting
        $excelTagsTextBox.Clear()
        
        # Add header in bold
        $excelTagsTextBox.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $excelTagsTextBox.AppendText("Keywords of `"$title`":")
        
        # Add newlines and keywords with commas
        $excelTagsTextBox.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10)
        if ($keywords) {
            $excelTagsTextBox.AppendText("`n`n")
            $excelTagsTextBox.AppendText(($keywords -join ", "))
        } else {
            $excelTagsTextBox.AppendText("`n`nNo keywords found")
        }
    }

    # Update the title textbox LostFocus event
    $titleTextBox.Add_LostFocus({
        if ($selectedFolderPathTextBox.Text -and (Test-Path (Join-Path $selectedFolderPathTextBox.Text "Metadata.xlsx"))) {
            try {
                $excelApp = New-Object -ComObject Excel.Application
                $excelApp.Visible = $false
                $excelApp.DisplayAlerts = $false
                
                $excelPath = Join-Path $selectedFolderPathTextBox.Text "Metadata.xlsx"
                $workbook = $excelApp.Workbooks.Open($excelPath)
                $worksheet = $workbook.Sheets.Item(1)
                
                $currentTitle = $titleTextBox.Text
                $usedRange = $worksheet.UsedRange
                $found = $false
                
                for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
                    if ($worksheet.Cells.Item($row, 1).Text -eq $currentTitle) {
                        $global:currentKeywords = $worksheet.Cells.Item($row, 4).Text
                        $found = $true
                        break
                    }
                }
                
                if (-not [string]::IsNullOrWhiteSpace($currentTitle)) {
                    if ($found) {
                        $keywords = $global:currentKeywords -split ',' | 
                            ForEach-Object { $_.Trim() } | 
                            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                        Format-KeywordsDisplay -title $currentTitle -keywords $keywords
                        Update-KeywordCount -keywords $keywords
                        $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(240,255,240)
                    } else {
                        Format-KeywordsDisplay -title $currentTitle -keywords @()
                        Update-KeywordCount -keywords @()
                        $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
                    }
                }
                
                # Cleanup
                $workbook.Close()
                $excelApp.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                
            } catch {
                $excelTagsTextBox.Text = "Error reading Excel file: $_"
                $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(255,240,240)
            }
        } else {
            if (-not [string]::IsNullOrWhiteSpace($titleTextBox.Text)) {
                Format-KeywordsDisplay -title $titleTextBox.Text -keywords @()
            } else {
                $excelTagsTextBox.Text = "Keywords of Metadata.xlsx, Select Folder to change the filepath"
            }
        }
    })

    # Function to update the Save Metadata button state
    function Update-SaveButtonState {
        if ([string]::IsNullOrWhiteSpace($fileNameTextBox.Text)) {
            $saveButton.Enabled = $false
        } else {
            $saveButton.Enabled = $true
        }
    }

    # Add TextChanged event to fileNameTextBox to update Save button state
    $fileNameTextBox.Add_TextChanged({
        Update-SaveButtonState
    })

    # Initial call to update the Save button state
    Update-SaveButtonState

    # Update the Save Metadata button click handler
    $saveButton.Add_Click({
        if (-not $selectedFolderPathTextBox.Text) {
            [System.Windows.Forms.MessageBox]::Show("Please select a folder to save the file.", "Error")
            return
        }

        if ([string]::IsNullOrWhiteSpace($fileNameTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a file name.", "Error")
            return
        }

        if (-not $selectedFilePath -or -not (Test-Path $selectedFilePath)) {
            [System.Windows.Forms.MessageBox]::Show("Selected file path is invalid or does not exist.", "Error")
            return
        }

        try {
            $selectedFolderPath = $selectedFolderPathTextBox.Text
            if (-not $selectedFolderPath.EndsWith("\")) {
                $selectedFolderPath += "\"
            }

            $newFileName = $fileNameTextBox.Text + $fileExtensionTextBox.Text
            $newFilePath = Join-Path $selectedFolderPath $newFileName

            $img = [System.Drawing.Image]::FromFile($selectedFilePath)
            
            switch ($global:currentRotation) {
                90  { $img.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone) }
                180 { $img.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipNone) }
                270 { $img.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone) }
            }

            $img.Save($newFilePath, $img.RawFormat)
            $img.Dispose()

            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace([System.IO.Path]::GetDirectoryName($newFilePath))
            $file = $folder.ParseName([System.IO.Path]::GetFileName($newFilePath))

            if ($file) {
                $vbsPath = [System.IO.Path]::GetTempFileName()
                $vbsPath = [System.IO.Path]::ChangeExtension($vbsPath, ".vbs")

                $vbsContent = @"
On Error Resume Next
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace("$([System.IO.Path]::GetDirectoryName($newFilePath))")
Set objFile = objFolder.ParseName("$([System.IO.Path]::GetFileName($newFilePath))")
objFolder.GetDetailsOf objFile, 21
objFolder.SetDetailsOf objFile, 21, "$($titleTextBox.Text)"
objFolder.GetDetailsOf objFile, 24
objFolder.SetDetailsOf objFile, 24, "$($subjectTextBox.Text)"
objFolder.GetDetailsOf objFile, 18
objFolder.SetDetailsOf objFile, 18, "$($descriptionTextBox.Text)"
objFolder.GetDetailsOf objFile, 25
objFolder.SetDetailsOf objFile, 25, "$($tagsTextBox.Text)"
objFolder.GetDetailsOf objFile, 12
objFolder.SetDetailsOf objFile, 12, "$($dateTimePicker.Value.ToString('yyyy-MM-dd HH:mm:ss'))"
Set objFile = Nothing
Set objFolder = Nothing
Set objShell = Nothing
"@
                
                $vbsContent | Out-File -FilePath $vbsPath -Encoding ASCII
                
                if (Test-Path $vbsPath) {
                    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $startInfo.FileName = "wscript.exe"
                    $startInfo.Arguments = "`"$vbsPath`""
                    $startInfo.UseShellExecute = $false
                    $startInfo.CreateNoWindow = $true
                    $startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
                    
                    $process = [System.Diagnostics.Process]::Start($startInfo)
                    $process.WaitForExit()
                    
                    if (Test-Path $vbsPath) {
                        Remove-Item $vbsPath -Force -ErrorAction SilentlyContinue
                    }
                }
            }

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            Show-TemporaryMessage "Metadata saved successfully!" 1000 ([System.Drawing.Color]::FromArgb(0,153,76))

            $currentIndex = $listBox.SelectedIndex
            $currentItem = $listBox.SelectedItem
            if ($currentItem) {
                $global:filePaths = $global:filePaths | Where-Object { 
                    [System.IO.Path]::GetFileName($_) -ne $currentItem 
                }
                $listBox.Items.RemoveAt($currentIndex)
            }

            if ($listBox.Items.Count -gt 0) {
                if ($currentIndex -ge $listBox.Items.Count) {
                    $listBox.SelectedIndex = $listBox.Items.Count - 1
                } else {
                    $listBox.SelectedIndex = $currentIndex
                }
            } else {
                $listBox.SelectedIndex = -1
                $pictureBox.Image = $null
                $titleTextBox.Text = ""
                $subjectTextBox.Text = ""
                $descriptionTextBox.Text = ""
                $tagsTextBox.Text = ""
                $dateTimePicker.Value = Get-Date
                $fileNameTextBox.Text = ""
                $displayedFileNameTextBox.Text = ""
                Show-TemporaryMessage "All images processed" 1000 ([System.Drawing.Color]::FromArgb(0,153,76))
            }

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving metadata: $_", "Error")
        }
    })

    # Reset button – Restart the tool
    $resetButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to reset? This will restart the tool and clear your selections.",
            "Confirmation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $form.Close()
            Start-Sleep -Milliseconds 200
            $startInfo = New-Object System.Diagnostics.ProcessStartInfo
            $startInfo.FileName = "powershell.exe"
            $startInfo.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`""
            $startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
            [System.Diagnostics.Process]::Start($startInfo)
        }
    })

    # Info button – More appealing guide with additional links
    $infoButton.Add_Click({
        $infoForm = New-Object System.Windows.Forms.Form
        $infoForm.Text = "Tool Guide"
        $infoForm.Size = New-Object System.Drawing.Size(800,600)  # Made form larger
        $infoForm.StartPosition = "CenterParent"
        $infoForm.BackColor = [System.Drawing.Color]::White
        $infoForm.FormBorderStyle = 'FixedSingle'
        $infoForm.MaximizeBox = $false
        $infoForm.TopMost = $true
        $infoForm.Focus()

        $headerLabel = New-Object System.Windows.Forms.Label
        $headerLabel.Text = "Image Metadata Editor - User Guide"
        $headerLabel.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
        $headerLabel.AutoSize = $true
        $headerLabel.Location = New-Object System.Drawing.Point(20,20)
        $infoForm.Controls.Add($headerLabel)

        $guideRichTextBox = New-Object System.Windows.Forms.RichTextBox
        $guideRichTextBox.Location = New-Object System.Drawing.Point(20,60)
        $guideRichTextBox.Size = New-Object System.Drawing.Size(740,420)  # Made textbox larger
        $guideRichTextBox.Font = New-Object System.Drawing.Font("Segoe UI",11)  # Increased font size
        $guideRichTextBox.ReadOnly = $true
        $guideRichTextBox.BorderStyle = 'None'
        $guideRichTextBox.BackColor = [System.Drawing.Color]::White
        $guideRichTextBox.Text = @"
This tool helps you manage and edit image metadata for Shutterstock submissions. Here's what you can do:

Image Management:
- View and select multiple images from your computer
- Preview images in the central panel
- Rotate images if needed using the rotation buttons
- Add more images to the list at any time

Metadata Editing:
- Edit basic file information:
  - Filename (with automatic .JPG extension)
  - Title
  - Subject
  - Description (limited to 200 characters)
  - Tags/Keywords
  - Date Taken

Keyword Management:
- Save keywords to Excel for reuse
- View existing keywords for similar images
- Get AI-powered keyword suggestions using Microsoft Copilot
- Track keyword count (maximum 50 keywords allowed)
- Keywords are automatically deduplicated

Saving and Organization:
- Save metadata directly to images
- Maintain a keyword database in Excel
- Automatically remove processed images from the list
- Reset the tool to start fresh

Tips:
- Always separate keywords with commas
- Use the 'AI Open (will prevent open a new tab)' checkbox to prevent multiple Copilot windows
- Check the keyword count to stay within Shutterstock's limits
- Save regularly to Excel to build your keyword database

For more information and updates:
Github Repository: https://github.com/marc-wyler/Shutterstock
"@
        $infoForm.Controls.Add($guideRichTextBox)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "Close"
        $okButton.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
        $okButton.Size = New-Object System.Drawing.Size(100,35)
        $okButton.Location = New-Object System.Drawing.Point(660,500)  # Adjusted position
        $okButton.BackColor = [System.Drawing.Color]::FromArgb(51,153,255)
        $okButton.ForeColor = [System.Drawing.Color]::White
        $okButton.FlatStyle = 'Flat'
        $okButton.FlatAppearance.BorderSize = 0
        $okButton.TextAlign = 'MiddleCenter'
        $okButton.Add_Click({ $infoForm.Close() })
        $infoForm.Controls.Add($okButton)

        $infoForm.ShowDialog()
    })

    # Update the Save to Excel button handler
    $saveToExcelButton.Add_Click({
        try {
            if (-not $selectedFolderPathTextBox.Text) {
                [System.Windows.Forms.MessageBox]::Show("Please select a save folder first.", "Warning")
                return
            }

            # Kill any existing Excel processes
            Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force

            $excelPath = Join-Path $selectedFolderPathTextBox.Text "Metadata.xlsx"
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            # Create or open workbook
            if (Test-Path $excelPath) {
                $workbook = $excel.Workbooks.Open($excelPath)
                $worksheet = $workbook.Sheets.Item(1)
            } else {
                $workbook = $excel.Workbooks.Add()
                $worksheet = $workbook.Sheets.Item(1)
                # Add headers for new file
                $headers = @("Title", "Subject", "Description", "Tags")
                for ($i = 0; $i -lt $headers.Count; $i++) {
                    $worksheet.Cells.Item(1, $i + 1) = $headers[$i]
                }
            }

            $lastRow = $worksheet.UsedRange.Rows.Count
            if ($lastRow -lt 1) { $lastRow = 1 }

            # Check if title exists
            $titleExists = $false
            $existingRow = 0
            for ($i = 2; $i -le $lastRow; $i++) {
                if ($worksheet.Cells.Item($i, 1).Text -eq $titleTextBox.Text) {
                    $titleExists = $true
                    $existingRow = $i
                    break
                }
            }

            if ($titleExists) {
                # Merge existing and new tags
                $existingTags = $worksheet.Cells.Item($existingRow, 4).Text
                $newTags = $tagsTextBox.Text
                
                # Combine tags, split by comma, trim, remove empties, and remove duplicates
                $allTags = @()
                if ($existingTags) { $allTags += $existingTags -split ',' | ForEach-Object { $_.Trim() } }
                if ($newTags) { $allTags += $newTags -split ',' | ForEach-Object { $_.Trim() } }
                $uniqueTags = $allTags | Where-Object { $_ } | Select-Object -Unique | Sort-Object

                # Update row with merged tags
                $worksheet.Cells.Item($existingRow, 2) = $subjectTextBox.Text
                $worksheet.Cells.Item($existingRow, 3) = $descriptionTextBox.Text
                $worksheet.Cells.Item($existingRow, 4) = ($uniqueTags -join ", ")
            } else {
                # Add new row
                $newRow = $lastRow + 1
                $worksheet.Cells.Item($newRow, 1) = $titleTextBox.Text
                $worksheet.Cells.Item($newRow, 2) = $subjectTextBox.Text
                $worksheet.Cells.Item($newRow, 3) = $descriptionTextBox.Text
                $worksheet.Cells.Item($newRow, 4) = $tagsTextBox.Text
            }

            # Auto-fit columns
            $worksheet.UsedRange.Columns.AutoFit()

            # Save and close properly
            $workbook.SaveAs($excelPath)
            $workbook.Close($true)
            $excel.Quit()

            # Clean up COM objects
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            Show-TemporaryMessage "Data saved to Excel successfully" 1000 ([System.Drawing.Color]::FromArgb(0,153,76))

            # Immediately load and display the updated keywords
            if (-not [string]::IsNullOrWhiteSpace($titleTextBox.Text)) {
                try {
                    $excelApp = New-Object -ComObject Excel.Application
                    $excelApp.Visible = $false
                    $excelApp.DisplayAlerts = $false
                    
                    $excelPath = Join-Path $selectedFolderPathTextBox.Text "Metadata.xlsx"
                    $workbook = $excelApp.Workbooks.Open($excelPath)
                    $worksheet = $workbook.Sheets.Item(1)
                    
                    $currentTitle = $titleTextBox.Text
                    $usedRange = $worksheet.UsedRange
                    $found = $false
                    
                    for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
                        if ($worksheet.Cells.Item($row, 1).Text -eq $currentTitle) {
                            $keywords = $worksheet.Cells.Item($row, 4).Text -split ',' | 
                                ForEach-Object { $_.Trim() } | 
                                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                            Format-KeywordsDisplay -title $currentTitle -keywords $keywords
                            Update-KeywordCount -keywords $keywords
                            $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(240,255,240)
                            $found = $true
                            break
                        }
                    }
                    
                    if (-not $found) {
                        Format-KeywordsDisplay -title $currentTitle -keywords @()
                        Update-KeywordCount -keywords @()
                        $excelTagsTextBox.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
                    }
                    
                    # Cleanup
                    $workbook.Close()
                    $excelApp.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    
                } catch {
                    Write-Debug "Error updating keywords display: $_"
                }
            }

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving to Excel: $_", "Error")
        }
    })

    # Update the keyword count display
    function Update-KeywordCount {
        param(
            [string[]]$keywords
        )
        
        $count = if ($keywords) { $keywords.Count } else { 0 }
        $maxKeywords = 50
        
        if ($count -gt $maxKeywords) {
            $keywordCountTextBox.ForeColor = [System.Drawing.Color]::Red
            $keywordCountTextBox.Text = "$count/$maxKeywords Keywords (OVER LIMIT!)"
        } else {
            $keywordCountTextBox.ForeColor = [System.Drawing.Color]::Black
            $keywordCountTextBox.Text = "$count/$maxKeywords Keywords"
        }
    }

    # Update the Find Tags button click handler
    $findTagsButton.Add_Click({
        try {
            if (-not $aiCheckBox.Checked) {
                Start-Process "https://copilot.microsoft.com"
            }
            
            $template = @"
Write me a description for an image which I want to sell on Shutterstock. The description should contain max. 200 characters and max. 50 keywords (keep it humanlike, perfect for a stock website, and do not repeat yourself while finding keywords).

File Name: "$($fileNameTextBox.Text)"
Title: "$($titleTextBox.Text)"
Subject: "$($subjectTextBox.Text)"
Description: "$($descriptionTextBox.Text)"
Tags: "$($tagsTextBox.Text)"

It's important that you support me with good keywords so the image can be perfectly advertised.
"@
            
            [System.Windows.Forms.Clipboard]::SetText($template)
            Show-TemporaryMessage "Template copied to clipboard" 2000 ([System.Drawing.Color]::FromArgb(51,153,255))
            $suggestedTagsTextBox.Text = $template
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error opening Copilot: $_", "Error")
        }
    })

    # Add click handler for Copy Suggestions button
    $copySuggestionsButton.Add_Click({
        $template = $suggestedTagsTextBox.Text
        [System.Windows.Forms.Clipboard]::SetText($template)
        Show-TemporaryMessage "Suggestions copied to clipboard" 2000 ([System.Drawing.Color]::FromArgb(51,153,255))
    })

    [System.Windows.Forms.Application]::Run($form)
}

# Main script execution
[System.Windows.Forms.Application]::EnableVisualStyles()

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff"
$openFileDialog.Multiselect = $true
$openFileDialog.Title = "Select Images"

if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $filePaths = $openFileDialog.FileNames
    Create-MetadataEditorForm -filePaths $filePaths
} else {
    [System.Windows.Forms.MessageBox]::Show("No files selected. Exiting.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}