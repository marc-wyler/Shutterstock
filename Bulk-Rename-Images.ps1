##########################################################################################
## Author: Marc Wyler                                                                   ##
## Created: 14.02.2025                                                                  ##
## Shutterstock Portfolio: https://www.shutterstock.com/g/Capture+Sunny?rid=460175361   ##
##########################################################################################

# Open an explorer window to select the folder
Add-Type -AssemblyName System.Windows.Forms

# Create a hidden form to act as the owner of the folder dialog
$form = New-Object System.Windows.Forms.Form -Property @{TopMost = $true; Width = 0; Height = 0; ShowInTaskbar = $false}

# Display introductory message
Write-Host "Welcome to the JPG Renaming Script!" -ForegroundColor Green
Write-Host "This script will help you rename JPG files in a selected folder to follow the pattern 'BaseName_X.jpg'." -ForegroundColor Green
Write-Host "You'll be asked to select a folder and enter a base name. Let's get started!" -ForegroundColor Green

# Show the form and open the folder dialog
$form.Show()

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the folder where the JPG files are located"

# Set default folder location
$defaultFolder = "C:\Shutterstock Upload\"
if (Test-Path -Path $defaultFolder -PathType Container) {
    $folderBrowser.SelectedPath = $defaultFolder
}

# Show the dialog with the hidden form as the owner
if ($folderBrowser.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
    $form.Close()
    Write-Host "No folder selected. Exiting..." -ForegroundColor Red
    exit
}
$folderPath = $folderBrowser.SelectedPath
$form.Close()

# Check if the folder exists
if (!(Test-Path -Path $folderPath -PathType Container)) {
    Write-Host "The specified folder does not exist. Exiting..." -ForegroundColor Red
    exit
}

# Function to show a textbox for name confirmation
function Get-BaseName {
    param ($defaultName)
    
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Confirm Base Name"
    $inputForm.Size = New-Object System.Drawing.Size(300,150)
    $inputForm.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter or confirm the base name:"
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.AutoSize = $true
    $inputForm.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $defaultName
    $textBox.Location = New-Object System.Drawing.Point(10,30)
    $textBox.Size = New-Object System.Drawing.Size(260,20)
    $inputForm.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(50,60)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputForm.Controls.Add($okButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Text = "Exit"
    $exitButton.Location = New-Object System.Drawing.Point(150,60)
    $exitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $exitButton.ForeColor = "Red"
    $inputForm.Controls.Add($exitButton)

    $inputForm.AcceptButton = $okButton
    $inputForm.CancelButton = $exitButton

    $result = $inputForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        Write-Host "Operation canceled. Exiting..." -ForegroundColor Red
        Write-Host "The final base name used was: $defaultName" -ForegroundColor Green
        for ($i = 5; $i -ge 0; $i--) {
            Write-Host "Press 'r' to rename again or wait for the countdown to finish. ($i seconds remaining)" -ForegroundColor Yellow
            Start-Sleep -Seconds 1
            if ([System.Console]::KeyAvailable -and [System.Console]::ReadKey($true).Key -eq 'R') {
                Write-Host "Renaming process started again..." -ForegroundColor Green
                return $null
            }
        }
        exit
    }
}

# Initial base name input
$baseName = Read-Host "Enter the base name for the files"

$allRenamed = $false

do {
    # Get all JPG files in the folder
    $jpgFiles = Get-ChildItem -Path $folderPath -Filter "*.jpg" | Sort-Object Name

    # Check if there are any JPG files
    if ($jpgFiles.Count -eq 0) {
        Write-Host "No JPG files found in the folder. Exiting..." -ForegroundColor Yellow
        exit
    }

    # Regex pattern to check if a file already matches "BaseName_X.jpg"
    $pattern = "^" + [regex]::Escape($baseName) + "_(\d+)\.jpg$"

    # Find existing numbers to determine the next available one
    $existingNumbers = @()
    foreach ($file in $jpgFiles) {
        if ($file.Name -match $pattern) {
            $existingNumbers += [int]$matches[1]
        }
    }

    # Get the first available number
    $counter = 1
    while ($existingNumbers -contains $counter) {
        $counter++
    }

    # Rename only incorrectly named files
    $allRenamed = $true
    foreach ($file in $jpgFiles) {
        if ($file.Name -match $pattern) {
            Write-Host "Skipping: $($file.Name) (Already named correctly)" -ForegroundColor Cyan
            continue
        }

        # Find the next available number
        while ($existingNumbers -contains $counter) {
            $counter++
        }

        $newName = "${baseName}_$counter.jpg"
        $newPath = Join-Path -Path $folderPath -ChildPath $newName

        # Attempt to rename the file and suppress errors
        Rename-Item -Path $file.FullName -NewName $newName -ErrorAction SilentlyContinue

        # Check if the file was successfully renamed
        if (Test-Path -Path $newPath) {
            Write-Host "Renamed: $($file.Name) -> $newName"
            $existingNumbers += $counter
        } else {
            Write-Host "Could not rename: $($file.Name)" -ForegroundColor Yellow
            $allRenamed = $false
            break
        }
    }

    if (-not $allRenamed) {
        Write-Host "Not all files could be renamed. Exiting..." -ForegroundColor Red
        exit
    }

    Write-Host "Renaming completed!" -ForegroundColor Green
    Write-Host "The final base name used was: $baseName" -ForegroundColor Green
    for ($i = 5; $i -ge 0; $i--) {
        Write-Host "Press 'r' to rename again or wait for the countdown to finish. ($i seconds remaining)" -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        if ([System.Console]::KeyAvailable -and [System.Console]::ReadKey($true).Key -eq 'R') {
            Write-Host "Renaming process started again..." -ForegroundColor Green
            $baseName = Read-Host "Enter a new base name for the files"
            break
        }
    }
    if ($i -lt 0) {
        exit
    }

    # Show confirmation dialog
    $newBaseName = Get-BaseName -defaultName $baseName

    if (-not $newBaseName) {
        Write-Host "Operation canceled. Exiting..." -ForegroundColor Red
        Write-Host "The final base name used was: $baseName" -ForegroundColor Green
        for ($i = 5; $i -ge 0; $i--) {
            Write-Host "Press 'r' to rename again or wait for the countdown to finish. ($i seconds remaining)" -ForegroundColor Yellow
            Start-Sleep -Seconds 1
            if ([System.Console]::KeyAvailable -and [System.Console]::ReadKey($true).Key -eq 'R') {
                Write-Host "Renaming process started again..." -ForegroundColor Green
                $baseName = Read-Host "Enter a new base name for the files"
                break
            }
        }
        if ($i -lt 0) {
            exit
        }
    }

    $baseName = $newBaseName

} while ($true)