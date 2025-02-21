# Load the necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to create a custom form with a checkbox list
function Show-CheckboxForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Files and Folders"
    $form.Size = New-Object System.Drawing.Size(600, 400)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = true

    $listBox = New-Object System.Windows.Forms.CheckedListBox
    $listBox.Size = New-Object System.Drawing.Size(560, 300)
    $listBox.Location = New-Object System.Drawing.Point(10, 10)
    $form.Controls.Add($listBox)

    $addFilesButton = New-Object System.Windows.Forms.Button
    $addFilesButton.Text = "Add Files"
    $addFilesButton.Location = New-Object System.Drawing.Point(10, 320)
    $addFilesButton.Size = New-Object System.Drawing.Size(100, 30)
    $addFilesButton.Add_Click({
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('Desktop')
            Filter = 'Audio Files (*.mp4,*.m4a,*.m4b,*.mp3)|*.mp4;*.m4a;*.m4b;*.mp3|All files (*.*)|*.*'
            Title = 'Select your audiobook files'
            Multiselect = $true
            CheckFileExists = $true
            CheckPathExists = $true
        }
        [void]$FileBrowser.ShowDialog()
        foreach ($file in $FileBrowser.FileNames) {
            if (-not $listBox.Items.Contains($file)) {
                $listBox.Items.Add($file, $true)
            }
        }
    })
    $form.Controls.Add($addFilesButton)

    $addFoldersButton = New-Object System.Windows.Forms.Button
    $addFoldersButton.Text = "Add Folders"
    $addFoldersButton.Location = New-Object System.Drawing.Point(120, 320)
    $addFoldersButton.Size = New-Object System.Drawing.Size(100, 30)
    $addFoldersButton.Add_Click({
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            Description = 'Select folders containing your audiobook files'
            RootFolder = [Environment+SpecialFolder]::Desktop
            SelectedPath = [Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
        }
        if ($FolderBrowser.ShowDialog() -eq 'OK') {
            $folder = $FolderBrowser.SelectedPath
            if (-not $listBox.Items.Contains($folder)) {
                $listBox.Items.Add($folder, $true)
            }
        }
    })
    $form.Controls.Add($addFoldersButton)

    $selectOutputFolderButton = New-Object System.Windows.Forms.Button
    $selectOutputFolderButton.Text = "Select Output Folder"
    $selectOutputFolderButton.Location = New-Object System.Drawing.Point(230, 320)
    $selectOutputFolderButton.Size = New-Object System.Drawing.Size(150, 30)
    $selectOutputFolderButton.Add_Click({
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            Description = 'Select a folder to save your output files'
            RootFolder = [Environment+SpecialFolder]::Desktop
            SelectedPath = [Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
        }
        if ($FolderBrowser.ShowDialog() -eq 'OK') {
            $global:outputFolder = $FolderBrowser.SelectedPath
        }
    })
    $form.Controls.Add($selectOutputFolderButton)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Execute"
    $okButton.Location = New-Object System.Drawing.Point(470, 320)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Add_Click({
        if ($listBox.CheckedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No files or folders selected. Exiting script.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        } elseif (-not $global:outputFolder) {
            $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
                Description = 'Select a folder to save your output files'
                RootFolder = [Environment+SpecialFolder]::Desktop
                SelectedPath = [Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
            }
            if ($FolderBrowser.ShowDialog() -eq 'OK') {
                $global:outputFolder = $FolderBrowser.SelectedPath
                $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            } else {
                $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            }
        } else {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        }
        $form.Close()
    })
    $form.Controls.Add($okButton)

    $form.ShowDialog()
    return $listBox.CheckedItems
}

# Show the custom form and get the selected files and folders
$selectedItems = Show-CheckboxForm

# Collect all audiobook files from the selected items
$audioBooks = @()
foreach ($item in $selectedItems) {
    if (Test-Path $item -PathType Container) {
        $audioBooks += Get-ChildItem -Path $item -Recurse -Include *.mp4, *.m4a, *.m4b, *.mp3 | ForEach-Object { $_.FullName }
    } elseif (Test-Path $item -PathType Leaf) {
        $audioBooks += $item
    }
}

# Remove duplicate paths
$audioBooks = $audioBooks | Sort-Object -Unique

# count and track number of jobs
$global:jobCount = $audioBooks.Length
if($jobCount -eq 0) {
    Write-Host 'No audiobook files selected. Exiting script.'
    exit
}
# to keep count of completed jobs
$global:doneCount= 0

# Check if output folder is selected
if (-not $global:outputFolder) {
    Write-Host 'No output folder selected, saving to desktop'
    $global:outputFolder = [Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
}

# Function to process each audiobook
function Process-Audiobook {
    param (
        [string]$audioBook,
        [string]$path
    )

    $Book = [System.IO.Path]::GetFileNameWithoutExtension($audioBook)

    # Extract book title using FFmpeg
    $booktitle = (ffmpeg -i "$audioBook" -f ffmetadata - 2>$null | Select-String -Pattern 'title=' | Select-Object -First 1).Line -replace 'title=', ''

    # Default output path for each book
    $defaultPath = Join-Path -Path $path -ChildPath $Book
    if (-not (Test-Path -Path $defaultPath)) {
        New-Item -ItemType Directory -Path $defaultPath -Force
    }

    # Get chapter end times from FFprobe
    $json = (ffprobe.exe -i "$audioBook" -show_chapters -print_format json) | Out-String | ConvertFrom-Json
    $Endtimes = $json.chapters.end_time
    $Starttime = 0.000000

    # Loop through chapters and create output files
    for ($i = 0; $i -lt $Endtimes.length; $i++) {
        $ChapterNum = $i + 1
        $endtime = $Endtimes[$i]

        # Format chapter number as 001, 002, ... 999
        if ($ChapterNum -le 9) {
            $ChapterNum = "00$ChapterNum"
        } elseif ($ChapterNum -le 99) {
            $ChapterNum = "0$ChapterNum"
        }
        $trackNum = $i + 1
        # Construct output filename
        $outputFile = Join-Path -Path $defaultPath -ChildPath "Chapter - $ChapterNum-$Book.mp3"

        # Run FFmpeg command
        ffmpeg -i "$audioBook" -ss $Starttime -to $endtime -f mp3 -metadata track="$trackNum" -metadata title="Chapter - $ChapterNum - $booktitle" -metadata album="$booktitle" -y "$outputFile"

        $Starttime = $endtime
    }

    Write-Host "Chapter splitting for '$Book' completed successfully!"
    $doneCount++
}
# Process each audiobook
$global:lastBook = $null
foreach ($audioBook in $audioBooks) {
    Process-Audiobook -audioBook $audioBook -path $global:outputFolder
    $lastBook=$audioBook
}

#completed. show status message and open folder if prompted

if ($doneCount -eq $jobCount -and $jobCount -gt 0) {
    $message = 'All chapter splitting completed successfully! Do you want to open the output folder?'
} elseif ($doneCount -lt $jobCount -and $jobCount -gt 0) {
    $message = "Completed: $doneCount out of $jobCount jobs. Do you want to open the output folder?"
}

if ($message -ne "") {
    Add-Type -AssemblyName System.Windows.Forms
    # Show a message box with the status message
    $result = [System.Windows.Forms.MessageBox]::Show($message, 'Open Folder?', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)

    # Open the output folder if "Yes" (Open) button is clicked
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$lastBook`""
    }
}
