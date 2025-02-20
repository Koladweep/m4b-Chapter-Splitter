# PowerShell script to split audiobook chapters using FFmpeg and FFprobe
# Ensure ffmpeg and ffprobe are added to the system PATH

# Open File Dialog to select the audiobook file
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'Audio Files (*.mp4,*.m4a,*.m4b,mp3)|*.mp4;*.m4a;*.m4b;*.mp3|All files (*.*)|*.*'
    Title = 'Select your audiobook file'
}
[void]$FileBrowser.ShowDialog()
$audioBook = $FileBrowser.FileName
$Book = [System.IO.Path]::GetFileNameWithoutExtension($audioBook)

# Extract book title using FFmpeg
$booktitle = (ffmpeg -i $audioBook -f ffmetadata - 2>$null | Select-String -Pattern 'title=' | Select-Object -First 1).Line -replace 'title=', ''

# Default output path
$defaultPath = Join-Path -Path ([Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)) -ChildPath $Book

# Open Folder Dialog to select output folder
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    Description = 'Select a folder to save your output files'
    RootFolder = [Environment+SpecialFolder]::MyComputer
    SelectedPath = [Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
}

$path = $defaultPath
if ($FolderBrowser.ShowDialog() -eq 'OK') {
    $path = $FolderBrowser.SelectedPath
}
else {
    Write-Host 'No folder selected, saving to desktop'
    if (-not (Test-Path -Path $path)) {
        New-Item -ItemType Directory -Path $path -Force
    }
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

    # Construct output filename
    $outputFile = Join-Path -Path $path -ChildPath "Chapter-$ChapterNum.mp3"

    # Run FFmpeg command
    ffmpeg -i "$audioBook" -ss $Starttime -to $endtime -f mp3 -metadata track="$ChapterNum" -metadata title="Chapter - $ChapterNum - $booktitle" -metadata album="$booktitle" -y "$outputFile"

    $Starttime = $endtime
}

Write-Host 'Chapter splitting completed successfully!'
