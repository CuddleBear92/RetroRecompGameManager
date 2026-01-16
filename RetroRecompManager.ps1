# RetroRecompManager.ps1
# A unified UI for managing and updating retro game recompilations from GitHub.

# Ensure TLS 1.2 is used for secure connections
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Load required assemblies for UI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

# Get the script's directory (robust for both script and EXE)
if ($MyInvocation.MyCommand.Path) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $assemblyLocation = [System.Reflection.Assembly]::GetExecutingAssembly().Location
    if (-not [string]::IsNullOrEmpty($assemblyLocation)) {
        $scriptDir = [System.IO.Path]::GetDirectoryName($assemblyLocation)
    } else {
        $scriptDir = [System.IO.Directory]::GetCurrentDirectory()  # Fallback to process current dir
        if ([string]::IsNullOrEmpty($scriptDir)) {
            $scriptDir = $env:USERPROFILE  # Ultimate fallback to user profile
        }
    }
}

# Path to config file
$configFile = Join-Path $scriptDir "manager_config.json"

# Default config if not exists (user can edit manually)
if (-not (Test-Path $configFile)) {
    $defaultPattern = if ($IsWindows) { "*Windows*.zip" } elseif ($IsLinux) { "*Linux*.zip" } else { "*Windows*.zip" }
    $defaultConfig = @{
        installRoot = "Games"  # Relative to scriptDir
        savesRoot = "Saves"  # Relative to scriptDir, for backup saves
        enableSaveBackup = $true  # Enable/disable save backups after play
        githubToken = $null  # Optional: Add your GitHub personal access token here to avoid rate limits
        games = @(
            @{
                assetPattern = "*.zip"
                title = "Chameleon Twist: Recompiled"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null  # Detected EXE name after install
                savePath = "%localappdata%\\ChameleonTwistRecompiled"  # Hardcoded save path
                repo = "Rainchus/ChameleonTwist1-JP-Recomp"
                cleanup = $true
            },
            @{
                assetPattern = "*Windows*.zip"
                title = "Dinosaur Planet: Recompiled"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = "%localappdata%\\DinoPlanetRecompiled"
                repo = "DinosaurPlanetRecomp/dino-recomp"
                cleanup = $true
            },
            @{
                assetPattern = "*Windows*.zip"
                title = "Dr Mario 64: Recompiled"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = "%localappdata%\\drmario64_recomp"
                repo = "theboy181/drmario64_recomp_plus"
                cleanup = $true
            },
            @{
                assetPattern = "*Windows*.zip"
                title = "Duke Nukem Zero Hour: Recompiled"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = "%localappdata%\\DNZHRecompiled"
                repo = "sonicdcer/DNZHRecomp"
                cleanup = $true
            },
            @{
                assetPattern = "*Windows*.zip"
                title = "Goemon 64: Recompiled"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = "%localappdata%\\Goemon64Recompiled"
                repo = "klorfmorf/Goemon64Recomp"
                cleanup = $true
            },
            @{
                assetPattern = "*Windows*.zip"
                title = "MarioKart 64: Recompiled"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = "%localappdata%\\MarioKart64Recompiled"
                repo = "sonicdcer/MarioKart64Recomp"
                cleanup = $true
            },
            @{
                assetPattern = "*x86_64-windows.zip"
                title = "Perfect Dark port"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = $null
                repo = "fgsfdsfgs/perfect_dark"
                cleanup = $true
            },
            @{
                assetPattern = "*Windows*.zip"
                title = "Sonic Unleashed Recompiled"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = $null
                repo = "hedge-dev/UnleashedRecomp"
                cleanup = $true
            },
            @{
                assetPattern = "*Windows*.zip"
                title = "Starfox 64: Recompiled"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = "%localappdata%\\Starfox64Recompiled"
                repo = "sonicdcer/Starfox64Recomp"
                cleanup = $true
            },
            @{
                assetPattern = "*Launcher*.zip"
                title = "Super Metroid Launcher"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = $null
                repo = "RadzPrower/Super-Metroid-Launcher"
                cleanup = $true
            },
            @{
                assetPattern = "*Windows*.zip"
                title = "WipeOut Phantom Edition"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = $null
                repo = "wipeout-phantom-edition/wipeout-phantom-edition"
                cleanup = $true
            },
            @{
                assetPattern = "*Windows*.zip"
                title = "Zelda 64: Recompiled"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = "%localappdata%\\Zelda64Recompiled"
                repo = "Zelda64Recomp/Zelda64Recomp"
                cleanup = $true
            },
            @{
                assetPattern = "*Launcher*.zip"
                title = "Zelda 3 Launcher"
                currentVersion = $null
                cachedLatestVersion = $null
                lastChecked = $null
                playExe = $null
                savePath = $null
                repo = "RadzPrower/Zelda-3-Launcher"
                cleanup = $true
            }
        )
    }
    $defaultConfig | ConvertTo-Json -Depth 3 | Set-Content $configFile
    Write-Host "Created default config file at $configFile. Edit it to add or modify games and settings."
}

# Read config file
try {
    $config = Get-Content $configFile | ConvertFrom-Json
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to parse manager_config.json: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

# Validate config
if (-not $config.games) {
    [System.Windows.Forms.MessageBox]::Show("Config file must contain a 'games' array.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}
if (-not $config.installRoot) {
    $config | Add-Member -MemberType NoteProperty -Name installRoot -Value "Games"
    $config | ConvertTo-Json -Depth 3 | Set-Content $configFile
}
if (-not $config.savesRoot) {
    $config | Add-Member -MemberType NoteProperty -Name savesRoot -Value "Saves"
    $config | ConvertTo-Json -Depth 3 | Set-Content $configFile
}
if ($null -eq $config.enableSaveBackup) {
    $config | Add-Member -MemberType NoteProperty -Name enableSaveBackup -Value $true
    $config | ConvertTo-Json -Depth 3 | Set-Content $configFile
}

# Compute root install path
$installRootPath = Join-Path $scriptDir $config.installRoot
if (-not (Test-Path $installRootPath)) {
    New-Item -ItemType Directory -Path $installRootPath -Force | Out-Null
}

# Compute saves root path
$savesRootPath = Join-Path $scriptDir $config.savesRoot
if (-not (Test-Path $savesRootPath)) {
    New-Item -ItemType Directory -Path $savesRootPath -Force | Out-Null
}

# Function to sanitize folder name
function Sanitize-FolderName {
    param ([string]$name)
    $name -replace '[^\w\- ]', '' -replace ' ', '_'
}

# Function to show GitHub token instructions
function Show-GitHubTokenInstructions {
    $message = "GitHub API rate limit exceeded. To fix this:`n`n" +
               "1. Go to GitHub > Settings > Developer settings > Personal access tokens > Tokens (classic) > Generate new token.`n" +
               "2. Name it, select 'repo' scope (or none for public), generate.`n" +
               "3. Copy the token (ghp_...).`n" +
               "4. Edit manager_config.json, add `"githubToken`": `"your_token_here`"` at the top level.`n" +
               "5. Save and restart the app."
    [System.Windows.Forms.MessageBox]::Show($message, "Rate Limit Exceeded", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
}

# Function to get latest release info from GitHub
function Get-LatestRelease {
    param (
        [string]$repo
    )
    $apiUrl = "https://api.github.com/repos/$repo/releases/latest"
    $headers = @{
        "Accept" = "application/vnd.github.v3+json"
        "User-Agent" = "PowerShell-Manager"
    }
    if ($config.githubToken) {
        $headers["Authorization"] = "token $($config.githubToken)"
    }
    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get
        return $response
    } catch {
        if ($_.Exception.Response.StatusCode.Value__ -eq 403 -and $_.ErrorDetails.Message -match "rate limit") {
            Show-GitHubTokenInstructions
        }
        return $null
    }
}

# Function to update a game with progress
function Update-Game {
    param (
        [int]$index
    )
    $game = $config.games[$index]
    $release = Get-LatestRelease -repo $game.repo
    if (-not $release) {
        [System.Windows.Forms.MessageBox]::Show("Failed to fetch latest release for $($game.title).", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $asset = $release.assets | Where-Object { $_.name -like $game.assetPattern } | Select-Object -First 1
    if (-not $asset) {
        [System.Windows.Forms.MessageBox]::Show("No asset found matching pattern '$($game.assetPattern)' for $($game.title).", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $downloadUrl = $asset.browser_download_url
    $assetName = $asset.name
    $outputFile = Join-Path $scriptDir $assetName  # Temp download in script dir

    $installPath = Join-Path $installRootPath (Sanitize-FolderName $game.title)
    if (-not (Test-Path $installPath)) {
        New-Item -ItemType Directory -Path $installPath -Force | Out-Null
    }

    # Show progress bar and status
    $statusLabel.Text = "Downloading $($game.title)..."
    $progressBar.Value = 0
    $progressBar.Visible = $true
    $statusLabel.Visible = $true

    # Async download with progress
    try {
        $httpClient = New-Object System.Net.Http.HttpClient
        $response = $httpClient.GetAsync($downloadUrl).Result
        $response.EnsureSuccessStatusCode()
        $totalBytes = $response.Content.Headers.ContentLength
        $stream = $response.Content.ReadAsStreamAsync().Result
        $fileStream = New-Object System.IO.FileStream($outputFile, [System.IO.FileMode]::Create)
        $buffer = New-Object byte[] 8192
        $bytesRead = 0
        $read = $stream.Read($buffer, 0, $buffer.Length)
        while ($read -gt 0) {
            $fileStream.Write($buffer, 0, $read)
            $bytesRead += $read
            if ($totalBytes -gt 0) {
                $progress = [int](($bytesRead / $totalBytes) * 100)
                $progressBar.Value = $progress
            }
            [System.Windows.Forms.Application]::DoEvents()  # Keep UI responsive
            $read = $stream.Read($buffer, 0, $buffer.Length)
        }
        $fileStream.Close()
        $stream.Close()
    } catch {
        $progressBar.Visible = $false
        $statusLabel.Visible = $false
        [System.Windows.Forms.MessageBox]::Show("Failed to download $assetName for $($game.title): $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Extraction with progress (approximate, since Expand-Archive doesn't report progress; use ZipFile for better control)
    $statusLabel.Text = "Extracting $($game.title)..."
    $progressBar.Value = 0
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($outputFile)
        $entries = $zip.Entries
        $totalEntries = $entries.Count
        $currentEntry = 0
        foreach ($entry in $entries) {
            $targetPath = Join-Path $installPath $entry.FullName
            if ($entry.FullName.EndsWith('/')) {
                New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
            } else {
                $parentDir = Split-Path $targetPath -Parent
                if (-not (Test-Path $parentDir)) {
                    New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
                }
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $targetPath, $true)
            }
            $currentEntry++
            $progress = [int](($currentEntry / $totalEntries) * 100)
            $progressBar.Value = $progress
            [System.Windows.Forms.Application]::DoEvents()  # Keep UI responsive
        }
        $zip.Dispose()
    } catch {
        $progressBar.Visible = $false
        $statusLabel.Visible = $false
        [System.Windows.Forms.MessageBox]::Show("Failed to extract $assetName for $($game.title): $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Remove-Item $outputFile -Force -ErrorAction SilentlyContinue
        return
    }

    # Detect playExe if not set (scan for .exe in installPath)
    if (-not $game.playExe) {
        $exes = Get-ChildItem -Path $installPath -Filter *.exe -Recurse | Select-Object -First 1
        if ($exes) {
            $config.games[$index].playExe = $exes.FullName.Replace($installPath, "").TrimStart('\')  # Store relative path
        }
    }

    # Update current version and cache
    $config.games[$index].currentVersion = $release.tag_name
    $config.games[$index].cachedLatestVersion = $release.tag_name
    $config.games[$index].lastChecked = [DateTime]::UtcNow.ToString("o")
    $config | ConvertTo-Json -Depth 3 | Set-Content $configFile

    # Cleanup
    if ($game.cleanup -eq $true) {
        Remove-Item $outputFile -Force -ErrorAction SilentlyContinue
    }

    # Hide progress
    $progressBar.Visible = $false
    $statusLabel.Visible = $false

    [System.Windows.Forms.MessageBox]::Show("Updated $($game.title) to version $($release.tag_name).", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    # Refresh UI
    Refresh-Grid
}

# Function to play the game
function Play-Game {
    param (
        [int]$index
    )
    $game = $config.games[$index]
    if (-not $game.playExe) {
        [System.Windows.Forms.MessageBox]::Show("No playable EXE found for $($game.title). Update the game first.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $installPath = Join-Path $installRootPath (Sanitize-FolderName $game.title)
    $exePath = Join-Path $installPath $game.playExe
    if (-not (Test-Path $exePath)) {
        [System.Windows.Forms.MessageBox]::Show("EXE not found at $exePath for $($game.title).", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Start process (no time tracking)
    $process = Start-Process -FilePath $exePath -WorkingDirectory $installPath -PassThru
    $process.WaitForExit()

    # Backup saves if enabled and savePath set
    if ($config.enableSaveBackup -and $game.savePath) {
        $expandedSavePath = [Environment]::ExpandEnvironmentVariables($game.savePath)
        if (Test-Path $expandedSavePath) {
            $backupSavePath = Join-Path $savesRootPath (Sanitize-FolderName $game.title)
            if (-not (Test-Path $backupSavePath)) {
                New-Item -ItemType Directory -Path $backupSavePath -Force | Out-Null
            }
            Copy-Item -Path "$expandedSavePath\*" -Destination $backupSavePath -Recurse -Force
        }
    }

    # Refresh UI (if needed)
    Refresh-Grid
}

# Function to open install folder
function Open-InstallFolder {
    param (
        [int]$index
    )
    $game = $config.games[$index]
    $installPath = Join-Path $installRootPath (Sanitize-FolderName $game.title)
    if (Test-Path $installPath) {
        Invoke-Item $installPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Install path not found for $($game.title).", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to remove a game
function Remove-Game {
    param (
        [int]$index
    )
    $game = $config.games[$index]
    $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove $($game.title)?", "Confirm Removal", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::No) {
        return
    }

    $deleteFiles = [System.Windows.Forms.MessageBox]::Show("Also delete the game files in the install folder?", "Delete Files?", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($deleteFiles -eq [System.Windows.Forms.DialogResult]::Yes) {
        $installPath = Join-Path $installRootPath (Sanitize-FolderName $game.title)
        if (Test-Path $installPath) {
            Remove-Item $installPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # Remove from config by index
    if ($config.games.Count -gt 1) {
        if ($index -eq 0) {
            $config.games = $config.games[1..($config.games.Count - 1)]
        } elseif ($index -eq ($config.games.Count - 1)) {
            $config.games = $config.games[0..($index - 1)]
        } else {
            $config.games = $config.games[0..($index - 1)] + $config.games[($index + 1)..($config.games.Count - 1)]
        }
    } else {
        $config.games = @()
    }

    $config | ConvertTo-Json -Depth 3 | Set-Content $configFile

    # Refresh UI
    Refresh-Grid
}

# Function to refresh the data grid with caching
function Refresh-Grid {
    $dataGrid.Rows.Clear()
    $needsSave = $false
    for ($i = 0; $i -lt $config.games.Count; $i++) {
        $game = $config.games[$i]
        $latestVersion = "Unknown"
        $release = $null
        $lastCheckedDate = if ($game.lastChecked) { [DateTime]::Parse($game.lastChecked) } else { [DateTime]::MinValue }
        if (([DateTime]::UtcNow - $lastCheckedDate).TotalDays -gt 1) {  # Check if more than 1 day old
            $release = Get-LatestRelease -repo $game.repo
            if ($release) {
                $latestVersion = $release.tag_name
                $config.games[$i].cachedLatestVersion = $latestVersion
                $config.games[$i].lastChecked = [DateTime]::UtcNow.ToString("o")
                $needsSave = $true
            }
        } else {
            $latestVersion = if ($game.cachedLatestVersion) { $game.cachedLatestVersion } else { "Unknown" }
        }

        $currentVersion = if ($game.currentVersion) { $game.currentVersion } else { "Not Installed" }
        $status = if ($currentVersion -eq $latestVersion -and $currentVersion -ne "Not Installed") { "Up to Date" } else { "Update Available" }

        $rowIndex = $dataGrid.Rows.Add($game.title, $currentVersion, $latestVersion, $status)
        if ($release) {
            $dataGrid.Rows[$rowIndex].Cells[2].Tag = $release.html_url  # Store release URL in Tag
        }
    }

    if ($needsSave) {
        $config | ConvertTo-Json -Depth 3 | Set-Content $configFile
    }

    # Auto-resize columns to fit content
    $dataGrid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)

    # Adjust form and grid width to fit the content (add padding for borders/scrollbars)
    $totalColumnWidth = ($dataGrid.Columns | ForEach-Object { $_.Width } | Measure-Object -Sum).Sum
    $newWidth = $totalColumnWidth + 40  # Extra padding for borders, potential vertical scrollbar, etc.
    $minWidth = 800  # Increased min width to accommodate progress elements stably
    $form.Width = [Math]::Max($newWidth, $minWidth)
    $dataGrid.Width = $form.Width - 20  # Adjust grid to fit form
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Retro Recomp Build Manager"
$form.Size = New-Object System.Drawing.Size(900, 400)  # Initial size; will be adjusted
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.ForeColor = [System.Drawing.Color]::White

# Create DataGridView
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Size = New-Object System.Drawing.Size(880, 300)  # Initial; will be adjusted
$dataGrid.Location = New-Object System.Drawing.Point(10, 10)
$dataGrid.ReadOnly = $true
$dataGrid.AllowUserToAddRows = $false
$dataGrid.AllowUserToDeleteRows = $false
$dataGrid.RowHeadersVisible = $false
$dataGrid.BackgroundColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$dataGrid.GridColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$dataGrid.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$dataGrid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dataGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$dataGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$dataGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dataGrid.EnableHeadersVisualStyles = $false

# Columns (initial widths; will auto-resize later)
$dataGrid.Columns.Add("Title", "Title") | Out-Null
$dataGrid.Columns[0].Width = 250
$dataGrid.Columns.Add("CurrentVersion", "Current Version") | Out-Null
$dataGrid.Columns[1].Width = 150

# Latest Version as Link Column
$latestVersionColumn = New-Object System.Windows.Forms.DataGridViewLinkColumn
$latestVersionColumn.HeaderText = "Latest Version"
$latestVersionColumn.Width = 150
$latestVersionColumn.LinkColor = [System.Drawing.Color]::LightBlue
$latestVersionColumn.ActiveLinkColor = [System.Drawing.Color]::Blue
$latestVersionColumn.VisitedLinkColor = [System.Drawing.Color]::Purple
$dataGrid.Columns.Add($latestVersionColumn) | Out-Null

$dataGrid.Columns.Add("Status", "Status") | Out-Null
$dataGrid.Columns[3].Width = 150

# Update button column
$updateBtnColumn = New-Object System.Windows.Forms.DataGridViewButtonColumn
$updateBtnColumn.HeaderText = "Update"
$updateBtnColumn.Text = "Update"
$updateBtnColumn.UseColumnTextForButtonValue = $true
$updateBtnColumn.Width = 100
$updateBtnColumn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$dataGrid.Columns.Add($updateBtnColumn) | Out-Null

# Play button column
$playBtnColumn = New-Object System.Windows.Forms.DataGridViewButtonColumn
$playBtnColumn.HeaderText = "Play"
$playBtnColumn.Text = "Play"
$playBtnColumn.UseColumnTextForButtonValue = $true
$playBtnColumn.Width = 100
$playBtnColumn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$dataGrid.Columns.Add($playBtnColumn) | Out-Null

# Open Folder button column
$openBtnColumn = New-Object System.Windows.Forms.DataGridViewButtonColumn
$openBtnColumn.HeaderText = "Open Folder"
$openBtnColumn.Text = "Open"
$openBtnColumn.UseColumnTextForButtonValue = $true
$openBtnColumn.Width = 100
$openBtnColumn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$dataGrid.Columns.Add($openBtnColumn) | Out-Null

# Remove button column
$removeBtnColumn = New-Object System.Windows.Forms.DataGridViewButtonColumn
$removeBtnColumn.HeaderText = "Remove"
$removeBtnColumn.Text = "Remove"
$removeBtnColumn.UseColumnTextForButtonValue = $true
$removeBtnColumn.Width = 100
$removeBtnColumn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$dataGrid.Columns.Add($removeBtnColumn) | Out-Null

# Handle cell formatting for status color
$dataGrid.Add_CellFormatting({
    if ($this.Columns[$_.ColumnIndex].Name -eq "Status") {
        $status = $_.Value
        if ($status -eq "Update Available") {
            $_.CellStyle.ForeColor = [System.Drawing.Color]::LimeGreen
        } else {
            $_.CellStyle.ForeColor = [System.Drawing.Color]::Gray
        }
    }
})

# Handle button and link clicks
$dataGrid.Add_CellContentClick({
    $columnIndex = $this.CurrentCell.ColumnIndex
    $rowIndex = $this.CurrentCell.RowIndex
    if ($columnIndex -eq 2) {  # Latest Version link
        $url = $this.Rows[$rowIndex].Cells[2].Tag
        if ($url) {
            Start-Process $url
        }
    } elseif ($columnIndex -eq 4) {  # Update button
        Update-Game -index $rowIndex
    } elseif ($columnIndex -eq 5) {  # Play button
        Play-Game -index $rowIndex
    } elseif ($columnIndex -eq 6) {  # Open Folder button
        Open-InstallFolder -index $rowIndex
    } elseif ($columnIndex -eq 7) {  # Remove button
        Remove-Game -index $rowIndex
    }
})

$form.Controls.Add($dataGrid)

# Status label (hidden initially, fixed left after buttons, reduced width)
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Size = New-Object System.Drawing.Size(200, 20)
$statusLabel.Location = New-Object System.Drawing.Point(230, 325)  # Y offset to center with buttons (buttons 30h, this 20h -> +5)
$statusLabel.ForeColor = [System.Drawing.Color]::White
$statusLabel.Visible = $false
$form.Controls.Add($statusLabel)

# Progress bar (hidden initially, closer after status label)
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Size = New-Object System.Drawing.Size(200, 20)
$progressBar.Location = New-Object System.Drawing.Point(440, 325)  # After status (230+200+10=440), Y centered
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Adjust positions on form resize (only Y, X fixed)
$form.Add_Resize({
    $refreshBtn.Location = New-Object System.Drawing.Point(10, ($form.Height - 80))
    $addBtn.Location = New-Object System.Drawing.Point(120, ($form.Height - 80))
    $statusLabel.Location = New-Object System.Drawing.Point(230, ($form.Height - 75))  # +5 for center
    $progressBar.Location = New-Object System.Drawing.Point(440, ($form.Height - 75))  # +5 for center
    $dataGrid.Height = $form.Height - 100
})

# Refresh button
$refreshBtn = New-Object System.Windows.Forms.Button
$refreshBtn.Text = "Refresh"
$refreshBtn.Size = New-Object System.Drawing.Size(100, 30)
$refreshBtn.Location = New-Object System.Drawing.Point(10, 320)
$refreshBtn.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$refreshBtn.ForeColor = [System.Drawing.Color]::White
$refreshBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$refreshBtn.Add_Click({ Refresh-Grid })
$form.Controls.Add($refreshBtn)

# Add Game button
$addBtn = New-Object System.Windows.Forms.Button
$addBtn.Text = "Add Game"
$addBtn.Size = New-Object System.Drawing.Size(100, 30)
$addBtn.Location = New-Object System.Drawing.Point(120, 320)
$addBtn.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$addBtn.ForeColor = [System.Drawing.Color]::White
$addBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$addBtn.Add_Click({ Add-NewGame })
$form.Controls.Add($addBtn)

# Initial refresh (which will resize)
Refresh-Grid

# Show form
$form.ShowDialog() | Out-Null