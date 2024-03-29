﻿<#
.SYNOPSIS
    Checks for or installs prerequisites, then uses BDFR and BDFR-HTML to generate a subreddit HTML archive.
    Creates root 'New-SubredditHTMLArchive' folder and subfolders under your %USERPROFILE% ($env:USERPROFILE) Documents folder.
    The reddit API returns a maximum of 1000 posts per BDFR pull, so only the newest 1000 posts will be included: https://github.com/reddit-archive/reddit/blob/master/r2/r2/lib/db/queries.py
    Script download URL from web browsers, so the code signature still works (Save As): https://raw.githubusercontent.com/mbarr564/powershell/master/New-SubredditHTMLArchive.ps1
.DESCRIPTION
    If you already have Python 3.9+, Git 2+, and GitHub CLI 2+ installed, you can skip this section.
    This script does NOT require administrator privileges to run, or to install the Python modules, WITHOUT the -InstallPackages parameter.
    On first run, you must include the -InstallPackages parameter, or manually install the below software packages before running this script.
    When installing these packages automatically, the user must confirm a UAC prompt for each package, allowing the installer to make changes to their computer.
        1. Git: https://github.com/git-for-windows/git/releases/ (only when manually installing)
        2. GitHub CLI: https://github.com/cli/cli/releases/ (only when manually installing)
            i. You'll need to launch cmd.exe and authenticate with 'gh auth login', and follow the prompts, pasting the OTP into your browser, after logging into your GitHub account (or make a new account).
        3. Python 3.9+ (includes pip): https://www.python.org/downloads/windows/ (only when manually installing)
            i. At beginning of install, YOU MUST CHECK 'Add Python 3.x to PATH'. (So PowerShell can call python.exe and pip.exe from anywhere)
    This script uses the following Python modules, which are detected and installed automatically via pip:
        1. BDFR: https://pypi.org/project/bdfr/
        2. BDFR-HTML: https://github.com/BlipRanger/bdfr-html
            i. When running setup.py to install BDFR-HTML (via script or manually), you may get an install error from Pillow about zlib being missing. You may need to run 'pip install pillow' from an elevated command prompt, so that Pillow installs correctly.
            ii. For manual BDFR-HTML install in case of Pillow install error: From an elevated CMD window, type these two quoted commands: 1) "cd %USERPROFILE%\Documents\BDFR\module_clone\bdfr-html", 2) "python.exe setup.py install"
            iii. https://stackoverflow.com/questions/64302065/pillow-installation-pypy3-missing-zlib
.PARAMETER Subreddit
    The name of the subreddit (as it appears after the /r/ in the URL) that will be archived.
.PARAMETER Subreddits
    An array of subreddit names (as they appear after the /r/ in the URL) that will be archived.
    Also generates a master index.html containing links to all of the other generated subreddit index.html files.
    All generated subreddit folders, files, and index pages, are automatically packaged into a ZIP file.
.PARAMETER InstallPackages
    The script will attempt to install ONLY MISSING pre-requisite packages: Python 3, GitHub, and Git
    When 'python.exe', 'gh.exe', or 'git.exe' are already in your $env:path, and executable from PowerShell, they will NOT be installed or modified.
.EXAMPLE
    PS> .\New-SubredditHTMLArchive.ps1 -Subreddit PowerShell
.EXAMPLE
    PS> .\New-SubredditHTMLArchive.ps1 -Subreddit PowerShell -InstallPackages
.EXAMPLE
    PS> .\New-SubredditHTMLArchive.ps1 -Subreddits 'PowerShell','Python','AmateurRadio','HackRF','GNURadio','OpenV2K','DataHoarder','AtheistHavens','Onions'
.EXAMPLE
    PS> .\New-SubredditHTMLArchive.ps1 -Subreddits (Get-Content "$($env:UserProfile)\Desktop\subreddit_list.txt")
.NOTES
    Last update: Sunday, March 13, 2022 6:03:29 PM
#>

param([string]$Subreddit, [ValidateCount(2,100)][string[]]$Subreddits, [switch]$InstallPackages)

## Subreddit parameters
if ($Subreddits)
{
    $Subreddits = $Subreddits | Select-Object -Unique #dupe removal
    [string[]]$badNames = @(); $Subreddits | ForEach-Object {if ($_ -notmatch "^[A-Z0-9_]{2,21}$"){$badNames += $_}}
    if ($badNames.count -gt 0){throw "Error: Subreddit name(s) failed regex validation: $($badNames -join ', ')"}
    [int]$maxMinutesPerSub = 100 #from experience BDFR can take about 1hr 40mins to finish pulling down the API maximum of 1000 records
    [int]$hoursToArchiveAll = ($maxMinutesPerSub * ($Subreddits.count)) / 60
    if ($hoursToArchiveAll -gt 24){$timeToArchiveAll = "over $($hoursToArchiveAll / 24) full day(s)"} else {$timeToArchiveAll = "$hoursToArchiveAll hours"}
    Write-Warning "Estimated maximum time to finish $($Subreddits.count) subreddits: $timeToArchiveAll ($((Get-Date).AddHours($hoursToArchiveAll)))"
}
else
{
    if (-not($Subreddit)){$Subreddit = Read-Host -Prompt 'Enter subreddit name to archive'}
    if (-not($Subreddit)){throw 'Error: Subreddit name is blank!'}
    if ($Subreddit -notmatch "^[A-Z0-9_]{2,21}$"){throw "Error: Subreddit name failed regex validation: $Subreddit"}
}

## Output / Logging paths
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Checking for root output folders ..."
[string]$rootFolder = "$($env:USERPROFILE)\Documents\New-SubredditHTMLArchive" #all script output is contained within this root directory
foreach ($folderName in @('JSON','HTML','ZIP','logs')){if (-not(Test-Path "$rootFolder\$folderName" -PathType Container)){New-Item -Path "$rootFolder\$folderName" -ItemType Directory -Force -ErrorAction Stop | Out-Null}}
[string]$logFolder = "$rootFolder\logs"; Get-ChildItem $logFolder | Where-Object  {$_.Name -like "*.1"} | ForEach-Object {Remove-Item -LiteralPath "$($_.FullName)" -ErrorAction SilentlyContinue} #clean empty log files with extension .1

## Check for command line utilities
[string[]]$missingExes = @()
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Checking for command line utilities ..."
foreach ($exeName in @('git','gh','python')){if (-not(Get-Command "$($exeName).exe" -ErrorAction SilentlyContinue | Where-Object {$_.Source -notlike "*\AppData\Local\Microsoft\WindowsApps\*"})){$missingExes += "$($exeName).exe"}} #exclude executable shortcuts under \WindowsApps\, that are really just placeholder executables that launch the Microsoft Store
if ($missingExes.count -gt 0)
{
    if ($InstallPackages)
    {
        ## Install command line utility packages
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Installing $($missingExes.count) missing package(s) via winget ..." -ForegroundColor Cyan
        [string]$OSName = ((Get-WmiObject -class Win32_OperatingSystem).Caption); [int]$OSBuild = [System.Environment]::OSVersion.Version.Build
        if ($OSBuild -lt 17763){throw "Error: this Windows OS build ($OSBuild) is older than build 17763, which is required for winget!"} #check that Windows version is new enough to support winget: https://github.com/microsoft/winget-cli#installing-the-client
        if (((Get-CimInstance -ClassName Win32_OperatingSystem).ProductType) -ne 1){throw "Error: this Windows OS is not a client/workstation edition, which is required for winget!"} #check that Windows is a client OS and not Server editions
        if (-not(Get-Command 'winget.exe' -ErrorAction SilentlyContinue)){
            switch -regex ($OSName){
                'Windows 10' {Start-Process 'https://www.microsoft.com/en-us/p/app-installer/9nblggh4nns1'; throw "Error: not found: winget.exe. Opened winget install link for Windows 10 in your default browser. Rerun script after install."}
                'Windows 11' {throw "Error: not found: winget.exe. Something is wrong. This should be included with Windows 11."}}}
        try {winget list | Out-Null} catch {throw "Error: running 'winget list' threw an exception:`n" + $error[0]} #trigger prompt to agree to the MS Store agreement terms (press Y, then Enter)
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Installing latest releases for: $($missingExes -join ', ') ..." -ForegroundColor Cyan
        switch ($missingExes){
            'python.exe' {try {winget install --id Python.Python.3} catch {winget install --id Python.Python.3}}
            'git.exe' {try {winget install --id Git.Git} catch {winget install --id Git.Git}}
            'gh.exe' {try {winget install --id GitHub.cli} catch {winget install --id GitHub.cli}}}

        ## Refresh PowerShell process environment (reloading %PATH% environment variables), otherwise this PowerShell session won't see the newly installed binary paths
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Refreshing PowerShell process environment variables ..."
        foreach ($hiveKey in @('HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment','HKCU:\Environment')){
            $variables = Get-Item $hiveKey #since PowerShell $env:path combines HKLM and HKCU environment variables
            $variables.GetValueNames() | ForEach-Object {
                if ($_ -ieq 'PATH'){
                    $value = $variables.GetValue($_)
                    switch -regex ($hiveKey){
                        "^HKLM" {$env:path = $value} #we have to read them seperately from the registry hives
                        "^HKCU" {$env:path += ";$value"}}}}} #then combine them into the PowerShell process path environment variable
        
        ## Recheck for command line utilities
        [string[]]$oldMissingExes = $missingExecs; [string[]]$missingExes = @()
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Rechecking for command line utilities after installs ..."
        foreach ($exeName in @('git','gh','python','pip'))
        {
            if (Get-Command "$($exeName).exe" -ErrorAction SilentlyContinue | Where-Object {$_.Source -notlike "*\AppData\Local\Microsoft\WindowsApps\*"})
            {
                [string]$exeVersion = (Invoke-Expression "$exeName --version") #invoke to process the string exe name and argument as a command, and string type to flatten any multi element string arrays / multi line returns
                Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Found $($exeName).exe version: $exeVersion"
            }
            else {$missingExes += "$($exeName).exe"}
        }
        if ($missingExes.count -gt 0){throw "[$(Get-Date -f HH:mm:ss.fff)] Error: Required command line utilities still missing: $($missingExes -join ', ')"}
        
        ## Authenticate GitHub CLI if newly installed
        if ('gh.exe' -in $oldMissingExes)
        {
            Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Launching GitHub CLI authentication process ..."
            $GitHubAuth = Start-Process "cmd.exe" -ArgumentList "/c gh auth login" -PassThru -Wait
            if ($GitHubAuth.ExitCode -ne 0)
            {
                Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Error: exit code $($GitHubAuth.ExitCode). Relaunching GitHub CLI authentication process ..." -ForegroundColor Yellow
                $GitHubAuthRetry = Start-Process "cmd.exe" -ArgumentList "/c gh auth login" -PassThru -Wait
                if ($GitHubAuthRetry.ExitCode -ne 0){throw "[$(Get-Date -f HH:mm:ss.fff)] Error: exit code $($GitHubAuthRetry.ExitCode). Failed to authenticate new GitHub CLI install! Please manually run cmd.exe, and complete the GitHub authentication process: gh auth login"}
            }
        }
    }
    else {throw "Error: Required command line utilities missing: $($missingExes -join ', ')! Rerun script with the -InstallPackages parameter?"}
}

## Check for Python modules BDFR and BDFR-HTML
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Checking for BDFR and BDFR-HTML Python modules ..."
[bool]$bdfrInstalled = $false
[bool]$bdfrhtmlInstalled = $false
[string[]]$installedPythonModules = @((pip list --disable-pip-version-check) | Where-Object {$_ -match "^[A-Z0-9-]{3,18}.*[0-9]{1,4}`.[0-9]{1,2}(`.[0-9]{1,2})?$"}) #strip name:value header
foreach ($installedPythonModule in $installedPythonModules)
{
    if ($installedPythonModule -like "bdfr *"){$bdfrInstalled = $true}
    if ($installedPythonModule -like "bdfrtohtml*"){$bdfrhtmlInstalled = $true}
}

## Install Python modules BDFR and BDFR-HTML
if (-not($bdfrInstalled))
{
    Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Installing BDFR Python module ..." -ForegroundColor Cyan
    $bdfrInstallProcess = Start-Process "python.exe" -ArgumentList "-m pip install bdfr --upgrade" -WindowStyle Hidden -RedirectStandardError "$logFolder\BDFR_installErrors.txt" -RedirectStandardOutput "$logFolder\BDFR_installStdOut.txt" -PassThru -Wait
    if ($bdfrInstallProcess.ExitCode -ne 0){Start-Process $logFolder; throw "[$(Get-Date -f HH:mm:ss.fff)] Error: Command: 'python.exe -m pip install bdfr --upgrade' returned exit code '$($bdfrInstallProcess.ExitCode)'! See opened logs folder."}
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Updating BDFR Python module ..." #these BDFR install commands have never failed for me.. but their error and output streams are redirected into log files anyway, just in case
    $bdfrUpdateProcess = Start-Process "python.exe" -ArgumentList "-m pip install bdfr --upgrade" -WindowStyle Hidden -RedirectStandardError "$logFolder\BDFR_updateErrors.txt" -RedirectStandardOutput "$logFolder\BDFR_updateStdOut.txt" -PassThru -Wait
    if ($bdfrUpdateProcess.ExitCode -ne 0){Start-Process $logFolder; throw "[$(Get-Date -f HH:mm:ss.fff)] Error: Command: 'python.exe -m pip install bdfr --upgrade' returned exit code '$($bdfrUpdateProcess.ExitCode)'! See opened logs folder."}
}
if (-not($bdfrhtmlInstalled))
{
    Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Installing BDFR-HTML Python module ..." -ForegroundColor Cyan
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Cloning GitHub repository for BDFR-HTML module ..."
    if (-not(Test-Path "$rootFolder\module_clone" -PathType Container)){New-Item -Path "$rootFolder\module_clone" -ItemType Directory -Force -ErrorAction Stop | Out-Null}
    $bdfrhtmlCloneProcess = Start-Process "gh.exe" -ArgumentList "repo clone BlipRanger/bdfr-html" -WorkingDirectory "$rootFolder\module_clone" -WindowStyle Hidden -RedirectStandardError "$logFolder\BDFRHTML_cloneErrors.txt" -RedirectStandardOutput "$logFolder\BDFRHTML_cloneStdOut.txt" -PassThru -Wait
    if ($bdfrhtmlCloneProcess.ExitCode -ne 0)
    {
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Warning: exit code $($bdfrhtmlCloneProcess.ExitCode) from GitHub during BDFR-HTML repository clone!" -ForegroundColor Yellow
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Retrying GitHub repository clone for BDFR-HTML module ..."
        Remove-Item -Path "$rootFolder\module_clone\bdfr-html" -Recurse -Force -ErrorAction Stop | Out-Null
        $bdfrhtmlCloneRetryProcess = Start-Process "gh.exe" -ArgumentList "repo clone BlipRanger/bdfr-html" -WorkingDirectory "$rootFolder\module_clone" -WindowStyle Hidden -RedirectStandardError "$logFolder\BDFRHTML_cloneRetryErrors.txt" -RedirectStandardOutput "$logFolder\BDFRHTML_cloneRetryStdOut.txt" -PassThru -Wait
        if ($bdfrhtmlCloneRetryProcess.ExitCode -ne 0){Start-Process $logFolder; throw "[$(Get-Date -f HH:mm:ss.fff)] Error while retrying BDFR-HTML repository clone. Command: 'gh.exe repo clone BlipRanger/bdfr-html' returned exit code '$($bdfrhtmlCloneRetryProcess.ExitCode)'! See opened logs folder."}
    }
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Upgrading pip module to latest version ..."
    Start-Process "python.exe" -ArgumentList "-m pip install --upgrade pip" -WindowStyle Hidden -Wait
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Running BDFR-HTML module setup script ..."
    $bdfrhtmlScriptProcess = Start-Process "python.exe" -ArgumentList "setup.py install" -WorkingDirectory "$rootFolder\module_clone\bdfr-html" -WindowStyle Hidden  -RedirectStandardError "$logFolder\BDFRHTML_installScriptErrors.txt" -RedirectStandardOutput "$logFolder\BDFRHTML_installScriptStdOut.txt" -PassThru -Wait
    if ($bdfrhtmlScriptProcess.ExitCode -ne 0)
    {
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Warning: exit code $($bdfrhtmlScriptProcess.ExitCode) from Python during BDFR-HTML module setup script!" -ForegroundColor Yellow
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Attempting alternate Pillow module install via pip ..."
        Start-Process "pip.exe" -ArgumentList "install pillow" -WindowStyle Hidden -Wait
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Rerunning BDFR-HTML module setup script ..."
        $bdfrhtmlScriptRetryProcess = Start-Process "python.exe" -ArgumentList "setup.py install" -WorkingDirectory "$rootFolder\module_clone\bdfr-html" -WindowStyle Hidden -RedirectStandardError "$logFolder\BDFRHTML_installScriptRetryErrors.txt" -RedirectStandardOutput "$logFolder\BDFRHTML_installScriptRetryStdOut.txt" -PassThru -Wait
        if ($bdfrhtmlScriptRetryProcess.ExitCode -ne 0){Start-Process $logFolder; throw "[$(Get-Date -f HH:mm:ss.fff)] Error while retrying BDFR-HTML module setup script. Command: 'python.exe $rootFolder\module_clone\setup.py install' returned exit code '$($bdfrhtmlScriptRetryProcess.ExitCode)'! See opened logs folder."}
    }
}

## Recheck for Python modules BDFR and BDFR-HTML (if modules weren't present)
if (-not($bdfrInstalled -and $bdfrhtmlInstalled))
{
    [string[]]$oldInstalledPythonModules = $installedPythonModules
    [string[]]$installedPythonModules =  @((pip list --disable-pip-version-check) | Where-Object {$_ -match "^[A-Z0-9-]{3,18}.*[0-9]{1,4}`.[0-9]{1,2}(`.[0-9]{1,2})?$"})
    if (($installedPythonModules.count) -gt ($oldinstalledPythonModules.count))
    {
        [string[]]$newPythonModuleNames = (Compare-Object -ReferenceObject $installedPythonModules -DifferenceObject $oldinstalledPythonModules).InputObject | ForEach-Object {($_ -split ' ')[0]} #grab first element after splitting on space, the module name, without version
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Installed $($newPythonModuleNames.count) Python modules: $($newPythonModuleNames -join ', ')"
    }
    foreach ($installedPythonModule in $installedPythonModules)
    {
        if ($installedPythonModule -like "bdfr *"){$bdfrInstalled = $true}
        if ($installedPythonModule -like "bdfrtohtml*"){$bdfrhtmlInstalled = $true}
    }
    if (-not($bdfrInstalled -and $bdfrhtmlInstalled)){throw "Error: Python modules are still not present: [ BDFR installed: $bdfrInstalled ] [ BDFR-HTML installed: $bdfrhtmlInstalled ]"}
}

## Subreddit archive loop init
[int]$subLoopCount = 1; $startDateTime = Get-Date
if ($Subreddit){[string[]]$Subreddits += $Subreddit} #single
foreach ($Sub in $Subreddits)
{
    ## Progress
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Initializing"
    [int]$percentComplete = ((($subLoopCount - 1) / ($Subreddits.count)) * 100) #count -1 so the first loop reports 0% complete when just starting
    Write-Progress -Activity "Archiving subreddit $subLoopCount of $($Subreddits.count) ..." -Status "Processing: /r/$Sub" -PercentComplete $percentComplete
    
    ## BDFR: Function for retries
    function Clone-Subreddit
    {
        ## BDFR: Check / Create / Clean output folders
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Checking, creating, cleaning output folders ..."
        foreach ($outputSubFolder in @('JSON','HTML')){
            if (-not(Test-Path "$rootFolder\$outputSubFolder\$Sub\" -PathType Container)){New-Item -Path "$rootFolder\$outputSubFolder\$Sub" -ItemType Directory -ErrorAction Stop | Out-Null}
            if (Get-ChildItem -Path "$rootFolder\$outputSubFolder\$Sub\*" -File -ErrorAction SilentlyContinue){Remove-Item -Path "$rootFolder\$outputSubFolder\$Sub\*" -Recurse -ErrorAction Stop | Out-Null}}
            
        ## BDFR: Clone subreddit to JSON
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Using BDFR to clone subreddit to disk in JSON ..."
        [string]$global:logPath = "$logFolder\bdfr_$($Sub)_$(Get-Date -f yyyyMMdd_HHmmss).log.txt" #global so these variables don't disappear each time the function exits
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Status (CTRL+C to retry): $logPath"
        $global:bdfrProcess = Start-Process "python.exe" -ArgumentList "-m bdfr clone $rootFolder\JSON --subreddit $Sub --disable-module Youtube --disable-module YoutubeDlFallback --verbose --log $logPath" -WindowStyle Hidden -PassThru
        
        ## Custom CTRL+C handling for retry and to taskkill PID for python.exe
        [console]::TreatControlCAsInput = $true #change the default behavior of CTRL+C so that the script can intercept and use it versus just terminating the script: https://blog.sheehans.org/2018/10/27/powershell-taking-control-over-ctrl-c/
        Start-Sleep -Seconds 1 #sleep for 1 second and then flush the key buffer so any previously pressed keys are discarded and the loop can monitor for the use of CTRL+C. The sleep command ensures the buffer flushes correctly
        $host.UI.RawUI.FlushInputBuffer() #flush the keyboard buffer (clear previous CTRL+C presses)
        [bool]$global:CTRLCUsedOnce = $false
        while (-not($bdfrProcess.HasExited)) #loop while the python BDFR process exists
        {
            if ($host.UI.RawUI.KeyAvailable -and ($key = $host.UI.RawUI.ReadKey("AllowCtrlC,NoEcho,IncludeKeyUp"))) #if a key was pressed during the loop execution, check to see if it was CTRL+C (aka "3"), and if so exit the script after clearing out any running python processes and setting CTRL+C back to normal
            {
                if ([int]$key.Character -eq 3){$global:CTRLCUsedOnce = $true; break} #CTRL+C pressed, exit while loop
                $host.UI.RawUI.FlushInputBuffer()
            }
        }

        ## CTRL+C pressed once - end python process
        if ($CTRLCUsedOnce)
        {
            [int]$waitLoopSeconds = 5 #seconds to wait for user to press CTRL+C the second time
            Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub] CTRL+C pressed: press again to exit script, do nothing to retry clone in $($waitLoopSeconds) secs ..." -ForegroundColor Yellow -NoNewLine
            Start-Process 'taskkill.exe' -ArgumentList "/F /PID $($bdfrProcess.ID)" -WindowStyle Hidden -ErrorAction Stop #removed -Wait here so second CTRL+C can register almost instantly, check status later after the loop
            [int]$secondsElapsed = 0; $loopStartTime = Get-Date #PS job running Stop-Process (instead of Start-Process spawning taskkill.exe), would also work here to speed up key response time
            $host.UI.RawUI.FlushInputBuffer()
            while ($waitLoopSeconds -ge 1)
            {
                if ($host.UI.RawUI.KeyAvailable -and ($key = $host.UI.RawUI.ReadKey("AllowCtrlC,NoEcho,IncludeKeyUp")))
                {
                    if ([int]$key.Character -eq 3)
                    {
                        ## CTRL+C pressed twice within 5 sec - end this script
                        [console]::TreatControlCAsInput = $false; [int]$pythonWaitMilliseconds = 1000
                        Write-Host "`n[$(Get-Date -f HH:mm:ss.fff)][$Sub] CTRL+C pressed twice: waiting up to $($pythonWaitMilliseconds)ms for python process to exit ..." -ForegroundColor Yellow #`n = line feed for the -NoNewLine periods
                        if (-not($bdfrProcess.WaitForExit($pythonWaitMilliseconds))){throw "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Failed to taskkill python process ID '$($bdfrProcess.ID)'!"}
                        exit
                    }
                    $host.UI.RawUI.FlushInputBuffer()
                }
                [int]$previousSecondsElapsed = $secondsElapsed #doing it this way instead of a Start-Sleep so the script is more responsive to CTRL+C being pressed..
                [int]$secondsElapsed = ((Get-Date) - $loopStartTime).TotalSeconds
                if ($secondsElapsed -gt $previousSecondsElapsed){Write-Host '.' -ForegroundColor Yellow -NoNewLine; $waitLoopSeconds--} #..while still writing one period per second to the console
            }
            Write-Host '.' -ForegroundColor Yellow #without -NoNewLine, to prevent the next output from wrapping onto the above periods
            if (Get-Process -ID $bdfrProcess.ID -ErrorAction SilentlyContinue | Where-Object {$_.ProcessName -eq 'python'}){Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Warning: failed to taskkill python process ID '$($bdfrProcess.ID)'!" -ForegroundColor Yellow} #only warn because an occasional still-running Python instance will not impact script execution
        }
        [console]::TreatControlCAsInput = $false #exit function
    }

    ## BDFR: Initial clone attempt and function retry loop
    Clone-Subreddit
    [int]$triesLeft = 4; [bool]$cloneSuccessful = $false
    while (($triesLeft -gt 0) -and (-not($cloneSuccessful)))
    {
        if (-not((Get-Content $logPath -Tail 1 -ErrorAction SilentlyContinue) -like "*INFO] - Program complete")) #this process often throws errors, so check for the program completion string in the tail of the log file
        {
            ## BDFR: Retry clone operation
            if ($CTRLCUsedOnce){Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] User cancelled BDFR clone operation. Retrying $triesLeft more times ..."}
            else {Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Error during BDFR clone operation. Retrying $triesLeft more times ..."}
            Clone-Subreddit
            $triesLeft--
        }
        else {$cloneSuccessful = $true}
    }
    if (-not($cloneSuccessful)){throw "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Error: Command: 'python.exe -m bdfr clone $rootFolder\JSON --subreddit $Sub --disable-module Youtube --disable-module YoutubeDlFallback --log $logPath' returned exit code '$($bdfrProcess.ExitCode)'! This was the final retry attempt."}
    if ($CTRLCUsedOnce){Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Information: the cancelled clone operation had already succeeded! Continuing ..." -ForegroundColor Green} #very rarely the clone operation succeeds around the same time CTRL+C is pressed, so in that case, acknowledge (but ignore CTRL+C not retrying or exiting)

    ## BDFR-HTML: Process Cloned Subreddit JSON into HTML
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Using BDFR-HTML to generate HTML pages from JSON archive ..."
    $bdfrhtmlProcess = Start-Process "python.exe" -ArgumentList "-m bdfrtohtml --input_folder $rootFolder\JSON\$Sub --output_folder $rootFolder\HTML\$Sub" -WindowStyle Hidden -PassThru -Wait
    if ($bdfrhtmlProcess.ExitCode -ne 0){throw "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Error: Command: 'python.exe -m bdfrtohtml --input_folder $rootFolder\JSON\$Sub --output_folder $rootFolder\HTML\$Sub' returned exit code '$($bdfrhtmlProcess.ExitCode)'!"}

    ## Replace generated index.html <title> with subreddit name
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Updating generated index.html title tag with subreddit name ..."
    (Get-Content "$rootFolder\HTML\$Sub\index.html").Replace("<title>BDFR Archive</title>","<title>/r/$Sub Archive</title>") | Set-Content "$rootFolder\HTML\$Sub\index.html" -Encoding 'UTF8' -Force
    
    ## Delete media files over 2MB threshold
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Deleting media folder files over 2MB ..."
    Get-ChildItem -Path "$rootFolder\HTML\$Sub\media" | Where-Object {(($_.Length)/1MB) -gt 2} | ForEach-Object {Remove-Item -LiteralPath "$($_.FullName)" -ErrorAction SilentlyContinue}

    ## Loop End
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Finished!"
    if (($Subreddits.count -eq 1) -or ($subLoopCount -eq $Subreddits.count)){Write-Progress -Activity "Archiving subreddit $subLoopCount of $($Subreddits.count) ..." -Completed} #if last loop, 100% progress
    else {$subLoopCount++}
}

## Script End / Package Output
if ($Subreddits.count -eq 1)
{
    ## Single subreddit: open HTML output folder
    Start-Process "$rootFolder\HTML\$Subreddit\"
}
else
{
    ## Multiple subreddits
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Completed all subreddit HTML archives!"
    [string]$zipOutputPath = "$rootFolder\ZIP\$(Get-Date -f yyyy-MM-dd_HHmmss)"
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Copying all HTML archive folders into ZIP prep folder ..."
    foreach ($Sub in $Subreddits)
    {
        ## Create ZIP prep directories and copy all HTML archive folders
        New-Item -Path "$zipOutputPath\$Sub\" -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Copy-Item -Path "$rootFolder\HTML\$Sub\*" -Destination "$zipOutputPath\$Sub\" -Recurse -ErrorAction Stop | Out-Null
    }

    ## Calculate size of raw JSON and finished HTML archive folders
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Calculating JSON clone and HTML archive folder sizes ..."
    $FSO = New-Object -ComObject 'Scripting.FileSystemObject' #speedy directory size calculation
    try {[string]$totalHTMLSize = "$([int]([decimal]($FSO.GetFolder($zipOutputPath).Size) / 1MB)) MB"} catch {[string]$totalHTMLSize = 'Unknown'}
    try {$Subreddits | ForEach-Object {$JSONMB += [int]([decimal]($FSO.GetFolder("$rootFolder\JSON\$_\").Size) / 1MB)}} catch {[string]$totalJSONSize = 'Unknown'} finally {if (-not($totalJSONSize)){[string]$totalJSONSize = "$JSONMB MB"}}
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($FSO)

    ## Generate master index.html for all archived subreddits
    [string[]]$indexContents = @(); $endDateTime = (Get-Date); $timeElapsed = $endDateTime - $startDateTime
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Generating master index.html for all archives ..."
    $indexContents += "<html><head><style>body {background-color: rgb(128, 128, 128);}</style><title>BDFR Archive Index</title></head><body><ul>"
    foreach ($subredditDirectoryName in @((Get-ChildItem $zipOutputPath -Directory).Name))
    {
        $indexContents += "<li><a href=`"./$($subredditDirectoryName)/index.html`"><h2>/r/$subredditDirectoryName</h2></a></li>" #unordered list
    }
    $indexContents += "</ul><hr>" #horizontal rule
    $indexContents += "Archive started: $startDateTime<br>"
    $indexContents += "Archive complete: $endDateTime<br>"
    $indexContents += "Time elapsed: $(($timeElapsed.TotalHours).ToString("###.##")) hours<br>"
    $indexContents += "JSON uncompressed size: $totalJSONSize<br>"
    $indexContents += "HTML uncompressed size: $totalHTMLSize"
    $indexContents += '</body></html>' #footer stats
    $indexContents | Set-Content "$zipOutputPath\index.html" -Encoding UTF8

    ## Compress ZIP & Open output
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Compressing everything into a ZIP archive ..."
    try {$zipDate = "$((Get-Culture).DateTimeFormat.GetMonthName((Get-Date).Month))-$(Get-Date -f dd-yyyy)"} catch {$zipDate = "$(Get-Date -f MM-dd-yyyy)"}
    Compress-Archive -Path "$zipOutputPath\*" -Destination "$zipOutputPath\indexed_subreddits_$($zipDate).zip"
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Finished! Time elapsed: $(($timeElapsed.TotalHours).ToString("###.##")) hours"
    Start-Process $zipOutputPath
}

# SIG # Begin signature block
# MIIVpAYJKoZIhvcNAQcCoIIVlTCCFZECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/47AvDaZi26lR8xXNkGcXTZs
# pSigghIFMIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
# AQwFADB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEh
# MB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTIxMDUyNTAwMDAw
# MFoXDTI4MTIzMTIzNTk1OVowVjELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEtMCsGA1UEAxMkU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5n
# IFJvb3QgUjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAjeeUEiIE
# JHQu/xYjApKKtq42haxH1CORKz7cfeIxoFFvrISR41KKteKW3tCHYySJiv/vEpM7
# fbu2ir29BX8nm2tl06UMabG8STma8W1uquSggyfamg0rUOlLW7O4ZDakfko9qXGr
# YbNzszwLDO/bM1flvjQ345cbXf0fEj2CA3bm+z9m0pQxafptszSswXp43JJQ8mTH
# qi0Eq8Nq6uAvp6fcbtfo/9ohq0C/ue4NnsbZnpnvxt4fqQx2sycgoda6/YDnAdLv
# 64IplXCN/7sVz/7RDzaiLk8ykHRGa0c1E3cFM09jLrgt4b9lpwRrGNhx+swI8m2J
# mRCxrds+LOSqGLDGBwF1Z95t6WNjHjZ/aYm+qkU+blpfj6Fby50whjDoA7NAxg0P
# OM1nqFOI+rgwZfpvx+cdsYN0aT6sxGg7seZnM5q2COCABUhA7vaCZEao9XOwBpXy
# bGWfv1VbHJxXGsd4RnxwqpQbghesh+m2yQ6BHEDWFhcp/FycGCvqRfXvvdVnTyhe
# Be6QTHrnxvTQ/PrNPjJGEyA2igTqt6oHRpwNkzoJZplYXCmjuQymMDg80EY2NXyc
# uu7D1fkKdvp+BRtAypI16dV60bV/AK6pkKrFfwGcELEW/MxuGNxvYv6mUKe4e7id
# FT/+IAx1yCJaE5UZkADpGtXChvHjjuxf9OUCAwEAAaOCARIwggEOMB8GA1UdIwQY
# MBaAFKARCiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBQy65Ka/zWWSC8oQEJw
# IDaRXBeF5jAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUE
# DDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0FBQUNlcnRpZmlj
# YXRlU2VydmljZXMuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQASv6Hvi3Sa
# mES4aUa1qyQKDKSKZ7g6gb9Fin1SB6iNH04hhTmja14tIIa/ELiueTtTzbT72ES+
# BtlcY2fUQBaHRIZyKtYyFfUSg8L54V0RQGf2QidyxSPiAjgaTCDi2wH3zUZPJqJ8
# ZsBRNraJAlTH/Fj7bADu/pimLpWhDFMpH2/YGaZPnvesCepdgsaLr4CnvYFIUoQx
# 2jLsFeSmTD1sOXPUC4U5IOCFGmjhp0g4qdE2JXfBjRkWxYhMZn0vY86Y6GnfrDyo
# XZ3JHFuu2PMvdM+4fvbXg50RlmKarkUT2n/cR/vfw1Kf5gZV6Z2M8jpiUbzsJA8p
# 1FiAhORFe1rYMIIGGjCCBAKgAwIBAgIQYh1tDFIBnjuQeRUgiSEcCjANBgkqhkiG
# 9w0BAQwFADBWMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVk
# MS0wKwYDVQQDEyRTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYw
# HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBUMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1Ymxp
# YyBDb2RlIFNpZ25pbmcgQ0EgUjM2MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIB
# igKCAYEAmyudU/o1P45gBkNqwM/1f/bIU1MYyM7TbH78WAeVF3llMwsRHgBGRmxD
# eEDIArCS2VCoVk4Y/8j6stIkmYV5Gej4NgNjVQ4BYoDjGMwdjioXan1hlaGFt4Wk
# 9vT0k2oWJMJjL9G//N523hAm4jF4UjrW2pvv9+hdPX8tbbAfI3v0VdJiJPFy/7Xw
# iunD7mBxNtecM6ytIdUlh08T2z7mJEXZD9OWcJkZk5wDuf2q52PN43jc4T9OkoXZ
# 0arWZVeffvMr/iiIROSCzKoDmWABDRzV/UiQ5vqsaeFaqQdzFf4ed8peNWh1OaZX
# nYvZQgWx/SXiJDRSAolRzZEZquE6cbcH747FHncs/Kzcn0Ccv2jrOW+LPmnOyB+t
# AfiWu01TPhCr9VrkxsHC5qFNxaThTG5j4/Kc+ODD2dX/fmBECELcvzUHf9shoFvr
# n35XGf2RPaNTO2uSZ6n9otv7jElspkfK9qEATHZcodp+R4q2OIypxR//YEb3fkDn
# 3UayWW9bAgMBAAGjggFkMIIBYDAfBgNVHSMEGDAWgBQy65Ka/zWWSC8oQEJwIDaR
# XBeF5jAdBgNVHQ4EFgQUDyrLIIcouOxvSK4rVKYpqhekzQwwDgYDVR0PAQH/BAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYD
# VR0gBBQwEjAGBgRVHSAAMAgGBmeBDAEEATBLBgNVHR8ERDBCMECgPqA8hjpodHRw
# Oi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RS
# NDYuY3JsMHsGCCsGAQUFBwEBBG8wbTBGBggrBgEFBQcwAoY6aHR0cDovL2NydC5z
# ZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdSb290UjQ2LnA3YzAj
# BggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEM
# BQADggIBAAb/guF3YzZue6EVIJsT/wT+mHVEYcNWlXHRkT+FoetAQLHI1uBy/YXK
# ZDk8+Y1LoNqHrp22AKMGxQtgCivnDHFyAQ9GXTmlk7MjcgQbDCx6mn7yIawsppWk
# vfPkKaAQsiqaT9DnMWBHVNIabGqgQSGTrQWo43MOfsPynhbz2Hyxf5XWKZpRvr3d
# MapandPfYgoZ8iDL2OR3sYztgJrbG6VZ9DoTXFm1g0Rf97Aaen1l4c+w3DC+IkwF
# kvjFV3jS49ZSc4lShKK6BrPTJYs4NG1DGzmpToTnwoqZ8fAmi2XlZnuchC4NPSZa
# PATHvNIzt+z1PHo35D/f7j2pO1S8BCysQDHCbM5Mnomnq5aYcKCsdbh0czchOm8b
# kinLrYrKpii+Tk7pwL7TjRKLXkomm5D1Umds++pip8wH2cQpf93at3VDcOK4N7Ew
# oIJB0kak6pSzEu4I64U6gZs7tS/dGNSljf2OSSnRr7KWzq03zl8l75jy+hOds9TW
# SenLbjBQUGR96cFr6lEUfAIEHVC1L68Y1GGxx4/eRI82ut83axHMViw1+sVpbPxg
# 51Tbnio1lB93079WPFnYaOvfGAA0e0zcfF/M9gXr+korwQTh2Prqooq2bYNMvUoU
# KD85gnJ+t0smrWrb8dee2CvYZXD5laGtaAxOfy/VKNmwuWuAh9kcMIIGcDCCBNig
# AwIBAgIQVdb9/JNHgs7cKqzSE6hUMDANBgkqhkiG9w0BAQwFADBUMQswCQYDVQQG
# EwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdv
# IFB1YmxpYyBDb2RlIFNpZ25pbmcgQ0EgUjM2MB4XDTIxMTIxNjAwMDAwMFoXDTIy
# MTIxNjIzNTk1OVowUDELMAkGA1UEBhMCVVMxEzARBgNVBAgMCldhc2hpbmd0b24x
# FTATBgNVBAoMDE1pY2hhZWwgQmFycjEVMBMGA1UEAwwMTWljaGFlbCBCYXJyMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAne6XW99iRvph0mHzkgX+e+6i
# mXxytFu35Vw4YC0TSeDqkUCc0PoSyojLc+MKLa/t+32ya1BWmSf1u5Hc55yo9BL3
# dvV7C9HisQ8gB3+Cb+04P+0b/buBor9M7Cu+rJe7RZOVS9bq+CuslCchBejc6tNe
# f+A8b1q9jzjgVvAUpv+dD4asi/KhMYdhDWxI23i0A9XOn8OBrfsu9zQBYGxFX7Is
# Wk+wunMNwN6PPeZ9gFVwHuh5OVXEDIXGVm+N7QTSdTTdLC6w5ttWzVrsKdQM6vZI
# yNuV5x1bQ32cbBdT2oB+R7ODSmuMTxMagfm4lrqjPZKNP91MCRVpbWbv/4/ealte
# KResVeIm+mQbXkWmFWIHgLkXToVDlyWOBFjG0I5rt2p9055FZ7Xpo36Vinvs+JWj
# fgDaYKPEeHJ3AFwdJD6gjVBH9xt0IJlZm7rWiqE+BpsgzxBKJGYzHqBwmWtLFZvG
# 5DdwVKCThFGyoIawT/POm7eBU9tyePv1g95xkzesqHGz854f+w+XXWW/qwAZBMAY
# QnAPLFI1ywJ1GHVkp7xZRaxAOEiId0WG57R/y4h5gtE12nPa07PUrtl3HPClZICE
# 6PP5UimZH2fF2ClwyAoaxXU70yblD6V+gzZ1wgDpDl1FYyDdZmtjtz6zh8MAp9b+
# /Rk2BS3SWH9iUjn0yTECAwEAAaOCAcAwggG8MB8GA1UdIwQYMBaAFA8qyyCHKLjs
# b0iuK1SmKaoXpM0MMB0GA1UdDgQWBBSmkRZEx8ANTjiACrZOUmUhtz5SvTAOBgNV
# HQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDAzAR
# BglghkgBhvhCAQEEBAMCBBAwSgYDVR0gBEMwQTA1BgwrBgEEAbIxAQIBAwIwJTAj
# BggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwCAYGZ4EMAQQBMEkG
# A1UdHwRCMEAwPqA8oDqGOGh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1
# YmxpY0NvZGVTaWduaW5nQ0FSMzYuY3JsMHkGCCsGAQUFBwEBBG0wazBEBggrBgEF
# BQcwAoY4aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNp
# Z25pbmdDQVIzNi5jcnQwIwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28u
# Y29tMCIGA1UdEQQbMBmBF21iYXJyNTY0QHByb3Rvbm1haWwuY29tMA0GCSqGSIb3
# DQEBDAUAA4IBgQBNf+rZW+Q5SiL5IjqhOgozqGFuotQ+kpZZz2GGBJoKldd1NT7N
# kEJWWzCd8ezH+iKAd78CpRXlngi/j006WbYuSZetz4Z2bBje1ex7ZcL00Hnh4pSv
# heAkiRgkS5cvbtVVDnK5+AdEre/L71qU1gpZGNs7eqEdp5tBiEcf9d18H0hLHMtd
# 5veYH2zXqwfXo8SNGYRz7CCgDiYSdHDsSE284a/CcUivte/jJe1YmZR/Zueuisti
# fkeqldgFrqc30JztyIU+EVXeNOivA5yihYj5WBz7zMVjnBsmEH0bUdrKImptWzCw
# 2x8dGzziG7jfeYs20gG05Xv4Jd0IBdoxhRMeznT8WhvwifG9aN4IZPDMyfYT9v1j
# 2zx8EbcmhD1aaio9gP18AvBWksa3KvOChA1BQvD7PR5YucZEzoljq10kIjKsLA3U
# te7JSxpXDFC7Ab/xeUYRGIG/x/wyCLRjENe+ryixRy6txVUDkxqDsqngzPVeyvYM
# fjlXjk9R0ZjWwNsxggMJMIIDBQIBATBoMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQK
# Ew9TZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUg
# U2lnbmluZyBDQSBSMzYCEFXW/fyTR4LO3Cqs0hOoVDAwCQYFKw4DAhoFAKB4MBgG
# CisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
# AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYE
# FNbFlBLsbeB2B1o/BMQ0tv4nq4XOMA0GCSqGSIb3DQEBAQUABIICABnW9STE4fig
# YmSMoj9TXb2GcqMZaQBTl6WTQ1J8TtasAxgaXNSGlVAZdORUYON299/HgurJojyc
# rDVjcqfx58+cheP1P5QqYRhDJbKwjJZwAVSsX1DYsVTCcx07ayMSHv2cyoxlT2U2
# ZJ1yfaXu+d8LdMkJ/9GYdCpVPcIAfeJlrB+w+/q8sSTwTpsKmmEBvopKbceL9IX7
# 5fWTYJB1O+HeJkJ6wu+zkp1jr1BVYY4HMthNwPabxWmxn4VZ3Yn5gofHTHlxUI1i
# MK333Eyedz3kC6FdiIRJGeQm+iZSpKpKQeWnxOgrJ/LkPCydJoXJq4kkXNm65zNM
# sDR8JReiCvFtrgZe8CE4WLr1UZ2v6iDRH+PVQaUV1PzQynLifByXdmwShpv+SLX8
# NwLuNSO+mdMs169k+NFxAGdKShEKzlutWbowFA9J9jUDB6U/FyDPXSPUAT7VCqRn
# 3bP3GUb7V5tft4p71I6jXEx+hv79hE+h+9amwGvhnFqzr62XfAPfEfZEhFPDiHH0
# noRsf9o773pZzcT6ZZ/frk8/nGfMOHImXF+1uglMc4HZNfUMeYkWxOKWIPTjSGrn
# xPZc4xUG+dXjVgUsBIk+4UbEf/+4DxUHV8qa/TxH4kGJnjaogq1D1e8LWe6oSQI+
# 3ys7VQwPctZdbv0MZ06BuwTQWwiyru0Z
# SIG # End signature block
