<#PSScriptInfo
.VERSION 2.2
.GUID 3ae5d1f9-f5be-4791-ab41-8a4c9e857e9c
.AUTHOR mbarr@tutanota.com
.PROJECTURI https://github.com/mbarr564/New-SubredditHTMLArchive
.DESCRIPTION Windows turnkey wrapper for BDFR and BDFR-HTML Python modules. Installs modules, and prerequisites, then creates offline/portable HTML archives of subreddit posts and comments.
#>
<#
.SYNOPSIS
    You do not need previous PowerShell, Python, or command line tool experience, to use this script. Either copy and paste from the README, or "Run as administrator" the included setup.cmd batch file.
    Checks for prerequisites (which this script can also install, when given permission), then uses the BDFR and BDFR-HTML Python modules to generate a portable subreddit HTML archive.
    Creates a root 'New-SubredditHTMLArchive' output folder, under your %USERPROFILE% ($env:USERPROFILE) Documents folder. The script handles all folder and data management, and writes verbose logs into the root logs folder.
    Clones all linked media (including huge videos) into the JSON folder, then deletes media files over 5MB in the HTML copy folder (keeps pictures only), to make a smaller ZIP archive. This can be overridden with the -NoMediaPurge parameter.
    Runs itself as a scheduled task as the current user, as an interactive console by default. The task can be run as a background task with the -Background parameter, allowing use of the lock screen.
    The reddit API returns a maximum of 1000 posts per request, so only the newest 1000 posts will be included: https://github.com/reddit-archive/reddit/blob/master/r2/r2/lib/db/queries.py
.DESCRIPTION
    If you run this script regularly, PAY FOR REDDIT PREMIUM, to offset their public API development and traffic costs: https://www.reddit.com/premium
    You can skip the rest of this section, if you already have Python 3 and Git 2 installed, and you don't care which prerequisite Python modules are installed automatically.
    This script does NOT require administrator privileges to run, or to install the Python modules, without the -InstallPackages parameter, which is usually only used once.
    However, when OVERWRITING existing scheduled tasks (as this script will do, when rerun), you must approve an administrator UAC prompt, because you're deleting and recreating the 'RunOnce' task.
    On first run, without the prerequisites installed, you must include the -InstallPackages parameter, or manually install the below software packages, before running this script.
    When installing these packages automatically, the user must confirm a UAC admin prompt for each package, allowing the installer to make changes to their computer.
        1. Git: https://github.com/git-for-windows/git/releases/ (only when manually installing)
        2. Python (includes pip): https://www.python.org/downloads/windows/ (only when manually installing)
            i. At beginning of install, YOU MUST CHECK 'Add Python 3.x to PATH'. (So PowerShell can call python.exe and pip.exe from anywhere)
    This script uses (and can install) the following Python modules, which are detected via pip (Package Installer for Python):
        1. BDFR: https://pypi.org/project/bdfr/
        2. BDFR-HTML: https://github.com/BlipRanger/bdfr-html
            i. Dependency Python modules installed by BDFR-HTML: click, markdown, appdirs, bs4, dict2xml, ffmpeg-python, praw, pyyaml, requests, youtube-dl, jinja2, pillow (some of which install their own dependencies)
            ii. When running setup.py to install BDFR-HTML, you may get an install error from Pillow about zlib being missing. You may need to run 'pip install pillow' from an elevated command prompt, so that Pillow installs correctly.
            iii. For manual BDFR-HTML install in case of Pillow install error: From an elevated CMD window, type these two quoted commands: 1) "cd %USERPROFILE%\Documents\New-SubredditHTMLArchive\tools\bdfr-html", 2) "python.exe setup.py install"
            iv. https://stackoverflow.com/questions/64302065/pillow-installation-pypy3-missing-zlib
            v. Something went wrong during install? To remove ALL Python modules: 1) pip freeze > requirements.txt, 2) pip uninstall -r requirements.txt -y
.PARAMETER Subreddit
    The name of a single subreddit that will be archived.
.PARAMETER Subreddits
    An array/list of subreddit names that will be archived.
    Also generates a master index.html containing links to all of the other generated subreddit index.html files.
    All generated subreddit folders, files, and index pages, are automatically packaged into a ZIP file.
.PARAMETER InstallPackages
    The script will attempt to install ONLY MISSING pre-requisite packages: Python 3 and/or Git
    When 'python.exe' or 'git.exe' are already in your $env:path, and executable from PowerShell, they will NOT be installed or modified.
.PARAMETER Background
    The script will spawn the scheduled task with S4U logon type instead of Interactive logon type. Requires approval of an admin UAC prompt to spawn the task.
    This switch allows the script to keep running in the background, regardless of the user's logon state (such as lock screens, when running overnight).
.PARAMETER NoMediaPurge
    The script will not purge media files over 5MB, which is helpful if you want your zipped HTML archive to include full video files (instead of only in the JSON folder).
    Note that this parameter will increase the final ZIP archive size by orders of magnitude, or often from hundreds of megabytes, to tens of gigabytes, for certain subreddits.
.EXAMPLE
    PS> .\New-SubredditHTMLArchive.ps1 -Subreddit PowerShell
.EXAMPLE
    PS> .\New-SubredditHTMLArchive.ps1 -Subreddit TestSubredditC -InstallPackages
.EXAMPLE
    PS> .\New-SubredditHTMLArchive.ps1 -Subreddits (Get-Content "$($env:USERPROFILE)\Desktop\subreddit_list.txt") -Background -NoMediaPurge
.EXAMPLE
    PS> .\New-SubredditHTMLArchive.ps1 -Subreddits 'PowerShell','Python','AmateurRadio','HackRF','GNURadio','OpenV2K','SignalIdentification','DataHoarder' -Background
.NOTES
    Last update: Wednesday, May 24, 2023 1:29:34 AM
#>

param([string]$Subreddit, [ValidateCount(2,200)][string[]]$Subreddits, [switch]$InstallPackages, [switch]$Background, [switch]$NoMediaPurge)

## Init
[string]$taskName = 'RunOnce' #slash characters in this string bad
[string]$scriptName = ((Split-Path $PSCommandPath -Leaf) -replace ".ps1",'')
[string]$rootFolder = "$($env:USERPROFILE)\Documents\$scriptName" #the leaf folder may not exist yet, but will be created.. parent folder must exist
[string]$zipOutputPath = "$rootFolder\ZIP\$(Get-Date -f yyyy-MM-dd_HHmmss)"
[string]$transcriptPath = "$rootFolder\logs\$($scriptName)_$($taskName)_transcript_-_$(Get-Date -f yyyy-MM-dd).txt"
Get-ChildItem "$rootFolder\logs" | Where-Object {$_.Name -like "*.1"} | ForEach-Object {Remove-Item -LiteralPath "$($_.FullName)" -ErrorAction SilentlyContinue} #clean empty log files with extension .1
if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){$isAdmin = '[ADMIN] '} else {$isAdmin = ''} #is this instance running as admin, string presence used as boolean

## Custom parameter validation
if ($Background){$logonType = 'S4U'} else {$logonType = 'Interactive'} #default Interactive (for viewable console window), -Background for S4U (non-interactive task). Not supported: Password (not tested), InteractiveOrPassword (not tested), Group (no profile), ServiceAccount (no profile), None (no profile)
if (-not(Test-Path (Split-Path $rootFolder -Parent) -PathType Container)){throw "Error: root output path doesn't exist or not a directory: $(Split-Path $rootFolder -Parent)"}
foreach ($folderName in @('JSON','HTML','ZIP','logs','tools')){if (-not(Test-Path "$rootFolder\$folderName" -PathType Container)){New-Item -Path "$rootFolder\$folderName" -ItemType Directory -Force -ErrorAction Stop | Out-Null}}
if ($InstallPackages -and $Background){throw 'Error: the -InstallPackages parameter cannot be used in conjunction with the -Background parameter! Use the -InstallPackages switch by itself first, then use the -Background switch on the next script run.'}
if ($Subreddits)
{
    [string[]]$badNames = @(); $Subreddits = $Subreddits | Select-Object -Unique; $Subreddits | ForEach-Object {if ($_ -notmatch "^[A-Z0-9_]{2,21}$"){$badNames += $_}}
    if ($badNames.count -gt 0){throw "Error: Subreddit name(s) failed regex validation: $($badNames -join ', ')"}
    [int]$hoursToArchiveAll = (60 * ($Subreddits.count)) / 60 #from experience BDFR usually takes 1-3hrs to finish pulling down the API maximum of 1000 records, per subreddit.. so 120 minutes (updated to 60 minutes because the API got way faster in 2023)
    if ($hoursToArchiveAll -gt 24){$timeToArchiveAll = "over $([int]($hoursToArchiveAll / 24)) full day(s)"} else {$timeToArchiveAll = "$hoursToArchiveAll hours"}
    if (-not($isAdmin)){Write-Host "Estimated maximum time to finish $($Subreddits.count) subreddits: $timeToArchiveAll ($((Get-Date).AddHours($hoursToArchiveAll)))" -ForegroundColor Cyan}
}
else
{
    if (-not($Subreddit)){$Subreddit = Read-Host -Prompt 'Enter subreddit name to archive'}
    if (-not($Subreddit)){throw 'Error: Subreddit name is blank!'}
    if ($Subreddit -notmatch "^[A-Z0-9_]{2,21}$"){throw "Error: Subreddit name failed regex validation: $Subreddit"}
}

## Relaunch script as Scheduled Task
if (-not((Get-ScheduledTask | Where-Object {$_.TaskPath -eq "\$scriptName\"}).State -eq 'Running')) #do not check task name, as it can be renamed and rescheduled
{
    ## Check / create Task Scheduler script name folder
    try {$scheduleObject = (New-Object -ComObject Schedule.Service); $scheduleObject.Connect(); $rootScheduleFolder = $scheduleObject.GetFolder('\')} catch {throw "Error: failed to connect to scheduling service!"}
    try {[void]$rootScheduleFolder.GetFolder($scriptName)} catch {try {[void]$rootScheduleFolder.CreateFolder($scriptName)} catch {throw "Error: failed to create scheduled tasks folder '$scriptName'!"}}
    
    ## Build script arguments string
    [string]$scriptArgs = ''
    if ($Subreddits){$scriptArgs = " -Subreddits '$($Subreddits -join "','")'"}
    elseif ($Subreddit){$scriptArgs = " -Subreddit $Subreddit"}
    if ($InstallPackages){$scriptArgs += ' -InstallPackages'}
    if ($Background){$scriptArgs += ' -Background'}
    if ($NoMediaPurge){$scriptArgs += ' -NoMediaPurge'}

    ## Build splatted scheduled task parameters
    $taskArguments = @{
        TaskName = $taskName
        TaskPath = $scriptName
        Trigger = (New-ScheduledTaskTrigger -At ((Get-Date).AddSeconds(5)) -Once)
        Action = (New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "$($PSCommandPath)$scriptArgs") #space before $scriptArgs included already in string build
        Settings = (New-ScheduledTaskSettingsSet -DisallowDemandStart -ExecutionTimeLimit (New-TimeSpan -Seconds 0)) #1) must be run with a new trigger datetime, instead of right-clicking the task and choosing run. 2) PT0S equiv for indefinite/disabled run time
        Principal = (New-ScheduledTaskPrincipal -UserID "$($env:COMPUTERNAME)\$($env:USERNAME)" -LogonType $logonType -RunLevel Limited)
        ErrorAction = 'Stop'
        Force = $true} #force overwrite of previous task with same name
    if ($Subreddits){$taskArguments.add('Description',"Subreddits: $($Subreddits -join ', ')    -=-    Background task transcript path: $rootFolder\logs\    -=-    Finished ZIP archive output path: $rootFolder\ZIP\")} #linebreaks not supported by the GUI for this field

    ## Register scheduled task, handle access errors, display log and zip paths
    if ($isAdmin){$foreColor = 'Yellow'} else {$foreColor = 'Cyan'}
    Write-Host "$($isAdmin)Creating task 'Task Scheduler Library > $scriptName > $taskName' with $logonType logon type ..." -ForegroundColor $foreColor #https://stackoverflow.com/questions/13965997/powershell-set-a-scheduled-task-to-run-when-user-isnt-logged-in
    try {Register-ScheduledTask @taskArguments | Out-Null} #can trigger UAC admin prompt, which will rerun script as admin to create the task, if task creation fails.. the created task will NOT run as admin
    catch [Microsoft.Management.Infrastructure.CimException]{if (-not($isAdmin)){Start-Process 'powershell.exe' -ArgumentList "$($PSCommandPath)$scriptArgs" -Verb RunAs -Wait} else {throw $error[0]}} #access denied.. rerun this script with same args, as admin (will also trigger if overwriting task with S4U LogonType)
    catch {throw $error[0]} #S4U type will trigger UAC admin prompt to create the task.. or.. user can create as Interactive, and manually change the 'Security options' to 'Run whether user is logged on or not', which does NOT trigger a UAC prompt.
    if (Get-ScheduledTask | Where-Object {($_.TaskPath -eq "\$scriptName\") -and ($_.TaskName -eq $taskName)}){if ($Background -and (-not($isAdmin))){Write-Host "Transcript logging for successfully spawned Task Scheduler background task (taskschd.msc):`n$transcriptPath" -ForegroundColor Cyan}}
    else {throw 'Error: failed to create scheduled task!'}
    if ($isAdmin){Start-Sleep -Seconds 2} #if running as admin, pause a moment so the user can see the administrator console output
    exit
}
else
{
    ## This script is already running as scheduled task.. is this instance the task?
    [string]$runningTaskName = ((Get-ScheduledTask | Where-Object {($_.TaskPath -eq "\$scriptName\") -and ($_.State -eq 'Running')}).TaskName) #task name can be changed when rescheduled
    [datetime]$taskLastRunTime = (((Get-ScheduledTask | Where-Object {($_.TaskPath -eq "\$scriptName\") -and ($_.TaskName -eq "$runningTaskName")}) | Get-ScheduledTaskInfo).LastRunTime)
    [datetime]$taskTriggerTime = ((Get-ScheduledTask | Where-Object {($_.TaskPath -eq "\$scriptName\") -and ($_.TaskName -eq "$runningTaskName")}).Triggers.StartBoundary) #this won't work if user reruns task later, but will if they set a new trigger.. added because LastRunTime wasn't being reliable
    if (($taskLastRunTime -lt ((Get-Date).AddSeconds(-45))) -or ($taskTriggerTime -lt ((Get-Date).AddSeconds(-5)))) #for reasons unknown, the spawned task's LastRunTime is about 30 seconds before the task was even created..
    {
        Write-Warning "Detected running script task! Last run: $($taskLastRunTime). Triggered: $($taskTriggerTime). Exiting ..."
        Write-Warning "Task Scheduler > Task Scheduler Library > $scriptName > $runningTaskName"
        Start-Process 'taskschd.msc'
        exit #exit so any background process is not interrupted
    }

    ## This instance IS the task (no exit above)
    if (-not($Background))
    {
        ## Interactive console rename and resize attempt
        Write-Host "Task running under interactive logon type, updating title and resizing console window ..." -ForegroundColor Cyan
        Write-Output "`n`0`n`0`n`0`n"; if ($Subreddit){Write-Output "`0`n"} #skip lines (with new lines and nulls) so the progress banner doesn't cover interactive console output ..and skip one more for a single subreddit (because no completion time estimate)
        $host.UI.RawUI.WindowTitle = "Task Scheduler Library > $scriptName > $runningTaskName" #based on size 16 Consolas font (right-click powershell.exe window title > properties > Font tab)
    }
    else
    {
        ## Background Task: Start Transcript
        Start-Transcript -LiteralPath $transcriptPath -Append -ErrorAction Stop | Out-Null #only start log once script is running as a task, and as a background task
        if (-not(Test-Path "$zipOutputPath" -PathType Container)){New-Item -Path "$zipOutputPath" -ItemType Directory -Force -ErrorAction Stop | Out-Null} #create ZIP output folder right away so user can open it in file explorer
    }
}

## Main: Check for command line utilities
[string[]]$missingExes = @()
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Checking for command line utilities ..."
foreach ($exeName in @('git','python')){if (-not(Get-Command "$($exeName).exe" -ErrorAction SilentlyContinue | Where-Object {$_.Source -notlike "*\AppData\Local\Microsoft\WindowsApps\*"})){$missingExes += "$($exeName).exe"}} #exclude executable shortcuts under \WindowsApps\, that are really just placeholder executables that launch the Microsoft Store
if ($missingExes.count -gt 0)
{
    if ($InstallPackages)
    {
        ## Install command line utility packages
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Installing $($missingExes.count) missing package(s) via winget ..." -ForegroundColor Cyan
        [string]$OSName = ((Get-WmiObject -class Win32_OperatingSystem).Caption); [int]$OSBuild = [System.Environment]::OSVersion.Version.Build
        if ($OSBuild -lt 17763){throw "Error: this Windows OS build ($OSBuild) is older than build 17763, which is required for winget!"} #check that Windows version is new enough to support winget: https://github.com/microsoft/winget-cli#installing-the-client
        if (((Get-CimInstance -ClassName Win32_OperatingSystem).ProductType) -ne 1){throw "Error: this Windows OS is not a client/workstation edition, which is required for winget!"} #check that Windows is a client OS and not Server editions
        if ((-not(Get-Command 'winget.exe' -ErrorAction SilentlyContinue)) -and ($OSName -like "*Windows 10*")){Start-Process 'https://www.microsoft.com/en-us/p/app-installer/9nblggh4nns1'; throw "Error: not found: winget.exe. Opened winget install link for Windows 10 in your default browser. Rerun script after install."}
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Installing latest releases for: $($missingExes -join ', ') ..." -ForegroundColor Cyan
        switch ($missingExes)
        {
            'python.exe' {try {winget install --id Python.Python.3.11 --location "$rootFolder\tools" --accept-source-agreements --accept-package-agreements} catch {winget install --id Python.Python.3.11 --location "$rootFolder\tools" --accept-source-agreements --accept-package-agreements}}
            'git.exe' {try {winget install --id Git.Git --location "$rootFolder\tools" --accept-source-agreements --accept-package-agreements} catch {winget install --id Git.Git --location "$rootFolder\tools" --accept-source-agreements --accept-package-agreements}} #try catch for simple single retries
        }

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
        [string[]]$missingExes = @()
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Rechecking for command line utilities after installs ..."
        foreach ($exeName in @('git','python','pip'))
        {
            if (Get-Command "$($exeName).exe" -ErrorAction SilentlyContinue | Where-Object {$_.Source -notlike "*\AppData\Local\Microsoft\WindowsApps\*"})
            {
                [string]$exeVersion = (Invoke-Expression "$exeName --version") #invoke to process the string exe name and argument as a command, and string type to flatten any multi element string arrays / multi line returns
                Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Found $($exeName).exe version: $exeVersion"
            }
            else {$missingExes += "$($exeName).exe"}
        }
        if ($missingExes.count -gt 0){throw "[$(Get-Date -f HH:mm:ss.fff)] Error: Required command line utilities still missing: $($missingExes -join ', ')"}
    }
    else {throw "Error: Required command line utilities missing: $($missingExes -join ', ')! Rerun script with the -InstallPackages parameter?"}
}

## Check for Python modules BDFR and BDFR-HTML
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Checking for BDFR and BDFR-HTML Python modules ..." #todo: version check/update
[bool]$bdfrInstalled = $false; [bool]$bdfrhtmlInstalled = $false
[string[]]$installedPythonModules = @(pip freeze --disable-pip-version-check)
foreach ($installedPythonModule in $installedPythonModules)
{
    if ($installedPythonModule -like "bdfr==*"){$bdfrInstalled = $true}
    if ($installedPythonModule -like "bdfrtohtml==*"){$bdfrhtmlInstalled = $true}
}

## Install Python modules BDFR and BDFR-HTML
if (-not($bdfrInstalled))
{
    Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Installing BDFR Python module ..." -ForegroundColor Cyan
    $bdfrInstallProcess = Start-Process "python.exe" -ArgumentList "-m pip install bdfr --upgrade" -WindowStyle Hidden -RedirectStandardError "$rootFolder\logs\BDFR_installErrors.txt" -RedirectStandardOutput "$rootFolder\logs\BDFR_installStdOut.txt" -PassThru -Wait
    if ($bdfrInstallProcess.ExitCode -ne 0){Start-Process "$rootFolder\logs"; throw "[$(Get-Date -f HH:mm:ss.fff)] Error: Command: 'python.exe -m pip install bdfr --upgrade' returned exit code '$($bdfrInstallProcess.ExitCode)'! See opened logs folder."}
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Updating BDFR Python module ..." #these BDFR install commands have never failed for me.. but their error and output streams are redirected into log files anyway, just in case
    $bdfrUpdateProcess = Start-Process "python.exe" -ArgumentList "-m pip install bdfr --upgrade" -WindowStyle Hidden -RedirectStandardError "$rootFolder\logs\BDFR_updateErrors.txt" -RedirectStandardOutput "$rootFolder\logs\BDFR_updateStdOut.txt" -PassThru -Wait
    if ($bdfrUpdateProcess.ExitCode -ne 0){Start-Process "$rootFolder\logs"; throw "[$(Get-Date -f HH:mm:ss.fff)] Error: Command: 'python.exe -m pip install bdfr --upgrade' returned exit code '$($bdfrUpdateProcess.ExitCode)'! See opened logs folder."}
}
if (-not($bdfrhtmlInstalled))
{
    Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Installing BDFR-HTML Python module ..." -ForegroundColor Cyan
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Cloning Git repository for BDFR-HTML module ..."
    if (-not(Test-Path "$rootFolder\tools\bdfr-html" -PathType Container)){New-Item -Path "$rootFolder\tools\bdfr-html" -ItemType Directory -Force -ErrorAction Stop | Out-Null}
    $bdfrhtmlCloneProcess = Start-Process "git.exe" -ArgumentList "clone https://github.com/BlipRanger/bdfr-html" -WorkingDirectory "$rootFolder\tools" -WindowStyle Hidden -RedirectStandardError "$rootFolder\logs\BDFRHTML_cloneErrors.txt" -RedirectStandardOutput "$rootFolder\logs\BDFRHTML_cloneStdOut.txt" -PassThru -Wait
    if ($bdfrhtmlCloneProcess.ExitCode -ne 0)
    {
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Warning: exit code $($bdfrhtmlCloneProcess.ExitCode) from GitHub during BDFR-HTML repository clone!" -ForegroundColor Yellow
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Retrying Git repository clone for BDFR-HTML module ..."
        Remove-Item -Path "$rootFolder\tools\bdfr-html" -Recurse -Force -ErrorAction Stop | Out-Null
        $bdfrhtmlCloneRetryProcess = Start-Process "git.exe" -ArgumentList "clone https://github.com/BlipRanger/bdfr-html" -WorkingDirectory "$rootFolder\tools" -WindowStyle Hidden -RedirectStandardError "$rootFolder\logs\BDFRHTML_cloneRetryErrors.txt" -RedirectStandardOutput "$rootFolder\logs\BDFRHTML_cloneRetryStdOut.txt" -PassThru -Wait
        if ($bdfrhtmlCloneRetryProcess.ExitCode -ne 0){Start-Process "$rootFolder\logs"; throw "[$(Get-Date -f HH:mm:ss.fff)] Error while retrying BDFR-HTML repository clone. Command: 'gh.exe repo clone BlipRanger/bdfr-html' returned exit code '$($bdfrhtmlCloneRetryProcess.ExitCode)'! See opened logs folder."}
    }
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Upgrading pip module to latest version ..."
    Start-Process "python.exe" -ArgumentList "-m pip install --upgrade pip" -WindowStyle Hidden -Wait
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Running BDFR-HTML module setup script ..."
    $bdfrhtmlScriptProcess = Start-Process "python.exe" -ArgumentList "setup.py install" -WorkingDirectory "$rootFolder\tools\bdfr-html" -WindowStyle Hidden -RedirectStandardError "$rootFolder\logs\BDFRHTML_installScriptErrors.txt" -RedirectStandardOutput "$rootFolder\logs\BDFRHTML_installScriptStdOut.txt" -PassThru -Wait
    if ($bdfrhtmlScriptProcess.ExitCode -ne 0)
    {
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)] Warning: exit code $($bdfrhtmlScriptProcess.ExitCode) from Python during BDFR-HTML module setup script!" -ForegroundColor Yellow
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Attempting alternate Pillow module install via pip ..."
        Start-Process "pip.exe" -ArgumentList "install pillow" -WindowStyle Hidden -Wait
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Rerunning BDFR-HTML module setup script ..."
        $bdfrhtmlScriptRetryProcess = Start-Process "python.exe" -ArgumentList "setup.py install" -WorkingDirectory "$rootFolder\tools\bdfr-html" -WindowStyle Hidden -RedirectStandardError "$rootFolder\logs\BDFRHTML_installScriptRetryErrors.txt" -RedirectStandardOutput "$rootFolder\logs\BDFRHTML_installScriptRetryStdOut.txt" -PassThru -Wait
        if ($bdfrhtmlScriptRetryProcess.ExitCode -ne 0){Start-Process "$rootFolder\logs"; throw "[$(Get-Date -f HH:mm:ss.fff)] Error while retrying BDFR-HTML module setup script. Command: 'python.exe $rootFolder\tools\bdfr-html\setup.py install' returned exit code '$($bdfrhtmlScriptRetryProcess.ExitCode)'! See opened logs folder."}
    }
}

## Recheck for Python modules BDFR and BDFR-HTML (if modules weren't present)
if (-not($bdfrInstalled -and $bdfrhtmlInstalled))
{
    [string[]]$installedPythonModules = @(pip freeze --disable-pip-version-check)
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Installed $($installedPythonModules.count) Python modules: $($installedPythonModules -join ', ')"
    foreach ($installedPythonModule in $installedPythonModules)
    {
        if ($installedPythonModule -like "bdfr==*"){$bdfrInstalled = $true}
        if ($installedPythonModule -like "bdfrtohtml==*"){$bdfrhtmlInstalled = $true}
    }
    if (-not($bdfrInstalled -and $bdfrhtmlInstalled)){throw "Error: Python modules are still not present: [ BDFR installed: $bdfrInstalled ] [ BDFR-HTML installed: $bdfrhtmlInstalled ]"}
}

## Subreddit archive loop init
$startDateTime = Get-Date
[int]$subLoopCount = 1; [int]$totalCloneRetries = 0; [int]$totalCloneSuccess = 0
if ($Subreddit){[string[]]$Subreddits += $Subreddit} #single
foreach ($Sub in $Subreddits)
{
    ## Progress
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Initializing"
    [int]$percentComplete = ((($subLoopCount - 1) / ($Subreddits.count)) * 100) #count -1 so the first loop reports 0% complete when just starting
    Write-Progress -Activity "Archiving subreddit $subLoopCount of $($Subreddits.count) ..." -Status "Processing: /r/$Sub" -PercentComplete $percentComplete
    
    ## Function for BDFR clone retries
    function Clone-Subreddit
    {
        param ([string[]]$IncludeIDsFilePath, [string[]]$ExcludeIDs)

        ## Check / Create / Clean output folders
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Checking, creating, cleaning output folders ..."
        foreach ($outputSubFolder in @('JSON','HTML')){
            if (-not(Test-Path "$rootFolder\$outputSubFolder\$Sub\" -PathType Container)){New-Item -Path "$rootFolder\$outputSubFolder\$Sub" -ItemType Directory -ErrorAction Stop | Out-Null}
            if (Get-ChildItem -Path "$rootFolder\$outputSubFolder\$Sub\*" -File -ErrorAction SilentlyContinue){Remove-Item -Path "$rootFolder\$outputSubFolder\$Sub\*" -Recurse -ErrorAction Stop | Out-Null}}
            
        ## BDFR: Clone subreddit to JSON
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Using BDFR to clone subreddit to disk in JSON ..."
        [string]$global:logPath = "$rootFolder\logs\bdfr_$($Sub)_$(Get-Date -f yyyyMMdd_HHmmss).log.txt" #global so these variables don't disappear each time the function exits
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Status (CTRL+C to retry): $logPath"
        if ($IncludeIDsFilePath){$global:bdfrProcess = Start-Process "python.exe" -ArgumentList "-m bdfr clone $rootFolder\JSON --subreddit $Sub --include-id-file $IncludeIDsFilePath --verbose --log $logPath" -WindowStyle Hidden -PassThru}
        elseif ($ExcludeIDs){$global:bdfrProcess = Start-Process "python.exe" -ArgumentList "-m bdfr clone $rootFolder\JSON --subreddit $Sub --exclude-id $($ExcludeIDs -join ' --exclude-id ') --verbose --log $logPath" -WindowStyle Hidden -PassThru}
        else {$global:bdfrProcess = Start-Process "python.exe" -ArgumentList "-m bdfr clone $rootFolder\JSON --subreddit $Sub --verbose --log $logPath" -WindowStyle Hidden -PassThru} #--disable-module Youtube --disable-module YoutubeDlFallback 
        
        ## Custom CTRL+C handling and timeout detection
        [int]$lastTotalCloneOutputGB = 0
        [bool]$global:CTRLCUsedOnce = $false; [bool]$global:cloneHangDetected = $false; [bool]$global:outputFolderGrowth = $false
        if (-not($Background)){[console]::TreatControlCAsInput = $true} #change the default behavior of CTRL+C so that the script can intercept and use it versus just terminating the script: https://blog.sheehans.org/2018/10/27/powershell-taking-control-over-ctrl-c/
        if (-not($Background)){Start-Sleep -Seconds 1} #sleep for 1 second and then flush the key buffer so any previously pressed keys are discarded and the loop can monitor for the use of CTRL+C. The sleep command ensures the buffer flushes correctly
        if (-not($Background)){$host.UI.RawUI.FlushInputBuffer()}
        $cloneTimeout = New-TimeSpan -Hours 4 #clone operation will retry after 4 hours elapsed, regardless of process status (unless output folder has grown in size by 1GB or more since last check)
        $cloneStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        while (-not($bdfrProcess.HasExited))
        {
            if (-not($Background)) #only read host keys if interactive
            {
                if ($host.UI.RawUI.KeyAvailable -and ($key = $host.UI.RawUI.ReadKey("AllowCtrlC,NoEcho,IncludeKeyUp"))) #if a key was pressed during the loop execution, check to see if it was CTRL+C (aka "3"), and if so exit the script after clearing out any running python processes and setting CTRL+C back to normal
                {
                    if ([int]$key.Character -eq 3){$global:CTRLCUsedOnce = $true; break} #CTRL+C pressed, exit while loop
                    $host.UI.RawUI.FlushInputBuffer()
                }
            }
            else {Start-Sleep -Seconds 10} #if not interactive, we don't need to loop quickly to catch CTRL+C presses
            if ($cloneStopwatch.Elapsed -gt $cloneTimeout)
            {
                [int]$totalCloneOutputGB = ((Get-ChildItem -Path "$rootFolder\JSON\$Sub\" | Measure-Object -Sum -Property Length).Sum / 1GB)
                if ($totalCloneOutputGB -gt $lastTotalCloneOutputGB) #timeout reached, but check if successfully downloading huge media files during the timespan
                {
                    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Timeout reached, but timer has been reset: output folder has grown by $($totalCloneOutputGB - $lastTotalCloneOutputGB)GB!"
                    $cloneStopwatch.Restart() #restart stopwatch from 0 to allow another 4 hours
                    [int]$lastTotalCloneOutputGB = $totalCloneOutputGB #continue if output folder size is increasing by > 1GB/4hrs
                    $global:outputFolderGrowth = $true #prevent retries.. and it is extremely unlikely an interactive console has been running for so long, so that CTRL+C case is NYI below
                }
                else {$global:cloneHangDetected = $true; break} #over 4 hours have passed, and output folder has not grown by atleast 1GB, exit while loop, and immediately re-attempt (also triggers when the completed process fails to exit, but does NOT retry in that case)
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
                        [console]::TreatControlCAsInput = $false
                        Write-Host "`n[$(Get-Date -f HH:mm:ss.fff)][$Sub] CTRL+C pressed twice: waiting up to 500ms for python process to exit ..." -ForegroundColor Yellow #`n = line feed for the -NoNewLine periods
                        if (-not($bdfrProcess.WaitForExit(500))){throw "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Failed to taskkill python process ID '$($bdfrProcess.ID)'!"}
                        exit 1
                    }
                    $host.UI.RawUI.FlushInputBuffer()
                }
                [int]$previousSecondsElapsed = $secondsElapsed #doing it this way instead of a Start-Sleep so the script is more responsive to CTRL+C being pressed..
                [int]$secondsElapsed = ((Get-Date) - $loopStartTime).TotalSeconds
                if ($secondsElapsed -gt $previousSecondsElapsed){Write-Host '.' -ForegroundColor Yellow -NoNewLine; $waitLoopSeconds--} #..while still writing one period per second to the console
            }
            Write-Host '.' -ForegroundColor Yellow
        }

        ## End function
        if (-not($Background)){[console]::TreatControlCAsInput = $false}
        if ($cloneHangDetected){Start-Process 'taskkill.exe' -ArgumentList "/F /PID $($bdfrProcess.ID)" -WindowStyle Hidden -ErrorAction Stop -Wait}
        if (Get-Process -ID ($bdfrProcess.ID) -ErrorAction SilentlyContinue)
        {
            Stop-Process $bdfrProcess -Force -ErrorAction SilentlyContinue #even after ($bdfrProcess.HasExited -eq $true), this Get-Process on the PID still sometimes returns the process ..so attempt another stop
            if (-not($bdfrProcess.WaitForExit(2000))){Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Warning: python process ID '$($bdfrProcess.ID)' still running!" -ForegroundColor Yellow} #only warn because an occasional still-running Python instance will not impact script execution
        }
    }

    ## BDFR: Initial clone attempt and function retry loop
    Clone-Subreddit
    [int]$triesLeft = 10
    [bool]$prawError = $false
    [bool]$cloneSuccessful = $false
    [string[]]$errorSubmissionIDs = @()
    while (($triesLeft -gt 0) -and (-not($cloneSuccessful)) -and (-not($outputFolderGrowth)))
    {
        $logLastLine = Get-Content $logPath -Tail 1 -ErrorAction SilentlyContinue
        if (-not($logLastLine -like "*INFO] - Program complete"))
        {
            ## Check log for recurring problem submission IDs (and --exclude-id param them on retries) (gci "$rootFolder\logs" | ? {$_.Name -like "bdfr*.log.txt"} | % {gc -Path $_.FullName -tail 1} | % {if ($_ -match "^OSError.*_(?<ID>[a-z0-9]{5,6})\.json'$"){$matches.ID}})
            if ($logLastLine -match "^OSError.*Invalid argument.*_(?<ID>[a-z0-9]{5,6})\.json'$"){$errorSubmissionIDs += $matches.ID}; $matches = $null #intended to match a newline/illegal character filesystem issue.. generates a $matches automatic variable with named group 'ID' for the submission ID
            if ($logLastLine -match "^praw\.exceptions\.InvalidURL\: Invalid URL\: (?<ID>[a-z0-9]{5,6})$"){$prawError = $true}; $matches = $null #if there's a praw error that spits the submission ID as the URL, make a best effort partial clone below
            [string[]]$excludeIDs = @(($errorSubmissionIDs | Group-Object | Where-Object {$_.Count -ge 2}).Name) #if the same submission ID has generated two errors, exclude the ID from the next BDFR clone attempt

            ## Check if excluded IDs are still generating errors, and try an inclusion list instead (for --include-id-file param)
            if (-not($prawError))
            {
                [string[]]$stillFailingExcludeIDs = @(($errorSubmissionIDs | Group-Object | Where-Object {$_.Count -ge 3}).Name) #on three or more errors (once post --exclude-id addition), instead include all successful submission IDs, before the error ID
                if ($stillFailingExcludeIDs)
                {
                    [string[]]$includeIDs = @() #workaround for archiver still attempting to clone --exclude-id post IDs (https://github.com/aliparlakci/bulk-downloader-for-reddit/issues/618)
                    Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub][Retry] Attempting clone with include ID list as excluded ID(s) still failing: $($stillFailingExcludeIDs -join ', ')" -ForegroundColor Yellow
                    [string]$includeIDsFilePath = "$rootFolder\logs\bdfr_$($Sub)_includedSubmissionIDs_$(Get-Date -f yyyyMMdd_HHmmss).txt"
                    Get-Content $logPath -ErrorAction SilentlyContinue | Where-Object {$_ -match "(archiver - INFO] - Record for entry item )(?<ID>[a-z0-9]{5,6})( written to disk)"} | ForEach-Object {$includeIDs += $matches.ID}
                    if ($includeIDs){$includeIDs | Set-Content $includeIDsFilePath}; $matches = $null
                }
                elseif ($excludeIDs){Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub][Retry] Excluding failing submission ID(s) from retry: $($excludeIDs -join ', ') ..." -ForegroundColor Yellow}
            }
            else {Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub][Retry] Encountered PRAW Invalid URL exception. Restarting once more for partial clone ..." -ForegroundColor Yellow}

            ## Retry clone operation
            $totalCloneRetries++
            if ($CTRLCUsedOnce){$retryReason = 'User cancelled'}
            if ($cloneHangDetected){$retryReason = 'Hang/timeout over 4 hours during'}
            if ((-not($CTRLCUsedOnce)) -and (-not($cloneHangDetected))){$retryReason = "Error '$($bdfrProcess.ExitCode)' during"}
            Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub][Retry] $retryReason BDFR clone operation -- retrying up to $triesLeft more times" -ForegroundColor Yellow
            if ($retryReason -like "Error*"){Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub][Retry] Last log line: $logLastLine" -ForegroundColor Yellow}
            [int]$sleepMinutes = $totalCloneRetries - ($totalCloneSuccess * 5) #sleep one minute for every retry/error/cancel, but remove five minutes for each success (and negative values act as credits toward future strings of errors)
            if ($sleepMinutes -ge 1){Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub][Retry] Sleeping $sleepMinutes minute(s) before trying again ..." -ForegroundColor Yellow; Start-Sleep -Seconds ($sleepMinutes * 60)}
            if ($prawError){Clone-Subreddit; break} #make a best effort partial clone attempt with no inclusions/exclusions because they're generating praw errors (https://github.com/aliparlakci/bulk-downloader-for-reddit/issues/620)
            elseif ($includeIDs){Clone-Subreddit -IncludeIDsFilePath $includeIDsFilePath}
            elseif ($excludeIDs){Clone-Subreddit -ExcludeIDs $excludeIDs}
            else {Clone-Subreddit}
            $triesLeft--
        }
        else {$cloneSuccessful = $true; $totalCloneSuccess++}
    }
    if (-not($cloneSuccessful))
    {
        ## BDFR: Clone did not complete but process successfully retrieved submissions anyway
        if ($includeIDs){$addlParams = "--include-id-file $includeIDsFilePath "} elseif ($excludeIDs){$addlParams = "--exclude-id $($ExcludeIDs -join ' --exclude-id ') "} else {$addlParams = ''} #so the error contains the full command string even with exclude/include ID params
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Unrecoverable errors: command: 'python.exe -m bdfr clone $rootFolder\JSON --subreddit $Sub $($addlParams)--verbose --log $logPath' returned exit code '$($bdfrProcess.ExitCode)'! This was the final retry attempt." -ForegroundColor Yellow
        [int]$retrievedSubmissions = (Get-ChildItem "$rootFolder\JSON\$Sub" -Recurse -Include "*.json" | Measure-Object).Count
        Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Processing partial clone of $retrievedSubmissions retrieved submissions ..." -ForegroundColor Yellow
    }
    if ($CTRLCUsedOnce -or $cloneHangDetected){Write-Host "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Information: the cancelled clone operation had already completed!" -ForegroundColor Cyan} #rarely the clone operation succeeds around the same time a clone is cancelled (or the process hangs after completion), so in that case, acknowledge (but ignore and don't retry or exit)

    ## BDFR-HTML: Process cloned subreddit JSON into HTML
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Using BDFR-HTML to generate HTML pages from JSON archive ..."
    $bdfrhtmlProcess = Start-Process "python.exe" -ArgumentList "-m bdfrtohtml --input_folder $rootFolder\JSON\$Sub --output_folder $rootFolder\HTML\$Sub --write_links_to_file All" -WindowStyle Hidden -PassThru -Wait
    if ($bdfrhtmlProcess.ExitCode -ne 0){throw "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Error: Command: 'python.exe -m bdfrtohtml --input_folder $rootFolder\JSON\$Sub --output_folder $rootFolder\HTML\$Sub' returned exit code '$($bdfrhtmlProcess.ExitCode)'!"} #this process/module has never failed on me, so there's no error handling

    ## Replace generated index.html <title> with subreddit name
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Updating generated index.html title tag with subreddit name ..."
    (Get-Content "$rootFolder\HTML\$Sub\index.html").Replace("<title>BDFR Archive</title>","<title>/r/$Sub Archive</title>") | Set-Content "$rootFolder\HTML\$Sub\index.html" -Encoding 'UTF8' -Force
    
    ## Delete media files over 5MB threshold (if switch not present)
    if (-not($NoMediaPurge))
    {
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Deleting HTML media folder files over 5MB ..."
        Get-ChildItem -Path "$rootFolder\HTML\$Sub\media" | Where-Object {(($_.Length)/1MB) -gt 5} | ForEach-Object {Remove-Item -LiteralPath "$($_.FullName)" -ErrorAction SilentlyContinue}
    }

    ## Loop end
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)][$Sub] Finished!"
    if (($Subreddits.count -eq 1) -or ($subLoopCount -eq $Subreddits.count)){Write-Progress -Activity "Archiving subreddit $subLoopCount of $($Subreddits.count) ..." -Completed} #if last loop, 100% progress
    else {$subLoopCount++}
}

## Script end / Package output
if ($Subreddits.count -eq 1)
{
    ## Single subreddit: open HTML output folder
    Start-Process "$rootFolder\HTML\$Subreddit\"
}
else
{
    ## Multiple subreddits
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Completed all subreddit HTML archives!"
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
    $indexContents += '<html><head><style>body {background-color: rgb(127, 127, 127);}</style><title>BDFR Archive Index</title></head><body><ul>'
    foreach ($subredditDirectoryName in @((Get-ChildItem $zipOutputPath -Directory).Name))
    {
        $indexContents += "<li><a href=`"./$($subredditDirectoryName)/index.html`"><h2>/r/$subredditDirectoryName</h2></a></li>" #unordered list
    }
    $indexContents += "</ul><hr>" #horizontal rule
    $indexContents += "Archive started: $startDateTime<br>"
    $indexContents += "Archive complete: $endDateTime<br>"
    $indexContents += "Time elapsed: $(($timeElapsed.TotalHours).ToString("###.##")) hours<br>"
    $indexContents += "JSON folder size: $totalJSONSize<br>"
    $indexContents += "HTML folder size: $totalHTMLSize"
    $indexContents += '</body></html>' #footer stats
    $indexContents | Set-Content "$zipOutputPath\index.html" -Encoding UTF8

    ## Compress ZIP & Open output
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Compressing everything into a ZIP archive ..."
    Compress-Archive -Path "$zipOutputPath\*" -Destination "$zipOutputPath\$(Get-Date -f yyyy-MM-dd)_indexed_HTML_archive_of_$($Subreddits.count)_subreddits.zip"
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Finished! Time elapsed: $(($timeElapsed.TotalHours).ToString("###.##")) hours"
    if (-not($Background)){Start-Process $zipOutputPath}
    else {Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Indexed archives and ZIP file output path:`n$zipOutputPath"}
}