<#
.SYNOPSIS
    Checks for prerequisites, then uses BDFR and BDFR-HTML to generate a subreddit HTML archive.
    Creates BDFR folder and subfolders under your %USERPROFILE% ($env:USERPROFILE) Documents folder.
    If running for the first time, run with administrator privileges. This is only needed once.
.DESCRIPTION
    This script uses the following Python modules, which are installed automatically:
        BDFR: https://pypi.org/project/bdfr/
        BDFR-HTML: https://github.com/BlipRanger/bdfr-html
            - When running setup.py to install BDFR-HTML (via script or manually), you may get an install error from Pillow about zlib being missing. You may need to run 'pip install pillow' from an elevated command prompt, so that Pillow installs correctly.
            - For manual BDFR-HTML install in case of Pillow install error: From an elevated CMD window, type these two quoted commands: 1) "cd %USERPROFILE%\Documents\BDFR\module_clone\bdfr-html", 2) "python.exe setup.py install"
            - https://stackoverflow.com/questions/64302065/pillow-installation-pypy3-missing-zlib
    Prerequisite tools that must be installed before running this script:
        1. Git: https://github.com/git-for-windows/git/releases/download/v2.33.0.windows.2/Git-2.33.0.2-64-bit.exe
        2. GitHub CLI: https://github.com/cli/cli/releases/download/v2.3.0/gh_2.3.0_windows_amd64.msi
            - You'll need to launch Git CMD and authenticate with 'gh auth login', and follow the prompts, pasting the OTP into your browser.
        3. Python 3.x (includes pip): https://www.python.org/ftp/python/3.10.0/python-3.10.0-amd64.exe
            - At beginning of install, YOU MUST CHECK 'Add Python 3.x to PATH'. (So PowerShell can call python.exe from anywhere)
.PARAMETER Subreddit
    The name of the subreddit (as it appears after the /r/ in the URL) that will be archived.
.EXAMPLE
    .\New-SubredditHTMLArchive.ps1 -Subreddit PowerShell
.NOTES
    The reddit API returns a maximum of 1000 posts per BDFR, so only the newest 1000 posts will be included:
    https://github.com/reddit-archive/reddit/blob/master/r2/r2/lib/db/queries.py
.NOTES
    Script URL: https://github.com/mbarr564/powershell/blob/master/New-SubredditHTMLArchive.ps1
.NOTES
    Last update: Monday, January 17, 2022 7:20:56 PM
#>

param([string]$Subreddit)

## Init
if (-not($Subreddit)){$Subreddit = Read-Host -Prompt 'Enter subreddit to archive'}
if (-not($Subreddit)){throw 'Error: Subreddit name is blank!'}
$stopWatch = New-Object System.Diagnostics.Stopwatch
$stopWatch.Start()

## Check for Command Line Utilities
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Checking for Command Line Utilities ..."
foreach ($exeName in @('git','gh','python','pip')){if (-not(Get-Command "$($exeName).exe" -ErrorAction SilentlyContinue)){throw "Error: Missing command line utility prerequisite: $($exeName).exe. See script comment header description for installers."}}
if ((&{git --version}) -notlike "*version 2.*"){throw 'Error: Git version 2 is required!'}
if ((&{gh --version})[0] -notlike "*version 2.*"){throw 'Error: GitHub CLI version 2 is required!'}
if ((&{python -V}) -notlike "*Python 3*"){throw 'Error: Python version 3 is required!'}
if ((&{pip -V}) -notlike "*pip 2*"){throw 'Error: Pip version 2 is required!'}

## Check/Create BDFR Output Folders
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Checking and cleaning BDFR output folders ..."
[string]$bdfrFolderRoot = "$($env:USERPROFILE)\Documents\BDFR"
[string]$bdfrJSONFolder = "$bdfrFolderRoot\JSON"; [string]$bdfrHTMLFolder = "$bdfrFolderRoot\HTML"
if (-not(Test-Path "$bdfrJSONFolder\$Subreddit\log" -PathType Container))
{
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Creating BDFR output folders at $bdfrFolderRoot ..."
    try {
        New-Item -Path "$bdfrFolderRoot\module_clone" -ItemType Directory -Force -ErrorAction Stop | Out-Null
        New-Item -Path "$bdfrJSONFolder\$Subreddit\log" -ItemType Directory -Force -ErrorAction Stop | Out-Null
        New-Item -Path "$bdfrHTMLFolder\$Subreddit" -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    catch {throw "Error: While creating output folders under: $bdfrFolderRoot"}
}

## Remove Existing Files in Output Folders
if (Get-ChildItem -Path "$bdfrJSONFolder\$Subreddit\*" -File -ErrorAction SilentlyContinue){Remove-Item -Path "$bdfrJSONFolder\$Subreddit\*" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null; New-Item -Path "$bdfrJSONFolder\$Subreddit\log" -ItemType Directory | Out-Null}
if (Get-ChildItem -Path "$bdfrHTMLFolder\$Subreddit\*" -File -ErrorAction SilentlyContinue){Remove-Item -Path "$bdfrHTMLFolder\$Subreddit\*" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null}

## Check for Python modules BDFR and BDFR-HTML
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Checking for BDFR and BDFR-HTML Python modules ..."
[boolean]$bdfrInstalled = $false
[boolean]$bdfrhtmlInstalled = $false
$installedPythonModules = @(pip list --disable-pip-version-check)
foreach ($installedPythonModule in $installedPythonModules)
{
    if ($installedPythonModule -like "bdfr *"){$bdfrInstalled = $true}
    if ($installedPythonModule -like "bdfrtohtml*"){$bdfrhtmlInstalled = $true}
}

## Install Python modules BDFR and BDFR-HTML
if (-not($bdfrInstalled))
{
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Installing BDFR Python module ..."
    $bdfrInstallProcess = Start-Process "python.exe" -ArgumentList "-m pip install bdfr --upgrade" -WindowStyle Hidden -PassThru -Wait
    if ($bdfrInstallProcess.ExitCode -ne 0){throw "Error: Command: 'python.exe -m pip install bdfr --upgrade' returned exit code '$($bdfrInstallProcess.ExitCode)'!"}
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Updating BDFR Python module ..."
    $bdfrUpdateProcess = Start-Process "python.exe" -ArgumentList "-m pip install bdfr --upgrade" -WindowStyle Hidden -PassThru -Wait
    if ($bdfrUpdateProcess.ExitCode -ne 0){throw "Error: Command: 'python.exe -m pip install bdfr --upgrade' returned exit code '$($bdfrUpdateProcess.ExitCode)'!"}
}
if (-not($bdfrhtmlInstalled))
{
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Installing BDFR-HTML Python module ..."
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Cloning GitHub repository for BDFR-HTML module ..."
    $bdfrhtmlCloneProcess = Start-Process "gh.exe" -ArgumentList "repo clone BlipRanger/bdfr-html" -WorkingDirectory "$bdfrFolderRoot\module_clone" -WindowStyle Hidden -PassThru -Wait
    if ($bdfrhtmlCloneProcess.ExitCode -ne 0){throw "Error: Command: 'gh.exe repo clone BlipRanger/bdfr-html' returned exit code '$($bdfrhtmlCloneProcess.ExitCode)'!"}
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Upgrading pip module to latest version ..."
    Start-Process "python.exe" -ArgumentList "-m pip install --upgrade pip" -WindowStyle Hidden -Wait
    Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Running BDFR-HTML module setup script ..."
    $bdfrhtmlScriptProcess = Start-Process "python.exe" -ArgumentList "setup.py install" -WorkingDirectory "$bdfrFolderRoot\module_clone\bdfr-html" -WindowStyle Hidden -PassThru -Wait
    if ($bdfrhtmlScriptProcess.ExitCode -ne 0)
    {
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Error during BDFR-HTML module setup script!"
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Attempting alternate Pillow module install via pip ..."
        Start-Process "pip.exe" -ArgumentList "install pillow" -WindowStyle Hidden -Wait
        Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Rerunning BDFR-HTML module setup script ..."
        $bdfrhtmlScriptRetryProcess = Start-Process "python.exe" -ArgumentList "setup.py install" -WorkingDirectory "$bdfrFolderRoot\module_clone\bdfr-html" -WindowStyle Hidden -PassThru -Wait
        if ($bdfrhtmlScriptRetryProcess.ExitCode -ne 0){throw "Error: Command: 'python.exe $bdfrFolderRoot\module_clone\setup.py install' returned exit code '$($bdfrhtmlScriptRetryProcess.ExitCode)'!"}
    }
}

## Recheck for Python modules BDFR and BDFR-HTML
$installedPythonModules = @(pip list --disable-pip-version-check)
foreach ($installedPythonModule in $installedPythonModules)
{
    if ($installedPythonModule -like "bdfr *"){$bdfrInstalled = $true}
    if ($installedPythonModule -like "bdfrtohtml*"){$bdfrhtmlInstalled = $true}
}
if (-not($bdfrInstalled -and $bdfrhtmlInstalled)){throw "Error: Python modules BDFR and/or BDFR-HTML are still not present!"}

## BDFR: Clone Subreddit to JSON
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Using BDFR to clone subreddit '$Subreddit' to disk ..."
[string]$logPath = "$bdfrJSONFolder\$Subreddit\log\bdfr_$($Subreddit)_$(Get-Date -f yyyyMMdd_HHmmss).log.txt"
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Status: $logPath"
$bdfrProcess = Start-Process "python.exe" -ArgumentList "-m bdfr clone $bdfrJSONFolder --subreddit $Subreddit --disable-module Youtube --disable-module YoutubeDlFallback --log $logPath" -WindowStyle Hidden -PassThru

## BDFR: custom CTRL+C handling to taskkill PID for python.exe
[console]::TreatControlCAsInput = $true #change the default behavior of CTRL-C so that the script can intercept and use it versus just terminating the script: https://blog.sheehans.org/2018/10/27/powershell-taking-control-over-ctrl-c/
Start-Sleep -Seconds 1 #sleep for 1 second and then flush the key buffer so any previously pressed keys are discarded and the loop can monitor for the use of CTRL-C. The sleep command ensures the buffer flushes correctly
$host.UI.RawUI.FlushInputBuffer()
while (-not($bdfrProcess.HasExited)) #loop while the python BDFR process exists
{
    if ($host.UI.RawUI.KeyAvailable -and ($key = $host.UI.RawUI.ReadKey("AllowCtrlC,NoEcho,IncludeKeyUp"))) #if a key was pressed during the loop execution, check to see if it was CTRL-C (aka "3"), and if so exit the script after clearing out any running python processes and setting CTRL-C back to normal
    {
        if ([int]$key.Character -eq 3)
        {
            Write-Warning "CTRL-C used: shutting down running python process before exiting ..."
            Start-Process "taskkill.exe" -ArgumentList "/F /PID $($bdfrProcess.ID)" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue #still executes upon CTRL+C, to clean up the running python process
            [console]::TreatControlCAsInput = $false
            exit 1
        }
        $host.UI.RawUI.FlushInputBuffer() #flush the key buffer again for the next loop
    }
}

## BDFR: Check log for success
if (-not((Get-Content $logPath -Tail 1) -like "*INFO] - Program complete")){throw "[$(Get-Date -f HH:mm:ss.fff)] Error: Command: 'python.exe -m bdfr clone $bdfrJSONFolder --subreddit $Subreddit --disable-module Youtube --disable-module YoutubeDlFallback --log $logPath' returned exit code '$($bdfrProcess.ExitCode)'!"} #this process often throws errors, so check for the program completion string in the tail of the log file

## BDFR-HTML: Process Cloned Subreddit to HTML
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Using BDFR-HTML to generate subreddit HTML pages ..."
$bdfrhtmlProcess = Start-Process "python.exe" -ArgumentList "-m bdfrtohtml --input_folder $bdfrJSONFolder\$Subreddit --output_folder $bdfrHTMLFolder\$Subreddit" -WindowStyle Hidden -PassThru -Wait
if ($bdfrhtmlProcess.ExitCode -ne 0){throw "Error: Command: 'python.exe -m bdfrtohtml --input_folder $bdfrJSONFolder\$Subreddit --output_folder $bdfrHTMLFolder\$Subreddit' returned exit code '$($bdfrhtmlProcess.ExitCode)'!"}

## Replace generated index.html <title> with subreddit name
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Updating generated index.html, adding subreddit to title ..."
[boolean]$indexLineFound = $false #skip -like operator
$indexFile = Get-Item "$bdfrHTMLFolder\$Subreddit\index.html"
$indexReader = New-Object -TypeName 'System.IO.StreamReader' -ArgumentList $indexFile
$indexOutput = New-Object -TypeName 'System.Collections.ArrayList'
while (-not($indexReader.EndOfStream))
{
    $indexLine = $indexReader.ReadLine()
    if ($indexLineFound){[void]$indexOutput.Add("$indexLine"); continue}
    if ($indexLine -like "*<title>*")
    {
        [void]$indexOutput.Add("        <title>/r/$Subreddit Archive</title>")
        $indexLineFound = $true
    }
    else {[void]$indexOutput.Add("$indexLine")}
}
$indexReader.Close(); $indexOutput | Set-Content $indexFile -Encoding 'UTF8' -Force

## Delete media files over 2MB threshold
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Deleting media folder files over 2MB in HTML output folder ..."
(Get-ChildItem -Path "$bdfrHTMLFolder\$Subreddit\media" | Where-Object {(($_.Length)/1MB) -gt 2}).FullName | ForEach-Object {if ($_){Remove-Item -Path $_ -Force -ErrorAction SilentlyContinue}}

## End
$stopWatch.Stop()
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Finished! Run time was $($stopWatch.Elapsed.Hours) hour(s) $($stopWatch.Elapsed.Minutes) minute(s) $($stopWatch.Elapsed.Seconds) second(s)."

## Open Completed HTML Folder
Write-Output "[$(Get-Date -f HH:mm:ss.fff)] Opening subreddit HTML output folder ..."
Start-Process "$bdfrHTMLFolder\$Subreddit\"
exit 0

# SIG # Begin signature block
# MIIVpAYJKoZIhvcNAQcCoIIVlTCCFZECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUygHVy9iuTJQl0NQoO/5FNcwv
# McmgghIFMIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
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
# FBp1nFom8p95bU2A66FFT2pRnNUrMA0GCSqGSIb3DQEBAQUABIICAHHkXJMZQXix
# LK6KhMRUERKX/O3lBdHV1pU5xc8yRz3f1T05MUWeO5GP3UbxDmYlH0HKZBRvYunx
# WtqPMCdJdRhANxYpa5AGbje1GbwAXS3fHMSW3zvwV2OfTBhHgJTKzED2VTP/ta4S
# KUkw28JTMEcMx+5ApvQ1e9UnX+GAI4NOoeT6fbXADHwFcSilZJlC0/OvK8hymEae
# B4UyMdz55LGJ1Dcm7aAqXDxuz3NRiYAyG2eYVvzwd/g+msxw+HqVH+QTJowVAulm
# Iq+f+HeQLYZ9I/xC8Y2+ZkV/QTTdojiKoU1iCr1D93VFPauNzZd9xVPPJ5lUki2U
# sDUO8Tdq7MFyNbc5eU7IJYBU3xPJgZmLVXywQKDuR4xvkFzBLrwWOgmRMeAGLfTS
# 7yMHhUr1VUa75bWtq5NbOsyIdgHAplhT4OuFP8nNXcVtEYIbB2wuz+NWJNwCqg2+
# ffaXJ9K+hBWwjX/277AUCkGsrvjfvMWaBve3B2l40ExN5w+ClPFlY6xSQeVhcsaZ
# G4TM6xuyvysIF52TKNZr0V4EX1X4KGAx/FW4tkptqzkuQpbPUOXzzH0UQuuCuHc8
# qmGnTQQa/7ZE1YP5wrmjcqK2XEAwEOeDxj5zQ1zbp+SaZ8TKQyPjO4pFZgBll5mw
# NPt3rtBxqI8939wZkIwLKkH8uX4GCDTI
# SIG # End signature block
