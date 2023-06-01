@echo off
title New-SubredditHTMLArchive 2.2.2 setup
echo This batch script completes all initial setup steps for the New-SubredditHTMLArchive
echo PowerShell script, with a right click and then a 'Run as administrator' click.
echo You will need to be present during install to approve multiple prompts.
echo.

REM ## Init / Check Admin
SET localPolicy=FALSE
SET changePolicy=FALSE
SET installNow=FALSE
SET isAdmin=TRUE
net.exe session 1>NUL 2>NUL || ( @SET isAdmin=FALSE )
IF %isAdmin%==FALSE (
    echo You MUST RIGHT CLICK this script and "Run as administrator".. Try again. Exiting.
    GOTO:END
)

REM ## Get-ExecutionPolicy
echo Checking PowerShell Execution Policy ...
FOR /f "skip=2 tokens=2*" %%a IN ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell" /v ExecutionPolicy') DO @SET "localPolicy=%%b"
IF %localPolicy%==Restricted ( @SET changePolicy=TRUE )
IF %localPolicy%==AllSigned ( @SET changePolicy=TRUE )
IF %localPolicy%==RemoteSigned ( @SET changePolicy=TRUE )
IF %localPolicy%==Default ( @SET changePolicy=TRUE )
IF %localPolicy%==Undefined ( @SET changePolicy=TRUE )
IF "%localPolicy%"=="" ( @SET changePolicy=TRUE )

REM ## Set-ExecutionPolicy
IF %changePolicy%==TRUE (
    echo Setting PowerShell Execution Policy ...
    reg add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell /v "ExecutionPolicy" /d "Unrestricted" /f
    IF ERRORLEVEL 1 (
        echo Error: Unable to set PowerShell execution policy. Exiting.
        GOTO:END
    )
)

REM ## Set-ItemProperty
IF %changePolicy%==TRUE (
    echo Setting LongPathsEnabled as MAX_PATH long post title workaround ...
    REM https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry
    reg add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem /v "LongPathsEnabled" /t REG_DWORD /d "1" /f
    IF ERRORLEVEL 1 (
        echo Error: Unable to set LongPathsEnabled as MAX_PATH workaround. Exiting.
        GOTO:END
    )
)

REM ## Check prerequisites
SET requiredPythonModules=bdfr bdfrtohtml click markdown appdirs bs4 dict2xml ffmpeg-python praw pyyaml requests youtube-dl jinja2 pillow
echo Checking if required tools and packages are installed ...
where /q pip
IF ERRORLEVEL 1 ( SET installNow=TRUE )
IF %installNow%==FALSE (
    FOR %%a IN (%requiredPythonModules%) DO (
        pip show %%a -q -q -q
        IF "%ERRORLEVEL%"=="1" (
            echo Missing Python module: %%a
            SET installNow=TRUE
        )
    )
)

REM ## Invoke-WebRequest
SET localDownloadPath=%HOMEDRIVE%\Users\%USERNAME%\Documents\New-SubredditHTMLArchive
IF %installNow%==TRUE (
    IF NOT EXIST %localDownloadPath%\New-SubredditHTMLArchive.ps1 (
        echo Downloading New-SubredditHTMLArchive script to Documents folder ...
        IF NOT EXIST %localDownloadPath% ( mkdir %localDownloadPath% )
        CALL bitsadmin.exe /transfer New-SubredditHTMLArchive /download "https://raw.githubusercontent.com/mbarr564/New-SubredditHTMLArchive/main/New-SubredditHTMLArchive.ps1" %localDownloadPath%\New-SubredditHTMLArchive.ps1
        IF NOT EXIST %localDownloadPath%\New-SubredditHTMLArchive.ps1 (
            echo Error: Unable to download the New-SubredditHTMLArchive PowerShell script from github using BITSAdmin! Exiting.
            GOTO:END
        )
    ) ELSE ( echo New-SubredditHTMLArchive script already exists locally, continuing ... )

    REM ## Installing New-SubredditHTMLArchive.ps1 script via PSGallery
    echo Installing New-SubredditHTMLArchive script to environment PATH via PSGallery ...
    CALL powershell -Command "Install-Script -Name New-SubredditHTMLArchive -Force"

    REM ## New-SubredditHTMLArchive.ps1 -InstallPackages -Subreddit TestSubredditC
    echo New-SubredditHTMLArchive: Installing required packages, running test archive ...
    CALL powershell -Command "%localDownloadPath%\New-SubredditHTMLArchive.ps1 -Subreddit TestSubredditC -InstallPackages" -Verb RunAs
    GOTO:END
)

REM ## Usage Examples - Previously Installed!
echo.
echo The New-SubredditHTMLArchive script install has previously completed successfully!
echo.
echo Now, anytime you want to archive a subreddit, it is this easy:
echo 1. Open PowerShell: Start menu, type "PowerShell", click "Windows PowerShell"
echo 2. Type "New-Sub", hit TAB, and the script name should autocomplete to "New-SubredditHTMLArchive.ps1"
echo 3. Continue to type (after script name), a space, then "-Subreddit", a space, then the subreddit name.
echo 4. Hit Enter to start the archive. ETA is one hour per subreddit, without retries, or video links.
echo.
echo TL;DR: WinKey+R, "powershell", ENTER, "New-Sub", TAB, SPACE, "-Su", TAB, SPACE, "mySubreddit", ENTER.
echo.
echo Example: PS^> New-SubredditHTMLArchive.ps1 -Subreddit AnySubredditName
echo Example: PS^> New-SubredditHTMLArchive.ps1 -Subreddit HugeVideoLinksSubreddit -Background -NoMediaPurge
echo Example: PS^> New-SubredditHTMLArchive.ps1 -Subreddits PowerShell,Python,GNURadio,HomeLab -Background

:END
echo.
PAUSE