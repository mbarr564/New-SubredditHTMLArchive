## GitHub Example Usage  
1) PS> Invoke-WebRequest -URI https://raw.githubusercontent.com/mbarr564/New-SubredditHTMLArchive/main/New-SubredditHTMLArchive.ps1 -OutFile .\New-SubredditHTMLArchive.ps1  
2) PS> .\New-SubredditHTMLArchive.ps1 -Subreddit 'PowerShell' -InstallPackages  
3) PS> .\New-SubredditHTMLArchive.ps1 -Subreddits 'GNURadio','SRAWeekend','Tails' -Background
  
## PowerShell Gallery Example Usage  
1) PS> Install-Script -Name New-SubredditHTMLArchive  
2) PS> New-SubredditHTMLArchive.ps1 -Subreddit 'Python' -InstallPackages  
3) PS> New-SubredditHTMLArchive.ps1 -Subreddits 'HackRF','DataHoarder','Onions' -Background  
4) PS> Update-Script -Name New-SubredditHTMLArchive  
  
## Comment Based Help  
See: [PSScriptInfo comment header breaking 'Get-Help .\Script.ps1 -Full'](https://stackoverflow.com/questions/71579241/powershell-gallery-psscriptinfo-comment-header-breaking-get-help-myscriptname/71579958#71579958)  
1) PS> Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process  
2) PS> Get-Content .\New-SubredditHTMLArchive.ps1 | Select -Skip 7 | Set-Content "$($env:temp)\s.ps1"  
3) PS> Get-Help -Name "$($env:temp)\s.ps1" -Full  
  
## Screenshots  
The script will run and make itself into a task called 'RunOnce' in Task Scheduler (taskschd.msc):  
  
![Task Scheduler Screenshot](./screenshots/screenshotTaskScheduler.png "Task Scheduler Screenshot")
  
Then seconds later, that created task will run, and by default will pop up an interactive console:  
  
![Interactive Screenshot](./screenshots/screenshotScript.png "Interactive Screenshot")
  
If run with the -Background switch parameter, you will instead see the path to the transcript log:  
  
![Background Task Screenshot](./screenshots/screenshotBackground.png "Background Task Screenshot")
  
The finished HTML archives and ZIP path are in the task description, and the end of the transcript.  
  
## Added features since initial release  
- Missing dependency package installation with the new -InstallPackages parameter.
    1. Outputs list of installed Python modules, and does prerequisite checks for winget.
- Support for arrays of subreddit names with the new -Subreddits parameter.
    1. Generates master index.html linking to all archived subreddit index files.
    2. Compresses master index and all HTML archive folders into a portable ZIP file.
- Assisted GitHub authentication step, progress bar, and subreddit input validation.
- Better BDFR-HTML module installation with error/standard output logging.
- BDFR clone operations are now retried up to 10 times, with cumulative sleep.
- Added logs folder and cleaned all folder management code.. all logs now retained.
- Added CTRL+C handling: once restarts clone, twice exits script. Added loop hang detection.
- Script spawns itself as a scheduled task, enabling background runs and rerun scheduling.
- Added -Background parameter to set spawned task LogonType to S4U (no stored password).
- PowerShell Gallery: https://www.powershellgallery.com/packages/New-SubredditHTMLArchive/
- Updated hang detection to not trigger if output JSON/media folder is growing by 1GB/4hrs.
- Checks BDFR logs for repeating errors from a submission ID, and excludes those IDs on retry.
- Now passes partial subreddit JSON clones to BDFR-HTML instead of terminating the script.