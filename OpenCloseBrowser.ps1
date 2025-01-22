### The script opens supplied URL(s) in a browser and waits, then closes 
### the browser and repeats the cycle. The script is useful for tasks like
### refreshing data sources in Excel Online file, keep-alive activity, etc.
### One or multiple URLs can be saved to a text file with each URL per line
### and then fetched to the script as '-Url (gc .\URL_file.txt)'

[CmdletBinding()]
param (
    # URL to open in a browser
    [Parameter(Mandatory=$true)]
    [String[]]
    $Url,

    # Open browser period in hours
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int]
    $UpdateInHours = 24,

    # Waiting time in seconds before closing the browser
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int]
    $WaitInSeconds = 60,

    # Browser to open Url with
    [Parameter(Mandatory=$false)]
    [ValidateSet("Chrome","Edge","Firefox")]
    [String]
    $Browser = "Default"
)

function DefaultBrowser {
    # Get the default browser from the registry
    $defaultBrowser = `
        (Get-ItemProperty `
        "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice"`
        ).ProgId
    $fullPath = `
        (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$defaultBrowser\shell\open\command"`
        ).'(default)'.split('"')[1]
    $processName = $fullPath.split('\')[-1].split('.')[0]
    return [System.Tuple]::Create($processName, $fullPath)
}

function GuessPath {
    param (
        [String]$Path
    )

    if (Test-Path $Path) {
        return $Path
    }
    if ($Path -match "(x86)") {
        $otherPath = $Path.replace("Program Files (x86)","Program Files")
    } else {
        $otherPath = $Path.replace("Program Files","Program Files (x86)")
    }
    if (Test-Path $otherPath) {
        return $otherPath
    }
    return "" # path not found
}

$startTime = Get-Date
$updatePeriod = $UpdateInHours * 3600
$nextUpdateCycle = 0

Write-Host "Start time: $startTime"
# Round start time to the nearest minute
if ($startTime.Second -lt 30) {
    $startTime = $startTime.AddSeconds(-$startTime.Second)
} else {
    $startTime = $startTime.AddSeconds(60 - $startTime.Second)
}
Write-Host "Start time (rounded): $startTime"


switch ($Browser) {
    "Chrome" {
        $processName = "chrome"
        $defaultPath = GuessPath `
            "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    }
    "Edge" {
        $processName = "msedge"
        $defaultPath = GuessPath `
            "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    }
    "Firefox" {
        $processName = "firefox"
        $defaultPath = GuessPath `
            "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        if ($defaultPath -eq "") { # probably Firefox was installed from Microsoft Store
            $defaultPath = "firefox"
        }
    }
    Default {
        $result = DefaultBrowser
        $processName = $result.Item1
        $defaultPath = $result.Item2
    }
}

if ($defaultPath -eq "") {
    Write-Host "'$Browser' browser not found!"
    exit
}

# Close all browser's instances
try {
    Get-Process -Name $processName -ErrorAction Stop | `
       Stop-Process -Force | `
       Out-Null
    # Wait for the browser to close completely
    Start-Sleep -Seconds 1
} catch {
    # No active browser's instances were found; continue
}

do {
    Write-Host "Open URL(s) in '$processName' browser"
    foreach ($item in $Url) {
        if ($item.Trim().Length -gt 0) {
            try {
                # Open URL in the browser
                Start-Process -FilePath $defaultPath -ArgumentList $item -ErrorAction Stop
                Start-Sleep -Seconds 1
            } catch {
                Write-Host "Cannot open '$processName' browser by calling $defaultPath"
                exit
            }
        }
    }

    # Wait a bit
    Write-Host "Wait $WaitInSeconds seconds to close browser..."
    Start-Sleep -Seconds $WaitInSeconds

    # Close all browser's instances
    try {
        Write-Host "Close '$processName' browser"
        Get-Process -Name $processName -ErrorAction Stop | `
        Stop-Process -Force | `
        Out-Null
    } catch {
        # No active browser's instances were found; continue
    }

    $nextUpdateCycle += $updatePeriod
    $nextUpdateTime = ($startTime.AddSeconds($nextUpdateCycle)).ToString()
    # Wait until next update cycle
    Write-Host "Wait $UpdateInHours hour(s) for the next update at $nextUpdateTime..."
    Start-Sleep -Seconds $updatePeriod

} while ($true) # loop forever; press Ctrl+C to exit
