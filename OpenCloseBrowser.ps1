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

    # Period to open browser in hours
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int]
    $PeriodInHours = 24,

    # Waiting time in seconds before closing the browser
    [Parameter(Mandatory=$false)]
    [ValidateRange(10, [int]::MaxValue)]
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
    return $null # our guess failed
}

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
        if ($null -eq $defaultPath) { # probably Firefox was installed from Microsoft Store
            $defaultPath = "firefox"
        }
    }
    Default {
        $result = DefaultBrowser
        $processName = $result.Item1
        $defaultPath = $result.Item2
    }
}

if ($null -eq $defaultPath) {
    Write-Host "'$Browser' browser not found!" -ForegroundColor Yellow
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

$pageRenderTimeout = 5 # reasonable time in seconds to fully render page in browser
$periodInSeconds = $PeriodInHours * 3600
$cycleCount = 1

do {
    Write-Host "OPEN-CLOSE CYCLE $cycleCount" -ForegroundColor Green
    $cycleCount++
    
    $startTime = $beginTime = Get-Date
    Write-Host "  Started at $($startTime.ToString())"
    
    $cycleDrift = $startTime.Second
    # Round start time to the nearest minute
    if ($cycleDrift -ge 30) {
        $cycleDrift = $cycleDrift - 60
    }
    $startTime = $startTime.AddSeconds(-$cycleDrift)

    Write-Host "  Open URL(s) in '$processName' browser"
    foreach ($item in $Url) {
        if ($item.Trim().Length -gt 0) {
            try {
                # Open URL in the browser
                Start-Process -FilePath $defaultPath -ArgumentList $item -ErrorAction Stop
                Start-Sleep -Seconds $pageRenderTimeout
            } catch {
                Write-Host "Cannot open '$processName' browser by calling $defaultPath" `
                   -ForegroundColor Yellow
                exit
            }
        }
    }

    # Wait a bit
    Write-Host "  Wait $WaitInSeconds seconds"
    Start-Sleep -Seconds $WaitInSeconds

    # Close all browser's instances
    try {
        Write-Host "  Close browser"
        Get-Process -Name $processName -ErrorAction Stop | `
        Stop-Process -Force | `
        Out-Null
    } catch {
        # No active browser's instances were found; continue
    }

    $nextCycleTime = ($startTime.AddSeconds($periodInSeconds)).ToString()
    $processingTime = [math]::round((Get-Date).Subtract($beginTime).TotalSeconds)
    $nextCycleTimeout = $periodInSeconds - $processingTime - $cycleDrift

    Write-Host "  Wait $PeriodInHours hour(s) for the next cycle at $nextCycleTime"
    Start-Sleep -Seconds $nextCycleTimeout

} while ($true) # loop forever; press Ctrl+C to exit
