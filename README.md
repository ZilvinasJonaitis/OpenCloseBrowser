# Open and close the URL in browser

The script launches a provided URL in a browser, waits for set duration, and then closes the browser. This action sequence repeats periodically and continues indefinitely until the script execution is manually stopped.

## About

The inspiration to write this script came in need to address the missing functionality of Microsoft Forms data sync to Excel. The problem lies in the limitation: "New responses will automatically sync when you open your Excel file for the web." This requires users to manually open the Excel file on web to retrieve latest form responses, which disrupts workflows, especially when Excel serves as a data source for Microsoft Power BI.

To mitigate this issue, a solution is to automatically open and close Excel file on web periodically using a PowerShell script. This ensures the data remains up-to-date without manual intervention, maintaining a smooth and efficient workflow.

## Features

- Open single or multiple URLs at once 
- Open/Close period can be set in hour step
- Open timeout can be set in seconds
- Option to select different browser
- Script was test on Windows 11 and Windows Server 2019

## Usage

1. Copy script file and URL text file to your PC/Server
2. Open PowerShell terminal and cd to script location
3. Ensure that execution of PowerShell scripts is enabled ('Unrestricted'):
```powershell
Get-ExecutionPolicy
Unrestricted
```
4. In case of restriction, enable execution on your own consideration and risk
```powershell
Set-ExecutionPolicy -ExecutionPolicy Unrestricted
```
5. Run a test:
```powershell
.\OpenCloseBrowser.ps1 -Url "https://google.com"
```
6. 'https://google.com' should open in default browser and close in 60 seconds.
7. Stop the script by pressing Ctrl+C or closing PowerShell terminal window.
8. To open one or multiple URLs at once, you can load them from file. If necessary, adjust waiting timeout.
```powershell
.\OpenCloseBrowser.ps1 -Url (gc .\URL_file.txt) -WaitInSeconds 10
```
9. Try out additional parameters:
```powershell
-PeriodInHours # set restart period in hours (default 24 hours)
-Browser [Chrome|Edge|Firefox] # use another browser (if installed)
```

## License

This project is licensed under the MIT License - see the LICENSE file for details. 
