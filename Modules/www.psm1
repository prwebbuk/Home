Function Log
{
Param(
    [Parameter(Mandatory=$true,HelpMessage="String to add to the log file")][string]$Strlog
)
    $Time = get-date -Format "yyy-MM-dd hh:mm:ss"

    If ($Strlog -like "Warning*") { $Color = "Yellow" }
    elseif ($Strlog -like "Error*") { $Color = "Red" }
    else { $color = "White" }
    Write-Host "$Time $strLog" -ForegroundColor $Color
}
