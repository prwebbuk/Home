<# .SYNOPSIS
    Set Outlook Contact Fields 
.DESCRIPTION
    Set outlook contact fields
.NOTES
    Author     : Phil Webb
#>
Clear-Host
$ScriptName = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name
$ScriptPath = Split-Path -Path (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Path
if (-not (Test-Path -Path ($ScriptPath + "\Modules")))
{
    Write-Host -ForegroundColor Yellow -Object ("Unable to find Modules folder.")
    exit
}
Log -Strlog ("Running...`nScriptName:`t'{0}'`nFolder:`t`t'{1}'" -f $ScriptName, $ScriptPath)

# Let's try and open the outlook object
Try
{
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop -ErrorVariable "OutlookError"

    $o = New-Object -ComObject Outlook.Application -ErrorAction Stop -ErrorVariable "ApplicationError"
    $ns = $o.GetNameSpace("MAPI")
}
Catch
{
    Log -Strlog $OutlookError
    Log -Strlog $ApplicationError
    Break
}

$ContactObj = $ns.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderContacts)
$ExistingContacts = $ContactObj.Items
$ExistingContacts = $ExistingContacts | Where-Object { $_.Categories -like "*Party*" }

Log -Strlog ("Number of Party Contacts: {0}" -f $ExistingContacts.count)
foreach ($contact in $ExistingContacts)
{
    Log ("Uppdating {0}" -f $contact.FileAs)
    $contact.User1 = "Party"
    $contact.Save()
}

Log "All done"