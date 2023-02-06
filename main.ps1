Add-Type -AssemblyName Microsoft.Office.Interop.Outlook

$Outlook = New-Object -ComObject Outlook.Application
$ns = $Outlook.GetNameSpace("MAPI")
$Inbox = $ns.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
$SubFolder = $Inbox.Folders | Where-Object { $_.Name -eq "Folder Name" }
$LatestEmail = $SubFolder.Items | Sort-Object -Property ReceivedTime -Descending | Select-Object -First 1

foreach ($Attachment in $LatestEmail.Attachments)
{
#this saves only xlsx file
    if ($Attachment.FileName -like "*.xlsx")
    {
        #$Attachment.SaveAsFile("\\nac-file\Eligibility\LeasingAssociates\testing")
        $Attachment.SaveAsFile((Join-Path $env:USERPROFILE "Downloads\$($Attachment.FileName)"))
        Write-Host "Attachment saved to folder."
    }
}
