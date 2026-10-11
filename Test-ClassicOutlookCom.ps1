# READ-ONLY prerequisite check. No email content, addresses or subjects printed.
$ErrorActionPreference = 'Stop'
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    $session = $app.GetNamespace('MAPI')
    $inbox = $session.GetDefaultFolder(6)
    if ($null -eq $inbox) { throw 'Inbox missing.' }
    $count = [int]$inbox.Items.Count
    Write-Host 'CLASSIC_OUTLOOK_COM=AVAILABLE'
    Write-Host ('INBOX_OBJECTS=' + $count)
    Write-Host 'No messages were modified.'
} catch {
    Write-Host 'CLASSIC_OUTLOOK_COM=UNAVAILABLE'
    Write-Host 'Use classic Outlook for Windows and start it first; new Outlook and Outlook Web do not expose this interface.'
    exit 2
}
