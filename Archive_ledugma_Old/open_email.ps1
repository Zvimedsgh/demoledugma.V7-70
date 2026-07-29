param($EntryID)
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $msg = $ns.GetItemFromID($EntryID)
    $msg.Display()
} catch {
    exit 1
}
