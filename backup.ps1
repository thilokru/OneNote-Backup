Set-StrictMode -Version "2.0"
Clear-Host

$PathToOneNote = "C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE" #OneNote executable
$BasePath = "C:\Path\To\Backup\Folder" #alternative: $evn:TEMP (for copying and deleting)


echo "Starting OneNote for API access"

Invoke-Item $PathToOneNote
Start-Sleep -Seconds 5

[void][reflection.assembly]::LoadWithPartialName("Microsoft.Office.Interop.Onenote")
$OneNote = New-Object Microsoft.Office.Interop.Onenote.ApplicationClass

[Xml]$Xml = $Null
$OneNote.GetHierarchy($Null, [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsNotebooks, [ref] $Xml)

$Date = Get-Date -Format "dd.MM.yyyy HH-mm"

echo ("Starting Backup, date: " + $Date)

ForEach($Notebook in ($Xml.Notebooks.Notebook)) {
    echo ("Starting export: " + $Notebook.name)
    $File = $BasePath + "\" + $Date + "\" + $Notebook.name + ".onepkg"
    $OneNote.Publish($Notebook.ID, $File, 1) #1 = .onepkg
    echo "Finished export"
    
    #Move to secure location (like ftp ore ssh server)

    #delete file (uncomment if remote backups are implemented)
    #Remove-Item $File
}