$updatesession =  New-Object -ComObject "Microsoft.Update.Session"
$updatesearcher = $updatesession.CreateUpdateSearcher() 
$searchresult = $updatesearcher.Search("IsInstalled=0")
$searchresult.Updates | ogv
#$searchresult.Updates | Export-Csv -LiteralPath "C:\Temp\OutstandingWindowsUpdates.csv" -NoTypeInformation
