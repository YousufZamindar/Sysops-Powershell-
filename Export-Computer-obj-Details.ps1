#function now prompts for the path to a text file containing hostnames,
#retrieves information for each hostname from Active Directory, & displays the information in an Out-GridView,
# and exports the results to a CSV file named "ComputerInformation.csv". Adjust the file path and name as needed.

Function OutputWithHeaders ($datasource, $OGVTitle) {
    [array]$header1info = $datasource | Select-Object -First 1 | ConvertTo-Csv -NoTypeInformation
    [array]$headerComponents = $header1info[0].split(",")
    $headerobject = New-Object PSObject 
    for ($index = 0; $index -lt $headerComponents.count; $index++) {
        $headerobject | Add-Member NoteProperty $headerComponents[$index].Replace('"','') $headerComponents[$index]
    }
    $outputtable = [array]$headerobject + $datasource
 
    if ($OGVTitle -ne "") {
        $outputtable | Out-GridView -Title $OGVTitle
    } else {
        $outputtable | Out-GridView
    }
}
 
$PathToFile = Read-Host "Enter the path to the file containing hostnames"
$Hostnames = Get-Content $PathToFile
 
$Results = foreach ($ComputerName in $Hostnames) {
    $ComputerInformation = Get-ADComputer -Filter {name -like $ComputerName} -Properties samaccountname, enabled, OperatingSystem, description, passwordexpired, passwordlastset, passwordneverexpires, lastlogondate, distinguishedname, created, modified, CanonicalName | Select samaccountname, enabled, OperatingSystem, description, passwordexpired, passwordlastset, passwordneverexpires, lastlogondate, distinguishedname, created, modified, CanonicalName
    $ComputerInformation
}
 
$Results | OutputWithHeaders -OGVTitle "Computer Information"
 
$Results | Export-Csv -Path "ComputerInformation.csv" -NoTypeInformation