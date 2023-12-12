<#
	Script			: Ping Multiple servers provided in Text file Serverlist.txt
	Purpose			: Ping Multiple servers publish the result in excel and highlight the dead ping
	Pre-requisite 	: Create a Text File Serverlist.txt
#>
$ServerListFile = "ServerList.txt"  
$ServerList = Get-Content "C:\ServerList.txt"
$Result = @() 
$user= whoami
$date = (get-Date).tostring()
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.add()
$excel.visible = $true
$s1 = $workbook.sheets | where {$_.name -eq 'Sheet1'}
$s1.Delete()
$s3 = $workbook.sheets | where {$_.name -eq 'Sheet3'}
$s3.Delete()
$s2 = $workbook.sheets | where {$_.name -eq 'Sheet2'}
$s2.name = "Ping Result"
$cells= $s2.Cells
$s2.range("A2:A2").cells="Hostname"
$s2.range("A2:A2").font.bold = "true"
$s2.range("A2:A2").interior.colorindex=48
$s2.range("A2:A2").HorizontalAlignment = -4108
$s2.range("B2:B2").cells="Ping Result"
$s2.range("B2:B2").font.bold = "true"
$s2.range("B2:B2").interior.colorindex=48
$s2.range("B2:B2").HorizontalAlignment = -4108
$s2.range("A3:b3").EntireColumn.autofit() | out-Null
$row=3
$col1=1
$col2=2
$s2.Cells.EntireColumn.AutoFilter()
write-host "Please wait..."
ForEach($computername in $ServerList)		
{
	if (test-Connection -ComputerName $computername -Count 3 -Quiet ) 
	{  
		$cells.item($row,$col1)=$computername
		$cells.item($row,$col2)="Server is alive and Pinging"
	} 
	else 
    { 
		$cells.item($row,$col1)=$computername
		$cells.item($row,$col1).Interior.ColorIndex = 46
		$cells.item($row,$col2)="Server seems dead not pinging"
		$cells.item($row,$col2).Interior.ColorIndex = 46
	}
	$row++
	$col1=1
	$col2=2
	$s2.range("A3:b3").EntireColumn.autofit() | out-Null
}
$row=$row+2
write-host "Script Completed !!!"
"`n"
$workbook.SaveAs("$env:userprofile\desktop\Multiple_Ping_Test.xlsx")