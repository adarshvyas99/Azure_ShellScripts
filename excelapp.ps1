#Creates Excel application
$excelObject = New-Object -ComObject excel.application -ErrorAction Stop
#Makes Excel Visable
$excelObject.Application.Visible = $true
$excelObject.DisplayAlerts = $true
#Creates Excel workBook
$title = "RG_Log_"
$book = $excelObject.Workbooks.Add()
$book.Title = ("$title " + (Get-Date -Format D))
$book.Author = "Kavitha"
#Adds worksheets


#----------------------Sheet1--------------------------
#gets the work sheet and Names it
$sheet1 = $book.Worksheets.Item(1)

$sheet1.name = 'RG_Logs'
$sheet2.name = 'RG_Logs2'
#Select a worksheet
$sheet1.Activate() | Out-Null

#Create a row and set it to Row 1
$row1 = 1
#Create a Column Variable and set it to column 1
$column1 = 1

#Add the word Information and change the Font of the word
$sheet1.Cells.Item($row1,$column1) = "Log Information of Resource Groups"
$sheet1.Cells.Item($row1,$column1).Font.Name = "Copperplate Gothic Bold"
$sheet1.Cells.Item($row1,$column1).Font.Size = 20
$sheet1.Cells.Item($row1,$column1).Font.ColorIndex = 16
$sheet1.Cells.Item($row1,$column1).Interior.ColorIndex = 2
$sheet1.Cells.Item($row1,$column1).HorizontalAlignment = -4108
$sheet1.Cells.Item($row1,$column1).Font.Bold = $true
#Merge the cells
$range = $sheet1.Range("A1:E1").Merge() | Out-Null
#Move to the next row
$row1++
#Create Intial row so you can add borders later
$initalRow = $row1
#create Headers for your sheet
$sheet1.Cells.Item($row1,$column1) = "Subscription"
$sheet1.Cells.Item($row1,$column1).Font.Size = 16
$sheet1.Cells.Item($row1,$column1).Font.ColorIndex = 1
$sheet1.Cells.Item($row1,$column1).Interior.ColorIndex = 48
$sheet1.Cells.Item($row1,$column1).Font.Bold = $true
$column1++
$sheet1.Cells.Item($row1,$column1) = "ResourceGroupName"
$sheet1.Cells.Item($row1,$column1).Font.Size = 16
$sheet1.Cells.Item($row1,$column1).Font.ColorIndex = 1
$sheet1.Cells.Item($row1,$column1).Interior.ColorIndex = 48
$sheet1.Cells.Item($row1,$column1).Font.Bold = $true
$column1++
$sheet1.Cells.Item($row1,$column1) = "Location"
$sheet1.Cells.Item($row1,$column1).Font.Size = 16
$sheet1.Cells.Item($row1,$column1).Font.ColorIndex = 1
$sheet1.Cells.Item($row1,$column1).Interior.ColorIndex = 48
$sheet1.Cells.Item($row1,$column1).Font.Bold = $true
$column1++
$sheet1.Cells.Item($row1,$column1) = "Owner"
$sheet1.Cells.Item($row1,$column1).Font.Size = 16
$sheet1.Cells.Item($row1,$column1).Font.ColorIndex = 1
$sheet1.Cells.Item($row1,$column1).Interior.ColorIndex = 48
$sheet1.Cells.Item($row1,$column1).Font.Bold = $true
$column1++
$sheet1.Cells.Item($row1,$column1) = "CreatedTime"
$sheet1.Cells.Item($row1,$column1).Font.Size = 16
$sheet1.Cells.Item($row1,$column1).Font.ColorIndex = 1
$sheet1.Cells.Item($row1,$column1).Interior.ColorIndex = 48
$sheet1.Cells.Item($row1,$column1).Font.Bold = $true
$column1++
$sheet1.Cells.Item($row1,$column1) = "LastEventTimeStamp"
$sheet1.Cells.Item($row1,$column1).Font.Size = 16
$sheet1.Cells.Item($row1,$column1).Font.ColorIndex = 1
$sheet1.Cells.Item($row1,$column1).Interior.ColorIndex = 48
$sheet1.Cells.Item($row1,$column1).Font.Bold = $true
#Now that the headers are done we go down a row and back to column 1
$row1++
$column1 = 1


$subs = Get-AzSubscription | select Name
foreach($sub in $subs){
    Set-AzContext -Subscription  $sub.Name
    #command you want to use to get infromation
    
    $groups = (Get-AzResourceGroup).ResourceGroupName
    foreach($group in $groups) {
        $Azlogcmd = Get-AzLog -Wa 0  -MaxRecord 1  -ResourceGroupName $group  -StartTime 2022-08-01T04:30:00  -EndTime 2022-09-05T04:30:00 

	$OwnerName= Get-AzADUser | select DisplayName| ForEach-Object {$_.DisplayName} 
    	$OwnerEmail= Get-Azcontext | select Account | ForEach-Object {$_.Account.Id}  
    	$rglocation = (Get-AzResourceGroup $group).Location
   
 	$sheet1.Cells.Item($row1,$column1) = $sub.Name
        $column1++ 
	$sheet1.Cells.Item($row1,$column1) = $Azlogcmd| %{$_.ResourceGroupName}
    	$column1++
    	$sheet1.Cells.Item($row1,$column1) = $rglocation
    	$column1++
    	$sheet1.Cells.Item($row1,$column1) = $OwnerName+"("+$OwnerEmail+")"
    	$column1++
	$sheet1.Cells.Item($row1,$column1) = "Rg creation time"
        $column1++    
    	$sheet1.Cells.Item($row1,$column1) = $Azlogcmd | %{$_.EventTimeStamp.DateTime}
    	$row1++
    	$column1 = 1
    }
}

$row1--
$dataRange = $sheet1.Range(("A{0}" -f $initalRow),("F{0}"  -f $row1))
7..12 | ForEach {
    $dataRange.Borders.Item($_).LineStyle = 1
    $dataRange.Borders.Item($_).Weight = 2
}
#Fits cells to size
$UsedRange = $sheet1.UsedRange
$UsedRange.EntireColumn.autofit() | Out-Null


#--------------------Sheet2--------------------------


$sheet2 = $book.Worksheets.Item(2)
$sheet2.Cells.Item($row1,$column1) = "Log Information 2"
$sheet2.Cells.Item($row1,$column1).Font.Name = "Copperplate Gothic Bold"
$sheet2.Cells.Item($row1,$column1).Font.Size = 20
$sheet2.Cells.Item($row1,$column1).Font.ColorIndex = 16
$sheet2.Cells.Item($row1,$column1).Interior.ColorIndex = 2
$sheet2.Cells.Item($row1,$column1).HorizontalAlignment = -4108
$sheet2.Cells.Item($row1,$column1).Font.Bold = $true
 

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) | Out-Null

[System.GC]::Collect() 
[System.GC]::WaitForPendingFinalizers()