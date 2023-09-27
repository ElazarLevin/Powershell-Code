$ExcelObj = New-Object -ComObject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open("$((Get-Location).ToString())\ExampleSheet.xlsx");
$ExcelWorkSheet = $ExcelWorkBook.Worksheets.Item("Sheet1")
Write-Host ([int]$ExcelWorkSheet.range("A1").text + [int]$ExcelWorkSheet.range("A2").text)
$ExcelWorkSheet.Cells.Item(5,5).Value2 = "Test"
$ExcelWorkBook.Save()
$ExcelWorkBook.Close($false)
$ExcelObj.Quit()
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWorkSheet)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWorkBook)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelObj)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()