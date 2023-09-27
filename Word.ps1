$word = New-Object -ComObject word.application
$word.Visible = $false
$doc = $word.documents.add()
$selection = $word.Selection
$FileName = "Example.docx"
$count = 1;
while(Test-Path -Path "$((Get-Location).ToString())\$FileName" -PathType Leaf){
   $FileName = "Example$count.docx"
   $count++
}
$SaveFile = "$((Get-Location).ToString())\$FileName"
$InputText = Read-Host "Enter some text"
$selection.TypeText("$InputText")
$selection.TypeParagraph()
$doc.SaveAs($SaveFile)
$doc.Close()
$word.Quit()

$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable word