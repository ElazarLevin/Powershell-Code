$ScriptBlock = {
    param($Name)
	Write-Host $Name
	Write-Host $pwd
	Set-Location C:\dev\Powershell Example Scripts
    New-Item -Path "C:\dev\Powershell Example Scripts\" -Name "$Name.txt" -ItemType File
}
$MaxThreads = 5
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
$RunspacePool.Open()
$Jobs = @()
1..10 | Foreach-Object {
	$PowerShell = [powershell]::Create()
	$PowerShell.RunspacePool = $RunspacePool
	$PowerShell.AddScript($ScriptBlock).AddArgument($_)
	$Jobs += @{powershell=$PowerShell;job=$PowerShell.BeginInvoke()}
}

while ($Jobs.job.IsCompleted -contains $false) {
	Start-Sleep 1
}

foreach($Job in $Jobs){
	$Job.powershell.EndInvoke($Job.job)
}