[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") |
out-null
$Outlook = New-Object -comobject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
$inbox =
  $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
$MyFolder1 =
  $namespace.Folders.Item('elazar.levin@lexisnexis.co.za').Folders.Item('Test')
$rules = $namespace.DefaultStore.GetRules()
$rule = $rules.create("My rule: Test",
	[Microsoft.Office.Interop.Outlook.OlRuleType]::olRuleReceive)
  
$rule_body = $rule.Conditions.Subject
$rule_body.Enabled = $true
$rule_body.Text = @('Test')
$action = $rule.Actions.CopyToFolder
$action.enabled = $true
  [Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction].InvokeMember(
    "Folder",
    [System.Reflection.BindingFlags]::SetProperty,
    $null,
    $action,
    $MyFolder1)
$rules.Save()