Connect-PnPOnline -Url "https://holidaywiser.sharepoint.com/sites/MainRoot/SubsiteA" -ClientId "3dce459d-605f-40f5-8573-4abd196e6220" -Interactive
$field = Get-PnPField -List "NewTaskList" -Identity "Status"

$field.ClientSideComponentId = "10e817ed-ae7f-4706-a467-6ff3385dc541"

$field.Update()

Invoke-PnPQuery