Connect-PnPOnline `
    -Url "https://apeda.sharepoint.com/sites/RootSite" `
    -ClientId "e4f4feeb-ec85-4db6-927b-5566740b4dc6" `
    -Interactive

$componentId = [Guid]::Parse("10e817ed-ae7f-4706-a467-6ff3385dc541")

Set-PnPField `
    -List "UserRegisterationDetailsList" `
    -Identity "Status" `
    -Values @{
        ClientSideComponentId = $componentId
    }

Write-Host "Field Customizer associated successfully!" -ForegroundColor Green