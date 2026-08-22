Connect-PnPOnline `
    -Url "https://apeda.sharepoint.com/sites/RootSite" `
    -ClientId "e4f4feeb-ec85-4db6-927b-5566740b4dc6" `
    -Interactive `
    -ForceAuthentication
Add-PnPField -List "TestList" -DisplayName "Address" -InternalName "Address" -Type Text -AddToDefaultView