Connect-PnPOnline -Url "https://holidaywiser.sharepoint.com/sites/MainRoot/SubsiteA" -ClientId "3dce459d-605f-40f5-8573-4abd196e6220" -Interactive

Add-PnPField -List "UserRegistrationDetails" -DisplayName "Joining Date" -InternalName "JoiningDate" -Type DateTime
Add-PnPField -List "UserRegistrationDetails" -DisplayName "Leaving Date" -InternalName "LeavingDate" -Type DateTime