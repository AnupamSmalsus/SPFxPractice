Connect-PnPOnline -Url "https://holidaywiser.sharepoint.com/sites/MainRoot/SubsiteA" -ClientId "3dce459d-605f-40f5-8573-4abd196e6220" -Interactive
$list = Get-PnPList -Identity "TopNavigation"

# Self Lookup XML
$fieldXml = @"
<Field
    Type="LookupMulti"
    DisplayName="Parent"
    Name="Parent"
    StaticName="Parent"
    List="{$($list.Id)}"
    ShowField="Title"
    Mult="TRUE" />
"@

# Add Self Lookup Column
Add-PnPFieldFromXml -FieldXml $fieldXml -List "TopNavigation"
