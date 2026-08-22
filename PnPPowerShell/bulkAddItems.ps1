Connect-PnPOnline `
    -Url "https://apeda.sharepoint.com/sites/RootSite" `
    -ClientId "e4f4feeb-ec85-4db6-927b-5566740b4dc6" `
    -Interactive `

$items = @(
    @{
        Title = "Test Task 1"
        Name = "Anupam"
        Address = "Test Address 1"
    },
    @{
        Title = "Test Task 2"
        Name= "Anju"
        Address = "Test Address 2"
    },
    @{
        Title = "Test Task 3"
        Name = "Ashwini"
        Address = "Test Address 3"
    },
    @{
        Title = "Test Task 4"
        Name = "Adarsh"
        Address = "Test Address 4"
    }
    
)

foreach($item in $items)
{
    Add-PnPListItem -List "TestList" -Values $item
}

Write-Host "Items Added Successfully"