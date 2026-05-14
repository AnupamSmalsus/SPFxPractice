Connect-PnPOnline -Url "https://holidaywiser.sharepoint.com/sites/MainRoot/SubsiteA" -ClientId "3dce459d-605f-40f5-8573-4abd196e6220" -Interactive

$items = @(
    @{
        Title = "Test Task 2"
        Status = "Approved"
        Agenda = "Discussion 1"
        SubmissionDate = [DateTime]::Now
    },
    @{
        Title = "Test Task 3"
        Status = "Rejected"
        Agenda = "Discussion 2"
        SubmissionDate = [DateTime]::Now
    },
    @{
        Title = "Test Task 4"
        Status = "Approved"
        Agenda = "Discussion 3"
        SubmissionDate = [DateTime]::Now
    },
    @{
        Title = "Test Task 5"
        Status = "Rejected"
        Agenda = "Discussion 2"
        SubmissionDate = [DateTime]::Now
    }
    
)

foreach($item in $items)
{
    Add-PnPListItem -List "NewTaskList" -Values $item
}

Write-Host "Items Added Successfully"