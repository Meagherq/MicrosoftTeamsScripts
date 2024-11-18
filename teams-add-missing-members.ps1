<#
    Description:
      Read an Excel file and add any missing members to the desired team.
      Uncomment the first line (Install-Module MicrosoftTeams) if you don't already have it installed.
      Change the $excelFile path for the Excel file to point to the path on your local machine.
      Change the $teamsGroupId to your desired Teams group id
#>

#Install-Module MicrosoftTeams

$ErrorActionPreference = "Stop"

Connect-MicrosoftTeams

$excelFile = ""
$teamsGroupId = ""

if (!(Test-Path -Path $excelFile))
{
    Write-Error "Excel file does not exist at the specified location!"
}

$excelObject = New-Object -ComObject Excel.Application
$excelWorkbook = $excelObject.Workbooks.Open($excelFile)
$excelWorksheet = $excelWorkbook.Sheets.Item(1)
$excelUsers = $excelWorksheet.UsedRange.Columns[4].Value2
$excelFileVerified = $false
$usersAdded = 0

$teamsUsers = Get-TeamUser -GroupId $teamsGroupId | select -Property User

foreach ($excelUser in $excelUsers)
{
    if ($excelUser -eq 'Email')
    {
        $excelFileVerified = $true
        continue
    }
    if (!($excelFileVerified))
    {
        Write-Error "Excel file format does not appear correct. Ensure email addresses in fourth column with 'Email' column heading"
    }

    $teamsUser = $teamsUsers | where { $_.User -eq $excelUser }
    if ($teamsUser -eq $null)
    {
        try {
            Write-Output "Adding user [$excelUser] to the team in Member role..."
            Add-TeamUser -GroupId $teamsGroupId -User $excelUser -Role Member
            $usersAdded = $usersAdded + 1
        }
        catch {
            Write-Output "An error occurred while attempting to add user [$excelUser] to the team."
            Write-Output $_
        }
    }
}

Write-Output "$usersAdded users were added to the team!"

$excelWorksheet = $null
$excelWorkbook = $null
$excelObject = $null
