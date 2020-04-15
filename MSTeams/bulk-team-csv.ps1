#requires -version 5.1
<#
.SYNOPSIS
  Bulk management of Teams through CSV files.
.DESCRIPTION
  Will take a CSV list of emails and compare against the selected team.
  Afterwards, you can bulk sync, bulk remove or bulk add.
.NOTES
  Version:        1.1
  Author:         Eric Wu
  Creation Date:  2020-04-15
  Purpose/Change: Initial script development

.EXAMPLE
  bulk-team-csv.ps1
#>
param (
    [string]$workemail = $( Read-Host "Please enter your work email address" )
 )

Install-Module -Name MicrosoftTeams -AllowClobber -Scope CurrentUser

Write-Host "Starting..."

Connect-MicrosoftTeams

$teamList = Get-Team -User $workemail

$i = 1;
$selected = 0;
$str = "";

Write-Host "Welcome to Team Bulk Import..."
Write-Host ""

if ($teamList.length -eq 0) {
   Write-Host "You need to be part of a team to use this functionality."
   return;
} elseif ($teamList.length -eq 1) {
   Write-Host "Only one team found, selecting it.";
   $selected = 0;
} else {
    Write-Host "Please choose a team to bulk import:"
    Write-Host ""
    foreach ($team in $teamList) {
        $str = "";
        if ($i -lt 10) {
            $str += " ";
        }

        $str += $i.ToString() + ". " + $team.DisplayName;
        Write-Host $str;
        $i++;
    }
    Write-Host ""

    [uint16]$selected = $( Read-Host "Select (1 - " $teamList.length ") >");
    while ($selected -lt 1 -or $selected -gt $teamList.length ) {
        [uint16]$selected = $( Read-Host "Select (1 - " $teamList.length ") >");
    }
    $selected--;
}

$teamSelected = $teamList[$selected];

Write-Host "Connecting to " $teamSelected.DisplayName "(Group ID:" $teamSelected.GroupId ")"

Write-Host ""
$yn = $( Read-Host "Are you sure you want to bulk import " $teamSelected.DisplayName " (Y/N)?" );
while ($yn -ne "Y" -and $yn -ne "N" ) {
    $yn = $( Read-Host "Are you sure you want to bulk import " $teamSelected.DisplayName " (Y/N)?" );
}

if ($yn -eq "N") {
    return;
}

Write-Host "Starting import " $teamSelected.DisplayName "(" $teamSelected.GroupId ")"
$entries = Import-Csv -Path .\email-list.csv

$i = 1;
$total = $entries.length;

foreach ($entry in $entries) {
    Add-TeamUser -GroupId $teamSelected.GroupId -User $entry.Email
    Write-Host "Importing " $i " of " $total  " - " $entry.Email
    $i++;
}
