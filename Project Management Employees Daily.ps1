<#Project Management Employees Daily.ps1#>

###################### initialize automation structure #############################

## Create functions and logging

Write-Host "Creating logging and global variables"

$logFile = $logFile = "C:\mycertificates\PS1\Project Managers Daily\logs\$(Get-Date -Format "MM-dd-yyyy").txt"
function Write-LogAndError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]  [string]$logMessage,
        [Parameter(Mandatory = $true)]  [string]$writeMessage,
        [Parameter(Mandatory = $false)] [string]$errorBody = ""
    )

    # Get the current time and format it
    $currentTime = Get-Date -Format "MM-dd-yyyy hh:mmtt"
    $fullLogMessage = "${currentTime}: ${logMessage}"

    # Log to file
    Add-Content -Path $logFile -Value $fullLogMessage
    if ($errorBody) {
        Add-Content -Path $logFile -Value $errorBody
    }

    # Write to host (visible to user, not pipeline)
    Write-Host $writeMessage
}

$Cred = Import-Clixml "C:\mycertificates\credential.xml"

###################### compile sql and sharepoint data #############################

try {


    Write-LogAndError `
        -logMessage "Compiling daily Project and Phase Manager data!" `
        -writeMessage "Compiling daily Project and Phase Manager data!" `

Import-Module sqlServer

## Change directory to fetch data from tables
Set-Location "SQLSERVER:\SQL\SQL99\DEFAULT\databases\BSTProd\Tables"

$projectQuery = @"

SELECT

       Project.[Code] As projectNumber
      ,Project.[Name] As projectName
      ,[StartDate] As startDate
      ,[FinishDate] As finishDate
      ,realClient.Name As clientName
      ,realOrganization.Name As office
      ,realWorkType.Name As workType
      ,realManager.FullName As managerName 	
      ,realManager.Code As managerENumber

  FROM [BSTProd].[Final].[Project] with (nolock)
  left join final.project_tasklocation As FTL with (nolock) on Project.Id = FTL.ParentId 
  left join final.Country_StateProvince As realLocation with (nolock) on RealLocation.Id = FTL.StateProvince
  left join final.client As realClient with (nolock) on realClient.Id = Project.Client
  left join final.client As realClientOwner with (nolock) on realClientOwner.Id = Project.ClientOwner
  left join final.Employee As realManager with (nolock) on realManager.Id = Project.Manager 
  left join final.Organization as realOrganization with (nolock) on realOrganization.Id = Project.Organization
  left join final.WorkType as realWorkType with (nolock) on realWorkType.Id = Project.WorkType
  left join final.MarketSegment as realMarketSegment with (nolock) on realMarketSegment.Id = Project.MarketSegment 
  
   WHERE Project.HasActiveChargeable = 1

"@

$phaseQuery = @"

SELECT
    phaseTable.Code As phaseCode
  , phaseTable.Name As phaseName
  , phaseTable.StartDate As start
  , phaseTable.FinishDate As finish
  , p.Name As projectName
  , p.Code As projectNumber
  , empTable.FullName As manager
  , empTable.Code As managerCode
  ,CONCAT(disciplineOrg.Code, ' - ', disciplineOrg.Name) AS discipline
  , phaseTable.EffectiveStatus As activeStatus
  , fFeeType.Name As feeType
  , officeOrg.Name As office
  , workTypeTable.Name As WorktypeName
FROM [BSTProd].[Final].[Project_Task] As phaseTable
left join final.Project As p with (nolock) on p.Id = phaseTable.ParentId
left join final.Organization As officeOrg with (nolock) on officeOrg.Id = p.Organization
left join final.Organization As disciplineOrg with (nolock) on disciplineOrg.Id = phaseTable.EnteredOrg
left join final.employee As empTable with (nolock) on empTable.Id = phaseTable.EffectiveManager
left join final.WorkType As workTypeTable with (nolock) on workTypeTable.Id = phaseTable.EffectiveWorkType
left join final.FeeType As fFeeType with (nolock) on fFeeType.Id = phaseTable.EnteredFeeType
WHERE phaseTable.LevelNbr = 2 

"@


$phases = Invoke-Sqlcmd -Query $phaseQuery
$projects  = Invoke-Sqlcmd -Query $projectQuery

$projectNums = $projects.projectNumber | Select-Object -Unique
$activeSqlPhases = $phases | Where-Object { $projectNums -contains $_.projectNumber }
$phases = $activeSqlPhases

$managerList = ($projects | Select-Object -Property managerEnumber -Unique) | Sort-Object managerEnumber
$phaseManagerList = ($phases | Select-Object -Property managerCode -Unique) | Sort-Object managerCode

## Connect to SharePoint online and get data

$shareConnect = @{

  "Url" = "Lochgroup.sharepoint.com/sites/Automate"
  "ClientId" = "f09dce88-8fad-49b1-86ee-eec9bd35d6de"
  "Tenant" = "Lochgroup.onmicrosoft.com"
  "CertificatePath" = "C:\mycertificates\PnP Shell.pfx"

 }

 # compile tables
Connect-PnPOnline @shareConnect
$cloudProjectManagers = Get-PNPListItem "Project Managers"
$cloudEmployee = Get-PNPListItem "Employee"

# index main employee table
$employeeIndex = @{}
$cloudEmployee | ForEach-Object { $employeeIndex[$_.FieldValues.ENumber] = $_.FieldValues.Employee }

# acquire project managers that def work here
$employedManagerList = $managerList | Where-Object { $employeeIndex.ContainsKey($_.managerENumber)  }

    Write-LogAndError `
        -logMessage "Compiled all needed data!" `
        -writeMessage "Compiled all needed data!" `

} catch {

      # log failure
    $errorBodyContent = $_ | Out-String
    Write-LogAndError `
        -logMessage "Failed to compile data." `
        -writeMessage "Failed to compile data." `
        -errorBody $errorBodyContent

    # exit block > exit 1

}


###################### update project manager list create/recycle #############################

$addedArray = @()
# add new managers to the manager list
foreach ( $manager in $employedManagerList ) {

         $managerToCheck = $cloudProjectManagers | Where-Object { $_.FieldValues.eNumber -eq $manager.managerEnumber }
         
         if ( -not $managerToCheck ) { 

         $addManager =  $employeeIndex[$manager.managerENumber]
         Add-PnPListItem -List "Project Managers" -Values @{"Title" = $addManager.LookupValue ; "Employee"= $addManager.Email ; "eNumber"= $manager.managerENumber}
         $addedArray += @{
                             Title    = $addManager.LookupValue
                             Email = $addManager.Email
                             eNumber  = $manager.managerENumber

         }

         }
        
}

$removalArray = @()


# if the index object does not exist in $employedManagerList use the recycle method to remove the from the Project Managers list
forEach ( $indexObject in $cloudProjectManagers )  {
    if ( $employedManagerList.managerENumber -notcontains $indexObject.FieldValues.eNumber ) {
        $indexObject.Title
        $removalArray += @{

                Title = $indexObject.FieldValues.Title
                Email = $indexObject.FieldValues.Employee.Email
                eNumber = $indexObject.FieldValues.eNumber     

        }
        Remove-PnPListItem -List "Project Managers" -Identity $indexObject.Id -Force -Recycle

    }
}


###################### compile phase manager data and patch #############################


## Phase managers

# recollect the project managers and inex them
$cloudProjectManagers = Get-PNPListItem "Project Managers"
$cloudPMIndex = @{}
$cloudProjectManagers | ForEach-Object { $cloudPMIndex[ $_.FieldValues.eNumber ] = $_ }

# collect existing phase managers and cross reference them to make sure they work here but are not on the PM list
$cloudPhaseManagers = Get-PNPListItem "Phase Managers"
# acquire phase managers who def work here
$employedPhaseManagerList = $phaseManagerList | Where-Object { $employeeIndex.ContainsKey($_.managerCode) }
# that are not on the PM list
$notProjectManagersList = $employedPhaseManagerList | Where-Object { -not $cloudPMIndex.ContainsKey($_.managerCode) }



# create phase managers
foreach ( $phaseManager in $notProjectManagersList ) {


      $phaseManagerToCheck = $cloudPhaseManagers | Where-Object { $_.FieldValues.eNumber -eq $phaseManager.managerCode }
    
      if ( -not $phaseManagerToCheck ) {
        $addPhaseManager =  $employeeIndex[$phaseManager.managerCode]
        Add-PnPListItem -List "Phase Managers" -Values @{"Title" = $addPhaseManager.LookupValue ; "Employee"= $addPhaseManager.Email ; "eNumber"= $phaseManager.managerCode }
        $addedArray += @{
                             Title    = $addPhaseManager.LookupValue
                             Email = $addPhaseManager.Email
                             eNumber  = $phaseManager.managerCode

         }
      }
}


###################### remove any phase managers that are no longer phase managers #############################

forEach ( $phaseIndexObject in $cloudPhaseManagers )  {
    if ( $notProjectManagersList.managerCode -notcontains $phaseIndexObject.FieldValues.eNumber ) {
        $removalArray += @{

                Title = $indexObject.FieldValues.Title
                Email = $indexObject.FieldValues.Employee.Email
                eNumber = $indexObject.FieldValues.eNumber     

        };

        Remove-PnPListItem -List "Phase Managers" -Identity $phaseIndexObject.Id -Force -Recycle

    }
}

###################### Email changes #############################

if( ( $addedArray.count -gt 0 ) -or ( $removalArray.count -gt 0 ) ) {
  



$emailBody = @"

Today's project and phase manager report had some additions and or removals.

Total Added: $($addedArray.count) <br>
Total Removed: $($removalArray.count) <br> 


"@


    $mailParams = @{
                            SmtpServer                 = 'smtp.azurecomm.net'
                            Port                       = '587'
                            UseSSL                     = $true
                            Credential                 = $Cred  
                            From                       = 'automate@lochgroup.com'
                            To                         = 'kevin.patenaude@lochgroup.com'
                            Subject                    = "Updates to project and phase managers > $(Get-Date)"
                            Body                       = $emailBody
                            BodyAsHtml                 = $true
                            BCC                        = "kevin.patenaude@lochgroup.com"
                           }
                        
                         ## Send the email   
                         Send-MailMessage @mailParams  

     }



     <#
     For later

     # Create the HTML table from the array of hashtables
$removalHtmlTable = $removalArray | ConvertTo-Html -Fragment

# Optional: Add a title or heading to the table
$removalHtmlTable = "<h2>Removal Report</h2>" + $removalHtmlTable


$emailBody = @"
Today's project and phase manager report had some additions and or removals.

Total Added: $($addedArray.count)
Total Removed: $($removalArray.count)

$removalHtmlTable

"@

# Example of sending the email with the HTML body
# Note: You'll need to configure your own parameters for Send-MailMessage
Send-MailMessage `
    -From "your.email@example.com" `
    -To "recipient@example.com" `
    -Subject "Project Manager Report" `
    -Body $emailBody `
    -BodyAsHtml `
    -SmtpServer "your.smtp.server"



    $removalHtmlTable = $removalArray | Select-Object Title, eNumber, Email | ConvertTo-Html -Fragment


    $css = @"
<style>
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
th { background-color: #f2f2f2; }
</style>
"@

$emailBody = $css + $emailBody



------ prolly won't use below here


$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <style>
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #f2f2f2; }
    </style>
</head>
<body>
"@

$htmlFooter = @"
</body>
</html>
"@

$removalHtmlTable = $removalArray | ConvertTo-Html -Fragment

$emailBody = @"
$htmlHeader

<p>Today's project and phase manager report had some additions and or removals.</p>

<p>Total Added: $($addedArray.count)</p>
<p>Total Removed: $($removalArray.count)</p>

$removalHtmlTable

$htmlFooter
"@
     
     
     #>