#This complete script is used to add monitoring to OneDrive User accounts It accepts user email as input. Please get approval from HR or concerned user and department to run this script.


function Construct-OneDriveString {
  <#
    .Description
    This function takes the email input as string and replaces the special characters with '_' to construct OneDrive URL string
  #>
    param(
        [string]$InputString,
        [string]$Replacement  = "_",
        [string]$SpecialChars = ".@"
    )

    $rePattern = ($SpecialChars.ToCharArray() |ForEach-Object { [regex]::Escape($_) }) -join "|"

   $InputString -replace $rePattern,$Replacement
    
}
function CheckUrl($urlparam) {
  <#
    .Description
    This function checks the One Drive URL after its build by function Construct-OneDriveString and returns the status
  #>
    try {
        Write-Host "verifying the url $urlparam" -ForegroundColor Yellow
        $CheckConnection = Invoke-WebRequest -Uri $urlparam
        if($CheckConnection.StatusCode -eq 200) {
        Write-Host "Connection Verified" -ForegroundColor Green
        $status="Success"
    }
}
catch [System.Net.WebException] {
    $ExceptionMessage = $Error[0].Exception
    if ($ExceptionMessage -match "403") {
    Write-Host "URL exists, but you are not authorized" -ForegroundColor Yellow
    Write-Log "URL $urlparam exists, but you are not authorized" -Severity Warning
    }
    elseif ($ExceptionMessage -match "503"){
    Write-Host "Error: Server Busy" -ForegroundColor Red
    Write-Log "URL $urlparam exists, but server is busy" -Severity Information
    }
    elseif ($ExceptionMessage -match "404"){
    Write-Host "Error: URL doesn't exists" -ForegroundColor Red
    Write-Log "URL $urlparam doesn't exists" -Severity Error
    }
    else{
    Write-Host "Error: There is an unknown error" -ForegroundColor Red
    Write-Log "URL $urlparam unknown error" -Severity Error
    }
    $status="Error Occured"
}
return $status
}
function Add-ServiceAccountToOneDrive($orgname, $onedriveurl,$email){
  <#
    .Description
    This function takes the service email account as input and adds it to targeted one drive user sites as Site Collection admin for monitoring.
  #>
       try{
           # connect to SharePoint admin site using SharePoint admin or glogal admin credentials
           Connect-SPOService -Url "https://$orgname-admin.sharepoint.com"
           #Connect to One drive site using current context           
           Write-Host "Connected to $onedriveurl using admin account" -ForegroundColor Green
           #adding the admin account as site collection admin
           $LegalHoldAccount = Read-Host "Enter the email address which needs to be added as admin to $email OneDrive for monitoring."
           Set-SPOUser -Site $onedriveurl -LoginName $LegalHoldAccount -IsSiteCollectionAdmin $true
           Write-Host "Successfully added Legal Hold Account for monitoring $email" -ForegroundColor Green
           Write-Log "Successfully added Legal Hold Account $LegalHoldAccount for monitoring $email" -Severity Information
       }
       catch{
           Write-Host "Error Occuured..." -f $_.Exception.Message
           Write-Log "Error occurred while adding $LegalHoldAccount to monitor $email" -Severity Error
       }
       
}
function Remove-ServiceAccountFromOneDrive($email){
  <#
    .Description
    This function takes the service email account as input and checks if the service account is present in Site Collection admins list of targeted one drive user if found then removes it from monitoring, if not found as site collection admin 
    then it gives message that service account is not found in targeted users one drive site permissions. 
  #>
    try{
           # connect to SharePoint admin site using SharePoint admin or glogal admin credentials           
           $AdminAccount = Read-Host 'Enter the admin ID that has either SharePoint admin or Global Admin Rights to connect to OneDrive url.'
           Connect-SPOService -Url "https://$orgname-admin.sharepoint.com" 
           #Connect to One drive site using current context           
           Set-SPOUser -Site $onedriveurl -LoginName $AdminAccount -IsSiteCollectionAdmin $true          
           Write-Host "Connected to $onedriveurl using admin account" -BackgroundColor Green           
           #removing the admin account as site collection admin
           $LegalHoldAccount = Read-Host "Enter the  service account that needs to be removed from $email OneDrive monitoring."
           $AdminsList=Get-SPOUser -Site $onedriveurl -Limit ALL | where {$_.IsSiteAdmin -eq $True} | select LoginName
           if($AdminsList -ne $null){
                if($AdminsList.LoginName -contains $LegalHoldAccount){
                    Write-Host "Found the user $LegalHoldAccount" -ForegroundColor Yellow
                    Set-SPOUser -Site $onedriveurl -LoginName $LegalHoldAccount -IsSiteCollectionAdmin $false
                    Set-SPOUser -Site $onedriveurl -LoginName $AdminAccount -IsSiteCollectionAdmin $false
                    Write-Host "Successfully removed Legal Hold Account $LegalHoldAccount from $email OneDrive" -ForegroundColor Green
                    Write-Log "Successfully removed Legal Hold Account $LegalHoldAccount from $email OneDrive" -Severity Information
                }
                else{
                    Write-Host "Could not find $LegalHoldAccount in monitoring list" -ForegroundColor Yellow           
                    Write-Log "Could not find $LegalHoldAccount in monitoring list" -Severity Warning
                }
           }
           else{
            Write-Host "No site collection admins exist" -ForegroundColor Red
            Write-Log "No site collection admins for request $email" -Severity Warning
           }
       }
       catch{
           Write-Host "Error Occuured..." -f $_.Exception.Message
           Write-Log "Error occurred while removing $LegalHoldAccount to monitor $email" -Severity Error
       }
}
function Write-Log {
   <#
    .Description
    This function logs the message to output CSV file which is found in C:\Temp folder. 
  #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Information','Warning','Error')]
        [string]$Severity = 'Information'
    )
 
    [pscustomobject]@{
        Time = (Get-Date -f g)
        Message = $Message
        Severity = $Severity
    } | Export-Csv -Path "C:\Temp\AddMonitoringLogs.csv" -Append -NoTypeInformation
 }

$Email = Read-host "Enter the Email ID of targeted OneDrive user" 
$OutputString=Construct-OneDriveString $Email 
$pattern = '(?<=\@).+?(?=\.)'
$OrgName = [regex]::Matches($Email, $pattern).Value
$OneDriveUrl = "https://$OrgName-my.sharepoint.com/personal/" + $OutputString
$OneDriveCheck = CheckUrl($OneDriveUrl)

if($OneDriveCheck -contains "Success"){
    Write-Host "What operation you want to perform" -ForegroundColor Yellow
    $UserOption = Read-Host "1) Add  account to OneDrive `n2) Remove account from  OneDrive"
        Switch($UserOption){
            1 {Write-Host 'You have selected to add monitoring to OneDrive account' -ForegroundColor Yellow
               Write-Host 'When prompted for login please use the account that has either SharePoint Admin Rights or Global Admin Rights' -ForegroundColor Yellow                
               Start-Sleep 5
               Add-ServiceAccountToOneDrive $OrgName $OneDriveUrl $Email}
            2 {Write-Host 'You have selected to remove monitoring from OneDrive account' -ForegroundColor Yellow
               Write-Host 'When prompted for login please use the account that has either SharePoint Admin Rights or Global Admin Rights' -ForegroundColor Yellow                
               Start-Sleep 5
               Remove-ServiceAccountFromOneDrive $Email}
            default {'Not a valid option.'}
        }
    }
else{
    Write-Host "Error Occured" -ForegroundColor Red
}
#Disconnecting SPO service
try {
Disconnect-SPOService
}
catch {
Write-Host "No Connection Exists" -ForegroundColor Red
}
