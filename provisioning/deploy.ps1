
$configFile = "config.json";

if((Test-Path $configFile) -eq $false) {
    $siteUrl = Read-Host -Prompt "Enter the site url"
    $username = Read-Host -Prompt "Enter the username"
    $securePassword = Read-Host -Prompt "Enter your tenant password" -AsSecureString | ConvertFrom-SecureString
    @{username=$username;securePassword=$securePassword;siteUrl=$siteUrl} | ConvertTo-Json | Out-File $configFile
}

$configObject = Get-Content $configFile | ConvertFrom-Json
$password = $configObject.securePassword | ConvertTo-SecureString
$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $configObject.username, $password
Connect-PnPOnline -url $configObject.siteUrl -Credentials $credentials

$web = Get-PnPWeb
$teamName = $web.Title
Write-Host $teamName

function ProvisionResources() {
    Write-Host ""
    Write-Host "Provisioning Site Columns, Content Types, & Lists" -ForegroundColor Yellow
    Write-Host "-------------------------------------------------" -ForegroundColor Yellow
    Write-Host "Content Type" -ForegroundColor Green
    Apply-PnPProvisioningTemplate ".\definition.xml"
}

#MS Graph Operations
Register-PnPManagementShellAccess -SiteUrl $configObject.siteUrl
Connect-PnPOnline -Scopes "Group.ReadWrite.All" -Credentials $credentials

#Create document library structure
if($false) {
Add-PnPTeamsChannel -Team $teamName -DisplayName "01. PD-Pre-Design"
Add-PnPTeamsChannel -Team $teamName -DisplayName "02. SD-Schematic Design"
Add-PnPTeamsChannel -Team $teamName -DisplayName "03. DD-Design Development"
Add-PnPTeamsChannel -Team $teamName -DisplayName "04. CD Construction Docs"
Add-PnPTeamsChannel -Team $teamName -DisplayName "05. BID-Bid"
Add-PnPTeamsChannel -Team $teamName -DisplayName "06. CA-Contract Administration"
Add-PnPTeamsChannel -Team $teamName -DisplayName "08. PO-Post Occupancy"
Add-PnPTeamsChannel -Team $teamName -DisplayName "09. Compliance"
Add-PnPTeamsChannel -Team $teamName -DisplayName "0S. Supplemental Services"
Add-PnPTeamsChannel -Team $teamName -DisplayName "None"
}

#SharePoint Operations
Connect-PnPOnline -url $configObject.siteUrl -Credentials $credentials

ProvisionResources

$viewFields = "Type","Name","Modified","Modified By","By Whom","Agencies","Document Status","Phase","Submittal Action","Trades","Action Date"

Add-PnPContentTypeToList -List "Documents" -ContentType "GA Document" -DefaultContentType
Set-PnPView -List "Documents" -Identity "All Documents" -Fields $viewFields

