[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True)]
    [string]$ResourceGroupName,
    [Parameter(Mandatory=$True)]
    [string]$SharePointSite
)

$creds = Get-Credential

$0 = $myInvocation.MyCommand.Definition
$CommandDirectory = [System.IO.Path]::GetDirectoryName($0)
Push-Location $CommandDirectory

Connect-AzureRmAccount -Credential $creds

if ($null -eq (Get-AzureRmResourceGroup -Name $ResourceGroupName -ErrorAction Ignore)) {
    throw "The Resources Group '$ResourceGroupName' doesn't exist."
}

Write-Host "Resource Group Name: $ResourceGroupName" -ForegroundColor Yellow
Write-Host "Deploying ARM Template ... " -ForegroundColor Yellow -NoNewline

Push-Location "..\liquid"
$liquidFile = Get-Content -Path "folder-to-process.liquid" | Out-String

Push-Location "..\arm"
$DeploymentName = "SharePointBatching-" + ((Get-Date).ToUniversalTime()).ToString("MMdd-HHmm")
$ArmParamsObject = @{"LiquidFile" = "$liquidFile";"SharePointSite" = "$SharePointSite";}
$DeploymentOutput = New-AzureRmResourceGroupDeployment -Mode Incremental -Name $DeploymentName -ResourceGroupName $ResourceGroupName -TemplateFile "sharepoint-batch.json" -TemplateParameterObject $ArmParamsObject | ConvertTo-Json -Compress
$ArmOutput = $DeploymentOutput | ConvertFrom-Json # COULD BE USEFUL
Write-Host "Done" -ForegroundColor Yellow

Push-Location "..\sp"
Write-Host "Connect to SharePoint Site: $SharePointSiteUrl" -ForegroundColor Yellow
Connect-PnPOnline -Url $SharePointSite -Credentials $creds
Write-Host "Deploying Site Template ... " -ForegroundColor Yellow -NoNewline
Apply-PnPProvisioningTemplate -Path "template.xml"
Write-Host "Done" -ForegroundColor Yellow

Push-Location $CommandDirectory