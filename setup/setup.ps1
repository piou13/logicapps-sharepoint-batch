[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True)]
    [string]$ResourceGroupName,
    [Parameter(Mandatory=$True)]
    [string]$SharePointSiteUrl
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

Push-Location "..\arm"

$DeploymentName = "SharePointBatching-" + ((Get-Date).ToUniversalTime()).ToString("MMdd-HHmm")
$liquidFile = Get-Content -Path ".\liquid\folder-to-process.liquid" | Out-String
$ArmParamsObject = @{"ResourceGroupName" = $ResourceGroupName;"LiquidFile" = "$liquidFile";}

$DeploymentOutput = New-AzureRmResourceGroupDeployment -Mode Incremental -Name $DeploymentName -ResourceGroupName $ResourceGroupName -TemplateFile ".\arm\sharepoint-batch.json" -TemplateParameterObject $ArmParamsObject | ConvertTo-Json -Compress
$ArmOutput = $DeploymentOutput | ConvertFrom-Json # COULD BE USEFUL

Write-Host "Done" -ForegroundColor Yellow

Write-Host "Connect to SharePoint Site: $SharePointSiteUrl" -ForegroundColor Yellow

Push-Location "..\sp"

if ($null -eq (Connect-PnPOnline -Url $SharePointSiteUrl -Credentials $creds -ErrorAction Ignore)) {
    throw "The connection to SharePoint has failed ('$SharePointSiteUrl'), please review youe configuration."
}

Write-Host "Deploying Site Template ... " -ForegroundColor Yellow -NoNewline

Apply-PnPProvisioningTemplate -Path ".\sp\template.xml"

Write-Host "Done" -ForegroundColor Yellow
