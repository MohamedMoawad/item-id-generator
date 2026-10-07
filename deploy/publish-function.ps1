# Publishes item-id-generator-func.zip to a new Node 22 Function App.
# Run in PowerShell after: az login
# The zip sits next to this script.

param(
    [Parameter(Mandatory = $true)][string]$ResourceGroup,
    [Parameter(Mandatory = $true)][string]$FunctionAppName,
    [Parameter(Mandatory = $true)][string]$StorageAccount,
    [string]$Location = 'eastus',
    [Parameter(Mandatory = $true)][string]$SharePointSiteUrl,
    [Parameter(Mandatory = $true)][string]$TenantId,
    [Parameter(Mandatory = $true)][string]$ClientId,
    [Parameter(Mandatory = $true)][string]$ClientSecret,
    [Parameter(Mandatory = $true)][string]$WebhookClientState,
    [Parameter(Mandatory = $true)][string]$SharePointOrigin
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw 'Install the Azure CLI, then run az login. https://learn.microsoft.com/cli/azure/install-azure-cli'
}

az account show | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw 'Run az login before this script.'
}

$zip = Join-Path $PSScriptRoot 'item-id-generator-func.zip'
if (-not (Test-Path $zip)) {
    throw "Missing $zip"
}

Write-Host "Creating resource group $ResourceGroup in $Location"
az group create --name $ResourceGroup --location $Location --output none
if ($LASTEXITCODE -ne 0) { throw 'az group create failed.' }

Write-Host "Creating storage account $StorageAccount"
az storage account create --name $StorageAccount --resource-group $ResourceGroup --location $Location --sku Standard_LRS --kind StorageV2 --output none
if ($LASTEXITCODE -ne 0) { throw 'az storage account create failed. Storage names must be 3-24 lowercase letters and numbers, globally unique.' }

Write-Host "Creating Function App $FunctionAppName"
az functionapp create `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --storage-account $StorageAccount `
    --consumption-plan-location $Location `
    --runtime node `
    --runtime-version 22 `
    --functions-version 4 `
    --output none
if ($LASTEXITCODE -ne 0) { throw 'az functionapp create failed.' }

$settingsFile = Join-Path ([System.IO.Path]::GetTempPath()) ("item-id-settings-" + [guid]::NewGuid().ToString() + '.json')
$settings = @(
    @{ name = 'FUNCTIONS_WORKER_RUNTIME'; value = 'node' },
    @{ name = 'WEBSITE_NODE_DEFAULT_VERSION'; value = '~22' },
    @{ name = 'NUMBERING_CONFIG_SITE_URL'; value = $SharePointSiteUrl.TrimEnd('/') },
    @{ name = 'NUMBERING_CONFIG_LIST_TITLE'; value = 'RequestNumberConfig' },
    @{ name = 'SHAREPOINT_TENANT_ID'; value = $TenantId },
    @{ name = 'SHAREPOINT_CLIENT_ID'; value = $ClientId },
    @{ name = 'SHAREPOINT_CLIENT_SECRET'; value = $ClientSecret },
    @{ name = 'SHAREPOINT_WEBHOOK_CLIENT_STATE'; value = $WebhookClientState }
)
$settings | ConvertTo-Json | Set-Content -Path $settingsFile -Encoding utf8
try {
    az functionapp config appsettings set --resource-group $ResourceGroup --name $FunctionAppName --settings "@$settingsFile" --output none
    if ($LASTEXITCODE -ne 0) { throw 'Setting app settings failed.' }
}
finally {
    Remove-Item $settingsFile -Force -ErrorAction SilentlyContinue
}

Write-Host 'Deploying the function zip. node_modules is already in the zip.'
az functionapp deployment source config-zip --resource-group $ResourceGroup --name $FunctionAppName --src $zip --output none
if ($LASTEXITCODE -ne 0) { throw 'Zip deploy failed.' }

function Get-FunctionKey([string]$FunctionName) {
    for ($attempt = 1; $attempt -le 8; $attempt++) {
        $key = az functionapp function keys list --resource-group $ResourceGroup --name $FunctionAppName --function-name $FunctionName --query 'default' --output tsv 2>$null
        if ($LASTEXITCODE -eq 0 -and $key) {
            return $key.Trim()
        }
        Start-Sleep -Seconds 15
    }
    throw "Could not read the $FunctionName function key. Open the function in the Azure portal and copy its key."
}

$webhookKey = Get-FunctionKey 'spoWebhook'
$registerKey = Get-FunctionKey 'RegisterWebhook'
$webhookUrl = "https://$FunctionAppName.azurewebsites.net/api/spoWebhook?code=$webhookKey"
$registerUrl = "https://$FunctionAppName.azurewebsites.net/api/RegisterWebhook?code=$registerKey"

az functionapp config appsettings set --resource-group $ResourceGroup --name $FunctionAppName --settings "SPO_WEBHOOK_NOTIFICATION_URL=$webhookUrl" --output none
if ($LASTEXITCODE -ne 0) { throw 'Saving SPO_WEBHOOK_NOTIFICATION_URL failed.' }

az functionapp cors add --resource-group $ResourceGroup --name $FunctionAppName --allowed-origins $SharePointOrigin --output none
if ($LASTEXITCODE -ne 0) { throw 'Adding CORS failed.' }

Write-Host ''
Write-Host 'Function app is deployed.'
Write-Host "spoWebhook (SharePoint calls this when a list item is added):"
Write-Host "  $webhookUrl"
Write-Host "RegisterWebhook (paste this into the web part property pane):"
Write-Host "  $registerUrl"
Write-Host ''
Write-Host 'Next: upload deploy/item-id-generator.sppkg to the SharePoint app catalog, add Request number config to the site collection, and save an active row. That registers the list webhook.'
