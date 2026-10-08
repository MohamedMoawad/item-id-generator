# Creates AutoGenFeatureConfiguration on the connected site.
# The configuration web part can do the same thing. This script is the PnP alternative.
# Connect first. Do not put secrets in this file.
#   Connect-PnPOnline -Url 'https://<tenant>.sharepoint.com/sites/<site>' -Interactive

$ListTitle = 'AutoGenFeatureConfiguration'
New-PnPList -Title $ListTitle -Template GenericList -ErrorAction SilentlyContinue | Out-Null

function Add-ConfigField {
    param([string] $FieldXml)
    Add-PnPFieldFromXml -List $ListTitle -FieldXml $FieldXml -ErrorAction SilentlyContinue | Out-Null
}

Add-ConfigField '<Field Type="Text" Name="TargetListUrl" StaticName="TargetListUrl" DisplayName="Target list URL" Required="TRUE" />'
Add-ConfigField '<Field Type="Text" Name="TargetListGuid" StaticName="TargetListGuid" DisplayName="Target list GUID" Required="TRUE" />'
Add-ConfigField '<Field Type="Text" Name="NumberColumnInternalName" StaticName="NumberColumnInternalName" DisplayName="Number column internal name" Required="TRUE" />'
Add-ConfigField '<Field Type="Text" Name="Formula" StaticName="Formula" DisplayName="Formula" Required="TRUE" />'
Add-ConfigField '<Field Type="Number" Name="CurrentCount" StaticName="CurrentCount" DisplayName="Current count" Decimals="0" Min="0" />'
Add-ConfigField '<Field Type="Choice" Name="ResetPeriod" StaticName="ResetPeriod" DisplayName="Reset period" Format="Dropdown"><CHOICES><CHOICE>None</CHOICE><CHOICE>Day</CHOICE><CHOICE>Month</CHOICE><CHOICE>Year</CHOICE></CHOICES><Default>Month</Default></Field>'
Add-ConfigField '<Field Type="DateTime" Name="LastResetDate" StaticName="LastResetDate" DisplayName="Last reset date" Format="DateTime" />'
Add-ConfigField '<Field Type="Boolean" Name="IsActive" StaticName="IsActive" DisplayName="Is active"><Default>1</Default></Field>'
Add-ConfigField '<Field Type="Number" Name="PadLength" StaticName="PadLength" DisplayName="Pad length" Decimals="0" Min="0" Max="12"><Default>4</Default></Field>'
Add-ConfigField '<Field Type="Text" Name="WebhookSubscriptionId" StaticName="WebhookSubscriptionId" DisplayName="Webhook subscription ID" />'

Write-Host 'AutoGenFeatureConfiguration columns are in place. Add a single-line text column such as RequestNumber on each target list, then add one config row per list.'
