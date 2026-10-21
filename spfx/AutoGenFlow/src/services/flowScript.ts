import { buildCreateFlowPayload, IListFlowTarget } from './siteFlow';

export function isFlowUnauthorized(note: string): boolean {
  return note.indexOf('HTTP 401') >= 0;
}

export function buildFlowPowerShell(input: {
  environmentId: string;
  environmentName: string;
  apiHost: string;
  target: IListFlowTarget;
  connectionName: string;
}): string {
  const connectionName = input.connectionName.trim();
  const payload = buildCreateFlowPayload(input.target, connectionName, connectionName ? 'Started' : 'Stopped');
  if (!connectionName) {
    payload.properties.connectionReferences = {};
  }
  const encoded = utf8Base64(JSON.stringify(payload));
  const environmentId = psQuote(input.environmentId.trim());
  const environmentName = psQuote(input.environmentName.trim() || 'Default');
  const apiHost = psQuote(flowApiHost(input.apiHost));
  return [
    '# Creates the AutoGen Flow cloud flow when the SharePoint app receives HTTP 401.',
    '# Run this in PowerShell. Sign in when asked. The script opens Edit flow in a new tab.',
    '$ErrorActionPreference = \'Stop\'',
    `$EnvironmentId = ${environmentId}`,
    `$EnvironmentName = ${environmentName}`,
    `$ApiHost = ${apiHost}`,
    'if (-not (Get-Module -ListAvailable -Name Microsoft.PowerApps.Administration.PowerShell)) {',
    '  Write-Host \'Installing Microsoft.PowerApps.Administration.PowerShell for the current user.\'',
    '  Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser -Force -AllowClobber',
    '}',
    'Import-Module Microsoft.PowerApps.Administration.PowerShell',
    'Add-PowerAppsAccount',
    '$token = $null',
    'if ($global:currentSession -and $global:currentSession.resourceTokens) {',
    '  foreach ($key in @(\'https://service.flow.microsoft.com/\', \'https://service.flow.microsoft.com\')) {',
    '    if ($global:currentSession.resourceTokens.ContainsKey($key)) {',
    '      $token = $global:currentSession.resourceTokens[$key]',
    '      break',
    '    }',
    '  }',
    '}',
    'if (-not $token) {',
    '  throw \'Sign-in did not return a Power Automate token. Use an account that can create flows in this environment.\'',
    '}',
    `$body = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('${encoded}'))`,
    '$headers = @{ Authorization = "Bearer $token"; \'Content-Type\' = \'application/json\'; Accept = \'application/json\' }',
    '$hosts = @($ApiHost, \'https://api.flow.microsoft.com\') | Select-Object -Unique',
    '$created = $null',
    '$lastError = \'\'',
    'foreach ($hostName in $hosts) {',
    '  $uri = "$hostName/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows?api-version=2016-11-01"',
    '  try {',
    '    $created = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body',
    '    if ($created.name) { break }',
    '  } catch {',
    '    $lastError = $_.Exception.Message',
    '    $status = 0',
    '    if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }',
    '    if ($status -ne 401) { throw }',
    '  }',
    '}',
    'if (-not $created -or -not $created.name) {',
    '  throw "The flow was not created. $lastError"',
    '}',
    '$flowId = $created.name',
    'if ($body -match \'"state":"Started"\') {',
    '  $startUri = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows/$flowId/start?api-version=2016-11-01"',
    '  try { Invoke-RestMethod -Method Post -Uri $startUri -Headers $headers | Out-Null } catch {',
    '    Write-Warning "The flow was created and did not start. Open Edit flow and turn it on. $($_.Exception.Message)"',
    '  }',
    '}',
    '$edit = "https://make.powerautomate.com/environments/$([uri]::EscapeDataString($EnvironmentId))/flows/$([uri]::EscapeDataString($flowId))"',
    'Write-Host "Environment: $EnvironmentName"',
    'Write-Host "Environment id: $EnvironmentId"',
    'Write-Host "Flow id: $flowId"',
    'Write-Host "Edit flow: $edit"',
    'Start-Process $edit'
  ].join('\n');
}

export function flowApiHost(host: string): string {
  const trimmed = host.trim().replace(/\/+$/, '');
  if (/^https:\/\/[a-z0-9.-]+\.flow\.microsoft\.com$/i.test(trimmed)) {
    return trimmed;
  }
  return 'https://api.flow.microsoft.com';
}

function psQuote(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

function utf8Base64(value: string): string {
  const bytes = utf8Bytes(value);
  let binary = '';
  for (let index = 0; index < bytes.length; index += 1) {
    binary += String.fromCharCode(bytes[index]);
  }
  return btoa(binary);
}

function utf8Bytes(value: string): number[] {
  const bytes: number[] = [];
  for (let index = 0; index < value.length; index += 1) {
    let code = value.charCodeAt(index);
    if (code >= 0xD800 && code <= 0xDBFF && index + 1 < value.length) {
      const next = value.charCodeAt(index + 1);
      if (next >= 0xDC00 && next <= 0xDFFF) {
        code = 0x10000 + ((code - 0xD800) << 10) + (next - 0xDC00);
        index += 1;
      }
    }
    if (code < 0x80) {
      bytes.push(code);
    } else if (code < 0x800) {
      bytes.push(0xC0 | (code >> 6), 0x80 | (code & 0x3F));
    } else if (code < 0x10000) {
      bytes.push(0xE0 | (code >> 12), 0x80 | ((code >> 6) & 0x3F), 0x80 | (code & 0x3F));
    } else {
      bytes.push(0xF0 | (code >> 18), 0x80 | ((code >> 12) & 0x3F), 0x80 | ((code >> 6) & 0x3F), 0x80 | (code & 0x3F));
    }
  }
  return bytes;
}
