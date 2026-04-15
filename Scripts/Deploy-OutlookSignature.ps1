[CmdletBinding()]
param(
    [string]$ConfigPath,
    [string]$ProjectRoot,
    [switch]$ForceRefresh,
    [switch]$SkipDefaultSignatureUpdate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Parent $scriptDirectory
}
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path -Path $ProjectRoot -ChildPath 'Config\signature-config.json'
}

$modulePath = Join-Path -Path $scriptDirectory -ChildPath 'OutlookSignature.Core.psm1'
Import-Module -Name $modulePath -Force

$result = Invoke-OutlookSignatureDeployment -ConfigPath $ConfigPath -ProjectRoot $ProjectRoot -ForceRefresh:$ForceRefresh -SkipDefaultSignatureUpdate:$SkipDefaultSignatureUpdate

if (-not $result.Success) {
    throw "Die Outlook-Signaturbereitstellung ist fehlgeschlagen. Details siehe Log: $($result.LogFilePath)"
}

Write-Host ("Logdatei: {0}" -f $result.LogFilePath)
