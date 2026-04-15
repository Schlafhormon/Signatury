[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$ForceRefresh,
    [switch]$SkipDefaultSignatureUpdate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDirectory
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path -Path $projectRoot -ChildPath 'Config\signature-config.json'
}

$deployScript = Join-Path -Path $scriptDirectory -ChildPath 'Deploy-OutlookSignature.ps1'
& $deployScript -ConfigPath $ConfigPath -ProjectRoot $projectRoot -ForceRefresh:$ForceRefresh -SkipDefaultSignatureUpdate:$SkipDefaultSignatureUpdate
