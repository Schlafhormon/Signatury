[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$modulePath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'Scripts\OutlookSignature.Core.psm1'
Import-Module -Name $modulePath -Force

$passed = 0
$failed = 0

function Assert-Equal {
    param(
        [string]$Name,
        $Expected,
        $Actual
    )

    if ($Expected -ceq $Actual) {
        Write-Host ("PASS  {0}" -f $Name)
        $script:passed++
        return
    }

    Write-Host ("FAIL  {0}`n  Expected: {1}`n  Actual:   {2}" -f $Name, $Expected, $Actual)
    $script:failed++
}

$projectRoot = Split-Path -Parent $PSScriptRoot

Assert-Equal -Name 'Resolve-ConfiguredPath relative' -Expected ([System.IO.Path]::GetFullPath((Join-Path $projectRoot '.\Templates'))) -Actual (Resolve-ConfiguredPath -Path '.\Templates' -ProjectRoot $projectRoot)
Assert-Equal -Name 'Normalize-WordParagraphText trims controls' -Expected 'Telefon:' -Actual (Normalize-WordParagraphText -Text ("Telefon:{0}{1}" -f [char]13, [char]7))
Assert-Equal -Name 'Get-StringSha256Hash deterministic' -Expected (Get-StringSha256Hash -Value 'abc') -Actual (Get-StringSha256Hash -Value 'abc')

$configuration = Import-OutlookSignatureConfiguration -ConfigPath (Join-Path $projectRoot 'Config\signature-config.json') -ProjectRoot $projectRoot
$definitions = Get-SignatureTemplateDefinitions -Configuration $configuration
Assert-Equal -Name 'One template definition present' -Expected 1 -Actual $definitions.Count
Assert-Equal -Name 'Template name from config' -Expected 'DK_Standard.docx' -Actual $definitions[0].TemplateFileName

Write-Host ''
Write-Host ("Tests bestanden: {0}" -f $passed)
Write-Host ("Tests fehlgeschlagen: {0}" -f $failed)

if ($failed -gt 0) {
    exit 1
}
