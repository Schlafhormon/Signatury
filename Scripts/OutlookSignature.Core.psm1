Set-StrictMode -Version Latest

function ConvertTo-HashtableRecursive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $InputObject
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $dictionary = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $dictionary[$key] = ConvertTo-HashtableRecursive -InputObject $InputObject[$key]
        }

        return $dictionary
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($item in $InputObject) {
            [void]$items.Add((ConvertTo-HashtableRecursive -InputObject $item))
        }

        return ,($items.ToArray())
    }

    if ($InputObject -is [pscustomobject]) {
        $dictionary = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $dictionary[$property.Name] = ConvertTo-HashtableRecursive -InputObject $property.Value
        }

        return $dictionary
    }

    return $InputObject
}

function Resolve-ConfiguredPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $expandedPath = [Environment]::ExpandEnvironmentVariables($Path)

    if ([System.IO.Path]::IsPathRooted($expandedPath)) {
        return [System.IO.Path]::GetFullPath($expandedPath)
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path $ProjectRoot -ChildPath $expandedPath))
}

function Get-StringValue {
    [CmdletBinding()]
    param(
        $Value,
        [string]$Default = ''
    )

    if ($null -eq $Value) {
        return $Default
    }

    return [string]$Value
}

function Get-BoolValue {
    [CmdletBinding()]
    param(
        $Value,
        [bool]$Default = $false
    )

    if ($null -eq $Value) {
        return $Default
    }

    return [System.Convert]::ToBoolean($Value)
}

function ConvertTo-ObjectArray {
    [CmdletBinding()]
    param(
        $Value
    )

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Array]) {
        return ,$Value
    }

    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string]) -and -not ($Value -is [System.Collections.IDictionary])) {
        return ,@($Value)
    }

    return ,$Value
}

function Get-JsonFileContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Konfigurationsdatei '$Path' wurde nicht gefunden."
    }

    $rawContent = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $jsonObject = $rawContent | ConvertFrom-Json
    return ConvertTo-HashtableRecursive -InputObject $jsonObject
}

function Import-OutlookSignatureConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,

        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $config = Get-JsonFileContent -Path $ConfigPath

    $resolved = [ordered]@{
        ProjectRoot                    = [System.IO.Path]::GetFullPath($ProjectRoot)
        ConfigPath                     = [System.IO.Path]::GetFullPath($ConfigPath)
        TemplateRootPath               = Resolve-ConfiguredPath -Path (Get-StringValue $config.TemplateRootPath '.\Templates') -ProjectRoot $ProjectRoot
        CachePath                      = Resolve-ConfiguredPath -Path (Get-StringValue $config.CachePath '.\Temp\Cache') -ProjectRoot $ProjectRoot
        LogPath                        = Resolve-ConfiguredPath -Path (Get-StringValue $config.LogPath '.\Logs') -ProjectRoot $ProjectRoot
        TempPath                       = Resolve-ConfiguredPath -Path (Get-StringValue $config.TempPath '.\Temp\Work') -ProjectRoot $ProjectRoot
        StatePath                      = Resolve-ConfiguredPath -Path (Get-StringValue $config.StatePath '.\Temp\State') -ProjectRoot $ProjectRoot
        SignatureOutputPath            = Resolve-ConfiguredPath -Path (Get-StringValue $config.SignatureOutputPath '%APPDATA%\Microsoft\Signatures') -ProjectRoot $ProjectRoot
        TemplateFileName               = Get-StringValue $config.TemplateFileName 'DK_Standard.docx'
        SignatureName                  = Get-StringValue $config.SignatureName 'DK_Standard'
        SetAsDefaultForNew             = Get-BoolValue $config.SetAsDefaultForNew $true
        SetAsDefaultForReplyForward    = Get-BoolValue $config.SetAsDefaultForReplyForward $true
        CreateTxt                      = Get-BoolValue $config.CreateTxt $true
        CreateRtf                      = Get-BoolValue $config.CreateRtf $true
        CleanupOldSignatureFiles       = Get-BoolValue $config.CleanupOldSignatureFiles $true
        OptionalRemoveEmptyLines       = Get-BoolValue $config.OptionalRemoveEmptyLines $true
        DisableRoamingSignatures       = Get-BoolValue $config.DisableRoamingSignatures $true
        VariableMappings               = if ($config.Contains('VariableMappings')) { $config.VariableMappings } else { [ordered]@{} }
        VariableOverrides              = if ($config.Contains('VariableOverrides')) { $config.VariableOverrides } else { [ordered]@{} }
        OptionalLineCleanupPatterns    = if ($config.Contains('OptionalLineCleanupPatterns')) { ConvertTo-ObjectArray -Value $config.OptionalLineCleanupPatterns } else { @('^\s*$', '^\s*Telefon:\s*$', '^\s*E-Mail:\s*$') }
        Templates                      = if ($config.Contains('Templates')) { ConvertTo-ObjectArray -Value $config.Templates } else { @() }
        RawConfiguration               = $config
    }

    return $resolved
}

function Initialize-OutlookSignatureDirectories {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration
    )

    $paths = @(
        $Configuration.CachePath,
        $Configuration.LogPath,
        $Configuration.TempPath,
        $Configuration.StatePath,
        $Configuration.SignatureOutputPath
    ) | Select-Object -Unique

    foreach ($path in $paths) {
        if (-not (Test-Path -LiteralPath $path)) {
            [void](New-Item -ItemType Directory -Path $path -Force)
        }
    }
}

function New-LogFilePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration
    )

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $fileName = 'OutlookSignature-{0}-{1}.log' -f $env:USERNAME, $timestamp
    return Join-Path -Path $Configuration.LogPath -ChildPath $fileName
}

function Write-SignatureLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = '{0} [{1}] {2}' -f $timestamp, $Level, $Message
    [System.IO.File]::AppendAllText($LogFilePath, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    Write-Host $line
}

function Get-CurrentUserContext {
    [CmdletBinding()]
    param()

    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()

    return [ordered]@{
        UserName           = $env:USERNAME
        UserDomain         = $env:USERDOMAIN
        IdentityName       = $identity.Name
        UserSid            = if ($identity.User) { $identity.User.Value } else { '' }
        ComputerName       = $env:COMPUTERNAME
        UserProfilePath    = $env:USERPROFILE
        LocalAppDataPath   = $env:LOCALAPPDATA
        RoamingAppDataPath = $env:APPDATA
    }
}

function Escape-LdapFilterValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return ($Value -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29' -replace [char]0, '\00')
}

function Get-CurrentUserActiveDirectoryAttributes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$AttributeNames,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    $result = [ordered]@{}
    foreach ($attributeName in $AttributeNames | Select-Object -Unique) {
        $result[$attributeName] = ''
    }

    try {
        $rootDse = [ADSI]'LDAP://RootDSE'
        $defaultNamingContext = [string]$rootDse.defaultNamingContext

        if ([string]::IsNullOrWhiteSpace($defaultNamingContext)) {
            throw 'Der aktuelle Benutzerkontext liefert keinen gültigen LDAP-Namenskontext.'
        }

        $directoryEntry = [ADSI]('LDAP://{0}' -f $defaultNamingContext)
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($directoryEntry)
        $searcher.PageSize = 1
        $searcher.SizeLimit = 1
        $searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
        $searcher.Filter = '(&(objectCategory=person)(objectClass=user)(sAMAccountName={0}))' -f (Escape-LdapFilterValue -Value $env:USERNAME)

        foreach ($attributeName in ($AttributeNames + 'distinguishedName' | Select-Object -Unique)) {
            [void]$searcher.PropertiesToLoad.Add($attributeName)
        }

        $match = $searcher.FindOne()
        if ($null -eq $match) {
            throw "Benutzer '$($env:USERNAME)' wurde in Active Directory nicht gefunden."
        }

        foreach ($attributeName in $AttributeNames | Select-Object -Unique) {
            $propertyName = $attributeName.ToLowerInvariant()
            if ($match.Properties.Contains($propertyName)) {
                $result[$attributeName] = Get-StringValue -Value ($match.Properties[$propertyName] | Select-Object -First 1)
            }
        }

        if ($match.Properties.Contains('distinguishedname')) {
            Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("AD-Benutzerobjekt gefunden: {0}" -f (Get-StringValue -Value ($match.Properties['distinguishedname'] | Select-Object -First 1)))
        }
    }
    catch {
        Write-SignatureLog -LogFilePath $LogFilePath -Level WARN -Message ("Active-Directory-Abfrage nicht erfolgreich. Es werden leere Werte oder konfigurierte Overrides verwendet. Fehler: {0}" -f $_.Exception.Message)
    }

    return $result
}

function Resolve-SignatureVariables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    $attributeNames = @()
    foreach ($mapping in $Configuration.VariableMappings.GetEnumerator()) {
        $attributeNames += [string]$mapping.Value
    }

    $adValues = Get-CurrentUserActiveDirectoryAttributes -AttributeNames $attributeNames -LogFilePath $LogFilePath
    $resolvedVariables = [ordered]@{}

    foreach ($mapping in $Configuration.VariableMappings.GetEnumerator()) {
        $placeholder = [string]$mapping.Key
        $attributeName = [string]$mapping.Value
        $resolvedValue = ''

        if ($adValues.Contains($attributeName)) {
            $resolvedValue = Get-StringValue -Value $adValues[$attributeName]
        }

        if (($resolvedValue -eq '') -and $Configuration.VariableOverrides.Contains($placeholder)) {
            $resolvedValue = Get-StringValue -Value $Configuration.VariableOverrides[$placeholder]
            Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Override für Platzhalter {0} verwendet." -f $placeholder)
        }

        $resolvedVariables[$placeholder] = $resolvedValue
        Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Variable {0} = '{1}' (AD-Attribut: {2})" -f $placeholder, $resolvedValue, $attributeName)
    }

    return $resolvedVariables
}

function Get-SignatureTemplateDefinitions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration
    )

    $definitions = New-Object System.Collections.Generic.List[hashtable]

    if ($Configuration.Templates.Count -gt 0) {
        foreach ($template in $Configuration.Templates) {
            $definition = [ordered]@{
                TemplateFileName            = Get-StringValue -Value $template.TemplateFileName -Default $Configuration.TemplateFileName
                SignatureName               = Get-StringValue -Value $template.SignatureName -Default $Configuration.SignatureName
                SetAsDefaultForNew          = if ($template.Contains('SetAsDefaultForNew')) { Get-BoolValue -Value $template.SetAsDefaultForNew -Default $Configuration.SetAsDefaultForNew } else { $Configuration.SetAsDefaultForNew }
                SetAsDefaultForReplyForward = if ($template.Contains('SetAsDefaultForReplyForward')) { Get-BoolValue -Value $template.SetAsDefaultForReplyForward -Default $Configuration.SetAsDefaultForReplyForward } else { $Configuration.SetAsDefaultForReplyForward }
                CreateTxt                   = if ($template.Contains('CreateTxt')) { Get-BoolValue -Value $template.CreateTxt -Default $Configuration.CreateTxt } else { $Configuration.CreateTxt }
                CreateRtf                   = if ($template.Contains('CreateRtf')) { Get-BoolValue -Value $template.CreateRtf -Default $Configuration.CreateRtf } else { $Configuration.CreateRtf }
                CleanupOldSignatureFiles    = if ($template.Contains('CleanupOldSignatureFiles')) { Get-BoolValue -Value $template.CleanupOldSignatureFiles -Default $Configuration.CleanupOldSignatureFiles } else { $Configuration.CleanupOldSignatureFiles }
                OptionalRemoveEmptyLines    = if ($template.Contains('OptionalRemoveEmptyLines')) { Get-BoolValue -Value $template.OptionalRemoveEmptyLines -Default $Configuration.OptionalRemoveEmptyLines } else { $Configuration.OptionalRemoveEmptyLines }
                OptionalLineCleanupPatterns = if ($template.Contains('OptionalLineCleanupPatterns')) { ConvertTo-ObjectArray -Value $template.OptionalLineCleanupPatterns } else { ConvertTo-ObjectArray -Value $Configuration.OptionalLineCleanupPatterns }
            }

            [void]$definitions.Add($definition)
        }
    }
    else {
        [void]$definitions.Add(([ordered]@{
            TemplateFileName            = $Configuration.TemplateFileName
            SignatureName               = $Configuration.SignatureName
            SetAsDefaultForNew          = $Configuration.SetAsDefaultForNew
            SetAsDefaultForReplyForward = $Configuration.SetAsDefaultForReplyForward
            CreateTxt                   = $Configuration.CreateTxt
            CreateRtf                   = $Configuration.CreateRtf
            CleanupOldSignatureFiles    = $Configuration.CleanupOldSignatureFiles
            OptionalRemoveEmptyLines    = $Configuration.OptionalRemoveEmptyLines
            OptionalLineCleanupPatterns = ConvertTo-ObjectArray -Value $Configuration.OptionalLineCleanupPatterns
        }))
    }

    return ,($definitions.ToArray())
}

function Get-FileSha256Hash {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Datei '$Path' wurde für die Hash-Bildung nicht gefunden."
    }

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Sync-TemplateToCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,

        [switch]$ForceRefresh
    )

    $sourceTemplatePath = Join-Path -Path $Configuration.TemplateRootPath -ChildPath $TemplateDefinition.TemplateFileName
    $cacheTemplateFolder = Join-Path -Path $Configuration.CachePath -ChildPath 'Templates'
    if (-not (Test-Path -LiteralPath $cacheTemplateFolder)) {
        [void](New-Item -ItemType Directory -Path $cacheTemplateFolder -Force)
    }

    $cachedTemplatePath = Join-Path -Path $cacheTemplateFolder -ChildPath $TemplateDefinition.TemplateFileName
    $sourceReachable = Test-Path -LiteralPath $sourceTemplatePath -PathType Leaf

    if ($sourceReachable) {
        $sourceHash = Get-FileSha256Hash -Path $sourceTemplatePath
        $cachedHash = if (Test-Path -LiteralPath $cachedTemplatePath -PathType Leaf) { Get-FileSha256Hash -Path $cachedTemplatePath } else { '' }

        if ($ForceRefresh -or ($sourceHash -ne $cachedHash)) {
            Copy-Item -LiteralPath $sourceTemplatePath -Destination $cachedTemplatePath -Force
            Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Vorlage '{0}' in lokalen Cache gespiegelt." -f $sourceTemplatePath)
        }
        else {
            Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Lokaler Cache für '{0}' ist aktuell." -f $TemplateDefinition.TemplateFileName)
        }

        return [ordered]@{
            SourceTemplatePath = $sourceTemplatePath
            CachedTemplatePath = $cachedTemplatePath
            TemplateHash       = $sourceHash
            UsedCacheOnly      = $false
        }
    }

    if (Test-Path -LiteralPath $cachedTemplatePath -PathType Leaf) {
        $cachedHash = Get-FileSha256Hash -Path $cachedTemplatePath
        Write-SignatureLog -LogFilePath $LogFilePath -Level WARN -Message ("Vorlagenquelle '{0}' ist nicht erreichbar. Es wird die lokale Cache-Kopie verwendet." -f $sourceTemplatePath)

        return [ordered]@{
            SourceTemplatePath = $sourceTemplatePath
            CachedTemplatePath = $cachedTemplatePath
            TemplateHash       = $cachedHash
            UsedCacheOnly      = $true
        }
    }

    throw "Vorlage '$sourceTemplatePath' ist weder an der Quelle noch im lokalen Cache verfügbar."
}

function Get-StringSha256Hash {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
        $hash = $sha256.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash)).Replace('-', '')
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-SignatureStateFilePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition
    )

    $safeName = ($TemplateDefinition.SignatureName -replace '[^a-zA-Z0-9._-]', '_')
    return Join-Path -Path $Configuration.StatePath -ChildPath ('{0}.json' -f $safeName)
}

function Get-SignatureFingerprint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition,

        [Parameter(Mandatory = $true)]
        [hashtable]$VariableValues,

        [Parameter(Mandatory = $true)]
        [string]$TemplateHash
    )

    $payload = [ordered]@{
        TemplateFileName            = $TemplateDefinition.TemplateFileName
        SignatureName               = $TemplateDefinition.SignatureName
        TemplateHash                = $TemplateHash
        VariableValues              = $VariableValues
        CreateTxt                   = $TemplateDefinition.CreateTxt
        CreateRtf                   = $TemplateDefinition.CreateRtf
        OptionalRemoveEmptyLines    = $TemplateDefinition.OptionalRemoveEmptyLines
        OptionalLineCleanupPatterns = @($TemplateDefinition.OptionalLineCleanupPatterns)
    } | ConvertTo-Json -Depth 10 -Compress

    return Get-StringSha256Hash -Value $payload
}

function Get-ExistingSignatureFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition
    )

    $basePath = Join-Path -Path $Configuration.SignatureOutputPath -ChildPath $TemplateDefinition.SignatureName
    $files = [ordered]@{
        Html = '{0}.htm' -f $basePath
        Txt  = '{0}.txt' -f $basePath
        Rtf  = '{0}.rtf' -f $basePath
    }

    return [hashtable]$files
}

function Get-SignatureAssetFolderPaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition
    )

    if (-not (Test-Path -LiteralPath $Configuration.SignatureOutputPath)) {
        return @()
    }

    $candidateNames = @(
        ('{0}_files' -f $TemplateDefinition.SignatureName),
        ('{0}.files' -f $TemplateDefinition.SignatureName),
        ('{0}-Dateien' -f $TemplateDefinition.SignatureName),
        ('{0}_Dateien' -f $TemplateDefinition.SignatureName)
    )

    $folders = Get-ChildItem -LiteralPath $Configuration.SignatureOutputPath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $candidateNames -contains $_.Name } |
        Select-Object -ExpandProperty FullName

    return @($folders)
}

function Get-HtmlReferencedAssetFolders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HtmlPath
    )

    if (-not (Test-Path -LiteralPath $HtmlPath -PathType Leaf)) {
        return @()
    }

    $htmlContent = Get-Content -LiteralPath $HtmlPath -Raw -Encoding Default
    $folderSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $htmlDirectory = Split-Path -Parent $HtmlPath

    foreach ($match in [regex]::Matches($htmlContent, '(?i)(?:src|href)\s*=\s*"([^"]+)"')) {
        $relativePath = ($match.Groups[1].Value -replace '/', '\')
        if ($relativePath -match '^(?<folder>[^\\]+)\\') {
            $folderName = $Matches['folder']
            if ($folderName -notin @('.', '..')) {
                $folderPath = Join-Path -Path $htmlDirectory -ChildPath $folderName
                [void]$folderSet.Add($folderPath)
            }
        }
    }

    return @($folderSet)
}

function Test-SignatureRefreshRequired {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition,

        [Parameter(Mandatory = $true)]
        [string]$Fingerprint
    )

    $statePath = Get-SignatureStateFilePath -Configuration $Configuration -TemplateDefinition $TemplateDefinition
    $signatureFiles = Get-ExistingSignatureFiles -Configuration $Configuration -TemplateDefinition $TemplateDefinition

    $requiredPaths = @($signatureFiles.Html)
    if ($TemplateDefinition.CreateTxt) {
        $requiredPaths += $signatureFiles.Txt
    }
    if ($TemplateDefinition.CreateRtf) {
        $requiredPaths += $signatureFiles.Rtf
    }

    foreach ($path in $requiredPaths) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            return $true
        }
    }

    if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) {
        return $true
    }

    $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
    return ([string]$state.Fingerprint -ne $Fingerprint)
}

function Save-SignatureState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition,

        [Parameter(Mandatory = $true)]
        [string]$Fingerprint,

        [Parameter(Mandatory = $true)]
        [string]$TemplateHash,

        [Parameter(Mandatory = $true)]
        [hashtable]$VariableValues,

        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$GeneratedFiles
    )

    $statePath = Get-SignatureStateFilePath -Configuration $Configuration -TemplateDefinition $TemplateDefinition
    $state = [ordered]@{
        SignatureName  = $TemplateDefinition.SignatureName
        TemplateFile   = $TemplateDefinition.TemplateFileName
        TemplateHash   = $TemplateHash
        Fingerprint    = $Fingerprint
        GeneratedAt    = (Get-Date).ToString('o')
        VariableValues = $VariableValues
        OutputFiles    = $GeneratedFiles
    }

    $state | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Remove-SignatureOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    $files = Get-ExistingSignatureFiles -Configuration $Configuration -TemplateDefinition $TemplateDefinition
    $assetFolders = Get-SignatureAssetFolderPaths -Configuration $Configuration -TemplateDefinition $TemplateDefinition
    foreach ($path in $files.Values + $assetFolders) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Recurse -Force
            Write-SignatureLog -LogFilePath $LogFilePath -Level DEBUG -Message ("Vorhandene Signaturdatei entfernt: {0}" -f $path)
        }
    }
}

function Get-WordConstants {
    [CmdletBinding()]
    param()

    return [ordered]@{
        AlertsNone         = 0
        ReplaceAll         = 2
        FindContinue       = 1
        FormatFilteredHtml = 10
        FormatUnicodeText  = 7
        FormatRtf          = 6
        DoNotSaveChanges   = 0
    }
}

function Invoke-WordFindReplaceAll {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Range,

        [Parameter(Mandatory = $true)]
        [string]$FindText,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$ReplacementText
    )

    $constants = Get-WordConstants
    $find = $Range.Find
    $find.ClearFormatting()
    $find.Replacement.ClearFormatting()
    $find.Text = $FindText
    $find.Replacement.Text = $ReplacementText
    $find.Forward = $true
    $find.Wrap = $constants.FindContinue
    $find.Format = $false
    $find.MatchCase = $false
    $find.MatchWholeWord = $false
    $find.MatchWildcards = $false
    $find.MatchSoundsLike = $false
    $find.MatchAllWordForms = $false

    [void]$find.Execute(
        $find.Text,
        $find.MatchCase,
        $find.MatchWholeWord,
        $find.MatchWildcards,
        $find.MatchSoundsLike,
        $find.MatchAllWordForms,
        $find.Forward,
        $find.Wrap,
        $find.Format,
        $find.Replacement.Text,
        $constants.ReplaceAll
    )
}

function Get-WordStoryRanges {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Document
    )

    $ranges = New-Object System.Collections.Generic.List[object]

    foreach ($storyRange in $Document.StoryRanges) {
        $currentRange = $storyRange
        while ($null -ne $currentRange) {
            [void]$ranges.Add($currentRange)
            try {
                $currentRange = $currentRange.NextStoryRange
            }
            catch {
                $currentRange = $null
            }
        }
    }

    return @($ranges.ToArray())
}

function Normalize-WordParagraphText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text
    )

    $normalized = $Text.Replace([string][char]13, '').Replace([string][char]7, '').Replace([string][char]160, ' ')
    return $normalized.Trim()
}

function Remove-OptionalEmptyParagraphs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Document,

        [Parameter(Mandatory = $true)]
        [string[]]$Patterns,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    foreach ($storyRange in (Get-WordStoryRanges -Document $Document)) {
        for ($index = $storyRange.Paragraphs.Count; $index -ge 1; $index--) {
            $paragraph = $storyRange.Paragraphs.Item($index)
            $paragraphText = Normalize-WordParagraphText -Text ([string]$paragraph.Range.Text)

            foreach ($pattern in $Patterns) {
                if ($paragraphText -match $pattern) {
                    [void]$paragraph.Range.Delete()
                    Write-SignatureLog -LogFilePath $LogFilePath -Level DEBUG -Message ("Leere oder optionale Zeile entfernt: '{0}'" -f $paragraphText)
                    break
                }
            }
        }
    }
}

function Export-WordDocumentFormat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $WordApplication,

        [Parameter(Mandatory = $true)]
        [string]$SourceDocumentPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath,

        [Parameter(Mandatory = $true)]
        [int]$FormatCode
    )

    $document = $null
    try {
        $document = $WordApplication.Documents.Open($SourceDocumentPath, $false, $true)
        [void]$document.SaveAs2($TargetPath, $FormatCode)
    }
    finally {
        if ($null -ne $document) {
            [void]$document.Close((Get-WordConstants).DoNotSaveChanges)
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        }
    }
}

function Release-ComObjectSafely {
    [CmdletBinding()]
    param(
        $ComObject
    )

    if ($null -ne $ComObject) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
        catch {
        }
    }
}

function Invoke-WordSignatureGeneration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition,

        [Parameter(Mandatory = $true)]
        [string]$TemplatePath,

        [Parameter(Mandatory = $true)]
        [hashtable]$VariableValues,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    $constants = Get-WordConstants
    $workingCopyPath = Join-Path -Path $Configuration.TempPath -ChildPath ('{0}-{1}.docx' -f [System.IO.Path]::GetFileNameWithoutExtension($TemplateDefinition.TemplateFileName), ([guid]::NewGuid().ToString('N')))
    Copy-Item -LiteralPath $TemplatePath -Destination $workingCopyPath -Force

    $word = $null
    $document = $null
    $generatedFiles = [hashtable](Get-ExistingSignatureFiles -Configuration $Configuration -TemplateDefinition $TemplateDefinition)

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = $constants.AlertsNone
        $word.ScreenUpdating = $false

        $document = $word.Documents.Open($workingCopyPath, $false, $false)
        foreach ($storyRange in (Get-WordStoryRanges -Document $document)) {
            foreach ($variable in $VariableValues.GetEnumerator()) {
                Invoke-WordFindReplaceAll -Range $storyRange -FindText ([string]$variable.Key) -ReplacementText ([string]$variable.Value)
            }
        }

        foreach ($variable in $VariableValues.GetEnumerator()) {
            Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Platzhalter ersetzt: {0} -> '{1}'" -f $variable.Key, $variable.Value)
        }

        if ($TemplateDefinition.OptionalRemoveEmptyLines) {
            Remove-OptionalEmptyParagraphs -Document $document -Patterns $TemplateDefinition.OptionalLineCleanupPatterns -LogFilePath $LogFilePath
        }

        [void]$document.Save()
        [void]$document.Close($constants.DoNotSaveChanges)
        Release-ComObjectSafely -ComObject $document
        $document = $null

        if ($TemplateDefinition.CleanupOldSignatureFiles) {
            Remove-SignatureOutput -Configuration $Configuration -TemplateDefinition $TemplateDefinition -LogFilePath $LogFilePath
        }

        Export-WordDocumentFormat -WordApplication $word -SourceDocumentPath $workingCopyPath -TargetPath $generatedFiles.Html -FormatCode $constants.FormatFilteredHtml
        $generatedFiles['HtmlAssetFolders'] = @(Get-HtmlReferencedAssetFolders -HtmlPath $generatedFiles.Html)
        Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("HTML-Signatur erzeugt: {0}" -f $generatedFiles.Html)

        if ($TemplateDefinition.CreateTxt) {
            Export-WordDocumentFormat -WordApplication $word -SourceDocumentPath $workingCopyPath -TargetPath $generatedFiles.Txt -FormatCode $constants.FormatUnicodeText
            Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("TXT-Signatur erzeugt: {0}" -f $generatedFiles.Txt)
        }

        if ($TemplateDefinition.CreateRtf) {
            try {
                Export-WordDocumentFormat -WordApplication $word -SourceDocumentPath $workingCopyPath -TargetPath $generatedFiles.Rtf -FormatCode $constants.FormatRtf
                Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("RTF-Signatur erzeugt: {0}" -f $generatedFiles.Rtf)
            }
            catch {
                $generatedFiles.Remove('Rtf')
                Write-SignatureLog -LogFilePath $LogFilePath -Level WARN -Message ("RTF-Export war nicht erfolgreich und wird für diesen Lauf übersprungen. Fehler: {0}" -f $_.Exception.Message)
            }
        }

        return [hashtable]$generatedFiles
    }
    catch {
        throw "Word-Verarbeitung für Vorlage '$($TemplateDefinition.TemplateFileName)' fehlgeschlagen. $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $document) {
            try {
                [void]$document.Close($constants.DoNotSaveChanges)
            }
            catch {
            }
        }

        if ($null -ne $word) {
            try {
                [void]$word.Quit()
            }
            catch {
            }
        }

        Release-ComObjectSafely -ComObject $document
        Release-ComObjectSafely -ComObject $word

        if (Test-Path -LiteralPath $workingCopyPath) {
            Remove-Item -LiteralPath $workingCopyPath -Force
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-InstalledOutlookOfficeVersion {
    [CmdletBinding()]
    param()

    $officeRoot = 'HKCU:\Software\Microsoft\Office'
    if (-not (Test-Path -LiteralPath $officeRoot)) {
        return $null
    }

    $versions = Get-ChildItem -LiteralPath $officeRoot -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' -and (Test-Path -LiteralPath (Join-Path -Path $_.PSPath -ChildPath 'Outlook')) } |
        Sort-Object { [version]$_.PSChildName } -Descending

    if ($versions.Count -eq 0) {
        return $null
    }

    return $versions[0].PSChildName
}

function Disable-OutlookRoamingSignatures {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OfficeVersion,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    $setupPath = 'HKCU:\Software\Microsoft\Office\{0}\Outlook\Setup' -f $OfficeVersion
    if (-not (Test-Path -LiteralPath $setupPath)) {
        [void](New-Item -Path $setupPath -Force)
    }

    New-ItemProperty -LiteralPath $setupPath -Name 'DisableRoamingSignaturesTemporaryToggle' -PropertyType DWord -Value 1 -Force | Out-Null
    New-ItemProperty -LiteralPath $setupPath -Name 'DisableRoamingSignatures' -PropertyType DWord -Value 1 -Force | Out-Null
    Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Roaming Signatures für Outlook {0} deaktiviert bzw. bestätigt." -f $OfficeVersion)
}

function Get-OutlookProfileConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OfficeVersion
    )

    $outlookRoot = 'HKCU:\Software\Microsoft\Office\{0}\Outlook' -f $OfficeVersion
    if (-not (Test-Path -LiteralPath $outlookRoot)) {
        return $null
    }

    $outlookProperties = Get-ItemProperty -LiteralPath $outlookRoot -ErrorAction SilentlyContinue
    $defaultProfile = Get-StringValue -Value $outlookProperties.DefaultProfile
    if ([string]::IsNullOrWhiteSpace($defaultProfile)) {
        return $null
    }

    $profileRoot = Join-Path -Path $outlookRoot -ChildPath ('Profiles\{0}' -f $defaultProfile)
    $accountsRoot = Join-Path -Path $profileRoot -ChildPath '9375CFF0413111d3B88A00104B2A6676'

    return [ordered]@{
        OutlookRoot    = $outlookRoot
        DefaultProfile = $defaultProfile
        ProfileRoot    = $profileRoot
        AccountsRoot   = $accountsRoot
    }
}

function Set-OutlookDefaultSignature {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateDefinition,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    $officeVersion = Get-InstalledOutlookOfficeVersion
    if ([string]::IsNullOrWhiteSpace($officeVersion)) {
        Write-SignatureLog -LogFilePath $LogFilePath -Level WARN -Message 'Es wurde keine unterstützte Outlook-Installation im Benutzerprofil gefunden. Die Signaturdateien wurden erzeugt, Standardwerte aber nicht gesetzt.'
        return
    }

    if ($Configuration.DisableRoamingSignatures) {
        Disable-OutlookRoamingSignatures -OfficeVersion $officeVersion -LogFilePath $LogFilePath
    }

    $mailSettingsPath = 'HKCU:\Software\Microsoft\Office\{0}\Common\MailSettings' -f $officeVersion
    if (-not (Test-Path -LiteralPath $mailSettingsPath)) {
        [void](New-Item -Path $mailSettingsPath -Force)
    }

    if ($TemplateDefinition.SetAsDefaultForNew) {
        New-ItemProperty -LiteralPath $mailSettingsPath -Name 'NewSignature' -PropertyType String -Value $TemplateDefinition.SignatureName -Force | Out-Null
        Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Standard-Signatur für neue Nachrichten in MailSettings gesetzt: {0}" -f $TemplateDefinition.SignatureName)
    }

    if ($TemplateDefinition.SetAsDefaultForReplyForward) {
        New-ItemProperty -LiteralPath $mailSettingsPath -Name 'ReplySignature' -PropertyType String -Value $TemplateDefinition.SignatureName -Force | Out-Null
        Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Standard-Signatur für Antworten/Weiterleitungen in MailSettings gesetzt: {0}" -f $TemplateDefinition.SignatureName)
    }

    $profileConfiguration = Get-OutlookProfileConfiguration -OfficeVersion $officeVersion
    if ($null -eq $profileConfiguration) {
        Write-SignatureLog -LogFilePath $LogFilePath -Level WARN -Message ("Standard-Outlook-Profil konnte für Office {0} nicht ermittelt werden." -f $officeVersion)
        return
    }

    Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Standard-Outlook-Profil erkannt: {0}" -f $profileConfiguration.DefaultProfile)

    if (-not (Test-Path -LiteralPath $profileConfiguration.AccountsRoot)) {
        Write-SignatureLog -LogFilePath $LogFilePath -Level WARN -Message ("Outlook-Kontenspeicher wurde nicht gefunden: {0}" -f $profileConfiguration.AccountsRoot)
        return
    }

    $accountKeys = Get-ChildItem -LiteralPath $profileConfiguration.AccountsRoot -Recurse -ErrorAction SilentlyContinue |
        Where-Object {
            $_.PSChildName -match '^[0-9A-F]{8}$'
        }

    foreach ($accountKey in $accountKeys) {
        $properties = Get-ItemProperty -LiteralPath $accountKey.PSPath -ErrorAction SilentlyContinue
        $propertyNames = @($properties.PSObject.Properties.Name)
        $looksLikeAccount = ($propertyNames -contains 'Account Name') -or
            ($propertyNames -contains 'Display Name') -or
            ($propertyNames -contains 'SMTP Email Address') -or
            ($propertyNames -contains 'New Signature') -or
            ($propertyNames -contains 'Reply-Forward Signature')

        if (-not $looksLikeAccount) {
            continue
        }

        if ($TemplateDefinition.SetAsDefaultForNew) {
            New-ItemProperty -LiteralPath $accountKey.PSPath -Name 'New Signature' -PropertyType String -Value $TemplateDefinition.SignatureName -Force | Out-Null
        }

        if ($TemplateDefinition.SetAsDefaultForReplyForward) {
            New-ItemProperty -LiteralPath $accountKey.PSPath -Name 'Reply-Forward Signature' -PropertyType String -Value $TemplateDefinition.SignatureName -Force | Out-Null
        }

        Write-SignatureLog -LogFilePath $LogFilePath -Level INFO -Message ("Outlook-Profilpfad aktualisiert: {0}" -f $accountKey.Name)
    }
}

function Invoke-OutlookSignatureDeployment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,

        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [switch]$ForceRefresh,

        [switch]$SkipDefaultSignatureUpdate
    )

    $configuration = Import-OutlookSignatureConfiguration -ConfigPath $ConfigPath -ProjectRoot $ProjectRoot
    Initialize-OutlookSignatureDirectories -Configuration $configuration
    $logFilePath = New-LogFilePath -Configuration $configuration

    Write-SignatureLog -LogFilePath $logFilePath -Level INFO -Message 'Outlook-Signaturbereitstellung gestartet.'

    try {
        $userContext = Get-CurrentUserContext
        Write-SignatureLog -LogFilePath $logFilePath -Level INFO -Message ("Benutzer: {0} | Domäne: {1} | SID: {2}" -f $userContext.UserName, $userContext.UserDomain, $userContext.UserSid)
        Write-SignatureLog -LogFilePath $logFilePath -Level INFO -Message ("Konfiguration geladen: {0}" -f $configuration.ConfigPath)

        $variableValues = Resolve-SignatureVariables -Configuration $configuration -LogFilePath $logFilePath
        $templateDefinitions = Get-SignatureTemplateDefinitions -Configuration $configuration

        foreach ($templateDefinition in $templateDefinitions) {
            Write-SignatureLog -LogFilePath $logFilePath -Level INFO -Message ("Verarbeite Vorlage '{0}' als Signatur '{1}'." -f $templateDefinition.TemplateFileName, $templateDefinition.SignatureName)
            $templateInfo = Sync-TemplateToCache -Configuration $configuration -TemplateDefinition $templateDefinition -LogFilePath $logFilePath -ForceRefresh:$ForceRefresh
            $fingerprint = Get-SignatureFingerprint -TemplateDefinition $templateDefinition -VariableValues $variableValues -TemplateHash $templateInfo.TemplateHash

            if ($ForceRefresh -or (Test-SignatureRefreshRequired -Configuration $configuration -TemplateDefinition $templateDefinition -Fingerprint $fingerprint)) {
                Write-SignatureLog -LogFilePath $logFilePath -Level INFO -Message ("Aktualisierung erforderlich für Signatur '{0}'." -f $templateDefinition.SignatureName)
                $generatedFiles = Invoke-WordSignatureGeneration -Configuration $configuration -TemplateDefinition $templateDefinition -TemplatePath $templateInfo.CachedTemplatePath -VariableValues $variableValues -LogFilePath $logFilePath
                Save-SignatureState -Configuration $configuration -TemplateDefinition $templateDefinition -Fingerprint $fingerprint -TemplateHash $templateInfo.TemplateHash -VariableValues $variableValues -GeneratedFiles $generatedFiles
            }
            else {
                Write-SignatureLog -LogFilePath $logFilePath -Level INFO -Message ("Keine Inhaltsänderung für Signatur '{0}' erkannt. Bestehende Ausgabe bleibt erhalten." -f $templateDefinition.SignatureName)
            }

            if (-not $SkipDefaultSignatureUpdate) {
                Set-OutlookDefaultSignature -Configuration $configuration -TemplateDefinition $templateDefinition -LogFilePath $logFilePath
            }
            else {
                Write-SignatureLog -LogFilePath $logFilePath -Level INFO -Message 'Setzen der Outlook-Standardsignatur wurde per Parameter übersprungen.'
            }
        }

        Write-SignatureLog -LogFilePath $logFilePath -Level INFO -Message 'Outlook-Signaturbereitstellung erfolgreich abgeschlossen.'
        return [ordered]@{
            Success     = $true
            LogFilePath = $logFilePath
        }
    }
    catch {
        Write-SignatureLog -LogFilePath $logFilePath -Level ERROR -Message ("Fehler bei der Signaturbereitstellung: {0}`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace)
        return [ordered]@{
            Success     = $false
            LogFilePath = $logFilePath
            Error       = $_.Exception.Message
        }
    }
}

Export-ModuleMember -Function *-*
