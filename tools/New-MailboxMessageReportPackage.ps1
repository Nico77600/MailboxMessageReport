<#
.SYNOPSIS
    Copies the files needed to run Mailbox Message Report into a separate folder, ready to be zipped.

.DESCRIPTION
    The package contains only what Invoke-MailboxMessageReport.ps1 needs at run time, plus the HTML guides:
        Invoke-MailboxMessageReport.ps1, MailboxMessageReport.psd1, MailboxMessageReport.psm1, src\, config\, templates\,
        docs\MailboxMessageReport-UserGuide.html, docs\MailboxMessageReport-Guide.html, README.md, CHANGELOG.md, LICENSE,
        THIRD-PARTY-NOTICES.md
    The HTML guides (user guide, developer guide) are copied as committed; they are rebuilt from their Markdown
    source only by tools\Build-Documentation.ps1, never by this script. They are self-contained (images inline), so the Markdown sources and the
    images are not copied.
    It never copies reports\, logs\, artifacts\, tests\ (with the simulated tenant) or tools\.

    The configuration is copied as delivered (empty tenant and application). The script checks the content
    of the package and that the module loads from it.

.PARAMETER Destination
    Package folder, named MailboxMessageReport-<version> for a release (it is the single root folder of the ZIP).
    Default: package\MailboxMessageReport-<version>, next to the repository folder.

.PARAMETER Unsigned
    Continuous integration or a developer build: the PowerShell files of package\ are not required to carry the owner
    signature. The committed package\PackageIntegrity.psd1, when it exists, is copied and verified with
    -DevelopmentMode. The result is never a release.

.PARAMETER Manifest
    Phase 2 of a release. Hashes the built folder (made from signed sources) and writes an unsigned PackageIntegrity.psd1
    in it. The owner signs it, then copies it to package\PackageIntegrity.psd1.

.PARAMETER Seal
    Phase 3 of a release. Verifies the signed manifest against the built folder and against package\, compares every
    byte of the files they share, then writes the release ZIP (one root folder) next to the built folder and prints
    its SHA-256.

.PARAMETER Force
    Replace the destination folder if it already contains a package. A folder that contains reports\ or
    logs\ (a package that has been run) is never replaced.

.EXAMPLE
    .\tools\New-MailboxMessageReportPackage.ps1
    Creates ..\package\MailboxMessageReport-<version>.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
#>
#Requires -Version 7.4
[CmdletBinding(DefaultParameterSetName = 'Build')]
param(
    [Parameter(ParameterSetName = 'Build')][string]$Destination,
    [Parameter(ParameterSetName = 'Build')][switch]$Force,
    [Parameter(ParameterSetName = 'Build')][switch]$Unsigned,
    [Parameter(ParameterSetName = 'Manifest', Mandatory)][string]$Manifest,
    [Parameter(ParameterSetName = 'Seal', Mandatory)][string]$Seal
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$root = Join-Path $repoRoot 'package'
$version = (Import-PowerShellDataFile -LiteralPath (Join-Path $root 'MailboxMessageReport.psd1')).ModuleVersion
$signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
$signedScripts = @('Invoke-MailboxMessageReport.ps1', 'MailboxMessageReport.psm1', 'MailboxMessageReport.psd1', 'src\Test-PackageIntegrity.ps1')
$signedScripts += @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.ps1' -File | Where-Object Name -ne 'Test-PackageIntegrity.ps1' |
    Sort-Object Name | ForEach-Object { "src\$($_.Name)" })

function Get-RelativeFile([string]$Base) {
    Get-ChildItem -LiteralPath $Base -Recurse -File -Force | ForEach-Object { $_.FullName.Substring($Base.Length + 1) } | Sort-Object { $_ } -CaseSensitive
}
function Assert-OwnerSigned([string]$Path) {
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne $signer -or $null -eq $signature.TimeStamperCertificate) {
        throw "Not signed with the trusted, timestamped owner certificate ($($signature.Status)): $Path"
    }
}
function Write-Manifest([string]$Package) {
    $entries = foreach ($relative in Get-RelativeFile $Package) {
        if ($relative -ceq 'PackageIntegrity.psd1' -or $relative -ceq 'config\MailboxMessageReport.config.psd1') { continue }
        [pscustomobject]@{ Path = $relative; Sha256 = (Get-FileHash -LiteralPath (Join-Path $Package $relative) -Algorithm SHA256).Hash }
    }
    $text = [Text.StringBuilder]::new()
    [void]$text.AppendLine('@{')
    [void]$text.AppendLine('    SchemaVersion = 1')
    [void]$text.AppendLine("    Algorithm = 'SHA256'")
    [void]$text.AppendLine("    SignerThumbprint = '$signer'")
    [void]$text.AppendLine("    Version = '$version'")
    [void]$text.AppendLine('    Files = @(')
    foreach ($entry in $entries) {
        if ($entry.Path.Contains("'")) { throw "A packaged path contains a quote: $($entry.Path)" }
        [void]$text.AppendLine("        @{ Path = '$($entry.Path)'; Sha256 = '$($entry.Sha256)' }")
    }
    [void]$text.AppendLine('    )')
    [void]$text.AppendLine('}')
    $target = Join-Path $Package 'PackageIntegrity.psd1'
    [IO.File]::WriteAllText($target, ($text.ToString() -replace "`r?`n", "`r`n"), [Text.UTF8Encoding]::new($false))
    $target
}

if ($PSCmdlet.ParameterSetName -eq 'Manifest') {
    $Manifest = [IO.Path]::GetFullPath($Manifest, (Get-Location).Path).TrimEnd('\')
    foreach ($relative in $signedScripts) { Assert-OwnerSigned (Join-Path $Manifest $relative) }
    $path = Write-Manifest $Manifest
    Write-Output "Unsigned manifest written: $path"
    Write-Output "Sign it, then copy the signed file to $(Join-Path $root 'PackageIntegrity.psd1') and run: .\tools\New-MailboxMessageReportPackage.ps1 -Seal `"$Manifest`""
    return
}

if ($PSCmdlet.ParameterSetName -eq 'Seal') {
    $Seal = [IO.Path]::GetFullPath($Seal, (Get-Location).Path).TrimEnd('\')
    if ((Split-Path $Seal -Leaf) -cne "MailboxMessageReport-$version") { throw "The built folder must be named MailboxMessageReport-$version (the root folder of the ZIP): $Seal" }
    foreach ($used in 'data', 'reports', 'logs', 'bin') {
        if (Test-Path -LiteralPath (Join-Path $Seal $used)) { throw "The built folder contains a $used folder (a package that has been run); build a new one: $Seal" }
    }
    $signedManifest = Join-Path $root 'PackageIntegrity.psd1'
    if (-not (Test-Path -LiteralPath $signedManifest)) { throw "Copy the signed manifest to $signedManifest first." }
    $builtManifest = Join-Path $Seal 'PackageIntegrity.psd1'
    if (-not (Test-Path -LiteralPath $builtManifest) -or (Get-FileHash -LiteralPath $signedManifest -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $builtManifest -Algorithm SHA256).Hash) {
        Copy-Item -LiteralPath $signedManifest -Destination $builtManifest -Force
    }
    $result = & (Join-Path $Seal 'src\Test-PackageIntegrity.ps1') -Root $Seal -Distribution
    if ($result.Status -cne 'Valid') { throw 'The built package is not Valid.' }
    $sourceResult = & (Join-Path $root 'src\Test-PackageIntegrity.ps1') -Root $root
    if ($sourceResult.Status -notlike 'Valid*') { throw 'The repository package folder is not Valid.' }
    foreach ($relative in Get-RelativeFile $Seal) {
        $sourceFile = Join-Path $root $relative
        if (-not (Test-Path -LiteralPath $sourceFile) -or $relative -ceq 'config\MailboxMessageReport.config.psd1') { continue }
        if ((Get-FileHash -LiteralPath $sourceFile -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath (Join-Path $Seal $relative) -Algorithm SHA256).Hash) {
            throw "The packaged file differs from the repository file: $relative"
        }
    }
    if ((Get-FileHash -LiteralPath (Join-Path $Seal 'CHANGELOG.md') -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath (Join-Path $repoRoot 'CHANGELOG.md') -Algorithm SHA256).Hash) {
        throw 'The packaged CHANGELOG.md differs from the repository file.'
    }
    $zip = "$Seal.zip"
    if (Test-Path -LiteralPath $zip) { throw "The ZIP already exists: $zip" }
    Compress-Archive -LiteralPath $Seal -DestinationPath $zip -CompressionLevel Optimal
    [pscustomobject]@{
        Package = $Seal; Zip = $zip; ZipBytes = (Get-Item -LiteralPath $zip).Length
        ZipSha256 = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
        Files = $result.Files; Signed = $true; Published = $false
    } | ConvertTo-Json
    return
}

# ---- Build -------------------------------------------------------------------------------------------
if (-not $Unsigned) { foreach ($relative in $signedScripts) { Assert-OwnerSigned (Join-Path $root $relative) } }
if (-not $Destination) { $Destination = Join-Path (Split-Path $repoRoot -Parent) "package\MailboxMessageReport-$version" }
$Destination = [IO.Path]::GetFullPath($Destination, (Get-Location).Path).TrimEnd('\')
$rootPrefix = [IO.Path]::GetFullPath($repoRoot).TrimEnd('\') + '\'
if (($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be outside the repository folder: $Destination"
}
if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) { throw "The destination already exists: $Destination. Use -Force to replace it." }
    if (-not (Test-Path -LiteralPath (Join-Path $Destination 'Invoke-MailboxMessageReport.ps1'))) { throw "The destination is not a Mailbox Message Report package, it is not replaced: $Destination" }
    foreach ($used in 'reports', 'logs') {
        if (Test-Path -LiteralPath (Join-Path $Destination $used)) { throw "The destination contains a $used folder (a package that has been run), it is not replaced: $Destination" }
    }
    Remove-Item -LiteralPath $Destination -Recurse -Force
}

# ---- Files needed at run time -------------------------------------------------------------------
$files = [Collections.Generic.List[string]]::new()
foreach ($f in 'Invoke-MailboxMessageReport.ps1', 'MailboxMessageReport.psd1', 'MailboxMessageReport.psm1', 'README.md', 'CHANGELOG.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md',
    'config\MailboxMessageReport.config.psd1', 'templates\Report.template.html', 'docs\MailboxMessageReport-UserGuide.html', 'docs\MailboxMessageReport-Guide.html') { $files.Add($f) }
Get-ChildItem -LiteralPath (Join-Path $root 'src') -File | Where-Object Extension -in '.ps1', '.cs' | ForEach-Object { $files.Add("src\$($_.Name)") }
foreach ($f in $files) {
    $sourceRoot = if ($f -eq 'CHANGELOG.md') { $repoRoot } else { $root }
    $source = Join-Path $sourceRoot $f
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

if ($Unsigned) {
    $committed = Join-Path $root 'PackageIntegrity.psd1'
    if (Test-Path -LiteralPath $committed) {
        Copy-Item -LiteralPath $committed -Destination (Join-Path $Destination 'PackageIntegrity.psd1')
        & (Join-Path $Destination 'src\Test-PackageIntegrity.ps1') -Root $Destination -DevelopmentMode 3> $null | Out-Null
    } else {
        Write-Warning 'package\PackageIntegrity.psd1 is not committed: this unsigned build carries no manifest.'
    }
}

# ---- Checks ---------------------------------------------------------------------------------------
$problems = [Collections.Generic.List[string]]::new()
foreach ($name in 'reports', 'logs', 'tests', 'artifacts', 'tools', 'docs\images') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
Get-ChildItem -LiteralPath $Destination -Recurse -File -Include '*.log', '*.csv', '*.json', '*.png', '*.Tests.ps1', '*FakeGraph*' |
    ForEach-Object { $problems.Add("Not a run-time file: $($_.Name)") }
foreach ($part in 'Console.ps1', 'Config.ps1', 'Graph.ps1', 'Mailboxes.ps1', 'Search.ps1', 'Report.ps1', 'Gui.ps1', 'Native.cs') {
    if (-not (Test-Path -LiteralPath (Join-Path $Destination "src\MailboxMessageReport.$part"))) { $problems.Add("Missing in the package: src\MailboxMessageReport.$part") }
}
foreach ($guide in 'MailboxMessageReport-UserGuide.html', 'MailboxMessageReport-Guide.html') {
    if (-not (Test-Path -LiteralPath (Join-Path $Destination "docs\$guide"))) { $problems.Add("Missing in the package: docs\$guide") }
}
$config = Import-PowerShellDataFile -LiteralPath (Join-Path $Destination 'config\MailboxMessageReport.config.psd1')
if ($config.Tenant.TenantId -or $config.Authentication.AppId -or $config.Authentication.CertificateThumbprint -or $config.Search.MailboxFile) { $problems.Add('The configuration of the package must not hold a tenant, an application, a certificate or a list of mailboxes.') }
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }

# The module loads from the package, and its reports go under the package folder.
$loaded = & pwsh -NoProfile -Command "`$WarningPreference = 'SilentlyContinue'; Import-Module '$Destination\MailboxMessageReport.psd1'; (Get-Command -Module MailboxMessageReport).Count; (Import-MmrConfiguration).OutputPath"
if ($LASTEXITCODE -ne 0 -or [int]$loaded[0] -lt 20 -or -not ([string]$loaded[1]).StartsWith($Destination)) { throw "The module does not load correctly from the package: $loaded" }

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  Mailbox Message Report $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host "  Check    : module loads ($($loaded[0]) commands), reports written under the package folder"
Write-Host "  Config   : empty tenant and application - fill them in before the first run (user guide, chapter 1)"
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,10:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }
