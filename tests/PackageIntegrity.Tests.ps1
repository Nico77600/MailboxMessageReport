#Requires -Version 7.4
<#
    Mailbox Message Report - package integrity tests (Pester 6.1 or later).
    Author  : Nicolas Fabert
    Version : 2.1.1

    The cases tagged RealSignature need the owner-signed files of a signed commit, or the release package named by
    MMR_RELEASE_PACKAGE (the folder written by New-MailboxMessageReportPackage.ps1 -Seal): they are skipped otherwise.
    The other cases run the real verifier on a fixture copy whose Authenticode checks are replaced by a no-op, so that
    the manifest, the inventory, the hashes and the folder rules are exercised before the owner signs.
#>
BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Source = Join-Path $script:RepoRoot 'package'
    $script:Pwsh = Join-Path $PSHOME 'pwsh.exe'
    $script:Signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
    $script:Work = Join-Path ([IO.Path]::GetTempPath()) ('mmr-integrity-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($script:Work)
    $script:FixtureNumber = 0
    $script:Fixture = $null
    $script:Version = (Import-PowerShellDataFile -LiteralPath (Join-Path $script:Source 'MailboxMessageReport.psd1')).ModuleVersion

    # The manifest writer of the package builder, used as is.
    $builderPath = Join-Path $script:RepoRoot 'tools\New-MailboxMessageReportPackage.ps1'
    $ast = [Management.Automation.Language.Parser]::ParseFile($builderPath, [ref]$null, [ref]$null)
    foreach ($name in 'Get-RelativeFile', 'Write-Manifest') {
        $function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name }, $true)
        if ($null -eq $function) { throw "The builder function $name is missing." }
        . ([scriptblock]::Create($function.Extent.Text))
    }
    $signer = $script:Signer
    $version = $script:Version

    function New-FixtureTemplate {
        $template = Join-Path $script:Work 'template'
        Copy-Item -LiteralPath $script:Source -Destination $template -Recurse
        foreach ($folder in 'logs', 'reports') { Remove-Item -LiteralPath (Join-Path $template $folder) -Recurse -Force -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath (Join-Path $template 'PackageIntegrity.psd1') -ErrorAction SilentlyContinue
        # A release package holds the HTML guides only.
        Remove-Item -LiteralPath (Join-Path $template 'docs\images') -Recurse -Force
        Get-ChildItem -LiteralPath (Join-Path $template 'docs') -Filter '*.md' | Remove-Item
        Copy-Item -LiteralPath (Join-Path $script:RepoRoot 'CHANGELOG.md') -Destination $template
        # The test copy of the verifier does not check Authenticode: everything else is the delivered code.
        $verifier = Join-Path $template 'src\Test-PackageIntegrity.ps1'
        $text = [IO.File]::ReadAllText($verifier)
        $stubbed = [regex]::Replace($text, '(?s)function Assert-OwnerSignature\(\[string\]\$Path\) \{.*?\r?\n\}\r?\n', "function Assert-OwnerSignature([string]`$Path) { }`r`n")
        if ($stubbed -ceq $text) { throw 'The owner-signature check was not found in the verifier.' }
        [IO.File]::WriteAllText($verifier, $stubbed, [Text.UTF8Encoding]::new($false))
        [void](Write-Manifest $template)
        $template
    }
    function Reset-Fixture([switch]$Source) {
        $script:FixtureNumber++
        $script:Fixture = Join-Path $script:Work ('package-{0:D4}' -f $script:FixtureNumber)
        Copy-Item -LiteralPath $script:Template -Destination $script:Fixture -Recurse
        # A source checkout: the file copied from the repository root at build time is absent.
        if ($Source) { Remove-Item -LiteralPath (Join-Path $script:Fixture 'CHANGELOG.md') }
    }
    function Invoke-Guard([string[]]$Switch = @()) {
        $verifier = Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1'
        $arguments = @{ Root = $script:Fixture }
        foreach ($name in $Switch) { $arguments[$name] = $true }
        & $verifier @arguments 3> $null
    }
    function Invoke-Child([string]$Command) {
        $output = & $script:Pwsh -NoProfile -NonInteractive -Command $Command 2>&1 | ForEach-Object { "$_" }
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Text = ($output -join "`n") }
    }
    function Edit-Protected([string]$Path) {
        if ([IO.Path]::GetExtension($Path) -in '.ps1', '.psm1', '.psd1') {
            [IO.File]::WriteAllText($Path, "# Evidence-only tamper`r`n" + [IO.File]::ReadAllText($Path), [Text.UTF8Encoding]::new($true))
        } else { [IO.File]::AppendAllBytes($Path, [Text.Encoding]::ASCII.GetBytes("`r`n# Evidence-only tamper")) }
    }
    $script:Template = New-FixtureTemplate
    $script:Entries = @((Import-PowerShellDataFile -LiteralPath (Join-Path $script:Template 'PackageIntegrity.psd1')).Files)
}

AfterAll {
    Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Signed manifest and verifier' {
    BeforeEach { Reset-Fixture }
    It 'accepts the distribution, the same files as a source checkout, and each in its own mode' {
        (Invoke-Guard @('Distribution')).Status | Should -Be 'Valid'
        (Invoke-Guard).Status | Should -Be 'Valid (files only)'
        (Invoke-Guard @('DevelopmentMode')).Status | Should -Be 'Valid (DevelopmentMode)'
        Reset-Fixture -Source
        (Invoke-Guard).Status | Should -Be 'Valid (files only)'
        { Invoke-Guard @('Distribution') } | Should -Throw
    }
    It 'refuses -Distribution together with -DevelopmentMode' {
        { Invoke-Guard @('Distribution', 'DevelopmentMode') } | Should -Throw -ExpectedMessage '*cannot be combined*'
    }
    It 'inventories exactly the protected files' {
        $script:Entries.Count | Should -Be 19
        @($script:Entries.Path | Select-Object -Unique).Count | Should -Be 19
        foreach ($path in 'Invoke-MailboxMessageReport.ps1', 'MailboxMessageReport.psm1', 'MailboxMessageReport.psd1', 'src\Test-PackageIntegrity.ps1',
            'src\MailboxMessageReport.Native.cs', 'templates\Report.template.html', 'docs\MailboxMessageReport-Guide.html', 'docs\MailboxMessageReport-UserGuide.html',
            'README.md', 'CHANGELOG.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md') {
            $script:Entries.Path | Should -Contain $path
        }
        foreach ($file in Get-ChildItem (Join-Path $script:Source 'src') -File) {
            $script:Entries.Path | Should -Contain ('src\' + $file.Name)
        }
        $script:Entries.Path | Should -Not -Contain 'config\MailboxMessageReport.config.psd1'
        $script:Entries.Path | Should -Not -Contain 'PackageIntegrity.psd1'
    }
    It 'keeps the verifier only in src and authenticates the package parent rather than src' {
        Test-Path -LiteralPath (Join-Path $script:Source 'Test-PackageIntegrity.ps1') | Should -BeFalse
        { & (Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1') -Root (Join-Path $script:Fixture 'src') } |
            Should -Throw -ExpectedMessage '*verifier must belong to this package root*'
        $outside = Join-Path $script:Work 'copied-verifier.ps1'
        Copy-Item -LiteralPath (Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1') -Destination $outside
        { & $outside -Root $script:Fixture } | Should -Throw -ExpectedMessage '*verifier must belong to this package root*'
    }
    It 'accepts the editable configuration and the generated folders' {
        Add-Content (Join-Path $script:Fixture 'config\MailboxMessageReport.config.psd1') '# Fictitious customization'
        foreach ($folder in 'logs', 'reports') {
            [void][IO.Directory]::CreateDirectory((Join-Path $script:Fixture $folder))
            Set-Content (Join-Path $script:Fixture "$folder\generated.txt") 'generated'
        }
        (Invoke-Guard @('Distribution')).Status | Should -Be 'Valid'
    }
    It 'rejects a changed byte in every protected file' {
        foreach ($entry in $script:Entries) {
            Reset-Fixture
            Edit-Protected (Join-Path $script:Fixture $entry.Path)
            { Invoke-Guard @('Distribution') } | Should -Throw -ExpectedMessage "*SHA256 mismatch: $($entry.Path)*"
            if ($entry.Path -cne 'CHANGELOG.md') {
                { Invoke-Guard } | Should -Throw -ExpectedMessage "*SHA256 mismatch: $($entry.Path)*"
                { Invoke-Guard @('DevelopmentMode') } | Should -Throw -ExpectedMessage "*SHA256 mismatch: $($entry.Path)*"
            }
        }
    }
    It 'rejects a missing protected file' {
        foreach ($entry in $script:Entries) {
            Reset-Fixture
            Remove-Item -LiteralPath (Join-Path $script:Fixture $entry.Path)
            { Invoke-Guard @('Distribution') } | Should -Throw
        }
    }
    It 'rejects an unexpected file anywhere in the package' {
        foreach ($relative in 'src\extra.cs', 'src\extra.ps1', 'templates\extra.html', 'docs\extra.ps1', 'docs\extra.html', 'docs\images\extra.png',
            'config\extra.psd1', 'extra.ps1', 'extra.dll', 'extra.cmd', 'tools\Test-PackageIntegrity.ps1', 'Test-PackageIntegrity.ps1') {
            Reset-Fixture
            [void][IO.Directory]::CreateDirectory((Split-Path (Join-Path $script:Fixture $relative) -Parent))
            Set-Content (Join-Path $script:Fixture $relative) 'not executable'
            { Invoke-Guard } | Should -Throw -ExpectedMessage '*unexpected stable file*Move your own files*MailboxMessageReport-data*'
        }
    }
    It 'rejects a junction in a protected subtree' {
        $outside = Join-Path $script:Work 'junction-target'
        [void][IO.Directory]::CreateDirectory($outside)
        foreach ($relative in 'src', 'templates', 'docs', 'config') {
            Reset-Fixture
            $link = Join-Path $script:Fixture $relative
            Remove-Item -LiteralPath $link -Recurse -Force
            New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
            try { { Invoke-Guard } | Should -Throw }
            finally { (Get-Item -LiteralPath $link -Force).Delete() }
        }
    }
    It 'rejects a junction used as a generated folder' {
        $outside = Join-Path $script:Work 'junction-reports'
        [void][IO.Directory]::CreateDirectory($outside)
        $link = Join-Path $script:Fixture 'reports'
        New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
        try { { Invoke-Guard } | Should -Throw -ExpectedMessage '*reparse point*' }
        finally { (Get-Item -LiteralPath $link -Force).Delete() }
    }
    It 'rejects an absent manifest and a manifest of the wrong shape' {
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1')
        { Invoke-Guard } | Should -Throw -ExpectedMessage '*PackageIntegrity.psd1 is missing. Use the complete release ZIP (https://github.com/Nico77600/MailboxMessageReport/releases/latest)*-DevelopmentMode*'
        Reset-Fixture
        Copy-Item -LiteralPath (Join-Path $script:Fixture 'MailboxMessageReport.psd1') -Destination (Join-Path $script:Fixture 'PackageIntegrity.psd1') -Force
        { Invoke-Guard } | Should -Throw -ExpectedMessage '*unsupported or incomplete manifest schema*'
    }
    It 'rejects a manifest that omits, duplicates, adds or retargets an entry' {
        $path = Join-Path $script:Fixture 'PackageIntegrity.psd1'
        $original = [IO.File]::ReadAllText($path)
        $variants = [ordered]@{
            omitted     = [regex]::Replace($original, "(?m)^\s*@\{ Path = 'README.md';.*\r?\n", '')
            duplicate   = $original.Replace("@{ Path = 'README.md';", "@{ Path = 'THIRD-PARTY-NOTICES.md';")
            unlisted    = $original.Replace("Path = 'README.md'", "Path = 'src\Extra.cs'")
            traversal   = $original.Replace("Path = 'README.md'", "Path = '..\README.md'")
            absolute    = $original.Replace("Path = 'README.md'", "Path = 'C:\Windows\win.ini'")
            wronghash   = $original.Replace("Path = 'README.md'; Sha256 = '", "Path = 'README.md'; Sha256 = 'g")
            wrongsigner = $original.Replace($script:Signer, '0' * 40)
        }
        foreach ($name in $variants.Keys) {
            [IO.File]::WriteAllText($path, $variants[$name], [Text.UTF8Encoding]::new($true))
            { Invoke-Guard @('Distribution') } | Should -Throw -ExpectedMessage '*Package integrity:*' -Because $name
        }
    }
}

Describe 'Integrity at the entry points' {
    BeforeEach { Reset-Fixture }
    It 'refuses to start the entry script from a changed package, before the module is loaded' {
        foreach ($relative in 'src\MailboxMessageReport.Native.cs', 'src\Test-PackageIntegrity.ps1', 'MailboxMessageReport.psm1', 'README.md') {
            Reset-Fixture
            Edit-Protected (Join-Path $script:Fixture $relative)
            $run = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MailboxMessageReport.ps1')' -Mailbox a@contoso.test; 'EXIT=' + `$LASTEXITCODE")
            $run.Text | Should -Match 'Package integrity.*owner signature'
            $run.Text | Should -Not -Match 'EXIT=0'
            Test-Path -LiteralPath (Join-Path $script:Fixture 'logs') | Should -BeFalse
            Test-Path -LiteralPath (Join-Path $script:Fixture 'reports') | Should -BeFalse
        }
    }
    It 'refuses to import the module from a changed package (the unsigned fixture verifier is refused first)' {
        Edit-Protected (Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1')
        $run = Invoke-Child ("Import-Module '$(Join-Path $script:Fixture 'MailboxMessageReport.psd1')' -Force -ErrorAction Stop")
        $run.Text | Should -Match 'Package integrity:'
    }
    It 'does not run a verifier that lacks the owner signature, even from an otherwise correct package' {
        # The fixture verifier is unsigned: the module and the entry script must refuse it instead of executing it.
        $import = Invoke-Child ("Import-Module '$(Join-Path $script:Fixture 'MailboxMessageReport.psd1')' -Force -ErrorAction Stop")
        $import.Text | Should -Match 'owner signature'
        $run = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MailboxMessageReport.ps1')' -Mailbox a@contoso.test -DevelopmentMode; 'EXIT=' + `$LASTEXITCODE")
        $run.Text | Should -Match 'owner signature'
    }
    It 'accepts an unverified source checkout only with -DevelopmentMode' {
        Reset-Fixture -Source
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1')
        $plain = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MailboxMessageReport.ps1')' -Mailbox a@contoso.test; 'EXIT=' + `$LASTEXITCODE")
        $plain.Text | Should -Match 'PackageIntegrity.psd1 is missing\. Use the complete release ZIP .*releases/latest.*-DevelopmentMode'
        $config = Join-Path $script:Work 'dev.config.psd1'
        "@{ Graph = @{ MaxConcurrency = 0 } }" | Set-Content $config
        $development = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MailboxMessageReport.ps1')' -Mailbox a@contoso.test -ConfigPath '$config' -DevelopmentMode 3>&1 | Out-String")
        $development.Text | Should -Match 'DevelopmentMode: no package manifest'
        # The run goes on to the configuration: the integrity block leaves the entry script able to start.
        $development.Text | Should -Match 'Graph.MaxConcurrency must be'
        $development.Text | Should -Not -Match 'positional|Package integrity'
        # One warning per process: the module does not repeat it at import.
        ([regex]::Matches($development.Text, 'no package manifest|no PackageIntegrity\.psd1')).Count | Should -Be 1
    }
}

Describe 'Compiled helpers' {
    It 'compiles the verified C# source in memory on each run and never writes or reads a cached assembly' {
        $module = [IO.File]::ReadAllText((Join-Path $script:Source 'MailboxMessageReport.psm1'))
        $module | Should -Match 'Add-Type -Path \(Join-Path \$PSScriptRoot ''src\\MailboxMessageReport\.Native\.cs''\)'
        foreach ($file in Get-ChildItem $script:Source -Recurse -File -Include '*.ps1', '*.psm1') {
            [IO.File]::ReadAllText($file.FullName) | Should -Not -Match 'OutputAssembly|Assembly\]::(LoadFrom|LoadFile)|Add-Type[^\r\n]*-Path[^\r\n]*\.dll|LOCALAPPDATA' -Because $file.Name
        }
        $child = Invoke-Child ("Import-Module '$(Join-Path $script:Source 'MailboxMessageReport.psd1')' -Force 3>`$null; `$a = ('MailboxMessageReportNative.Fast' -as [type]).Assembly; 'LOCATION=[' + `$a.Location + ']'")
        $child.Text | Should -Match 'LOCATION=\[\]'
        Get-ChildItem $script:Source -Recurse -File -Filter '*.dll' | Should -BeNullOrEmpty
    }
}

Describe 'Start of the entry script' {
    It 'reaches the configuration loading with -DevelopmentMode and reports a missing configuration file' {
        Reset-Fixture -Source
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1')
        $missing = Join-Path $script:Work 'no-such-config.psd1'
        $run = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MailboxMessageReport.ps1')' -Mailbox a@contoso.test -ConfigPath '$missing' -DevelopmentMode 3>`$null; 'EXIT=' + `$LASTEXITCODE")
        $run.Text | Should -Match 'Configuration file not found'
        $run.Text | Should -Match 'EXIT=1'
        $run.Text | Should -Not -Match 'positional|Package integrity'
    }
}

Describe 'Builder' {
    It 'rebuilds the HTML guides to the same bytes' {
        $guides = Get-ChildItem (Join-Path $script:Source 'docs') -Filter '*.html' -File
        $before = @($guides | Get-FileHash | ForEach-Object Hash)
        & $script:Pwsh -NoProfile -File (Join-Path $script:RepoRoot 'tools\Build-Documentation.ps1') | Out-Null
        @($guides | Get-FileHash | ForEach-Object Hash) | Should -Be $before
    }
    It 'passes the builder switch to the package stage of the CI as a named parameter' {
        $ci = [IO.File]::ReadAllText((Join-Path $script:RepoRoot '.github\scripts\Invoke-CI.ps1'))
        $ci | Should -Match '\$builderArguments = @\{ Destination = \$destination \}'
        $ci | Should -Match '\$builderArguments.Unsigned = \$true'
        $ci | Should -Match '& \$builder @builderArguments'
    }
    It 'refuses to seal a built folder that holds reports, before it writes a ZIP' {
        $built = Join-Path $script:Work 'seal-check\MailboxMessageReport-2.1.1'
        [void][IO.Directory]::CreateDirectory((Join-Path $built 'reports'))
        $run = Invoke-Child ("& '$(Join-Path $script:RepoRoot 'tools\New-MailboxMessageReportPackage.ps1')' -Seal '$built' 2>&1 | Out-String")
        $run.Text | Should -Match 'contains a reports folder'
        Test-Path -LiteralPath "$built.zip" | Should -BeFalse
    }
    It 'refuses to seal a folder that holds reports or logs, and compares the shipped CHANGELOG' {        $text = [IO.File]::ReadAllText((Join-Path $script:RepoRoot 'tools\New-MailboxMessageReportPackage.ps1'))
        $text | Should -Match 'foreach \(\$used in ''data'', ''reports'', ''logs'', ''bin''\)'
        $text | Should -Match 'The packaged CHANGELOG.md differs from the repository file'
    }
    It 'writes a manifest that the verifier parses, with a stable order and the five keys' {
        Reset-Fixture
        $manifest = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1')
        $manifest.Keys | Sort-Object | Should -Be @('Algorithm', 'Files', 'SchemaVersion', 'SignerThumbprint', 'Version')
        $manifest.Version | Should -Be $script:Version
        $paths = @($manifest.Files.Path)
        $paths | Should -Be @($paths | Sort-Object { $_ } -CaseSensitive)
    }
    It 'offers the three modes of the release procedure and refuses unsigned scripts for a release build' {
        $text = [IO.File]::ReadAllText((Join-Path $script:RepoRoot 'tools\New-MailboxMessageReportPackage.ps1'))
        $text | Should -Match '\[switch\]\$Unsigned'
        $text | Should -Match '\[string\]\$Manifest'
        $text | Should -Match '\[string\]\$Seal'
        $text | Should -Match 'if \(-not \$Unsigned\) \{ foreach \(\$relative in \$signedScripts\) \{ Assert-OwnerSigned'
    }
}

Describe 'Owner signatures' -Tag 'RealSignature' {
    It 'requires valid timestamped owner signatures on the PowerShell files and the manifest' {
        if (-not (Test-Path -LiteralPath (Join-Path $script:Source 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'package\PackageIntegrity.psd1 is not committed: this tree is not signed'
            return
        }
        $names = @('PackageIntegrity.psd1', 'Invoke-MailboxMessageReport.ps1', 'MailboxMessageReport.psm1', 'MailboxMessageReport.psd1') +
            @(Get-ChildItem (Join-Path $script:Source 'src') -Filter '*.ps1' -File | ForEach-Object { 'src\' + $_.Name })
        foreach ($name in $names) {
            $signature = Get-AuthenticodeSignature -LiteralPath (Join-Path $script:Source $name)
            [string]$signature.Status | Should -Be 'Valid' -Because $name
            $signature.SignerCertificate.Thumbprint | Should -BeExactly $script:Signer
            $signature.TimeStamperCertificate | Should -Not -BeNullOrEmpty
        }
        (& (Join-Path $script:Source 'src\Test-PackageIntegrity.ps1') -Root $script:Source).Status | Should -Be 'Valid (files only)'
    }
    It 'rejects an altered manifest before parsing it' {
        if (-not (Test-Path -LiteralPath (Join-Path $script:Source 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'package\PackageIntegrity.psd1 is not committed: this tree is not signed'
            return
        }
        $fixture = Join-Path $script:Work 'real-signature'
        Copy-Item -LiteralPath $script:Source -Destination $fixture -Recurse
        $path = Join-Path $fixture 'PackageIntegrity.psd1'
        [IO.File]::WriteAllText($path, "# Changed`r`n" + [IO.File]::ReadAllText($path), [Text.UTF8Encoding]::new($true))
        { & (Join-Path $fixture 'src\Test-PackageIntegrity.ps1') -Root $fixture } | Should -Throw -ExpectedMessage '*expected trusted, timestamped owner signature*'
    }
    It 'starts the sealed release package up to its configuration and refuses one changed template byte' {
        $package = $env:MMR_RELEASE_PACKAGE
        if (-not $package -or -not (Test-Path -LiteralPath (Join-Path $package 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'MMR_RELEASE_PACKAGE names no sealed release package folder (New-MailboxMessageReportPackage.ps1 -Seal)'
            return
        }
        $missing = Join-Path $script:Work 'no-such-config.psd1'
        $run = Invoke-Child ("& '$(Join-Path $package 'Invoke-MailboxMessageReport.ps1')' -Mailbox a@contoso.test -ConfigPath '$missing'; 'EXIT=' + `$LASTEXITCODE")
        $run.Text | Should -Match 'Configuration file not found'
        $run.Text | Should -Not -Match 'Package integrity|positional'
        $copy = Join-Path $script:Work 'sealed-template'
        Copy-Item -LiteralPath $package -Destination $copy -Recurse
        [IO.File]::AppendAllBytes((Join-Path $copy 'templates\Report.template.html'), [byte[]](32))
        $run = Invoke-Child ("& '$(Join-Path $copy 'Invoke-MailboxMessageReport.ps1')' -Mailbox a@contoso.test -ConfigPath '$missing'; 'EXIT=' + `$LASTEXITCODE")
        $run.Text | Should -Match 'SHA256 mismatch: templates'
    }
    It 'rejects a cryptographically valid manifest with its timestamp removed' {
        $package = $env:MMR_RELEASE_PACKAGE
        if (-not $package -or -not (Test-Path -LiteralPath (Join-Path $package 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'MMR_RELEASE_PACKAGE names no sealed release package folder (New-MailboxMessageReportPackage.ps1 -Seal)'
            return
        }
        Add-Type -AssemblyName System.Security.Cryptography.Pkcs
        $fixture = Join-Path $script:Work 'real-timestamp'
        Copy-Item -LiteralPath $package -Destination $fixture -Recurse
        $path = Join-Path $fixture 'PackageIntegrity.psd1'
        $encoding = [Text.UTF8Encoding]::new($false, $true)
        $text = $encoding.GetString([IO.File]::ReadAllBytes($path))
        $match = [regex]::Match($text, '(?ms)^# SIG # Begin signature block\r?\n(?<Data>.*?)^# SIG # End signature block\r?\n?\z')
        $match.Success | Should -BeTrue
        $cms = [Security.Cryptography.Pkcs.SignedCms]::new()
        $cms.Decode([Convert]::FromBase64String((($match.Groups['Data'].Value -replace '(?m)^#\s?', '') -replace '\s', '')))
        foreach ($attribute in @($cms.SignerInfos[0].UnsignedAttributes)) {
            foreach ($value in $attribute.Values) {
                $cms.SignerInfos[0].RemoveUnsignedAttribute([Security.Cryptography.AsnEncodedData]::new($attribute.Oid, $value.RawData))
            }
        }
        $base64 = [Convert]::ToBase64String($cms.Encode())
        $lines = for ($i = 0; $i -lt $base64.Length; $i += 64) { '# ' + $base64.Substring($i, [Math]::Min(64, $base64.Length - $i)) }
        [IO.File]::WriteAllBytes($path, $encoding.GetBytes($text.Substring(0, $match.Index) +
            "# SIG # Begin signature block`r`n" + ($lines -join "`r`n") + "`r`n# SIG # End signature block`r`n"))
        [string](Get-AuthenticodeSignature -LiteralPath $path).Status | Should -Be 'Valid'
        (Get-AuthenticodeSignature -LiteralPath $path).TimeStamperCertificate | Should -BeNullOrEmpty
        { & (Join-Path $fixture 'src\Test-PackageIntegrity.ps1') -Root $fixture -Distribution } | Should -Throw -ExpectedMessage '*expected trusted, timestamped owner signature*'
    }
    It 'accepts the sealed release package and refuses a copy with one changed byte' {
        $package = $env:MMR_RELEASE_PACKAGE
        if (-not $package -or -not (Test-Path -LiteralPath (Join-Path $package 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'MMR_RELEASE_PACKAGE names no sealed release package folder (New-MailboxMessageReportPackage.ps1 -Seal)'
            return
        }
        (& (Join-Path $package 'src\Test-PackageIntegrity.ps1') -Root $package -Distribution).Status | Should -Be 'Valid'
        $copy = Join-Path $script:Work 'sealed-copy'
        Copy-Item -LiteralPath $package -Destination $copy -Recurse
        [IO.File]::AppendAllBytes((Join-Path $copy 'README.md'), [byte[]](10))
        { & (Join-Path $copy 'src\Test-PackageIntegrity.ps1') -Root $copy -Distribution } | Should -Throw -ExpectedMessage '*SHA256 mismatch: README.md*'
        $run = Invoke-Child ("& '$(Join-Path $copy 'Invoke-MailboxMessageReport.ps1')' -Mailbox a@contoso.test; 'EXIT=' + `$LASTEXITCODE")
        $run.Text | Should -Match 'Package integrity'
    }
}
