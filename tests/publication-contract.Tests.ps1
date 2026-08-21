#Requires -Version 7.0
<#
.SYNOPSIS
    publication policy / VBA source contract / VERSION整合をPester形式で検証する。
#>

BeforeAll {
    $script:Root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $script:Policy = Get-Content -LiteralPath (Join-Path $Root 'release\publication-policy.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $script:Inventory = Get-Content -LiteralPath (Join-Path $Root 'release\vba-entrypoint-inventory-policy.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $script:ManifestSchemaPath = Join-Path $Root 'release\release-manifest.schema.json'
    $script:ManifestFixturePath = Join-Path $Root 'tests\fixtures\release-manifest.sample.json'
}

Describe 'publication-policy.json' -Tag 'PublicationPolicy' {
    It 'is schemaVersion 1' {
        $Policy.schemaVersion | Should -Be 1
    }

    It 'lists a file that actually exists, for every entry' {
        foreach ($entry in $Policy.files) {
            $full = Join-Path $Root ($entry.path -replace '/', '\')
            Test-Path -LiteralPath $full -PathType Leaf | Should -BeTrue -Because "$($entry.path) is listed in publication-policy.json"
        }
    }

    It 'produces zero UNCLASSIFIED_FILE via Test-PublicationCandidate.ps1' {
        $output = & pwsh -NoProfile -File (Join-Path $Root 'scripts\Test-PublicationCandidate.ps1') 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
    }
}

Describe 'VBA source byte contract' -Tag 'VbaSourceContract' {
    $sourceNames = @('modFolderTree.bas', 'modOperationPlan.bas', 'modFileOperations.bas', 'modRecycleBin.bas', 'modWorkbookIo.bas')

    It '<_> has no UTF-8 BOM' -ForEach $sourceNames {
        $bytes = [IO.File]::ReadAllBytes((Join-Path $Root "src\$_"))
        $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $hasBom | Should -BeFalse
    }

    It '<_> uses LF-only line endings' -ForEach $sourceNames {
        $bytes = [IO.File]::ReadAllBytes((Join-Path $Root "src\$_"))
        $hasCr = $false
        for ($i = 0; $i -lt $bytes.Length; $i++) { if ($bytes[$i] -eq 13) { $hasCr = $true; break } }
        $hasCr | Should -BeFalse
    }

    It '<_> does not hardcode #Const TEST_BUILD' -ForEach $sourceNames {
        $text = Get-Content -LiteralPath (Join-Path $Root "src\$_") -Raw -Encoding UTF8
        $text | Should -Not -Match '(?m)^#Const TEST_BUILD'
    }
}

Describe 'VBA entrypoint inventory' -Tag 'VbaSourceContract' {
    It 'every listed test wrapper name is present in its source module' {
        foreach ($moduleName in $Inventory.modules.PSObject.Properties.Name) {
            $module = $Inventory.modules.$moduleName
            $text = Get-Content -LiteralPath (Join-Path $Root "src\$moduleName.bas") -Raw -Encoding UTF8
            foreach ($wrapper in $module.testWrappers) {
                $text | Should -Match ([regex]::Escape($wrapper.name)) -Because "$moduleName.$($wrapper.name) is listed in the inventory policy"
            }
        }
    }

    It 'production accessors are not enclosed in #If TEST_BUILD' {
        foreach ($moduleName in $Inventory.modules.PSObject.Properties.Name) {
            $module = $Inventory.modules.$moduleName
            if ($module.productionAccessors.Count -eq 0) { continue }
            $text = Get-Content -LiteralPath (Join-Path $Root "src\$moduleName.bas") -Raw -Encoding UTF8
            $blocks = [regex]::Matches($text, '(?m)^#If TEST_BUILD Then$([\s\S]*?)^#End If$')
            foreach ($accessorName in $module.productionAccessors) {
                foreach ($block in $blocks) {
                    $block.Groups[1].Value | Should -Not -Match "\bPublic (Sub|Function) $([regex]::Escape($accessorName))\b" -Because "$moduleName.$accessorName must stay outside #If TEST_BUILD"
                }
            }
        }
    }
}

Describe 'release safety subset contract' -Tag 'ManifestSchema' {
    BeforeAll {
        $script:ManifestSchema = Get-Content -LiteralPath (Join-Path $Root 'release\release-manifest.schema.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $script:SubsetSchema = $ManifestSchema.properties.releaseSafetySubset
        $script:BundleScript = Get-Content -LiteralPath (Join-Path $Root 'scripts\New-PublicReleaseBundle.ps1') -Raw -Encoding UTF8
    }

    It 'schema keeps recycleBinUnavailableFallback as a non-dynamic, static-review-only record (OD-11)' {
        # OD-11で動的検証を対象外にした判断が、後からschemaへ静かに戻らないよう固定する。
        @($SubsetSchema.properties.checks.properties.recycleBinUnavailableFallback.enum) |
            Should -Be @('static-source-review-only')
    }

    It 'every recycleBinUnavailableFallback literal in the bundle script is allowed by the schema' {
        $allowed = @($SubsetSchema.properties.checks.properties.recycleBinUnavailableFallback.enum)
        $literals = @([regex]::Matches($BundleScript, "recycleBinUnavailableFallback\s*=\s*'([^']+)'") |
            ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
        $literals.Count | Should -BeGreaterThan 0 -Because 'the bundle script must record this field explicitly'
        foreach ($literal in $literals) {
            $allowed | Should -Contain $literal -Because "New-PublicReleaseBundle.ps1 emits '$literal'"
        }
    }

    It 'schema status enum does not carry a deferred/partial state' {
        # 3項目が揃えばsubsetはpass。「4項目目待ち」を表す中間stateは持たせない。
        @($SubsetSchema.properties.status.enum) | Should -Be @('skipped-by-flag', 'pass', 'fail')
    }

    It 'SECURITY-MODEL.md discloses that the recycle-bin fallback is not dynamically verified' {
        $securityModel = Get-Content -LiteralPath (Join-Path $Root 'SECURITY-MODEL.md') -Raw -Encoding UTF8
        $securityModel | Should -Match '## 安全性検査の検証範囲'
        $securityModel | Should -Match '動的な実地検証を行っていません'
        $securityModel | Should -Match ([regex]::Escape('ResolveRecycleVolumeConfig'))
    }
}

Describe 'VERSION consistency' -Tag 'ManifestSchema' {
    It 'VERSION is semver' {
        (Get-Content -LiteralPath (Join-Path $Root 'VERSION') -Raw).Trim() | Should -Match '^\d+\.\d+\.\d+$'
    }

    It 'CHANGELOG.md references the current VERSION' {
        $version = (Get-Content -LiteralPath (Join-Path $Root 'VERSION') -Raw).Trim()
        $changelog = Get-Content -LiteralPath (Join-Path $Root 'CHANGELOG.md') -Raw -Encoding UTF8
        $changelog | Should -Match ([regex]::Escape("[$version]"))
    }
}

Describe 'release-manifest.schema.json' -Tag 'ManifestSchema' {
    It 'declares a $schema pointing at json-schema.org' {
        $schema = Get-Content -LiteralPath $ManifestSchemaPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $schema.'$schema' | Should -Match 'json-schema.org'
    }

    It 'validates a representative release manifest instance' {
        (Test-Json -LiteralPath $ManifestFixturePath -SchemaFile $ManifestSchemaPath) | Should -BeTrue
    }

    It 'requires known limitations to be bound to the performance waiver' {
        $bundleScript = Get-Content -LiteralPath (Join-Path $Root 'scripts\New-PublicReleaseBundle.ps1') -Raw -Encoding UTF8
        $bundleScript | Should -Match 'knownLimitations'
        $bundleScript | Should -Match 'CF3CC8B3D32D063FBCB3ED7BD594EE3C68F5429B34BF2A11BA96523529B494C8'
    }

    It 'rejects an unknown top-level manifest property' {
        $manifest = Get-Content -LiteralPath $ManifestFixturePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $manifest | Add-Member -NotePropertyName unexpected -NotePropertyValue 'must fail'
        $json = $manifest | ConvertTo-Json -Depth 20
        { Test-Json -Json $json -SchemaFile $ManifestSchemaPath -ErrorAction Stop } | Should -Throw
    }
}
