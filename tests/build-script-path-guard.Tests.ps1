#Requires -Version 7.0
<#
.SYNOPSIS
    scripts/build-xlsm.ps1 の OutputPath guard（root外・reparse・既存directory）をPester形式で検証する。

.DESCRIPTION
    いずれもExcel COMへ到達する前に拒否される経路であり、GitHub Actions（windows-latest、
    Excel未インストール）上でも実行できる。正常系ビルド（Excel COMを要する）はこのファイルの
    対象外とし、対話型Windows環境でのrun-release-validation.ps1が担当する。
#>

BeforeAll {
    $script:Root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $script:BuildScript = Join-Path $Root 'scripts\build-xlsm.ps1'
    $script:Pwsh = (Get-Command pwsh -ErrorAction Stop).Source
    $script:GuardWorkRoot = Join-Path $Root '.build-work\pester-path-guard'
}

AfterAll {
    if (Test-Path -LiteralPath $GuardWorkRoot) { Remove-Item -LiteralPath $GuardWorkRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'build-xlsm.ps1 OutputPath guard' -Tag 'PathGuard' {
    BeforeEach {
        if (Test-Path -LiteralPath $GuardWorkRoot) { Remove-Item -LiteralPath $GuardWorkRoot -Recurse -Force -ErrorAction SilentlyContinue }
        New-Item -ItemType Directory -Path $GuardWorkRoot -Force | Out-Null
    }

    It 'rejects an OutputPath outside the repository root' {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        & $Pwsh -NoProfile -File $BuildScript -CandidateOnly -Profile Release -OutputPath (Join-Path $env:TEMP 'ftexcel-guard-escape.xlsm') 2>&1 | Out-Null
        $exit = $LASTEXITCODE
        $sw.Stop()
        $exit | Should -Not -Be 0
        $sw.ElapsedMilliseconds | Should -BeLessThan 15000 -Because 'the guard must fail before Excel COM is touched'
    }

    It 'rejects an OutputPath reached through a reparse point (junction) inside the repository root' {
        $junctionLink = Join-Path $GuardWorkRoot 'out-junction'
        New-Item -ItemType Junction -Path $junctionLink -Target $env:TEMP -Force | Out-Null
        $sw = [Diagnostics.Stopwatch]::StartNew()
        & $Pwsh -NoProfile -File $BuildScript -CandidateOnly -Profile Release -OutputPath (Join-Path $junctionLink 'escape.xlsm') 2>&1 | Out-Null
        $exit = $LASTEXITCODE
        $sw.Stop()
        $exit | Should -Not -Be 0
        $sw.ElapsedMilliseconds | Should -BeLessThan 15000 -Because 'the guard must fail before Excel COM is touched'
    }

    It 'rejects an OutputPath that already exists as a directory' {
        $dirTarget = Join-Path $GuardWorkRoot 'existing-dir.xlsm'
        New-Item -ItemType Directory -Path $dirTarget -Force | Out-Null
        $sw = [Diagnostics.Stopwatch]::StartNew()
        & $Pwsh -NoProfile -File $BuildScript -CandidateOnly -Profile Release -OutputPath $dirTarget 2>&1 | Out-Null
        $exit = $LASTEXITCODE
        $sw.Stop()
        $exit | Should -Not -Be 0
        $sw.ElapsedMilliseconds | Should -BeLessThan 15000 -Because 'the guard must fail before Excel COM is touched'
    }

    It 'rejects being invoked from outside a valid repository root (VERSION/src missing as siblings)' {
        # build-xlsm.ps1 は $PSScriptRoot の1階層上を root とみなすため、
        # 実際の scripts/ 配下と同じ深さ（isolated/scripts/build-xlsm.ps1）へ配置し、
        # isolated 直下に VERSION/src が存在しない状態を作る。
        $isolated = Join-Path $GuardWorkRoot 'isolated-copy'
        $isolatedScripts = Join-Path $isolated 'scripts'
        New-Item -ItemType Directory -Path $isolatedScripts -Force | Out-Null
        Copy-Item -LiteralPath $BuildScript -Destination (Join-Path $isolatedScripts 'build-xlsm.ps1') -Force
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $output = & $Pwsh -NoProfile -File (Join-Path $isolatedScripts 'build-xlsm.ps1') -CandidateOnly -Profile Release 2>&1
        $exit = $LASTEXITCODE
        $sw.Stop()
        $exit | Should -Not -Be 0
        ($output | Out-String) | Should -Match 'Repository root identity check failed' -Because 'the error must name the root identity check, not an unrelated failure'
        $sw.ElapsedMilliseconds | Should -BeLessThan 15000 -Because 'the guard must fail before Excel COM is touched'
    }
}
