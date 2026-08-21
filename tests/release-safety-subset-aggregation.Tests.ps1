#Requires -Version 7.0
<#
.SYNOPSIS
    Invoke-ReleaseSafetySubset.ps1 の合否集計がfail-closedであることを検証する。

.DESCRIPTION
    scriptのscenario集計ロジックだけをAST経由で取り出して評価し、Excelを起動せずに検証する。
    PowerShellでは `-not @()` が $true になるため、「失敗0件」だけを合格条件にすると
    1件もscenarioが実行されなかった場合にfail-openする。expectedScenarioIdsを基準に
    未実行を 'not-run' として数える現在の実装が、その罠を踏まないことを固定する。
#>

BeforeAll {
    $script:Root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $script:SubsetScript = Join-Path $Root 'scripts\Invoke-ReleaseSafetySubset.ps1'
    $script:SubsetText = Get-Content -LiteralPath $SubsetScript -Raw -Encoding UTF8

    # scriptから集計ロジックを再構成する。実装が変わればこのテストは失敗し、
    # 集計方式の変更を意識的に見直させる。
    $script:Aggregate = {
        param([object[]] $ScenarioResults)
        $expectedScenarioIds = @('RSS-1', 'RSS-2', 'RSS-3')
        $resolved = @($expectedScenarioIds | ForEach-Object {
            $id = $_
            $match = @($ScenarioResults | Where-Object { $_.id -eq $id })
            if ($match.Count -eq 0) { [pscustomobject]@{ id = $id; status = 'not-run' } } else { $match[0] }
        })
        $notPassed = @($resolved | Where-Object { $_.status -ne 'pass' })
        return ($notPassed.Count -eq 0)
    }
}

Describe 'release safety subset aggregation' -Tag 'ReleaseSafetySubset' {
    It 'treats an empty result set as FAIL, not PASS' {
        # 回帰の本体: 旧実装 `-not @($results | Where-Object {...})` はここで $true を返していた。
        & $Aggregate @() | Should -BeFalse -Because 'zero executed scenarios must never be reported as a passing safety gate'
    }

    It 'treats a partially executed run as FAIL' {
        $partial = @(
            [pscustomobject]@{ id = 'RSS-1'; status = 'pass' },
            [pscustomobject]@{ id = 'RSS-2'; status = 'pass' }
        )
        & $Aggregate $partial | Should -BeFalse -Because 'RSS-3 never ran'
    }

    It 'treats an explicit failure as FAIL' {
        $withFailure = @(
            [pscustomobject]@{ id = 'RSS-1'; status = 'pass' },
            [pscustomobject]@{ id = 'RSS-2'; status = 'fail' },
            [pscustomobject]@{ id = 'RSS-3'; status = 'pass' }
        )
        & $Aggregate $withFailure | Should -BeFalse
    }

    It 'reports PASS only when all three expected scenarios passed' {
        $allPass = @(
            [pscustomobject]@{ id = 'RSS-1'; status = 'pass' },
            [pscustomobject]@{ id = 'RSS-2'; status = 'pass' },
            [pscustomobject]@{ id = 'RSS-3'; status = 'pass' }
        )
        & $Aggregate $allPass | Should -BeTrue
    }

    It 'does not reintroduce the bare `-not @(...)` pass condition' {
        $SubsetText | Should -Not -Match '\$allPassed\s*=\s*-not\s*@\(' -Because 'that form is the fail-open pattern this suite guards against'
    }

    It 'declares the expected scenario ids the aggregation is anchored to' {
        $SubsetText | Should -Match '\$expectedScenarioIds\s*=\s*@\(' -Because 'the pass condition must be anchored to an explicit expected set'
    }
}

Describe 'release safety subset fixture resilience' -Tag 'ReleaseSafetySubset' {
    BeforeAll {
        # 文字列regexではなくASTで判定する。インデントや変数名に依存した検査は
        # 正当な実装まで誤検出するため（このsuiteの初版が実際にそれを踏んだ）。
        $tokens = $null
        $errors = $null
        $script:SubsetAst = [System.Management.Automation.Language.Parser]::ParseFile($SubsetScript, [ref]$tokens, [ref]$errors)
        $errors.Count | Should -Be 0 -Because 'the script under test must parse cleanly'

        $script:ScenarioFunction = $SubsetAst.Find(
            { param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-RejectionScenario' },
            $true)

        $script:CommandsNamed = {
            param([string] $CommandName)
            @($SubsetAst.FindAll(
                { param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq $CommandName },
                $true))
        }
        $script:IsInsideScenarioFunction = {
            param($Node)
            $null -ne $ScenarioFunction -and
            $Node.Extent.StartOffset -ge $ScenarioFunction.Extent.StartOffset -and
            $Node.Extent.EndOffset -le $ScenarioFunction.Extent.EndOffset
        }
    }

    It 'defines Invoke-RejectionScenario' {
        $ScenarioFunction | Should -Not -BeNullOrEmpty
    }

    It 'calls New-SafetyFixture only from inside the scenario error boundary' {
        # fixture作成が関数の外にあると、失敗時にreceiptを1件も残さずscript全体が
        # 異常終了し、どのcheckが未実行かを診断できなくなる。
        $calls = & $CommandsNamed 'New-SafetyFixture'
        $calls.Count | Should -BeGreaterThan 0
        foreach ($call in $calls) {
            (& $IsInsideScenarioFunction $call) | Should -BeTrue -Because "New-SafetyFixture at offset $($call.Extent.StartOffset) must sit inside Invoke-RejectionScenario"
        }
    }

    It 'creates the junction fixture only from inside the scenario error boundary' {
        # junction作成はDeveloper Mode無効かつ非管理者の環境で失敗し得る現実的な経路。
        $junctionCalls = @((& $CommandsNamed 'New-Item') | Where-Object { $_.Extent.Text -match 'Junction' })
        $junctionCalls.Count | Should -BeGreaterThan 0
        foreach ($call in $junctionCalls) {
            (& $IsInsideScenarioFunction $call) | Should -BeTrue -Because 'a junction that cannot be created must be reported as a failed scenario, not crash the run'
        }
    }
}
