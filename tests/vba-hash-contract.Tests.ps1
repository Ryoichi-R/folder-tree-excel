#Requires -Version 7.0
BeforeAll {
    $builder = Join-Path $PSScriptRoot '../scripts/build-xlsm.ps1'
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($builder, [ref]$null, [ref]$parseErrors)
    if ($parseErrors.Count) { throw 'Builder did not parse.' }
    foreach ($name in @('Text-Hash', 'Normalize-Vba', 'Configure-OperationSheet')) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        . ([scriptblock]::Create($node.Extent.Text))
    }
}

Describe 'Embedded VBA semantic hash' -Tag 'VbaSourceContract' {
    It 'detects case changes inside a string literal' {
        (Text-Hash 'value = "ABC"') | Should -Not -Be (Text-Hash 'value = "abc"')
    }
    It 'allows VBE identifier capitalization without changing literals' {
        (Text-Hash 'Value = "ABC"') | Should -Be (Text-Hash 'value = "ABC"')
    }
    It 'preserves escaped quotes inside a VBA string' {
        (Text-Hash 'Value = "A""B"') | Should -Not -Be (Text-Hash 'value = "A""b"')
    }
    It 'does not treat apostrophes or Rem inside a string as comments' {
        (Text-Hash 'Value = "Rem '' ABC"') | Should -Not -Be (Text-Hash 'value = "rem '' abc"')
    }
    It 'detects changes in comments' {
        (Text-Hash "' ABC") | Should -Not -Be (Text-Hash "' abc")
        (Text-Hash 'Rem ABC') | Should -Not -Be (Text-Hash 'Rem abc')
    }
    It 'normalizes allowed profile and line ending differences' {
        $source = "Attribute VB_Name = `"Example`"`nOption Explicit`nValue = `"ABC`"`n"
        $embedded = "#Const TEST_BUILD = False`r`nOption Explicit`r`nvalue = `"ABC`"`r`n"
        (Text-Hash (Normalize-Vba $source)) | Should -Be (Text-Hash (Normalize-Vba $embedded))
    }
    It 'still detects a changed literal after profile normalization' {
        $source = "#Const TEST_BUILD = True`nValue = `"ABC`""
        $embedded = "#Const TEST_BUILD = False`nvalue = `"abc`""
        (Text-Hash (Normalize-Vba $source)) | Should -Not -Be (Text-Hash (Normalize-Vba $embedded))
    }
}
