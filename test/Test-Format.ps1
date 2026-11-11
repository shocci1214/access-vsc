# AccessVsc.Format.ps1 のテスト（Access 不要）。test/format の入力と期待値は vba-format.js のテストと共通
$ErrorActionPreference = 'Stop'
$scripts = Join-Path $PSScriptRoot '..\scripts'
. (Join-Path $scripts 'AccessVsc.Common.ps1')
. (Join-Path $scripts 'AccessVsc.Lint.ps1')
. (Join-Path $scripts 'AccessVsc.Format.ps1')

$fail = 0
function Check([string]$Name, [bool]$Cond, $Detail) {
    if ($Cond) { Write-Output "PASS $Name" } else { Write-Output "FAIL $Name $Detail"; $script:fail++ }
}
$dir = Join-Path $PSScriptRoot 'format'
$in = Read-SourceText (Join-Path $dir 'basic.in.bas')
$exp = Read-SourceText (Join-Path $dir 'basic.out.bas')
$got = Format-VbaText $in
$diff = ''
if ($got -cne $exp -and $got) {
    $a = $got -split "`n"; $b = $exp -split "`n"
    for ($i = 0; $i -lt [Math]::Max($a.Count, $b.Count); $i++) { if ($a[$i] -cne $b[$i]) { $diff = "line $($i + 1): got [$($a[$i])] expected [$($b[$i])]"; break } }
}
Check 'basic indentation (same as JS)' ($got -ceq $exp) $diff
Check 'idempotent' ((Format-VbaText $exp) -ceq $exp)
Check 'keeps CRLF' ((Format-VbaText ($in.Replace("`n", "`r`n"))) -ceq $exp.Replace("`n", "`r`n"))
Check 'unbalanced -> null' ($null -eq (Format-VbaText (Read-SourceText (Join-Path $dir 'unbalanced.in.bas'))))
Check 'tab indent' ((Format-VbaText "Sub A()`nx = 1`nEnd Sub`n" '' "`t") -ceq "Sub A()`n`tx = 1`nEnd Sub`n")
$form = "Version =21`nBegin Form`n    Caption =`"x`"`nEnd`nCodeBehindForm`nAttribute VB_GlobalNameSpace = False`nPrivate Sub cmdOK_Click()`nDoCmd.Close`nEnd Sub`n"
Check 'form: only the code part' ((Format-VbaText $form 'forms/frm.txt') -ceq $form.Replace("`nDoCmd.Close", "`n    DoCmd.Close"))
Check 'Next j, i' ((Format-VbaText "Sub A()`nFor i = 1 To 2`nFor j = 1 To 2`nNext j, i`nEnd Sub`n") -ceq "Sub A()`n    For i = 1 To 2`n        For j = 1 To 2`n    Next j, i`nEnd Sub`n")

# ソースフォルダ全体（-Check と書き込み、lint の format / block-structure）
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("accessvsc-fmt-" + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Force (Join-Path $tmp 'modules') | Out-Null
    Write-SourceText (Join-Path $tmp 'modules\modA.bas') $in
    Write-SourceText (Join-Path $tmp 'modules\modB.bas') (Read-SourceText (Join-Path $dir 'unbalanced.in.bas'))
    Write-SourceText (Join-Path $tmp 'modules\modC.bas') $exp
    $lint = Invoke-SourceLint $tmp $null
    $rules = @($lint.problems | ForEach-Object { "$($_.file):$($_.rule)" })
    Check 'lint: format (info) for modA' ($rules -contains 'modules/modA.bas:format') ($rules -join ', ')
    Check 'lint: block-structure for modB' ($rules -contains 'modules/modB.bas:block-structure') ($rules -join ', ')
    Check 'lint: nothing for formatted modC' (-not ($rules | Where-Object { $_ -like 'modules/modC.bas:*format*' })) ($rules -join ', ')
    $chk = Invoke-SourceFormat $tmp $null $true
    Check 'format -Check lists modA only' ((@($chk.changed) -join ',') -eq 'modules/modA.bas' -and @($chk.skipped).Count -eq 1) (ConvertTo-Json $chk -Compress)
    Check 'format -Check writes nothing' ((Read-SourceText (Join-Path $tmp 'modules\modA.bas')) -ceq $in)
    $res = Invoke-SourceFormat $tmp $null $false
    Check 'format writes modA' ((Read-SourceText (Join-Path $tmp 'modules\modA.bas')) -ceq $exp)
    Check 'format leaves unbalanced modB' ((Read-SourceText (Join-Path $tmp 'modules\modB.bas')) -ceq (Read-SourceText (Join-Path $dir 'unbalanced.in.bas')))
}
finally { Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue }

if ($fail) { Write-Output "$fail failed"; exit 1 } else { Write-Output 'all passed' }
