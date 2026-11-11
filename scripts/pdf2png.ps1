<#
  pdf2png.ps1
  PDF の各ページを PNG にする（Windows 標準の Windows.Data.Pdf を使用。追加ソフト不要）。
  WinRT のスレッドが残ってプロセスが終わらないことがあるため、最後に [Environment]::Exit で必ず終了する。

  -Pdf <file>        1 つの PDF   → 出力: {"pageCount":N,"images":[...]}
  -ListFile <file>   PDF のパスを 1 行に 1 つ → 出力: {"results":[{"pdf":..,"pageCount":N,"images":[..],"error":..}]}
#>
param(
    [string]$Pdf,
    [string]$ListFile,
    [int]$MaxPages = 3,
    [int]$Width = 1200
)
$code = 0
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Convert-One([string]$Path) {
    $file = Wait-Op ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
    $doc = Wait-Op ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
    $base = $Path.Substring(0, $Path.Length - [IO.Path]::GetExtension($Path).Length)
    $images = @()
    $n = [Math]::Min([int]$doc.PageCount, $MaxPages)
    for ($i = 0; $i -lt $n; $i++) {
        $page = $doc.GetPage($i)
        $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
        $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
        $opts.DestinationWidth = $Width
        $opts.BitmapEncoderId = [Windows.Graphics.Imaging.BitmapEncoder]::PngEncoderId
        Wait-Action ($page.RenderToStreamAsync($stream, $opts))
        $png = "${base}_p$($i + 1).png"
        $in = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream.GetInputStreamAt(0))
        $fs = [IO.File]::Create($png)
        try { $in.CopyTo($fs) } finally { $fs.Dispose(); $in.Dispose(); $stream.Dispose(); $page.Dispose() }
        $images += $png
    }
    return [ordered]@{ pdf = $Path; pageCount = [int]$doc.PageCount; images = $images }
}

try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapEncoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $methods = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 }
    $script:asTaskOp = $methods | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
    $script:asTaskAction = $methods | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
    function Wait-Op($op, [type]$t) { $task = $script:asTaskOp.MakeGenericMethod($t).Invoke($null, @($op)); [void]$task.Wait(-1); $task.Result }
    function Wait-Action($op) { $task = $script:asTaskAction.Invoke($null, @($op)); [void]$task.Wait(-1) }

    if ($ListFile) {
        $results = @()
        foreach ($p in [IO.File]::ReadAllLines($ListFile, [Text.Encoding]::UTF8)) {
            if (-not $p.Trim()) { continue }
            try { $results += Convert-One $p.Trim() }
            catch { $results += [ordered]@{ pdf = $p.Trim(); error = $_.Exception.Message } }
        }
        [Console]::Out.WriteLine((ConvertTo-Json -InputObject ([ordered]@{ results = $results }) -Compress -Depth 5))
    }
    else {
        $r = Convert-One $Pdf
        [Console]::Out.WriteLine((ConvertTo-Json -InputObject ([ordered]@{ pageCount = $r.pageCount; images = $r.images }) -Compress))
    }
}
catch {
    [Console]::Out.WriteLine((ConvertTo-Json -InputObject ([ordered]@{ error = $_.Exception.Message }) -Compress))
    $code = 1
}
[Console]::Out.Flush()
[Environment]::Exit($code)
