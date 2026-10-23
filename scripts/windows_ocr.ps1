param([Parameter(Mandatory=$true)][string]$InputJson,
      [Parameter(Mandatory=$true)][string]$OutputJson,
      [string]$Language='zh-Hans-CN')
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null=[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
$null=[Windows.Storage.Streams.IRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]
$null=[Windows.Graphics.Imaging.BitmapDecoder,Windows.Foundation,ContentType=WindowsRuntime]
$null=[Windows.Graphics.Imaging.SoftwareBitmap,Windows.Foundation,ContentType=WindowsRuntime]
$null=[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$null=[Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime]
$asTask=[System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
} | Select-Object -First 1
function Await-Operation($Operation,[Type]$Type) {
    $task=$asTask.MakeGenericMethod($Type).Invoke($null,@($Operation))
    $task.Wait()
    return $task.Result
}
$lang=New-Object Windows.Globalization.Language($Language)
$engine=[Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if ($null -eq $engine) {throw "Windows OCR language unavailable: $Language"}
$items=Get-Content -LiteralPath $InputJson -Raw -Encoding UTF8 | ConvertFrom-Json
$output=@()
foreach($item in $items) {
    $file=Await-Operation ([Windows.Storage.StorageFile]::GetFileFromPathAsync([string]$item.image)) ([Windows.Storage.StorageFile])
    $stream=Await-Operation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder=Await-Operation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap=Await-Operation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    try {
        $result=Await-Operation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $lines=@()
        foreach($line in $result.Lines) {
            $words=@()
            foreach($word in $line.Words) {
                $r=$word.BoundingRect
                $words+=@{text=$word.Text;box=@($r.X,$r.Y,$r.X+$r.Width,$r.Y+$r.Height)}
            }
            $lines+=@{text=$line.Text;words=$words}
        }
        $output+=@{id=$item.id;lines=$lines;angle=$result.TextAngle;language=$Language}
    } finally {$bitmap.Dispose();$stream.Dispose()}
}
$json=ConvertTo-Json -InputObject @($output) -Depth 12
[System.IO.File]::WriteAllText($OutputJson,$json,(New-Object System.Text.UTF8Encoding($false)))
