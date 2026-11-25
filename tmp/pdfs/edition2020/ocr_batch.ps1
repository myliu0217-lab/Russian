param([Parameter(Mandatory=$true)][string]$Directory)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
  Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1 } |
  Select-Object -First 1
function Await($operation, $type) {
  $task = $method.MakeGenericMethod($type).Invoke($null, @($operation))
  $task.Wait()
  return $task.Result
}
$memoryType = [Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]
$writerType = [Windows.Storage.Streams.DataWriter,Windows.Storage.Streams,ContentType=WindowsRuntime]
$decoderType = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
$bitmapType = [Windows.Graphics.Imaging.SoftwareBitmap,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
$resultType = [Windows.Media.Ocr.OcrResult,Windows.Foundation,ContentType=WindowsRuntime]
$languageType = [Windows.Globalization.Language,Windows.Foundation,ContentType=WindowsRuntime]
$engineType = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$language = New-Object $languageType -ArgumentList 'ru'
$engine = $engineType::TryCreateFromLanguage($language)

Get-ChildItem -LiteralPath $Directory -Filter '*.png' -File | ForEach-Object {
  $stream = New-Object $memoryType
  $writer = New-Object $writerType -ArgumentList $stream
  $writer.WriteBytes([System.IO.File]::ReadAllBytes($_.FullName))
  $stored = Await ($writer.StoreAsync()) ([uint32])
  $stream.Seek(0)
  $decoder = Await ($decoderType::CreateAsync($stream)) $decoderType
  $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) $bitmapType
  $result = Await ($engine.RecognizeAsync($bitmap)) $resultType
  [System.IO.File]::WriteAllText([System.IO.Path]::ChangeExtension($_.FullName,'.txt'),$result.Text,[System.Text.UTF8Encoding]::new($false))
  $lines = @($result.Lines | ForEach-Object {
    $first = @($_.Words)[0]
    $rect = $first.BoundingRect
    ("{0:F0}`t{1:F0}`t{2}" -f $rect.X,$rect.Y,$_.Text)
  })
  [System.IO.File]::WriteAllLines([System.IO.Path]::ChangeExtension($_.FullName,'.tsv'),$lines,[System.Text.UTF8Encoding]::new($false))
  Write-Output $_.Name
  $bitmap.Dispose()
  $stream.Dispose()
}
