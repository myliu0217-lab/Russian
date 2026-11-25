param([Parameter(Mandatory=$true)][string]$ImagePath)
Add-Type -AssemblyName System.Runtime.WindowsRuntime

function Await($operation, $type) {
  $method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1 } |
    Select-Object -First 1
  $task = $method.MakeGenericMethod($type).Invoke($null, @($operation))
  try { $task.Wait() } catch { Write-Error $_.Exception.ToString(); throw }
  return $task.Result
}

$memoryType = [Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]
$writerType = [Windows.Storage.Streams.DataWriter,Windows.Storage.Streams,ContentType=WindowsRuntime]
$decoderType = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
$bitmapType = [Windows.Graphics.Imaging.SoftwareBitmap,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
$resultType = [Windows.Media.Ocr.OcrResult,Windows.Foundation,ContentType=WindowsRuntime]
$languageType = [Windows.Globalization.Language,Windows.Foundation,ContentType=WindowsRuntime]
$engineType = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]

$stream = New-Object $memoryType
$writer = New-Object $writerType -ArgumentList $stream
$writer.WriteBytes([System.IO.File]::ReadAllBytes((Resolve-Path $ImagePath).Path))
$stored = Await ($writer.StoreAsync()) ([uint32])
$stream.Seek(0)
$decoder = Await ($decoderType::CreateAsync($stream)) $decoderType
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) $bitmapType
$language = New-Object $languageType -ArgumentList 'ru'
$engine = $engineType::TryCreateFromLanguage($language)
$result = Await ($engine.RecognizeAsync($bitmap)) $resultType
$result.Text
