Add-Type -AssemblyName System.Runtime.WindowsRuntime
$langs = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]::AvailableRecognizerLanguages
$langs | ForEach-Object { $_.LanguageTag }
