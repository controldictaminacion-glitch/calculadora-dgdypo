param(
  [Parameter(Mandatory = $true)][string[]]$Path
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
[Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.DataReader, Windows.Storage.Streams, ContentType=WindowsRuntime] | Out-Null

$operationMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
} | Select-Object -First 1
$actionMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and -not $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
} | Select-Object -First 1

function Await-Operation($Operation, [Type]$ResultType) {
  $task = $operationMethod.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
  $task.Wait()
  return $task.Result
}
function Await-Action($Operation) {
  $task = $actionMethod.Invoke($null, @($Operation))
  $task.Wait()
}

foreach ($inputPath in $Path) {
  $absolute = [System.IO.Path]::GetFullPath($inputPath)
  $file = Await-Operation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($absolute)) ([Windows.Storage.StorageFile])
  $pdf = Await-Operation ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
  Write-Output ((Split-Path $absolute -Leaf) + ': ' + $pdf.PageCount + ' pages')
  for ($index = 0; $index -lt $pdf.PageCount; $index++) {
    $page = $pdf.GetPage($index)
    $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $options.DestinationWidth = [uint32][Math]::Ceiling($page.Size.Width * 1.5)
    $options.DestinationHeight = [uint32][Math]::Ceiling($page.Size.Height * 1.5)
    Await-Action ($page.RenderToStreamAsync($stream, $options))
    $reader = New-Object Windows.Storage.Streams.DataReader($stream.GetInputStreamAt(0))
    $length = Await-Operation ($reader.LoadAsync([uint32]$stream.Size)) ([uint32])
    $bytes = New-Object byte[] $length
    $reader.ReadBytes($bytes)
    $outputPath = Join-Path ([System.IO.Path]::GetDirectoryName($absolute)) ([System.IO.Path]::GetFileNameWithoutExtension($absolute) + '-page-' + ($index + 1) + '.png')
    [System.IO.File]::WriteAllBytes($outputPath, $bytes)
    Write-Output $outputPath
    $reader.Dispose()
    $stream.Dispose()
    $page.Dispose()
  }
}
# Windows.Data.Pdf can fault while CLR/WinRT objects unload after a successful render.
# The streams and pages above are already closed; exit the standalone renderer directly.
[Console]::Out.Flush()
[Console]::Error.Flush()
[Environment]::Exit(0)
