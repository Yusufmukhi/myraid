$src = 'C:\Users\yusuf\Downloads\Apollo\exam-calc\exam-calc'
$outZip = 'C:\Users\yusuf\Downloads\Apollo\exam-calc-production.zip'
$outHtml = 'C:\Users\yusuf\Downloads\Apollo\exam-calc-hostable.html'

Copy-Item -Path "$src\dist\index.html" -Destination $outHtml -Force

if (Test-Path $outZip) {
  Remove-Item $outZip -Force
}

$items = Get-ChildItem -Path $src | Where-Object { $_.Name -ne 'node_modules' -and $_.Name -ne '.git' }
Compress-Archive -Path $items.FullName -DestinationPath $outZip -Force

Get-Item $outHtml, $outZip | Select-Object Name, Length, LastWriteTime
