# Enable TLSv1.2 for compatibility with older clients
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

$DownloadURL = 'https://raw.githubusercontent.com/BsNgChiThanh/Kich-hoat-Office/KichHoatOffice/Mondo%2C2021%2C2019%2C2016.cmd'

$rand = Get-Random -Maximum 1000
$FilePath = "$env:TEMP\Mondo%2C2021%2C2019%2C2016_$rand.cmd"

try {
    Invoke-WebRequest -Uri $DownloadURL -UseBasicParsing -OutFile $FilePath
} catch {
    Write-Error $_
	Return
}

if (Test-Path $FilePath) {
    Start-Process $FilePath -Wait
    $item = Get-Item -LiteralPath $FilePath
    $item.Delete()
}
