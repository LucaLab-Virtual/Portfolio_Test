<# 
This code is written in PowerShell scripting language. 
It uses the Google Chart API to generate the QR code image 
and the WebClient class to download the image file from the 
generated URL. 
#>

# Assigns the base URL for the Google Chart API
$baseUrl = 'https://chart.googleapis.com/chart'
# Sets the desired size of the QR code image
$size = '300x300'
# Sets the link that will be encoded in the QR code.
$link = "https://www.google.com"

<# 
Constructs the URL for generating the QR code using the Google Chart API
It uses string formatting ('-f') to substiture the values of
$baseUrl, $size, $link 
#>
$qrCodeUrl = "{0}?cht=qr&chs={1}&chl={2}" -f $baseUrl, $size,
[System.Uri]::EscapeDataString($link)

# Sets the path where the downloaded QR code image will be saved.
$downloadPath = 'C:\Users\Working\Desktop\QRCode.png'

<# 
Creates a new instance of the 'WebClient' class,
which allows downloading files from a URL 
#>
$webClient = New-Object System.Net.WebClient
<# 
Uses the 'DownloadFile' method ot the 'WebClient' object
to download the QR code image from the URL specified in 'qrCodeUrl' 
#>
$webClient.DownloadFile($qrCodeUrl,$downloadPath)

Write-Host "QR code image downloaded at: $downloadPath"