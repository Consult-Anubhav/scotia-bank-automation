$usdExchangeRate = 0

# Set the API endpoint and parameters
$apiUrl = "https://api.fxratesapi.com/historical"
$date = "2022-11-30"
$currencies = "USD"
$baseCurrency = "CAD"
$apiKey = "fxr_live_71c57a5422d7df4405908aefe15afe4bfc3f"

# Construct the URI using UriBuilder
$uriBuilder = New-Object System.UriBuilder
$uriBuilder.Scheme = "https"
$uriBuilder.Host = "api.fxratesapi.com"
$uriBuilder.Path = "/historical"
$uriBuilder.Query = "date=$date&currencies=$currencies&base=$baseCurrency&api_key=$apiKey"

# Make the API request using Invoke-RestMethod
$response = Invoke-RestMethod -Uri $uriBuilder.Uri -Method Get

# Check if the request was successful
if ($response.success -eq $true) {
    # Store the USD exchange rate in a variable
    $usdExchangeRate = $response.rates.USD

    # Display the result
    Write-Host ("Exchange rate for USD on {0}: {1}" -f $date, $usdExchangeRate)
} else {
    # Display an error message if the request was not successful
    Write-Host "API request failed. Check the error message or try again."
}

	
Import-Module ImportExcel

$ExcelData = Open-ExcelPackage -path "/Users/yashdesai/Downloads/Was resume 2/OPICS Scotia Investments Jamaica Limited/ForexRates.xlsx"

$Data = $ExcelData.Workbook.WorkSheets["YAVG"].Cells

$Data[7,5].Value = $usdExchangeRate #prev:1.37067723 new:0.745505904
$Data[7,5] | Select -ExpandProperty Value
Close-ExcelPackage $ExcelData
