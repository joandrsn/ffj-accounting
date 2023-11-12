#Requires -Modules ImportExcel

Add-Type -AssemblyName System.Windows.Forms

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = "$HOME\Downloads"
$openFileDialog.Filter = "Excel (*.xlsx)|*.xlsx"
if ($openFileDialog.ShowDialog() -ne "OK") {
  return
}

$headers = "Date", "Description", "Amount"
$data = Import-Excel $openFileDialog.FileName -StartRow 6 -AsDate "Date" -HeaderName $headers

$folder = Join-Path "$env:APPDATA" "ffj-accounting"
if (-not (Test-Path $folder)) {
  $null = New-Item -Path $folder -ItemType Directory
}
$docnofile = Join-Path $folder "bank.txt"
if (-not (Test-Path $docnofile)) {
  $null = New-Item -Path $docnofile -ItemType File
}
$nextdocno = Get-Content $docnofile
if (-not $nextdocno) {
  $nextdocno = 1
}
$prompttext = "Hvad er næste bilagsnr.?`nForslag: {0}`nTryk Enter for at acceptere, ellers skriv hvad næste bilagsnr. skal være og tryk enter." -f $nextdocno
$promptvalue = Read-Host -Prompt $prompttext
if([string]::IsNullOrWhiteSpace($promptvalue))
{
    $documentNo = $nextdocno
}
else
{
    $documentNo = [int]$promptvalue
}

$data = $data | Where-Object {
  $PSItem.Description -notlike "DKFLX*" -and
  $PSItem.Description -ne "FI89136954,INDB.KA71" -and
  $PSItem.Description -notlike "Nets * 8650061" -and 
  $PSItem.Description -ne "danmark"
} | Sort-Object -Property "Date" | ForEach-Object {
  [PSCustomObject]@{
    "Date"        = $PSItem.Date.ToString("dd-MM-yyyy")
    "Description" = $PSItem.Description
    "Amount"      = $PSItem.Amount
    "Document No" = $documentNo++
  }
}

Set-Content $docnofile $documentNo

$filename = "bank_udtog_{0}.csv" -f $data[0].Date

$saveCsvFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveCsvFileDialog.InitialDirectory = "$HOME\Downloads"
$saveCsvFileDialog.Filter = "CSV (*.csv)|*.csv"
$saveCsvFileDialog.FileName = $filename
if ($saveCsvFileDialog.ShowDialog() -ne "OK") {
  return
}
$data | Export-Csv -Path $saveCsvFileDialog.FileName -Encoding UTF8 -Delimiter ";" -NoTypeInformation


