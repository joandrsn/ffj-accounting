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

$documentNo = 1
$data = $data | Where-Object {
  $PSItem.Description -notlike "DKFLX*" -and
  $PSItem.Description -ne "FI89136954,INDB.KA71" -and
  $PSItem.Description -notlike "Nets * 8650061"
} | Sort-Object -Property "Date" | ForEach-Object {
  [PSCustomObject]@{
    "Date"        = $PSItem.Date.ToString("dd-MM-yyyy")
    "Description" = $PSItem.Description
    "Amount"      = $PSItem.Amount
    "Document No" = $documentNo++
  }
}

$filename = "bank_udtog_{0}.csv" -f $data[0].Date

$saveCsvFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveCsvFileDialog.InitialDirectory = "$HOME\Downloads"
$saveCsvFileDialog.Filter = "CSV (*.csv)|*.csv"
$saveCsvFileDialog.FileName = $filename
if ($saveCsvFileDialog.ShowDialog() -ne "OK") {
  return
}
$data | Export-Csv -Path $saveCsvFileDialog.FileName -Encoding UTF8 -Delimiter ";" -NoTypeInformation


