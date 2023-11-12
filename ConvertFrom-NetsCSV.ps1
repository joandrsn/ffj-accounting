Add-Type -AssemblyName System.Windows.Forms

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = "$HOME\Downloads"
$openFileDialog.Filter = "CSV (*.csv)|*.csv"
if ($openFileDialog.ShowDialog() -ne "OK") {
    return
}
$culture = [cultureinfo]::GetCultureInfo('da-DK')

$data = Get-Content $openFileDialog.FileName | ForEach-Object {
    $line = $_.Split(';')
    $section = $line[0].Trim('"')
    # Skip until we reach the section we want
    if ($section -ne "Settlement Information") {
        return
    }
    # Skip the header
    if ($line[1].Trim('"') -eq "Date") {
        return
    }
    [PSCustomObject]@{
        "Date"                    = $line[1].Trim('"')
        "Acquirer"                = $line[2].Trim('"')
        "Transaction Amount"      = [System.Convert]::ToDecimal($line[3], $culture)
        "Subscription Fees"       = [System.Convert]::ToDecimal($line[4], $culture)
        "Service Fees"            = [System.Convert]::ToDecimal($line[5], $culture)
        "Chargeback / Adjustment" = [System.Convert]::ToDecimal($line[6], $culture)
        "Settled"                 = [System.Convert]::ToDecimal($line[7], $culture)
        "Currency"                = $line[8].Trim('"')
        "Bank Account"            = $line[9].Trim('"')
        "Payment Reference"       = $line[10].Trim('"')
        "Number Of Transactions"  = $line[11].Trim('"')
        "Number Of Batches"       = $line[12].Trim('"')
        "Chargeback Amount"       = $line[13].Trim('"')
        "Adjustment Amount"       = $line[14].Trim('"')
    }
}
$data = $data | Where-Object { $PSItem."Service Fees" -lt 0 -or $PSItem."Subscription Fees" -lt 0 }
[array]::Reverse($data)

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

$finaldata = $data | ForEach-Object {
    $description = $PSItem."Payment Reference"
    if ($PSItem.Acquirer -eq "International Cards") {
        $description = $description + " / Internationale kort"
    }
    if ($PSItem.Acquirer -eq "Dankort" -and $description -match "^\d+$") {
        $description = "Dankort - abonnement"
    }
    [PSCustomObject]@{
        "Date"        = $PSItem."Date"
        "Amount"      = $PSItem."Service Fees" + $PSItem."Subscription Fees"
        "Description" = $description
        "Document No" = $documentno
    }
    $documentno += 1
}

Set-Content $docnofile $documentNo

$filename = "kreditkort-gebyrer_{0}.csv" -f $finaldata[0].Date

$saveCsvFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveCsvFileDialog.InitialDirectory = "$HOME\Downloads"
$saveCsvFileDialog.Filter = "CSV (*.csv)|*.csv"
$saveCsvFileDialog.FileName = $filename
if ($saveCsvFileDialog.ShowDialog() -ne "OK") {
    return
}
$finaldata | Export-Csv -Path $saveCsvFileDialog.FileName -Encoding UTF8 -Delimiter ";" -NoTypeInformation


