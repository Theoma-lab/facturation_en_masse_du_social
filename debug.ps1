function Get-Delimiter($path) {
    $line = Get-Content $path -TotalCount 1
    if (($line -split ';').Count -gt ($line -split ',').Count) { return ';' } else { return ',' }
}

$delimFact = Get-Delimiter "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\factures.csv"
$delimProd = Get-Delimiter "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\produits.csv"

$factures = Import-Csv "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\factures.csv" -Delimiter $delimFact
$produits = Import-Csv "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\produits.csv" -Delimiter $delimProd

$out = @{
    FacturesHeaders = ($factures[0].PSObject.Properties | ForEach-Object { $_.Name })
    ProduitsHeaders = ($produits[0].PSObject.Properties | ForEach-Object { $_.Name })
    FacturesFirst = $factures[0]
    ProduitsFirst = $produits[0]
}

$out | ConvertTo-Json -Depth 3 | Out-File -FilePath "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\debug.json" -Encoding utf8
