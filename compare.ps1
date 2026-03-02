$lineF = Get-Content "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\factures.csv" -TotalCount 1
$delimFact = if ($lineF -match ';') { ';' } else { ',' }

$lineP = Get-Content "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\produits.csv" -TotalCount 1
$delimProd = if ($lineP -match ';') { ';' } else { ',' }

$factures = Import-Csv "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\factures.csv" -Delimiter $delimFact
$produits = Import-Csv "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\produits.csv" -Delimiter $delimProd

$prodDict = @{}
foreach ($p in $produits) {
  $props = $p.PSObject.Properties | Select-Object -ExpandProperty Value
  if ($props.Count -gt 5) {
      $name = [string]$props[1]
      $name = $name.Trim()
      
      $priceStr = [string]$props[5]
      $priceStr = $priceStr -replace ',', '.' -replace ' ', '' -replace '\$|\€|\£', ''
      $price = if ($priceStr -as [double]) { [double]$priceStr } else { 0.0 }
      
      if ($name) { $prodDict[$name] = [math]::Round($price, 2) }
  }
}

$results = @()
foreach ($f in $factures) {
  $props = $f.PSObject.Properties | Select-Object -ExpandProperty Value
  if ($props.Count -gt 17) {
      $client = ([string]$props[4]).Trim()
      $prodName = ([string]$props[9]).Trim()
      
      if (-not $client -or -not $prodName) { continue }
      
      $qtyStr = [string]$props[13] -replace ',', '.' -replace ' ', ''
      $qty = if ($qtyStr -as [double]) { [double]$qtyStr } else { 1.0 }
      if ($qty -eq 0) { $qty = 1.0 }

      $montantStr = [string]$props[17] -replace ',', '.' -replace ' ', '' -replace '\$|\€|\£', ''
      $montant = if ($montantStr -as [double]) { [double]$montantStr } else { 0.0 }
      $unitPrice = [math]::Round($montant / $qty, 2)
      
      if ($prodDict.ContainsKey($prodName)) {
        $basePrice = $prodDict[$prodName]
        $isCustom = ($unitPrice -ne $basePrice)
        $results += [PSCustomObject] @{
          Client = $client
          Produit = $prodName
          PrixFacture = $unitPrice
          PrixBase = $basePrice
          IsCustom = $isCustom
        }
      } else {
        $results += [PSCustomObject] @{
          Client = $client
          Produit = $prodName
          PrixFacture = $unitPrice
          PrixBase = $null
          IsCustom = "Inconnu"
        }
      }
  }
}

$grouped = $results | Group-Object Client
$report = @{}
foreach ($g in $grouped) {
  $customCount = ($g.Group | Where-Object { $_.IsCustom -eq $true }).Count
  $baseCount = ($g.Group | Where-Object { $_.IsCustom -eq $false }).Count
  $unknownCount = ($g.Group | Where-Object { $_.IsCustom -is [string] }).Count
  $report[$g.Name] = @{
    TotalLignes = $g.Count
    PrixCustom = $customCount
    PrixBase = $baseCount
    ProduitsInconnus = $unknownCount
    Details = $g.Group
  }
}

$report | ConvertTo-Json -Depth 5 | Out-File -FilePath "c:\Users\melanie.v\.gemini\antigravity\scratch\facturation-du-social\report.json" -Encoding utf8
