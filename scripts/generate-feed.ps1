$ErrorActionPreference = "Stop"
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent $scriptDir
$feedUrl = "https://www.inmediaclean.sk/export-full-products-fk7V9zgku9.xml"
$feedLocal = Join-Path $scriptDir "feed.xml"
$outputXml = Join-Path $repoRoot "shoptet-import-produkty.xml"

Write-Host "Downloading supplier feed from $feedUrl ..."
Invoke-WebRequest -Uri $feedUrl -OutFile $feedLocal -UseBasicParsing

# --- load category map ---
$catMap = @{}
Get-Content (Join-Path $scriptDir "category-map.tsv") -Encoding UTF8 | ForEach-Object {
    if ($_ -match "^\s*$") { return }
    $parts = $_ -split "`t", 2
    if ($parts.Count -eq 2) { $catMap[$parts[0]] = $parts[1] }
}

# --- load brand list ---
$brands = Get-Content (Join-Path $scriptDir "brands.txt") -Encoding UTF8 | Where-Object { $_.Trim() -ne "" }

function Get-Manufacturer($title) {
    $tokens = $title -split '\s+' | Select-Object -First 4
    foreach ($t in $tokens) {
        $tClean = $t.Trim(',', '.', ':' )
        if ($tClean -ieq "HP") { continue }
        foreach ($b in $brands) {
            if ($tClean -ieq $b) { return $b }
        }
    }
    return ""
}

function Get-DefaultCategory($codes) {
    $withSub = $null
    $bare = $null
    foreach ($c in $codes) {
        if ($catMap.ContainsKey($c)) {
            $val = $catMap[$c]
            if ($val -like "* > *") { if (-not $withSub) { $withSub = $val } }
            else { if (-not $bare) { $bare = $val } }
        }
    }
    if ($withSub) { return $withSub }
    if ($bare) { return $bare }
    return $null
}

# --- load keyword -> category-path map ---
$kwMap = @{}
Get-Content (Join-Path $scriptDir "keyword-map.tsv") -Encoding UTF8 | ForEach-Object {
    if ($_ -match "^\s*$") { return }
    $parts = $_ -split "`t", 2
    if ($parts.Count -eq 2) { $kwMap[$parts[0]] = $parts[1] }
}

$keywordRules = [ordered]@{
    'toaletn.{0,2}\s*papier|toal\.?\s*papier'    = 'KW_TOALETNY_PAPIER'
    'vlhcen|vlhčen'                               = 'KW_VLHCENE_UTIERKY'
    'obrusk|obrúsk|utierk.*rol|papierov. utierk' = 'KW_PAPIEROVE_UTIERKY'
    'alobal|folia|fólia|na pecenie|na pečenie'   = 'KW_FOLIA'
    '\bmop\b|mopy|mopov'                          = 'KW_MOP'
    'rukavic'                                     = 'KW_RUKAVICE'
    'hubk|drotenk|drôtenk|spong|špong'           = 'KW_HUBKY'
    'metl|zmeta|zmetá|kefa|kefy|kief'             = 'KW_METLY'
    'vedro|vedra|vedrá|zmykac|žmýkač|vozik|vozík' = 'KW_VEDRA'
    'stierk|stahovac|sťahovač|skrabk|škrabk'      = 'KW_STIERKY'
    'zasobnik|zásobník|davkovac|dávkovač'         = 'KW_ZASOBNIKY'
    'mikrovlakn|mikrovlákn|handr|utierk'          = 'KW_UTIERKY_MIKRO'
    '\bwc\b|pisoar|pisoár|zachod|záchod'          = 'KW_WC'
    'dezinfek|bielidl|\bsavo\b'                   = 'KW_DEZINFEKCIA'
    'odmast|odmasť|vodneho kamen|vodného kameň|hrdz' = 'KW_ODMASTOVACE'
    'okna|okná|okien|sklo|zrkadl'                 = 'KW_OKNA'
    'podlah'                                      = 'KW_PODLAHY'
    'abraz|cistiac. pasta|čistiac. pasta|prasok|prášok' = 'KW_ABRAZIVA'
    'osviezovac|osviežovač|difuzer|difuzér'       = 'KW_OSVIEZOVACE'
    'repelent|hmyz|komar|komár|proti much'        = 'KW_HMYZ'
    'riad|umyvanie riadu|umývanie riadu'          = 'KW_RIAD'
    'univerzaln. cistic|univerzáln. čistič'       = 'KW_UNIVERZALNE'
    'praci|prací|pracie|na pranie|gel na pranie|gél na pranie' = 'KW_PRACIE'
    'aviva|avivaž|škvŕn|skvrn'                    = 'KW_AVIVAZ'
    'menu box|jednorazov. obal|misk'              = 'KW_MENUBOX'
    'vrec|sack|sáčk|taska|taška'                  = 'KW_VRECIA'
    'dezodorant|antiperspirant'                   = 'KW_DEZODORANT'
    'sampon|šampón|vlas'                          = 'KW_VLASY'
    'zubn. pasta|zubnu kefk|zubn[yý]'             = 'KW_ZUBNA'
    'holiac|holenie|ziletk|žiletk'                = 'KW_HOLENIE'
    'mydlo|mydla'                                 = 'KW_MYDLO'
    'sprchov. gel|sprchov. gél|telov. mliek|krem na ruky|krém na ruky' = 'KW_TELO'
    'vlozk|vložk|tampon|tampón'                   = 'KW_DAMSKA_HYGIENA'
    'plet\b|seru\b|séru\b'                        = 'KW_PLET'
}

$diacriticMap = @{
    [char]0x00E1='a'; [char]0x00E4='a'; [char]0x010D='c'; [char]0x010F='d'
    [char]0x00E9='e'; [char]0x00ED='i'; [char]0x013E='l'; [char]0x013A='l'
    [char]0x0148='n'; [char]0x00F3='o'; [char]0x00F4='o'; [char]0x0155='r'
    [char]0x0161='s'; [char]0x0165='t'; [char]0x00FA='u'; [char]0x00FD='y'
    [char]0x017E='z'
    [char]0x00C1='A'; [char]0x00C4='A'; [char]0x010C='C'; [char]0x010E='D'
    [char]0x00C9='E'; [char]0x00CD='I'; [char]0x013D='L'; [char]0x0139='L'
    [char]0x0147='N'; [char]0x00D3='O'; [char]0x00D4='O'; [char]0x0154='R'
    [char]0x0160='S'; [char]0x0164='T'; [char]0x00DA='U'; [char]0x00DD='Y'
    [char]0x017D='Z'
}
function Remove-Diacritics($s) {
    if (-not $s) { return $s }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $s.ToCharArray()) {
        if ($diacriticMap.ContainsKey($ch)) { [void]$sb.Append($diacriticMap[$ch]) }
        else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

# extra ASCII-only stems added after auditing the "Ostatné" residual list
$keywordRulesExtra = [ordered]@{
    'pletov|vrask|micelarn|obocie'          = 'KW_PLET'
    'kondicioner|\bmaska\b'                 = 'KW_VLASY'
    'zubn'                                   = 'KW_ZUBNA'
    'corega|ustna voda'                      = 'KW_ZUBNA'
    'holen'                                   = 'KW_HOLENIE'
    'umyvack|sol do umyvac'                 = 'KW_RIAD'
    'glade|\bwick\b|sviecka|vonna'          = 'KW_OSVIEZOVACE'
    '\bbiolit\b|\btrap\b'                   = 'KW_HMYZ'
    '\bcoccolino\b|\bbeckmann\b|\blenor\b'  = 'KW_AVIVAZ'
    '\blibresse\b'                          = 'KW_DAMSKA_HYGIENA'
    'umyvacia pena|mliek. opal|balzam|\bkrem\b|\bmast\b' = 'KW_TELO'
    'vodn. kamen|odvapnovac'                = 'KW_ODMASTOVACE'
    '\bdeo\b|\bstr8\b|\bezo\b'              = 'KW_DEZODORANT'
    'tuzidlo|wellaflex|melir'               = 'KW_VLASY'
    'cistic|cistiac|prostriedok na povrchy' = 'KW_UNIVERZALNE'
    'mys\b|potkan|\bmol\b|molam|osiam|srsn|mucholapk|na muchy|proti muc|hmy\b|postipan|\braid\b|odparovac' = 'KW_HMYZ'
    'naplast|kompres steril|tehotensky test' = 'KW_TELO'
    'odlicov|tonikum|pod oci|tvarov. seru|vankusiky pod' = 'KW_PLET'
    'pena do kupela|umyvacia emulzia|gel na nohy|francovka|intim gel' = 'KW_TELO'
    'dratenk'                                = 'KW_HUBKY'
    'tekuty piesok'                          = 'KW_ABRAZIVA'
    'vona do pradla'                         = 'KW_AVIVAZ'
    'gillette|\brazor\b'                     = 'KW_HOLENIE'
}

function Get-KeywordCategory($title) {
    if (-not $title) { return $kwMap['KW_OSTATNE'] }
    $plain = Remove-Diacritics $title
    foreach ($k in $keywordRules.Keys) {
        if ($plain -match $k) { return $kwMap[$keywordRules[$k]] }
    }
    foreach ($k in $keywordRulesExtra.Keys) {
        if ($plain -match $k) { return $kwMap[$keywordRulesExtra[$k]] }
    }
    return $kwMap['KW_OSTATNE']
}

Write-Host "Loading XML..."
[xml]$xml = Get-Content $feedLocal -Encoding UTF8
$products = $xml.PRODUCTS.PRODUCT
Write-Host "Products loaded: $($products.Count)"

$exportedCount = 0
$excludedCount = 0

$xmlSettings = New-Object System.Xml.XmlWriterSettings
$xmlSettings.Encoding = New-Object System.Text.UTF8Encoding($false)
$xmlSettings.Indent = $true
$xmlSettings.IndentChars = "  "
$xw = [System.Xml.XmlWriter]::Create($outputXml, $xmlSettings)
$xw.WriteStartDocument()
$xw.WriteStartElement("SHOP")

foreach ($p in $products) {

    $desc = $p.DESCRIPTIONS.DESCRIPTION
    if ($desc -is [System.Array]) { $desc = $desc[0] }
    $title = if ($desc) { $desc.TITLE } else { "" }

    $priceNode = $p.SelectSingleNode("PRICES/PRICE/PRICELISTS/PRICELIST/PRICE_WITH_VAT")
    $price = 0.0
    if ($priceNode -and $priceNode.InnerText) {
        [void][double]::TryParse($priceNode.InnerText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$price)
    }

    if ($price -le 0) {
        $excludedCount++
        continue
    }

    $code = $p.CODE
    $ean = $p.EAN
    $unit = $p.UNIT
    $stockRaw = 0
    [void][int]::TryParse($p.STOCK, [ref]$stockRaw)
    $stock = if ($stockRaw -gt 0) { $stockRaw } else { 0 }
    $availability = $p.AVAILABILITY

    $priceListCurrency = $p.SelectSingleNode("PRICES/PRICE/CURRENCY")
    $currency = if ($priceListCurrency) { $priceListCurrency.InnerText } else { "EUR" }

    $shortDesc = if ($desc -and $desc.SHORT_DESCRIPTION) { $desc.SHORT_DESCRIPTION.InnerText } else { "" }
    $longDesc  = if ($desc -and $desc.LONG_DESCRIPTION)  { $desc.LONG_DESCRIPTION.InnerText }  else { "" }

    $imgNodes = $p.SelectNodes("IMAGES/IMAGE/URL")
    $images = @()
    foreach ($n in $imgNodes) { $images += $n.InnerText }

    $catNodes = $p.SelectNodes("CATEGORIES/CATEGORY/CODE")
    $catCodes = @()
    foreach ($n in $catNodes) { if ($n.InnerText -ne "K00208") { $catCodes += $n.InnerText } }
    $defaultCategory = Get-DefaultCategory $catCodes
    if (-not $defaultCategory) { $defaultCategory = Get-KeywordCategory $title }

    $manufacturer = Get-Manufacturer $title

    $paramNodes = $p.SelectNodes("PARAMETERS/PARAMETER")
    $textProperty = ""
    if ($paramNodes -and $paramNodes.Count -gt 0) {
        $pn = $paramNodes[0]
        $pname = $pn.SelectSingleNode("NAME").InnerText
        $pval  = $pn.SelectSingleNode("VALUE").InnerText
        $textProperty = "$pname;$pval"
    }

    $relNodes = $p.SelectNodes("RELATED_PRODUCTS/CODE")
    $rel = @()
    foreach ($n in $relNodes) { $rel += $n.InnerText; if ($rel.Count -ge 5) { break } }

    $altNodes = $p.SelectNodes("ALTERNATIVE_PRODUCTS/CODE")
    $alt = @()
    foreach ($n in $altNodes) { $alt += $n.InnerText; if ($alt.Count -ge 5) { break } }

    $newFlagActive = ""
    $labelNodes = $p.SelectNodes("LABELS/LABEL")
    foreach ($ln in $labelNodes) {
        if ($ln.NAME -eq "Novinka") { $newFlagActive = $ln.ACTIVE_YN; break }
    }

    $xw.WriteStartElement("SHOPITEM")
    $xw.WriteElementString("NAME", $title)
    if ($shortDesc) { $xw.WriteElementString("SHORT_DESCRIPTION", $shortDesc) }
    if ($longDesc)  { $xw.WriteElementString("DESCRIPTION", $longDesc) }
    if ($manufacturer) { $xw.WriteElementString("MANUFACTURER", $manufacturer) }
    $xw.WriteElementString("ITEM_TYPE", "product")
    if ($unit) { $xw.WriteElementString("UNIT", $unit) }
    $xw.WriteElementString("CODE", $code)
    if ($ean) { $xw.WriteElementString("EAN", $ean) }

    $xw.WriteStartElement("CATEGORIES")
    $xw.WriteElementString("CATEGORY", $defaultCategory)
    $xw.WriteEndElement()

    if ($images.Count -gt 0) {
        $xw.WriteStartElement("IMAGES")
        foreach ($imgUrl in $images) { $xw.WriteElementString("IMAGE", $imgUrl) }
        $xw.WriteEndElement()
    }

    if ($textProperty) {
        $tpParts = $textProperty -split ";", 2
        $xw.WriteStartElement("TEXT_PROPERTIES")
        $xw.WriteStartElement("TEXT_PROPERTY")
        $xw.WriteElementString("NAME", $tpParts[0])
        $xw.WriteElementString("VALUE", $tpParts[1])
        $xw.WriteEndElement()
        $xw.WriteEndElement()
    }

    if ($rel.Count -gt 0) {
        $xw.WriteStartElement("RELATED_PRODUCTS")
        foreach ($rc in $rel) { $xw.WriteElementString("CODE", $rc) }
        $xw.WriteEndElement()
    }
    if ($alt.Count -gt 0) {
        $xw.WriteStartElement("ALTERNATIVE_PRODUCTS")
        foreach ($ac in $alt) { $xw.WriteElementString("CODE", $ac) }
        $xw.WriteEndElement()
    }

    $xw.WriteStartElement("FLAGS")
    $xw.WriteStartElement("FLAG")
    $xw.WriteElementString("CODE", "new")
    $xw.WriteElementString("ACTIVE", $(if ($newFlagActive -eq "1") { "1" } else { "0" }))
    $xw.WriteEndElement()
    $xw.WriteEndElement()

    $xw.WriteElementString("VISIBILITY", "visible")
    $xw.WriteElementString("CURRENCY", $currency)
    $xw.WriteElementString("PRICE", ([Math]::Round($price, 2)).ToString("F2", [System.Globalization.CultureInfo]::InvariantCulture))

    $xw.WriteStartElement("STOCK")
    $xw.WriteStartElement("WAREHOUSES")
    $xw.WriteStartElement("WAREHOUSE")
    $xw.WriteElementString("NAME", $kwMap['WAREHOUSE_NAME'])
    $xw.WriteElementString("VALUE", $stock.ToString([System.Globalization.CultureInfo]::InvariantCulture))
    $xw.WriteEndElement()
    $xw.WriteEndElement()
    $xw.WriteEndElement()

    if ($availability) { $xw.WriteElementString("AVAILABILITY", $availability) }
    $xw.WriteEndElement()

    $exportedCount++
}

$xw.WriteEndElement()
$xw.WriteEndDocument()
$xw.Flush()
$xw.Close()

Remove-Item $feedLocal -ErrorAction SilentlyContinue

Write-Host "Exported: $exportedCount  Excluded (zero price): $excludedCount"
Write-Host "Output: $outputXml"
