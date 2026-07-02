# ============================================================
# gen_datos.ps1 — genera datos.js para el dashboard VivoLife
# Fuente: suscripciones_anuales.xlsx (40 cols, sin dependientes individuales)
# Exclusion: deleted_at (col 12) no vacío, company_agreement_id (col 8) no vacío
# Status filtro: active + overdue (excluye presale, reactivation, cancelados)
# Una fila por suscripción; dep = total_adicionales (col 26)
# ============================================================

$xlPath  = "C:\Users\Maria Gutierrez\Downloads\suscripciones_anuales.xlsx"
$outPath = "C:\Users\Maria Gutierrez\renovaciones-vivolife\datos.js"
$hoy     = [datetime]::Today

Write-Host "Abriendo Excel..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($xlPath)
$ws = $wb.Sheets.Item(1)

Write-Host "Leyendo datos en bloque..."
$rng  = $ws.UsedRange
$data = $rng.Value2
$rows = $rng.Rows.Count
$cols = $rng.Columns.Count
Write-Host "Filas: $rows  Cols: $cols"

$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "Excel cerrado. Procesando..."

# ── helpers ──────────────────────────────────────────────────
function Cell($r,$c){
    if($c -le $script:cols){
        $v = $script:data[$r,$c]
        if($v -eq $null){ "" } else { "$v".Trim() }
    } else { "" }
}

function ParseDate($v){
    if(-not $v){ return $null }
    try{
        if($v -is [double]){ return [datetime]::FromOADate($v) }
        $d = 0.0
        if([double]::TryParse("$v",[ref]$d)){ return [datetime]::FromOADate($d) }
        return [datetime]::Parse("$v")
    } catch { return $null }
}

function EscJs($s){
    $s = "$s" -replace '\\','\\' -replace '"','\"' -replace "`r`n",' ' -replace "`n",' ' -replace "`r",' '
    return $s
}

function GetPlan($id){
    switch($id){
        "4"  { return "Vivolife Silver" }
        "16" { return "Vivolife Plus" }
        default { return "Vivolife" }
    }
}

function GetCom($id){ if($id -eq "16"){0.06} else {0.05} }

# ── Serializar ───────────────────────────────────────────────
$sinList = [System.Collections.Generic.List[string]]::new()
$conList = [System.Collections.Generic.List[string]]::new()
$cM=0; $cP=0; $cA=0; $cS=0; $cC=0; $cT=0; $cSkip=0

for($i=2; $i -le $rows; $i++){
    $id = Cell $i 1
    if(-not $id){ continue }

    # Excluir cancelados (deleted_at no vacío)
    if(Cell $i 12){ $cSkip++; continue }

    # Filtro status: solo active / overdue
    $status = (Cell $i 6).ToLower()
    if($status -ne "active" -and $status -ne "overdue"){ continue }

    # Excluir empresariales
    if(Cell $i 8){ continue }

    $planId  = Cell $i 7
    $vid     = Cell $i 19   # titular_personal_id = vivo_id titular
    $nom     = "$(Cell $i 21) $(Cell $i 22)".Trim()
    $tel     = Cell $i 23
    $email   = Cell $i 24
    $dep     = [int]("$(Cell $i 26)" -replace '[^0-9]','')   # total_adicionales
    $valPlan = 0.0; [double]::TryParse((Cell $i 27),[ref]$valPlan) | Out-Null
    $valAdd  = 0.0; [double]::TryParse((Cell $i 28),[ref]$valAdd)  | Out-Null
    $tok     = if((Cell $i 29) -eq "1"){ 1 } else { 0 }

    $vDate    = ParseDate (Cell $i 17)   # expiration_date
    $texpDate = ParseDate (Cell $i 31)   # expiracion_tarjeta_default
    $diasExp  = if($vDate){ [math]::Ceiling(($vDate-$hoy).TotalDays) } else { 9999 }
    $texpd    = if($texpDate){ [math]::Ceiling(($texpDate-$hoy).TotalDays) } else { 9999 }
    $vStr     = if($vDate){ $vDate.ToString("dd/MM/yyyy") } else { "" }
    $texpStr  = if($texpDate){ $texpDate.ToString("dd/MM/yyyy") } else { "" }

    # Estado basado en FECHA de vencimiento
    $s = if($diasExp -lt 0){ "mora" } elseif($diasExp -le 60){ "proximo" } else { "activo" }

    $vtot = $valPlan + ($valAdd * $dep)
    $com  = [math]::Round($valPlan * (GetCom $planId), 2)
    $upago = Cell $i 36

    # Sin dependientes individuales en este archivo
    $rec = "{id:$id,vid:`"$(EscJs $vid)`",c:`"$(EscJs $nom)`",t:`"$(EscJs $tel)`",e:`"$(EscJs $email)`",p:`"$(EscJs (GetPlan $planId))`",v:`"$vStr`",s:`"$s`",diasExp:$diasExp,m:$valPlan,vadd:$valAdd,dep:$dep,vtot:$vtot,com:$com,u:`"$(EscJs $upago)`",texp:`"$texpStr`",texpd:$texpd,deps:[]}"

    if($s -eq "mora"){ $cM++ } elseif($s -eq "proximo"){ $cP++ } else { $cA++ }

    if($tok -eq 0){
        $cS++
        $sinList.Add($rec)
    } else {
        $cC++
        $conList.Add($rec)
        if($texpd -le 45){ $cT++ }
    }
}

Write-Host "Cancelados omitidos: $cSkip"

# Ordenar SIN TOKEN: mora primero, luego proximo, activo
$sinSorted = $sinList | Sort-Object {
    $l = $_
    $sm = if($l -match 's:"mora"'){ 0 } elseif($l -match 's:"proximo"'){ 1 } else { 2 }
    $dm = if($l -match 'diasExp:(-?\d+)'){ [int]$Matches[1] + 99999 } else { 999999 }
    $sm * 1000000 + $dm
}

$conSorted = $conList | Sort-Object {
    $l = $_
    if($l -match 'diasExp:(-?\d+)'){ [int]$Matches[1] + 99999 } else { 999999 }
}

$content = "window.SIN_TOKEN=[" + ($sinSorted -join ',') + "];`r`nwindow.CON_TOKEN=[" + ($conSorted -join ',') + "];"
[System.IO.File]::WriteAllText($outPath, $content, [System.Text.Encoding]::UTF8)

Write-Host "=== RESULTADO ==="
Write-Host "SIN TOKEN  : $cS"
Write-Host "CON TOKEN  : $cC"
Write-Host "Mora       : $cM"
Write-Host "Proximos   : $cP"
Write-Host "Activos    : $cA"
Write-Host "Texp<=45d  : $cT"
Write-Host "datos.js generado OK"
