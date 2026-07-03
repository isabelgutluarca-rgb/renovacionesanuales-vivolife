# ============================================================
# gen_datos.ps1 — genera datos.js para el dashboard VivoLife
# Fuente: suscripciones_anuales_adicionales_final.xlsx (48 cols)
# Col 8: plan (nombre), col 9: company_agreement_id
# Col 13: deleted_at, col 18: expiration_date
# Cols 42-46: usuarios activos (vivo_id, nombre, apellido, tel, rol)
# Col 48: Monto de oportunidad
# Exclusion: deleted_at (col 13) no vacío, company_agreement_id (col 9) no vacío
# Status filtro: active + overdue (excluye presale, reactivation)
# ============================================================

$xlPath  = "C:\Users\Maria Gutierrez\Downloads\suscripciones_anuales_adicionales_final.xlsx"
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

function GetCom($planNom){
    if($planNom -match "Plus"){ return 0.06 }
    return 0.05
}

# ── PASO 1: recolectar filas por suscripción ─────────────────
# Los datos de suscripción (cols 1-41) son iguales en todas las filas del mismo ID.
# Se toma la primera fila como base y se acumulan los deps de filas adicionales.
$raw     = [System.Collections.Generic.Dictionary[string,hashtable]]::new()
$seenDep = [System.Collections.Generic.HashSet[string]]::new()

for($i=2; $i -le $rows; $i++){
    $id = Cell $i 1
    if(-not $id){ continue }

    # Excluir cancelados (deleted_at col 13)
    if(Cell $i 13){ continue }

    # Filtro status: solo active / overdue
    $status = (Cell $i 6).ToLower()
    if($status -ne "active" -and $status -ne "overdue"){ continue }

    # Excluir empresariales (company_agreement_id col 9)
    if(Cell $i 9){ continue }

    # Primera vez que vemos este ID → guardar datos de suscripción
    if(-not $raw.ContainsKey($id)){
        $raw[$id] = @{
            id=$id; status=$status
            planNom=(Cell $i 8)
            vid=(Cell $i 20)
            nom="$(Cell $i 22) $(Cell $i 23)".Trim()
            tel=(Cell $i 24); email=(Cell $i 25)
            expDate=(Cell $i 18)
            valPlan=(Cell $i 28); valAdd=(Cell $i 29)
            token=(Cell $i 30)
            texpRaw=(Cell $i 32)
            upago=(Cell $i 37)
            monto=(Cell $i 48)
            deps=[System.Collections.Generic.List[hashtable]]::new()
        }
    }

    # Acumular deps: filas donde el miembro NO es el titular
    $titVid = Cell $i 20
    $memVid = Cell $i 42
    $rol    = (Cell $i 46).ToUpper()
    $isDep  = ($memVid -ne "" -and $memVid -ne $titVid -and $rol -ne "TITULAR")
    if($isDep){
        $dkey = "${id}_${memVid}"
        if($seenDep.Add($dkey)){
            $raw[$id].deps.Add(@{
                vid=$memVid
                c="$(Cell $i 43) $(Cell $i 44)".Trim()
                t=(Cell $i 45)
            })
        }
    }
}

Write-Host "Suscripciones raw: $($raw.Count)"

# ── PASO 2: serializar ───────────────────────────────────────
$sinList = [System.Collections.Generic.List[string]]::new()
$conList = [System.Collections.Generic.List[string]]::new()
$cM=0; $cP=0; $cA=0; $cS=0; $cC=0; $cT=0; $cSkip=0

foreach($kv in $raw.GetEnumerator()){
    $r = $kv.Value

    $vDate    = ParseDate $r.expDate
    $texpDate = ParseDate $r.texpRaw
    $diasExp  = if($vDate){ [math]::Ceiling(($vDate-$hoy).TotalDays) } else { 9999 }
    $texpd    = if($texpDate){ [math]::Ceiling(($texpDate-$hoy).TotalDays) } else { 9999 }
    $vStr     = if($vDate){ $vDate.ToString("dd/MM/yyyy") } else { "" }
    $texpStr  = if($texpDate){ $texpDate.ToString("dd/MM/yyyy") } else { "" }

    # Estado basado en fecha de vencimiento (más preciso que el status del Excel)
    $s = if($diasExp -lt 0){ "mora" } elseif($diasExp -le 60){ "proximo" } else { "activo" }

    $planNom = if($r.planNom){ $r.planNom } else { "Vivolife" }
    $valPlan = 0.0; [double]::TryParse($r.valPlan,[ref]$valPlan) | Out-Null
    $valAdd  = 0.0; [double]::TryParse($r.valAdd,[ref]$valAdd)  | Out-Null
    $dep     = $r.deps.Count
    $tok     = if($r.token -eq "1"){ 1 } else { 0 }

    $montRaw = $r.monto
    $vtot    = if($montRaw -and [double]$montRaw -gt 0){ [double]$montRaw } else { $valPlan + ($valAdd * $dep) }
    $com     = [math]::Round($valPlan * (GetCom $planNom), 2)

    $dj = "["; $f = $true
    foreach($d in $r.deps){
        if(-not $f){ $dj += "," }
        $dj += "{vid:`"$(EscJs $d.vid)`",c:`"$(EscJs $d.c)`",t:`"$(EscJs $d.t)`"}"
        $f = $false
    }
    $dj += "]"

    $rec = "{id:$($r.id),vid:`"$(EscJs $r.vid)`",c:`"$(EscJs $r.nom)`",t:`"$(EscJs $r.tel)`",e:`"$(EscJs $r.email)`",p:`"$(EscJs $planNom)`",v:`"$vStr`",s:`"$s`",diasExp:$diasExp,m:$valPlan,vadd:$valAdd,dep:$dep,vtot:$vtot,com:$com,u:`"$(EscJs $r.upago)`",texp:`"$texpStr`",texpd:$texpd,deps:$dj}"

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

Write-Host "Omitidos: $cSkip"

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
