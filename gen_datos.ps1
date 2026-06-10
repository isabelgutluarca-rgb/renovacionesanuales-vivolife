# ============================================================
# gen_datos.ps1 — genera datos.js para el dashboard VivoLife
# Lineamiento v6: granularidad = 1 fila por usuario activo
# Exclusion: company_agreement_id (col H=8) no vacío
# Status filtro: active + overdue
# Titular = fila donde col_41(vivo_id) = col_19(titular_personal_id) o rol=TITULAR
# Dedup: pares únicos (id, vivo_id)
# ============================================================

$xlPath  = "C:\Users\Maria Gutierrez\Downloads\suscripciones_anuales_v6_con adicionales.xlsx"
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

# ── PASO 1: recolectar todas las filas válidas por suscripción ─
# Estructura: $raw[id] = @{ titular=@{}, deps=@{vid=>@{}} }
$raw  = [System.Collections.Generic.Dictionary[string,hashtable]]::new()
$seen = [System.Collections.Generic.HashSet[string]]::new()  # "id_vid" para dedup

for($i=2; $i -le $rows; $i++){
    $id = Cell $i 1
    if(-not $id){ continue }

    # Filtro: solo active / overdue
    $status = (Cell $i 6).ToLower()
    if($status -ne "active" -and $status -ne "overdue"){ continue }

    # Excluir empresariales — SOLO col 8 (company_agreement_id)
    if(Cell $i 8){ continue }

    $titVid = Cell $i 19   # titular_personal_id (col S)
    $memVid = Cell $i 41   # usuarios_activos.vivo_id (col AO)
    $rol    = (Cell $i 45).ToUpper()   # usuarios_activos.rol (col AS)

    # Deduplicar por (id, vivo_id)
    $dkey = "${id}_${memVid}"
    if(-not $seen.Add($dkey)){ continue }

    # ¿Es fila del titular?
    $esTit = ($memVid -ne "" -and $memVid -eq $titVid) -or ($rol -eq "TITULAR")

    if(-not $raw.ContainsKey($id)){
        $raw[$id] = @{ titular=$null; deps=[System.Collections.Generic.List[hashtable]]::new() }
    }

    if($esTit){
        if($raw[$id].titular -eq $null){
            # Guardar datos del titular desde esta fila
            $raw[$id].titular = @{
                id=$id; status=$status
                titVid=$titVid
                planId=(Cell $i 7)
                vid=(Cell $i 19)          # titular vivo_id
                nom="$(Cell $i 21) $(Cell $i 22)".Trim()
                tel=(Cell $i 23); email=(Cell $i 24)
                expDate=(Cell $i 17)      # expiration_date
                valPlan=(Cell $i 27); valAdd=(Cell $i 28)
                token=(Cell $i 29)        # tiene_tarjeta_tokenizada
                texpRaw=(Cell $i 31)      # expiracion_tarjeta_default
                upago=(Cell $i 36)        # ultimo_pago_estado
                monto=(Cell $i 47)        # monto_oportunidad (AU)
            }
        }
    } else {
        # Dependiente
        if($memVid -and $memVid -ne $titVid){
            $raw[$id].deps.Add(@{
                vid=$memVid
                c="$(Cell $i 42) $(Cell $i 43)".Trim()
                t=(Cell $i 44)
            })
        }
    }
}

Write-Host "Suscripciones raw: $($raw.Count)"

# ── PASO 2: para suscripciones sin fila titular explícita,
#    usar primera fila dependiente como proxy (caso raro) ─────
# (Ya cubierto: la primera vez que se ve el id y NO es titular,
#  raw[$id].titular queda null — se descarta al serializar)

# ── PASO 3: serializar ───────────────────────────────────────
$sinList = [System.Collections.Generic.List[string]]::new()
$conList = [System.Collections.Generic.List[string]]::new()
$cM=0; $cP=0; $cA=0; $cS=0; $cC=0; $cT=0; $cSkip=0

foreach($kv in $raw.GetEnumerator()){
    $r = $kv.Value.titular
    if(-not $r){ $cSkip++; continue }   # sin fila titular → omitir

    $vDate    = ParseDate $r.expDate
    $texpDate = ParseDate $r.texpRaw
    $diasExp  = if($vDate){ [math]::Ceiling(($vDate-$hoy).TotalDays) } else { 9999 }
    $texpd    = if($texpDate){ [math]::Ceiling(($texpDate-$hoy).TotalDays) } else { 9999 }
    $vStr     = if($vDate){ $vDate.ToString("dd/MM/yyyy") } else { "" }
    $texpStr  = if($texpDate){ $texpDate.ToString("dd/MM/yyyy") } else { "" }

    # Estado basado en FECHA (no en status Excel)
    $s = if($diasExp -lt 0){ "mora" } elseif($diasExp -le 60){ "proximo" } else { "activo" }

    $planId  = $r.planId
    $valPlan = [double]$r.valPlan
    $valAdd  = [double]$r.valAdd
    $dep     = $kv.Value.deps.Count
    $tok     = if($r.token -eq "1"){ 1 } else { 0 }

    # vtot = Monto de oportunidad (AU) si existe, si no calcular
    $montRaw = $r.monto
    $vtot    = if($montRaw -and [double]$montRaw -gt 0){ [double]$montRaw } else { $valPlan + ($valAdd * $dep) }
    $com     = [math]::Round($valPlan * (GetCom $planId), 2)

    # Dependientes JSON
    $dj = "["; $f = $true
    foreach($d in $kv.Value.deps){
        if(-not $f){ $dj += "," }
        $dj += "{vid:`"$(EscJs $d.vid)`",c:`"$(EscJs $d.c)`",t:`"$(EscJs $d.t)`"}"
        $f = $false
    }
    $dj += "]"

    $rec = "{id:$($r.id),vid:`"$(EscJs $r.vid)`",c:`"$(EscJs $r.nom)`",t:`"$(EscJs $r.tel)`",e:`"$(EscJs $r.email)`",p:`"$(EscJs (GetPlan $planId))`",v:`"$vStr`",s:`"$s`",diasExp:$diasExp,m:$valPlan,vadd:$valAdd,dep:$dep,vtot:$vtot,com:$com,u:`"$(EscJs $r.upago)`",texp:`"$texpStr`",texpd:$texpd,deps:$dj}"

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

Write-Host "Sin titular (omitidos): $cSkip"

# Ordenar SIN TOKEN: mora primero (diasExp más negativo), luego proximo, activo
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
