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
$data = $rng.Value2          # 2D array: $data[fila,col] — base 1
$rows = $rng.Rows.Count
$cols = $rng.Columns.Count
Write-Host "Filas: $rows  Cols: $cols"

$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "Excel cerrado. Procesando..."

# ── helpers ──────────────────────────────────────────────────
function Cell($r,$c){ if($c -le $script:cols){ $v=$script:data[$r,$c]; if($v -eq $null){""} else{"$v".Trim()} } else {""} }

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
    $s = "$s" -replace '"','\"' -replace "`r`n",' ' -replace "`n",' ' -replace "`r",' '
    return $s
}

function GetPlan($id){
    switch($id){ "4"{"Vivolife Silver"} "16"{"Vivolife Plus"} default{"Vivolife"} }
}

function GetCom($id){ if($id -eq "16"){0.06} else {0.05} }

# ── procesar filas ────────────────────────────────────────────
$subs = [System.Collections.Generic.Dictionary[string,hashtable]]::new()

for($i=2; $i -le $rows; $i++){
    $id     = Cell $i 1
    if(-not $id){ continue }

    $status  = (Cell $i 6).ToLower()
    if($status -ne "active" -and $status -ne "overdue"){ continue }

    # Excluir empresariales (company_agreement_id col8 o category col9)
    if((Cell $i 8) -or (Cell $i 9)){ continue }

    $rol    = (Cell $i 45).ToUpper()
    $memVid = Cell $i 41

    if(-not $subs.ContainsKey($id)){
        $vDate    = ParseDate (Cell $i 17)
        $texpDate = ParseDate (Cell $i 31)
        $diasExp  = if($vDate){ [math]::Ceiling(($vDate-$hoy).TotalDays) } else { 9999 }
        $texpd    = if($texpDate){ [math]::Ceiling(($texpDate-$hoy).TotalDays) } else { 9999 }
        $vStr     = if($vDate){ $vDate.ToString("dd/MM/yyyy") } else {""}
        $texpStr  = if($texpDate){ $texpDate.ToString("dd/MM/yyyy") } else {""}

        $planId   = Cell $i 7
        $valPlan  = [double](Cell $i 27)
        $valAdd   = [double](Cell $i 28)
        $monto    = Cell $i 47
        $token    = Cell $i 29
        $tok      = if($token -eq "0" -or $token -eq "" -or $token -eq "0"){0} else {1}

        # Estado POR FECHA (ignorar status Excel)
        $s = if($diasExp -lt 0){"mora"} elseif($diasExp -le 60){"proximo"} else {"activo"}

        $subs[$id] = @{
            id=$id; vid=(Cell $i 19)
            c="$(Cell $i 21) $(Cell $i 22)".Trim()
            t=(Cell $i 23); e=(Cell $i 24)
            p=(GetPlan $planId); v=$vStr; s=$s
            diasExp=$diasExp; m=$valPlan; vadd=$valAdd
            dep=0; vtot=0
            com=[math]::Round($valPlan*(GetCom $planId),2)
            u=(Cell $i 36); texp=$texpStr; texpd=$texpd; tok=$tok
            monto=$monto; titVid=(Cell $i 19)
            deps=[System.Collections.Generic.List[hashtable]]::new()
        }
    } else {
        # Adicional: agregar si Vivo ID diferente y no es TITULAR duplicado
        $rec = $subs[$id]
        if($memVid -and $memVid -ne $rec.titVid -and $rol -ne "TITULAR"){
            $dup=$false
            foreach($d in $rec.deps){ if($d.vid -eq $memVid){$dup=$true;break} }
            if(-not $dup){
                $rec.deps.Add(@{
                    vid=$memVid
                    c="$(Cell $i 42) $(Cell $i 43)".Trim()
                    t=(Cell $i 44)
                })
                $rec.dep++
            }
        }
    }
}

Write-Host "Suscripciones: $($subs.Count)"

# ── serializar ────────────────────────────────────────────────
$sinList=[System.Collections.Generic.List[string]]::new()
$conList=[System.Collections.Generic.List[string]]::new()
$cM=0;$cP=0;$cA=0;$cS=0;$cC=0;$cT=0

foreach($kv in $subs.GetEnumerator()){
    $r=$kv.Value
    $vtot=if($r.monto -and [double]$r.monto -gt 0){[double]$r.monto} else {$r.m+($r.vadd*$r.dep)}

    $dj="["; $f=$true
    foreach($d in $r.deps){
        if(-not $f){$dj+=","}
        $dj+="{vid:`"$(EscJs $d.vid)`",c:`"$(EscJs $d.c)`",t:`"$(EscJs $d.t)`"}"
        $f=$false
    }
    $dj+="]"

    $rec="{id:$($r.id),vid:`"$(EscJs $r.vid)`",c:`"$(EscJs $r.c)`",t:`"$(EscJs $r.t)`",e:`"$(EscJs $r.e)`",p:`"$(EscJs $r.p)`",v:`"$($r.v)`",s:`"$($r.s)`",diasExp:$($r.diasExp),m:$($r.m),vadd:$($r.vadd),dep:$($r.dep),vtot:$vtot,com:$($r.com),u:`"$(EscJs $r.u)`",texp:`"$($r.texp)`",texpd:$($r.texpd),deps:$dj}"

    if($r.s -eq "mora"){$cM++} elseif($r.s -eq "proximo"){$cP++} else {$cA++}
    if($r.tok -eq 0){$cS++;$sinList.Add($rec)}
    else{$cC++;$conList.Add($rec);if($r.texpd -le 45){$cT++}}
}

# Ordenar SIN: mora(más vencida primero), proximo, activo
$sinSorted=$sinList|Sort-Object {
    $l=$_
    $sm=if($l-match 's:"mora"'){0}elseif($l-match 's:"proximo"'){1}else{2}
    $dm=if($l-match 'diasExp:(-?\d+)'){[int]$Matches[1]+99999}else{999999}
    $sm*1000000+$dm
}
$conSorted=$conList|Sort-Object {
    $l=$_
    if($l-match 'diasExp:(-?\d+)'){[int]$Matches[1]+99999}else{999999}
}

$content="window.SIN_TOKEN=["+($sinSorted-join',')+"];`r`nwindow.CON_TOKEN=["+($conSorted-join',')+"];"
[System.IO.File]::WriteAllText($outPath,$content,[System.Text.Encoding]::UTF8)

Write-Host "=== RESULTADO ==="
Write-Host "SIN TOKEN : $cS"
Write-Host "CON TOKEN : $cC"
Write-Host "Mora      : $cM"
Write-Host "Proximos  : $cP"
Write-Host "Activos   : $cA"
Write-Host "Texp<=45d : $cT"
Write-Host "datos.js generado OK"
