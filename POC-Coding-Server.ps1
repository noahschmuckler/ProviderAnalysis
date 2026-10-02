#requires -Version 5.1
<#
Point-of-Care Coding Localhost Analyzer
Run: powershell -ExecutionPolicy Bypass -File .\POC-Coding-Server.ps1
Requirements: Windows, Microsoft Excel desktop, and an .xlsx export with the Sheet1-style headers.
Open http://localhost:8765, select a file, and click Analyze.
Press Ctrl+C in PowerShell to stop the server.
#>
param([int]$Port = 8765)

Add-Type -AssemblyName System.Web
$TargetDepartment = 'CRH PC EMERALD PLACE [1006002036]'
$Required = @('Service Prov','Guarantor','Service Dt','Procedure','Department','Tx ID','Diagnosis Code','Encounter Num','Work RVU','Procedure Code','Proc Mod')

function H([object]$v) { [System.Net.WebUtility]::HtmlEncode([string]$v) }
function Rate($n,$d) { if($d -eq 0){ return $null }; return [double]$n/$d }
function Pct($v) { if($null -eq $v){ return '' }; return ('{0:P1}' -f $v) }
function Has25([string]$m) { return $m -match '(^|[,;\s])25($|[,;\s])' }
function IsEM([string]$c) { return $c -match '^9920[2-5]$|^9921[1-5]$' }
function IsPrev([string]$c) { return $c -match '^9938[0-9]$|^9939[0-9]$' }
function IsAWV([string]$c) { return @('G0402','G0438','G0439') -contains $c }
function IsFocus([string]$c) { return (IsEM $c) -or (IsPrev $c) -or (IsAWV $c) -or @('G2211','G0447','99401','99402','99403','99404','99406','99407','G0446','G0537','G0538','G0444','96127') -contains $c }

function Read-Xlsx([string]$Path) {
  $xl=$null;$book=$null
  try {
    $xl=New-Object -ComObject Excel.Application;$xl.Visible=$false;$xl.DisplayAlerts=$false
    $book=$xl.Workbooks.Open($Path,$null,$true)
    foreach($ws in $book.Worksheets){
      $used=$ws.UsedRange;$vals=$used.Value2
      if($used.Rows.Count -lt 2 -or $used.Columns.Count -lt 11){continue}
      $map=@{};for($c=1;$c -le $used.Columns.Count;$c++){$map[[string]$vals[1,$c]]=$c}
      $ok=$true;foreach($h in $Required){if(-not $map.ContainsKey($h)){$ok=$false;break}}
      if(-not $ok){continue}
      $rows=New-Object System.Collections.Generic.List[object]
      for($r=2;$r -le $used.Rows.Count;$r++){
        $o=[ordered]@{};foreach($h in $Required){$o[$h]=$vals[$r,$map[$h]]}
        $rows.Add([pscustomobject]$o)
      }
      return @{Sheet=$ws.Name;Rows=$rows}
    }
    throw 'No worksheet contains all required Sheet1-style headers.'
  } finally {
    if($book){$book.Close($false)};if($xl){$xl.Quit()}
    foreach($x in @($used,$ws,$book,$xl)){if($x){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
  }
}

function Analyze($Rows) {
  $scope=@($Rows|Where-Object{[string]$_.Department -eq $TargetDepartment})
  $diag=@{};foreach($r in $Rows){$e=[string]$r.'Encounter Num';if($e){$diag[$e]+=' '+[string]$r.'Diagnosis Code'+' '+[string]$r.Procedure}}
  $encs=New-Object System.Collections.Generic.List[object];$confEM=0;$confPrev=0;$base=0.0
  foreach($eg in ($scope|Group-Object 'Encounter Num')){
    $active=New-Object System.Collections.Generic.List[object]
    foreach($cg in ($eg.Group|Group-Object {$_.('Procedure Code').ToString().Trim().ToUpper()+'|'+$_.('Proc Mod').ToString().Trim().ToUpper()})){
      $a=@($cg.Group|Sort-Object {[double]$_.'Tx ID'} -Descending)
      $nz=@($a|Where-Object{[double]$_.'Work RVU' -ne 0}|Select-Object -First 1)
      $code=[string]$a[0].'Procedure Code';$code=$code.Trim().ToUpper()
      if($nz){$active.Add($nz[0])}elseif(IsFocus $code){$active.Add($a[0])}
    }
    $ems=@($active|Where-Object{IsEM ([string]$_.'Procedure Code')});if($ems.Count -gt 1){$confEM++;$keep=$ems|Sort-Object {[double]$_.'Tx ID'} -Descending|Select-Object -First 1;$active=@($active|Where-Object{(-not (IsEM ([string]$_.'Procedure Code'))) -or $_ -eq $keep})}
    $pvs=@($active|Where-Object{IsPrev ([string]$_.'Procedure Code')});if($pvs.Count -gt 1){$confPrev++;$keep=$pvs|Sort-Object {[double]$_.'Tx ID'} -Descending|Select-Object -First 1;$active=@($active|Where-Object{(-not (IsPrev ([string]$_.'Procedure Code'))) -or $_ -eq $keep})}
    foreach($x in $active){$base += [double]$x.'Work RVU'}
    $new=$eg.Group|Sort-Object {[double]$_.'Tx ID'} -Descending|Select-Object -First 1
    $em=$active|Where-Object{IsEM ([string]$_.'Procedure Code')}|Sort-Object {[double]$_.'Tx ID'} -Descending|Select-Object -First 1
    $pv=$active|Where-Object{IsPrev ([string]$_.'Procedure Code')}|Sort-Object {[double]$_.'Tx ID'} -Descending|Select-Object -First 1
    $codes=@($active|ForEach-Object{([string]$_.'Procedure Code').Trim().ToUpper()});$txt=[string]$diag[$eg.Name]
    $ob=$txt -match '(?i)\bE66\b|obesity';$tob=$txt -match '(?i)\bF17\b|\bZ72\.0\b|nicotine|tobacco';$asc=$txt -match '(?i)\bI10\b|\bE78\b|\bE11\b|\bE66\b|\bF17\b|hypertension|hyperlipidemia|diabetes|obesity|nicotine|tobacco';$mh=$txt -match '(?i)\bF32\b|\bF33\b|\bF41\b|\bZ13\.31\b|\bZ13\.32\b|depression|anxiety'
    $obc=@($codes|Where-Object{$_ -eq 'G0447' -or (($ob) -and @('99401','99402','99403','99404') -contains $_)});$tc=@($codes|Where-Object{@('99406','99407') -contains $_});$ac=@($codes|Where-Object{@('G0446','G0537','G0538') -contains $_});$mc=@($codes|Where-Object{@('G0444','96127') -contains $_})
    $emmod=if($em){[string]$em.'Proc Mod'}else{''};$has25=if($em){Has25 $emmod}else{$false};$hasg=$codes -contains 'G2211'
    $provider=([string]$new.'Service Prov') -replace '\s*\[[^\]]*\]\s*$',''
    $encs.Add([pscustomobject]@{Encounter=$eg.Name;Date=[DateTime]::FromOADate([double]$new.'Service Dt');Provider=$provider.Trim();EM=if($em){[string]$em.'Procedure Code'}else{''};EMMod=$emmod;Prev=if($pv){[string]$pv.'Procedure Code'}else{''};AWV=@($codes|Where-Object{IsAWV $_}) -join ', ';G2211=$hasg;P25=($pv -and $em -and $has25);P25G=($pv -and $em -and $has25 -and $hasg);ObCode=($obc.Count -gt 0);TobCode=($tc.Count -gt 0);AscCode=($ac.Count -gt 0);MHCode=($mc.Count -gt 0);Ob=$ob;Tob=$tob;Asc=$asc;MH=$mh})
  }
  return @{Raw=$Rows.Count;Scope=$scope.Count;Enc=$encs;Base=$base;EMConf=$confEM;PrevConf=$confPrev}
}

function Build-Page($A,[string]$File) {
  $e=@($A.Enc);$providers=@($e.Provider|Sort-Object -Unique);$em=@($e|Where-Object{$_.EM}).Count;$g=@($e|Where-Object{$_.EM -and $_.G2211}).Count;$pv=@($e|Where-Object{$_.Prev}).Count;$p25=@($e|Where-Object{$_.P25}).Count;$p25g=@($e|Where-Object{$_.P25G}).Count
  $score=@(foreach($p in $providers){$x=@($e|Where-Object{$_.Provider -eq $p});$xe=@($x|Where-Object{$_.EM});$xg=@($xe|Where-Object{$_.G2211});$other=@($e|Where-Object{$_.Provider -ne $p -and $_.EM});$og=@($other|Where-Object{$_.G2211});$xp=@($x|Where-Object{$_.Prev});[pscustomobject]@{Provider=$p;Enc=$x.Count;EM=$xe.Count;G=$xg.Count;Rate=$(Rate $xg.Count $xe.Count);Peer=$(Rate $og.Count $other.Count);Gap=if($xe.Count -and $other.Count){(Rate $xg.Count $xe.Count)-(Rate $og.Count $other.Count)}else{$null};Prev=$xp.Count;P25=@($x|Where-Object{$_.P25}).Count;P25G=@($x|Where-Object{$_.P25G}).Count;ObC=@($x|Where-Object{$_.ObCode}).Count;Ob=@($x|Where-Object{$_.Ob}).Count;TobC=@($x|Where-Object{$_.TobCode}).Count;Tob=@($x|Where-Object{$_.Tob}).Count;AscC=@($x|Where-Object{$_.AscCode}).Count;Asc=@($x|Where-Object{$_.Asc}).Count;MHC=@($x|Where-Object{$_.MHCode}).Count;MH=@($x|Where-Object{$_.MH}).Count}})|Sort-Object Rate -Descending
  $rows=($score|ForEach-Object{"<tr><td>$(H $_.Provider)</td><td>$($_.Enc)</td><td>$($_.EM)</td><td>$($_.G)</td><td>$(Pct $_.Rate)</td><td>$(Pct $_.Peer)</td><td class='$(if($_.Gap -ge 0){'pos'}else{'neg'})'>$(Pct $_.Gap)</td><td>$($_.Prev)</td><td>$($_.P25)</td><td>$(Pct (Rate $_.P25 $_.Prev))</td><td>$($_.P25G)</td><td>$(Pct (Rate $_.P25G $_.Prev))</td><td>$($_.ObC)/$($_.Ob)</td><td>$($_.TobC)/$($_.Tob)</td><td>$($_.AscC)/$($_.Asc)</td><td>$($_.MHC)/$($_.MH)</td></tr>"}) -join ''
  $opp1=@($e|Where-Object{($_.Prev -or $_.AWV) -and $_.EM -and (Has25 $_.EMMod) -and -not $_.G2211}).Count;$opp2=@($e|Where-Object{-not $_.Prev -and -not $_.AWV -and $_.EM -eq '99214' -and -not $_.G2211}).Count
  return @"
<!doctype html><html><head><meta charset='utf-8'><title>POC Coding Behavior</title><style>
:root{--navy:#17365d;--teal:#1f6d7a;--pale:#eaf3f5;--ink:#243447}*{box-sizing:border-box}body{margin:0;background:#f4f7f9;font:14px Segoe UI,Arial;color:var(--ink)}header{background:var(--navy);color:white;padding:24px 4vw}main{padding:22px 4vw}.caveat{background:#fff6da;border-left:5px solid #d7a900;padding:12px;margin:14px 0}.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.tile,.card{background:white;border-radius:8px;padding:16px;box-shadow:0 1px 5px #ccd}.tile b{display:block;font-size:26px;color:var(--teal)}.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:16px}h2{color:var(--navy)}table{width:100%;border-collapse:collapse;background:white}th{background:var(--teal);color:white;position:sticky;top:0}th,td{padding:8px;border-bottom:1px solid #dde5e8;text-align:right;white-space:nowrap}th:first-child,td:first-child{text-align:left}.scroll{overflow:auto;max-height:520px}.pos{color:#087830}.neg{color:#b42318}input[type=range]{width:100%}@media(max-width:900px){.kpis,.grid{grid-template-columns:1fr 1fr}}</style></head><body>
<header><h1>Point-of-Care Coding Behavior</h1><div>$(H $File) · $($A.Scope) in-scope rows · $($e.Count) encounters · $($providers.Count) providers</div></header><main>
<div class='caveat'>Diagnosis-text proxies are broad discussion signals only—not confirmed eligibility, medical necessity, denial risk, compliance findings, or proof of incorrect coding. Latest active state uses Tx ID ordering; zero wRVU alone does not erase an earlier non-zero transaction.</div>
<section class='kpis'><div class='tile'>Eligible encounters<b>$($e.Count)</b></div><div class='tile'>E/M encounters<b>$em</b></div><div class='tile'>G2211 with E/M<b>$g</b></div><div class='tile'>Group G2211 rate<b>$(Pct (Rate $g $em))</b></div><div class='tile'>Preventive encounters<b>$pv</b></div><div class='tile'>Prev + E/M/25<b>$p25</b></div><div class='tile'>+ G2211 nested<b>$p25g</b></div><div class='tile'>Active wRVUs<b>$('{0:N1}'-f$A.Base)</b></div></section>
<section class='grid'><div class='card'><h2>Preventive adoption</h2><table><tr><th>Stage</th><th>Encounters</th><th>% preventive</th></tr><tr><td>Any annual preventive</td><td>$pv</td><td>$(Pct (Rate $pv $pv))</td></tr><tr><td>Preventive + E/M with 25</td><td>$p25</td><td>$(Pct (Rate $p25 $pv))</td></tr><tr><td>Preventive + E/M/25 + G2211</td><td>$p25g</td><td>$(Pct (Rate $p25g $pv))</td></tr></table></div>
<div class='card'><h2>G2211 increased-usage calculator</h2><label>Increase dial: <b id='dialOut'>100%</b></label><input id='dial' type='range' min='0' max='100' value='100'><label>wRVU per G2211 <input id='wrvu' type='number' step='.01' value='.33'></label><table><tr><th>Opportunity</th><th>Without</th><th>Added G2211</th><th>Added wRVU</th></tr><tr><td>Preventive/AWV + E/M/25</td><td>$opp1</td><td id='a1'></td><td id='w1'></td></tr><tr><td>Non-preventive 99214</td><td>$opp2</td><td id='a2'></td><td id='w2'></td></tr><tr><td>Total</td><td>$($opp1+$opp2)</td><td id='at'></td><td id='wt'></td></tr></table></div></section>
<h2>Provider Coding Scorecard</h2><div class='scroll'><table><tr><th>Provider</th><th>Encounters</th><th>E/M</th><th>G2211</th><th>G2211 rate</th><th>Peer rate</th><th>Gap</th><th>Preventive</th><th>Prev+E/M/25</th><th>Adoption</th><th>+G2211</th><th>Nested rate</th><th>Obesity code/proxy</th><th>Tobacco code/proxy</th><th>ASCVD code/proxy</th><th>Mental health code/proxy</th></tr>$rows</table></div>
<script>const d=document.querySelector('#dial'),w=document.querySelector('#wrvu');function calc(){let p=d.value/100,v=+w.value,x1=Math.round($opp1*p),x2=Math.round($opp2*p);dialOut.textContent=d.value+'%';a1.textContent=x1;a2.textContent=x2;at.textContent=x1+x2;w1.textContent=(x1*v).toFixed(1);w2.textContent=(x2*v).toFixed(1);wt.textContent=((x1+x2)*v).toFixed(1)}d.oninput=calc;w.oninput=calc;calc()</script></main></body></html>
"@
}

$Home=@"
<!doctype html><html><head><meta charset='utf-8'><title>POC Analyzer</title><style>body{font:16px Segoe UI;background:#f4f7f9;color:#243447;display:grid;place-items:center;height:100vh}.box{background:white;padding:32px;border-radius:10px;box-shadow:0 2px 12px #bbc;max-width:620px}h1{color:#17365d}button{background:#1f6d7a;color:white;border:0;border-radius:5px;padding:10px 18px}</style></head><body><form class='box' method='post' action='/analyze' enctype='multipart/form-data'><h1>Point-of-Care Coding Analyzer</h1><p>Select an .xlsx export containing the required Sheet1-style columns.</p><input type='file' name='file' accept='.xlsx' required><button>Analyze</button></form></body></html>
"@

function Send($ctx,[string]$html,[int]$status=200){$b=[Text.Encoding]::UTF8.GetBytes($html);$ctx.Response.StatusCode=$status;$ctx.Response.ContentType='text/html; charset=utf-8';$ctx.Response.ContentLength64=$b.Length;$ctx.Response.OutputStream.Write($b,0,$b.Length);$ctx.Response.Close()}
function Save-Upload($req){
  $ct=$req.ContentType;if($ct -notmatch 'boundary=(.+)$'){throw 'Invalid upload.'};$boundary='--'+$Matches[1].Trim('"');$ms=New-Object IO.MemoryStream;$req.InputStream.CopyTo($ms);$raw=$ms.ToArray();$latin=[Text.Encoding]::GetEncoding(28591);$txt=$latin.GetString($raw);$start=$txt.IndexOf("`r`n`r`n");$end=$txt.LastIndexOf("`r`n$boundary");if($start -lt 0 -or $end -lt 0){throw 'Could not parse upload.'};$head=$txt.Substring(0,$start);if($head -notmatch 'filename="([^"]+)"'){throw 'No file selected.'};$name=[IO.Path]::GetFileName($Matches[1]);if([IO.Path]::GetExtension($name) -ne '.xlsx'){throw 'Only .xlsx files are accepted.'};$bytes=$raw[($start+4)..($end-1)];$path=Join-Path $env:TEMP ("POC_"+[guid]::NewGuid()+'.xlsx');[IO.File]::WriteAllBytes($path,$bytes);return @($path,$name)
}

$listener=[Net.HttpListener]::new();$listener.Prefixes.Add("http://localhost:$Port/");$listener.Start();Start-Process "http://localhost:$Port/";Write-Host "POC analyzer running at http://localhost:$Port/  (Ctrl+C to stop)"
try{while($listener.IsListening){$ctx=$listener.GetContext();try{if($ctx.Request.HttpMethod -eq 'GET'){Send $ctx $Home}else{$up=Save-Upload $ctx.Request;$data=Read-Xlsx $up[0];$a=Analyze $data.Rows;Send $ctx (Build-Page $a $up[1]);Remove-Item $up[0] -Force -ErrorAction SilentlyContinue}}catch{Send $ctx ("<h1>Analysis error</h1><pre>"+(H $_.Exception.Message)+"</pre><p><a href='/'>Try another file</a></p>") 400}}}finally{$listener.Stop();$listener.Close()}
