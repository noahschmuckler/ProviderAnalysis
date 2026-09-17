#requires -Version 5.1
# ProviderAnalysisServer.ps1 - Draft 5.3 (ranked outreach lists, per-provider patient flags, Communication tab, unattended-run recovery)
# Adds provider indexes, guided cross-source mapping, fuzzy suggestions, NPI profiles, prepared job tracking, ranked outreach lists,
# a flag mode (/flag?jobId=...) that excludes flagged patients from every list and panel on every later report for that provider,
# and a Communication tab (/communication) that emails every provider in a risk pool a fresh report PDF via Outlook drafts.
[CmdletBinding()]
param([int]$PreferredPort=8765,[switch]$NoBrowser,[switch]$SkipModuleInstallPrompt,[string]$RunJobId='')
Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'
$script:Root=Split-Path -Parent $MyInvocation.MyCommand.Path
if([string]::IsNullOrWhiteSpace($script:Root)){$script:Root=(Get-Location).Path}
Set-Location -LiteralPath $script:Root
$script:IndexVersion=7
$script:IndexMemo=@{}
$script:TokenCache=@{}
$script:RunNext=$false
$script:Paths=[ordered]@{
 CanonicalCurrent=Join-Path $script:Root 'canonical-current'
 CanonicalArchive=Join-Path $script:Root 'canonical-archive'
 ProvidersCurrent=Join-Path $script:Root 'providers-current'
 ProvidersArchive=Join-Path $script:Root 'providers-archive'
 Profiles=Join-Path $script:Root 'profiles'
 State=Join-Path $script:Root 'state'
 Staging=Join-Path $script:Root 'staging'
 Logs=Join-Path $script:Root 'logs'
 Contacts=Join-Path $script:Root 'contacts'
}
$script:ConfigPath=Join-Path $script:Root 'config.json'
$script:ManifestPath=Join-Path $script:Paths.State 'manifest.json'
$script:Stop=$false;$script:Listener=$null;$script:Mutex=$null
function Safe([object]$v){if($null -eq $v){return ''};$s=([string]$v) -replace '[\r\n\t]+',' ';if($s.Length -gt 500){$s=$s.Substring(0,500)};return $s}
function Log([string]$Action,[string]$Result='OK',[string]$Message='',[string]$LogSourceKey=''){
 if(!(Test-Path $script:Paths.Logs)){return}
 $e=[ordered]@{utc=[DateTime]::UtcNow.ToString('o');windowsUser=[Environment]::UserName;action=Safe $Action;sourceKey=Safe $LogSourceKey;result=Safe $Result;message=Safe $Message}
 Add-Content -LiteralPath (Join-Path $script:Paths.Logs ((Get-Date -Format 'yyyy-MM-dd')+'.jsonl')) -Value ($e|ConvertTo-Json -Compress) -Encoding UTF8
}
function Initialize-AppFolders{foreach($p in $script:Paths.Values){if(!(Test-Path $p)){New-Item -ItemType Directory -Path $p -Force|Out-Null}}}
function Source([string]$Key,[string]$Name,[string]$Providers,[string]$Npi,[string]$Headers){
 [ordered]@{sourceKey=$Key;displayName=$Name;canonicalFileName=($Key+'.xlsx');headerScan=25;providerColumns=@($Providers.Split('|'));npiColumn=$(if($Npi){$Npi}else{$null});orderedHeaders=@($Headers.Split('|'))}
}
function Registry{
 @(
  (Source 'RiskPopulationOutreach' 'Risk Population Outreach' 'Epic Pcp' '' 'Cdo|Payer|Patient|Epic Pcp|Payer File Pcp|Member ID|Birth Date|Mrn|Last Acv|Next Acv|Last A1c|Last A1c Value'),
  (Source 'SerialScheduling' 'Serial Scheduling' 'Provider Name' '' 'Payer|Product|Provider Name|Risk Pool|Member ID|MRN|First Name|Last Name|DOB|GSD|BCS|COLO|CBP|Last A1c|A1c Date|DM Eye Exam|Kidney Evaluation for Diabetes|Med Adherence DM|Med Adherence Statin|Med Adherence HTN|SUPD|Completed PCP Visits 2025|Completed Endo Visits 2025|Completed PCP Visits 2026|Future PCP Visits 2026|PCP Visit Dates|Completed Endo Visits 2026|Future Endo Visits 2026|Endo Visit Dates'),
  (Source 'DiabetesScorecard' 'Diabetes Scorecard' 'Provider' '' 'Cdo|Location|Provider|Patient|Member ID|Mrn|Product|Payer|KED|EED|Eye Exam Gap Status|Eye Exam Date|Next Appt Date|Next Appt Specialty|Next Appt Location|Risk|GSD|Med Adherence DM|MAD Days Supply|Dx Date|% Saw PCP last 6 mo.|Avg PCP visits / Patient last 12 mo.|% Saw Endo last 6 mo.|Avg Endo visits / Patient last 12 mo.|Avg Last A1c|% eGFR last 12 mo.|% uACR last 12 mo.|% DM Rx w/ Ext Supply|% DM Inj w/ Ext Supply'),
  (Source 'HR-CRH' 'HR-CRH' 'PCP Name' '' 'PCP Name|Patient First Name|Patient Last Name|DOB|Patient Insurance ID|Number of QEMs with PCP|Risk Level in source'),
  (Source 'HR-RIVPHNYCMM' 'HR-RIVPHNYCMM' 'PCP Name' '' 'PCP Name|Patient First Name|Patient Last Name|DOB|Patient Insurance ID|Number of QEMs with PCP|Risk Level in source'),
  (Source 'PtListQuality' 'PtListQuality' 'Provider Name' 'NPI' 'MemberID|First Name|Last Name|DOB|Population|Risk|Risk Pool|Payer|Product|PPM|Practice Name|Provider Name|NPI|TIN|BCS|COLO|EED|GSD|CBP|OMW|KED|SPC|MAD|MAC|MAH|SUPD|COB|POLY|PCR - DENOMINATOR|PCR - NUMERATOR|FMC - DENOMINATOR|FMC - NUMERATOR|TRCM - DENOMINATOR|TRCM - NUMERATOR|TRCE - DENOMINATOR|TRCE - NUMERATOR|OMW Critical Due Date|Fracture Date|MAC Days Supply|MAD Days Supply|MAH Days Supply|COL_Sent|COL_Returned|COL_NPI|GSD_Sent|GSD_Returned|GSD_NPI|KED_Sent|KED_Returned|KED_NPI|BCS_CR|CBP_CR'),
  (Source 'Export' 'Export' 'Provider Name' 'NPI' 'Population|Payer|Product|PPM|Practice Name|Provider Name|Risk Pool|Risk|NPI|TIN|MemberID|First Name|Last Name|Date of Birth|Home Address 1|Home Address 2|City|State|Zip|Phone Number|SRF|Dual Eligible Status|Lis Status|Disabled|HCC UNCOVERED RATIO|Total Care Gaps|# of Part C Care Gaps|# of Part D Care Gaps|Last QEM date with non-PCP|Last ACV Date|Last QEM Visit Date with any PCP in assigned TIN|TCM|Completed Attestations|Incompleted Attestations|Open ICDs|New Patient|Active Date')
 )
}
function Initialize-AppConfiguration{
 $existing=Json $script:ConfigPath;$created=$(if($existing -and $existing.createdUtc){[string]$existing.createdUtc}else{[DateTime]::UtcNow.ToString('o')})
 $c=[ordered]@{configVersion=2;createdUtc=$created;updatedUtc=[DateTime]::UtcNow.ToString('o');loopbackOnly=$true;preferredPort=$PreferredPort;sourceHeaderRule='Exact ordered nonblank headers after trimming outer whitespace';exportDateRule='Use each incoming XLSX Windows CreationTime as its source export date';sources=@(Registry)}
 $c|ConvertTo-Json -Depth 12|Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8;if($existing){Log 'CONFIG_MIGRATED' 'OK' 'Draft 3.1 registry enforced; RPO uses Epic Pcp only'}else{Log 'CONFIG_CREATED'}
 if(!(Test-Path $script:ManifestPath)){[ordered]@{manifestVersion=1;updatedUtc=[DateTime]::UtcNow.ToString('o');sources=[ordered]@{}}|ConvertTo-Json -Depth 8|Set-Content $script:ManifestPath -Encoding UTF8;Log 'MANIFEST_CREATED'}
}
function Json([string]$Path){if(Test-Path $Path){Get-Content $Path -Raw -Encoding UTF8|ConvertFrom-Json}else{$null}}

function Test-ImportExcelModule{
 $m=Get-Module -ListAvailable ImportExcel|Sort-Object Version -Descending|Select-Object -First 1
 if(!$m){
  if($SkipModuleInstallPrompt){throw 'ImportExcel is required but was not found.'}
  $a=Read-Host 'ImportExcel was not found. Install for CurrentUser now? [Y/N]'
  if($a -notmatch '^(?i)y(es)?$'){throw 'ImportExcel installation was declined.'}
  Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber;Log 'MODULE_INSTALL'
 }
 Import-Module ImportExcel -ErrorAction Stop
 Log 'MODULE_CHECK' 'FOUND' ('ImportExcel '+(Get-Module ImportExcel).Version)
}
function Lock{
 $b=[Text.Encoding]::UTF8.GetBytes($script:Root.ToLowerInvariant());$s=[Security.Cryptography.SHA256]::Create()
 try{$h=([BitConverter]::ToString($s.ComputeHash($b))).Replace('-','').Substring(0,16)}finally{$s.Dispose()}
 $new=$false;$script:Mutex=New-Object Threading.Mutex($true,('Local\ProviderAnalysis_'+$h),[ref]$new)
 if(!$new){throw 'Another instance is already running for this application folder.'}
}
function FreePort([int]$Start){foreach($n in $Start..($Start+100)){$t=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,$n);try{$t.Start();return $n}catch{}finally{try{$t.Stop()}catch{}}};throw 'No free loopback port found.'}
function Status{
 $c=Json $script:ConfigPath;$m=Json $script:ManifestPath;$out=@()
 foreach($s in $c.sources){
  $f=Get-Item (Join-Path $script:Paths.CanonicalCurrent $s.canonicalFileName) -ErrorAction SilentlyContinue;$me=$null
  if($m -and $m.sources){$p=$m.sources.PSObject.Properties[$s.sourceKey];if($p){$me=$p.Value}}
  $sha=$(if($me){[string]$me.sha256}else{''});$index=$(if($f){Get-ProviderIndexState $s.sourceKey $f $sha}else{[ordered]@{state='NoFile';count=$null;builtUtc=$null}})
  $out+=[ordered]@{sourceKey=$s.sourceKey;displayName=$s.displayName;canonicalFileName=$s.canonicalFileName;present=($null -ne $f);byteLength=$(if($f){$f.Length}else{$null});fileCreationTime=$(if($f){$f.CreationTime.ToString('o')}else{$null});lastWriteTime=$(if($f){$f.LastWriteTime.ToString('o')}else{$null});importedUtc=$(if($me){$me.importedUtc}else{$null});rowCount=$(if($me){$me.rowCount}else{$null});sha256=$(if($me){$me.sha256}else{$null});indexState=$index.state;indexCount=$index.count;indexBuiltUtc=$index.builtUtc}
 };return $out
}
function Get-ProviderIndexState([string]$SourceKey,$File,[string]$CurrentSha){
 if($script:IndexMemo.ContainsKey($SourceKey)){
  $m=$script:IndexMemo[$SourceKey];$stamp=$(if($File){[string]$File.Length+'|'+$File.LastWriteTimeUtc.Ticks}else{''})
  if($m.stamp -eq $stamp -or ($CurrentSha -and $m.hash -eq $CurrentSha)){return [ordered]@{state='Ready';count=$m.count;builtUtc=$m.builtUtc}}
 }
 if(!$CurrentSha){return [ordered]@{state='Unknown';count=$null;builtUtc=$null}}
 $cached=Json (Join-Path $script:Paths.State ('provider-index-'+$SourceKey+'.json'))
 if($cached){
  $p=$cached.PSObject.Properties
  $shaOk=($p['sha256'] -and [string]$p['sha256'].Value -eq $CurrentSha);$versionOk=($p['indexVersion'] -and [int]$p['indexVersion'].Value -eq $script:IndexVersion)
  if($shaOk -and $versionOk){return [ordered]@{state='Ready';count=$(if($p['count']){[int]$p['count'].Value}else{$null});builtUtc=$(if($p['builtUtc']){[string]$p['builtUtc'].Value}else{$null})}}
 }
 return [ordered]@{state='Stale';count=$null;builtUtc=$null}
}
function Update-ProviderIndexAfterImport([string]$SourceKey,[switch]$Force){
 try{$sw=[Diagnostics.Stopwatch]::StartNew();$items=@(Get-ProviderIndex $SourceKey -Force:$Force);return @{count=$items.Count;seconds=[Math]::Round($sw.Elapsed.TotalSeconds,1);error=$null}}
 catch{Log 'PROVIDER_INDEX' 'FAILED' $_.Exception.Message $SourceKey;return @{count=$null;seconds=$null;error=$_.Exception.Message}}
}
function Initialize-ProviderIndexes{
 $config=Json $script:ConfigPath;Write-Host 'Preparing provider indexes (only rebuilt when a source file changed)...' -ForegroundColor Cyan
 foreach($s in $config.sources){
  if(!(Test-Path -LiteralPath (Join-Path $script:Paths.CanonicalCurrent $s.canonicalFileName))){continue}
  $r=Update-ProviderIndexAfterImport ([string]$s.sourceKey)
  if($r.error){Write-Warning ($s.displayName+': provider index unavailable - '+$r.error)}else{Write-Host ('  '+$s.displayName+': '+$r.count+' providers ('+$r.seconds+'s)')}
 }
}





















function Serve{
 $port=FreePort $PreferredPort;$url='http://127.0.0.1:'+$port+'/';$script:Listener=New-Object Net.HttpListener;$script:Listener.Prefixes.Add($url);$script:Listener.Start();Log 'SERVER_START' 'OK' ('Loopback port '+$port);Write-Host ('Provider Analysis Server: '+$url) -ForegroundColor Green
 if(!$NoBrowser){Start-Process $url}
 while(-not $script:Stop -and $script:Listener.IsListening){$a=$script:Listener.BeginGetContext($null,$null);while(-not $a.AsyncWaitHandle.WaitOne(250)){if($script:Stop){break};if(([DateTime]::UtcNow-$script:LastHeartbeat).TotalSeconds -ge $script:HeartbeatSeconds){Invoke-Heartbeat}};if($script:Stop){break};try{Request ($script:Listener.EndGetContext($a))}catch{Log 'HTTP_LOOP' 'FAILED' $_.Exception.Message};if($script:RunNext){$script:RunNext=$false;try{Invoke-NextPendingJob}catch{Log 'JOB_LAUNCH' 'FAILED' $_.Exception.Message}}}
}
function Save-JsonAtomic([string]$Path,[object]$Value){$tmp=$Path+'.tmp';$Value|ConvertTo-Json -Depth 15|Set-Content -LiteralPath $tmp -Encoding UTF8;Move-Item -LiteralPath $tmp -Destination $Path -Force}
function Get-FileHash256([string]$Path){return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()}
function Test-HeaderEqual([object[]]$Actual,[object[]]$Expected){if($Actual.Count -ne $Expected.Count){return $false};for($i=0;$i -lt $Expected.Count;$i++){if(([string]$Actual[$i]).Trim() -cne ([string]$Expected[$i]).Trim()){return $false}};return $true}
function Test-WorkbookSchema([string]$Path){
 $config=Json $script:ConfigPath;$package=$null;$schemaMatches=@();$observed=@()
 try{$package=Open-ExcelPackage -Path $Path -ErrorAction Stop;foreach($ws in $package.Workbook.Worksheets){if($null -eq $ws.Dimension){continue};$maxCol=$ws.Dimension.End.Column;$limit=[Math]::Min(25,$ws.Dimension.End.Row);for($row=1;$row -le $limit;$row++){$headers=@();for($col=1;$col -le $maxCol;$col++){$text=([string]$ws.Cells[$row,$col].Text).Trim();if($text){$headers+=$text}};if($headers.Count -eq 0){continue};$dupes=@($headers|Group-Object -CaseSensitive|Where-Object{$_.Count -gt 1}|ForEach-Object{$_.Name});$observed+=[ordered]@{worksheet=$ws.Name;row=$row;headers=$headers;duplicates=$dupes};if($dupes.Count -gt 0){continue};foreach($source in $config.sources){if(Test-HeaderEqual $headers @($source.orderedHeaders)){$schemaMatches+=[ordered]@{sourceKey=$source.sourceKey;displayName=$source.displayName;canonicalFileName=$source.canonicalFileName;worksheet=$ws.Name;headerRow=$row;rowCount=[Math]::Max(0,$ws.Dimension.End.Row-$row)}}}}}}
 finally{if($package){Close-ExcelPackage $package -NoSave}}
 if($schemaMatches.Count -ne 1){$summary=@($observed|Where-Object{$_.headers.Count -gt 1}|Select-Object -First 8|ForEach-Object{[ordered]@{worksheet=$_.worksheet;row=$_.row;headerCount=$_.headers.Count;firstHeaders=@($_.headers|Select-Object -First 6);duplicates=$_.duplicates}});throw ('Schema recognition requires exactly one match; found '+$schemaMatches.Count+'. Observed: '+($summary|ConvertTo-Json -Compress -Depth 5))};return $schemaMatches[0]
}
function Receive-Upload($Context){
 $name=[Net.WebUtility]::UrlDecode([string]$Context.Request.Headers['X-File-Name']);if([string]::IsNullOrWhiteSpace($name)){throw 'Missing file name.'};if([IO.Path]::GetExtension($name) -ine '.xlsx'){throw 'Only XLSX files are accepted.'};$token=[Guid]::NewGuid().ToString('N');$path=Join-Path $script:Paths.Staging ($token+'.xlsx');$metaPath=Join-Path $script:Paths.Staging ($token+'.json')
 $stream=[IO.File]::Create($path);try{$buffer=New-Object byte[] 1048576;$total=0;while(($read=$Context.Request.InputStream.Read($buffer,0,$buffer.Length)) -gt 0){$total+=$read;if($total -gt 2147483648){throw 'Upload exceeds 2 GB limit.'};$stream.Write($buffer,0,$read)}}finally{$stream.Dispose()}
 try{$clientDate=[string]$Context.Request.Headers['X-File-Timestamp'];if($clientDate){$dt=[DateTimeOffset]::FromUnixTimeMilliseconds([int64]$clientDate).UtcDateTime;(Get-Item $path).CreationTimeUtc=$dt};$match=Test-WorkbookSchema $path;$hash=Get-FileHash256 $path;$file=Get-Item $path;$meta=[ordered]@{token=$token;sourceKey=$match.sourceKey;displayName=$match.displayName;canonicalFileName=$match.canonicalFileName;originalName=[IO.Path]::GetFileName($name);stagedPath=$path;sha256=$hash;byteLength=$file.Length;worksheet=$match.worksheet;headerRow=$match.headerRow;rowCount=$match.rowCount;sourceFileCreatedUtc=$file.CreationTimeUtc.ToString('o');stagedUtc=[DateTime]::UtcNow.ToString('o');uploadedBy=[Environment]::UserName};Save-JsonAtomic $metaPath $meta;Log 'UPLOAD_VALIDATED' 'OK' ('Identified '+$match.sourceKey+'; '+$match.rowCount+' rows.') $match.sourceKey;return $meta}catch{Remove-Item $path,$metaPath -Force -ErrorAction SilentlyContinue;Log 'UPLOAD_VALIDATION' 'FAILED' $_.Exception.Message;throw}
}
function Publish-StagedFile([string]$Token){
 if($Token -notmatch '^[a-f0-9]{32}$'){throw 'Invalid confirmation token.'};$metaPath=Join-Path $script:Paths.Staging ($Token+'.json');$meta=Json $metaPath;if($null -eq $meta){throw 'Staged upload was not found or expired.'};$sourceKey=[string]$meta.sourceKey;$target=Join-Path $script:Paths.CanonicalCurrent ([string]$meta.canonicalFileName);$archiveDir=Join-Path $script:Paths.CanonicalArchive $sourceKey;if(!(Test-Path $archiveDir)){New-Item -ItemType Directory -Path $archiveDir -Force|Out-Null};$archive=$null;$promoted=$false;$stamp=Get-Date -Format 'yyyyMMdd-HHmmss'
 try{if(Test-Path $target){$oldHash=Get-FileHash256 $target;$archive=Join-Path $archiveDir ($sourceKey+'__'+$stamp+'__'+$oldHash.Substring(0,8)+'.xlsx');Move-Item -LiteralPath $target -Destination $archive -ErrorAction Stop};Move-Item -LiteralPath ([string]$meta.stagedPath) -Destination $target -ErrorAction Stop;$promoted=$true;$manifest=Json $script:ManifestPath;if($null -eq $manifest){$manifest=[pscustomobject]@{manifestVersion=1;updatedUtc=$null;sources=[pscustomobject]@{}}};if($null -eq $manifest.sources){$manifest|Add-Member -NotePropertyName sources -NotePropertyValue ([pscustomobject]@{}) -Force};$entry=[ordered]@{sourceKey=$sourceKey;currentPath=$target;sha256=[string]$meta.sha256;byteLength=[int64]$meta.byteLength;rowCount=[int]$meta.rowCount;worksheet=[string]$meta.worksheet;headerRow=[int]$meta.headerRow;fileLastWriteUtc=(Get-Item $target).LastWriteTimeUtc.ToString('o');sourceFileCreatedUtc=[string]$meta.sourceFileCreatedUtc;importedUtc=[DateTime]::UtcNow.ToString('o');importedBy=[Environment]::UserName;previousArchivePath=$archive};$manifest.sources|Add-Member -NotePropertyName $sourceKey -NotePropertyValue $entry -Force;$manifest.updatedUtc=[DateTime]::UtcNow.ToString('o');Save-JsonAtomic $script:ManifestPath $manifest;Remove-Item $metaPath -Force;Log 'CANONICAL_REPLACED' 'OK' ('Archived prior file and imported '+$meta.originalName) $sourceKey;return $entry}catch{$failure=$_.Exception.Message;try{if($promoted -and (Test-Path $target)){Move-Item $target ([string]$meta.stagedPath) -Force};if($archive -and (Test-Path $archive)){Move-Item $archive $target -Force}}catch{$failure+='; rollback error: '+$_.Exception.Message};Log 'CANONICAL_REPLACE' 'FAILED' $failure $sourceKey;throw $failure}
}
function Remove-ExpiredStaging{Get-ChildItem $script:Paths.Staging -File -ErrorAction SilentlyContinue|Where-Object{$_.LastWriteTimeUtc -lt [DateTime]::UtcNow.AddDays(-2)}|Remove-Item -Force -ErrorAction SilentlyContinue}
function ConvertTo-HtmlSafe([object]$Value){if($null -eq $Value){''}else{[Net.WebUtility]::HtmlEncode([string]$Value)}}
function Page{
 $rows='';foreach($r in (Status)){$state=$(if($r.present){'Present'}else{'Missing'});$size=$(if($r.byteLength){'{0:N1} MB' -f ($r.byteLength/1MB)}else{'-'});$hash=$(if($r.sha256){([string]$r.sha256).Substring(0,8)}else{'-'});$index=$(switch([string]$r.indexState){'Ready'{'Ready ('+$r.indexCount+' providers)'}'Stale'{'Needs rebuild'}'Unknown'{'Not built'}default{'-'}});$rows+='<tr><td>'+(ConvertTo-HtmlSafe $r.displayName)+'</td><td>'+$state+'</td><td>'+(ConvertTo-HtmlSafe $r.canonicalFileName)+'</td><td>'+$size+'</td><td>'+(ConvertTo-HtmlSafe $r.fileCreationTime)+'</td><td>'+(ConvertTo-HtmlSafe $r.lastWriteTime)+'</td><td>'+(ConvertTo-HtmlSafe $r.importedUtc)+'</td><td>'+(ConvertTo-HtmlSafe $r.rowCount)+'</td><td>'+$hash+'</td><td>'+(ConvertTo-HtmlSafe $index)+'</td><td><button onclick="pickFile()">Refresh</button></td></tr>'}
 $html=@'
<!doctype html><html><head><meta charset="utf-8"><title>Provider Analysis</title>
<style>body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:white;padding:22px 22px 12px}main{padding:22px}.card{background:white;border:1px solid #dce4ef;overflow:auto}__TABCSS__table{border-collapse:collapse;width:100%;min-width:1250px}th,td{padding:9px;border-bottom:1px solid #dce4ef;text-align:left;white-space:nowrap}th{background:#eaf1f8}button{padding:7px 11px;cursor:pointer}.stop{background:#a61b1b;color:white}.modal{position:fixed;inset:0;background:#0008;display:none;align-items:center;justify-content:center}.box{background:white;padding:22px;border-radius:8px;max-width:650px;white-space:pre-wrap}</style></head>
<body><header><h1>Provider Analysis Automation</h1><p>Draft 4.7 - import the seven source exports; each is validated, fingerprinted and indexed</p></header>__NAV__
<main><button onclick="location.reload()">Reload</button> <button onclick="reindex()">Rebuild provider indexes</button> <button class="stop" onclick="stopServer()">Stop server</button><input id="file" type="file" accept=".xlsx" hidden>
<h2>Canonical sources</h2><div class="card"><table><thead><tr><th>Source</th><th>Status</th><th>File</th><th>Size</th><th>Export/File created</th><th>Last write</th><th>Imported UTC</th><th>Rows</th><th>SHA-256</th><th>Provider index</th><th>Action</th></tr></thead><tbody>__ROWS__</tbody></table></div>
<p>Uploads are staged and fingerprinted before replacement. Nothing changes until you confirm. The provider index for a source is rebuilt automatically after each import, so the wizard opens its name lists instantly.</p></main>
<div id="modal" class="modal"><div class="box"><div id="message"></div><p><button id="confirm" style="display:none">Confirm replacement</button> <button onclick="closeModal()">Close</button></p></div></div>
<script>
let token='';const file=document.getElementById('file');
async function call(url,opts){const r=await fetch(url,opts);const text=await r.text();let j=null;try{j=text?JSON.parse(text):null}catch(e){throw new Error('Server returned an unreadable response: '+text.slice(0,200))}if(!r.ok)throw new Error((j&&j.error)||('Request failed ('+r.status+')'));return j}
function pickFile(){file.value='';file.click()}
file.onchange=async()=>{if(!file.files.length)return;const f=file.files[0];show('Uploading and validating '+f.name+' ...',false);try{const j=await call('/api/stage',{method:'POST',headers:{'X-File-Name':encodeURIComponent(f.name),'X-File-Timestamp':String(f.lastModified)},body:f});token=j.token;show('Recognized source: '+j.displayName+'\nWorksheet/header: '+j.worksheet+' / row '+j.headerRow+'\nData rows: '+j.rowCount+'\nSHA-256: '+j.sha256+'\n\nConfirm to archive the current canonical file and replace it.',true)}catch(e){show('Error: '+e.message,false)}};
document.getElementById('confirm').onclick=async()=>{show('Replacing canonical file and building its provider index (large files can take a minute) ...',false);try{const j=await call('/api/confirm?token='+encodeURIComponent(token),{method:'POST'});show('Replacement completed.\nProvider index: '+(j.indexError?'ERROR - '+j.indexError:(j.providerCount+' providers')),false);setTimeout(()=>location.reload(),j.indexError?4000:1200)}catch(e){show('Error: '+e.message,false)}};
async function reindex(){show('Rebuilding provider indexes for every imported source. Large files can take a minute each ...',false);try{const j=await call('/api/reindex',{method:'POST'});show((j.map(x=>x.displayName+': '+(x.error?'ERROR - '+x.error:(x.count+' providers in '+x.seconds+'s'))).join('\n'))||'No imported sources to index.',false);setTimeout(()=>location.reload(),3000)}catch(e){show('Error: '+e.message,false)}}
function show(t,c){document.getElementById('message').textContent=t;document.getElementById('confirm').style.display=c?'inline-block':'none';document.getElementById('modal').style.display='flex'}
function closeModal(){document.getElementById('modal').style.display='none'}
async function stopServer(){if(confirm('Stop server?')){await fetch('/api/stop',{method:'POST'});document.body.innerHTML='<main><h2>Server stopped.</h2></main>'}}
</script></body></html>
'@
 return $html.Replace('__ROWS__',$rows).Replace('__NAV__',(Get-NavHtml 'sources')).Replace('__TABCSS__',$script:TabCss)
}
function Send($Context,[int]$Code,[string]$Type,[string]$Body){$b=[Text.Encoding]::UTF8.GetBytes($Body);$Context.Response.StatusCode=$Code;$Context.Response.ContentType=$Type;$Context.Response.ContentLength64=$b.Length;$Context.Response.Headers['Cache-Control']='no-store';$Context.Response.OutputStream.Write($b,0,$b.Length);$Context.Response.Close()}
function Request($Context){
 $remote=$null;try{$remote=$Context.Request.RemoteEndPoint}catch{}
 if($null -eq $remote -or -not ([Net.IPAddress]::IsLoopback($remote.Address))){try{Send $Context 403 'text/plain' 'Forbidden'}catch{};return}
 $method=$Context.Request.HttpMethod;$path=$Context.Request.Url.AbsolutePath.TrimEnd('/');if(!$path){$path='/'}
 try{if($method -eq 'GET' -and $path -eq '/'){Send $Context 200 'text/html; charset=utf-8' (Page);return};if($method -eq 'GET' -and $path -eq '/providers'){Send $Context 200 'text/html; charset=utf-8' (ProviderPage);return};if($method -eq 'GET' -and $path -eq '/provider-index'){Send $Context 200 'text/html; charset=utf-8' (ProviderIndexPage);return};if($method -eq 'GET' -and $path -eq '/api/locations'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-KnownLocations ([string]$Context.Request.QueryString['riskPool'])) -Depth 3);return};if($method -eq 'GET' -and $path -eq '/api/provider-index'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-ProviderIndexModel) -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/profile-job'){$result=New-JobFromProfile ([string]$Context.Request.QueryString['npi']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/health'){Send $Context 200 'application/json' '{"status":"ok","draft":4}';return};if($method -eq 'GET' -and $path -eq '/api/status'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Status) -Depth 6);return};if($method -eq 'GET' -and $path -eq '/api/provider-sources'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @((Json $script:ConfigPath).sources|Select-Object sourceKey,displayName,providerColumns) -Depth 5);return};if($method -eq 'GET' -and $path -eq '/api/providers'){$sourceKey=[string]$Context.Request.QueryString['sourceKey'];$sw=[Diagnostics.Stopwatch]::StartNew();$result=@(Get-ProviderIndexForPool $sourceKey ([string]$Context.Request.QueryString['riskPool']));Log 'PROVIDER_LIST' 'OK' ($result.Count.ToString()+' names in '+[Math]::Round($sw.Elapsed.TotalSeconds,2)+'s') $sourceKey;Send $Context 200 'application/json' (ConvertTo-Json -InputObject $result -Depth 6);return};if($method -eq 'GET' -and $path -eq '/api/suggest'){$names=@($Context.Request.QueryString.GetValues('name')|Where-Object{$_});$result=@(Get-ProviderSuggestions ([string]$Context.Request.QueryString['sourceKey']) $names ([string]$Context.Request.QueryString['riskPool']));Send $Context 200 'application/json' (ConvertTo-Json -InputObject $result -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/reindex'){$out=@();foreach($s in (Json $script:ConfigPath).sources){if(!(Test-Path -LiteralPath (Join-Path $script:Paths.CanonicalCurrent $s.canonicalFileName))){continue};$r=Update-ProviderIndexAfterImport ([string]$s.sourceKey) -Force;$out+=[ordered]@{sourceKey=$s.sourceKey;displayName=$s.displayName;count=$r.count;seconds=$r.seconds;error=$r.error}};Send $Context 200 'application/json' (ConvertTo-Json -InputObject @($out) -Depth 5);return};if($method -eq 'GET' -and $path -eq '/api/profiles'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-ProviderProfiles) -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/profile'){$result=Save-ProviderProfile (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/html'){Send-JobOutput $Context ([string]$Context.Request.QueryString['jobId']) 'html';return};if($method -eq 'GET' -and $path -eq '/pdf'){Send-JobOutput $Context ([string]$Context.Request.QueryString['jobId']) 'pdf';return};if($method -eq 'GET' -and $path -eq '/flag'){Send $Context 200 'text/html; charset=utf-8' (ConvertTo-AnalysisHtml (Get-FlagReportModel ([string]$Context.Request.QueryString['jobId'])) -Interactive);return};if($method -eq 'GET' -and $path -eq '/communication'){Send $Context 200 'text/html; charset=utf-8' (CommunicationPage);return};if($method -eq 'GET' -and $path -eq '/api/communication'){Send $Context 200 'application/json' ((Get-CommunicationModel)|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/contact-list'){$result=Select-ContactList ([string]$Context.Request.QueryString['path']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/contact-list-upload'){$result=Receive-ContactListUpload $Context;Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 6);return};if($method -eq 'GET' -and $path -eq '/api/recipients'){Send $Context 200 'application/json' ((Get-CommunicationRecipients ([string]$Context.Request.QueryString['riskPool']))|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/campaign'){$result=New-Campaign (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/campaign'){Send $Context 200 'application/json' ((Get-CampaignModel ([string]$Context.Request.QueryString['id']))|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/campaigns'){Send $Context 200 'application/json' ([ordered]@{queue=(Get-QueueStatus);campaigns=@(Get-CampaignList)}|ConvertTo-Json -Depth 9);return};if($method -eq 'GET' -and $path -eq '/api/queue'){Send $Context 200 'application/json' ((Get-QueueStatus)|ConvertTo-Json -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/job-reset'){$result=Reset-RunningJob ([string]$Context.Request.QueryString['jobId']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/campaign-cancel'){$result=Stop-Campaign ([string]$Context.Request.QueryString['id']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/campaign-drafts'){$result=Invoke-CampaignDrafts ([string]$Context.Request.QueryString['id']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/campaign-resume'){$result=Resume-Campaign ([string]$Context.Request.QueryString['id']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/flags'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-ProviderFlags ([string]$Context.Request.QueryString['npi'])) -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/flag'){$result=Save-FlagFromRequest (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/run-next'){$script:RunNext=$true;Send $Context 202 'application/json' '{"status":"scheduled"}';return};if($method -eq 'GET' -and $path -eq '/api/jobs'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-Jobs) -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/job'){$result=New-PreparedJob (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/stage'){$result=Receive-Upload $Context;Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/confirm'){$result=Publish-StagedFile ([string]$Context.Request.QueryString['token']);$index=Update-ProviderIndexAfterImport ([string]$result.sourceKey) -Force;$result['providerCount']=$index.count;$result['indexSeconds']=$index.seconds;$result['indexError']=$index.error;Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/stop'){$script:Stop=$true;Send $Context 202 'application/json' '{"status":"stopping"}';return};Send $Context 404 'application/json' '{"error":"not found"}'}catch{$msg=Safe $_.Exception.Message;Log 'HTTP_REQUEST' 'FAILED' $msg;try{Send $Context 400 'application/json' (([ordered]@{error=$msg}|ConvertTo-Json -Compress))}catch{}}
}


# --- Draft 3: provider indexes, profiles, guided mapping, suggestions, and prepared jobs ---
function Read-BodyJson($Context){$reader=New-Object IO.StreamReader($Context.Request.InputStream,$Context.Request.ContentEncoding);try{$raw=$reader.ReadToEnd()}finally{$reader.Dispose()};if([string]::IsNullOrWhiteSpace($raw)){return $null};return $raw|ConvertFrom-Json}
function Get-SourceConfig([string]$Key){$config=Json $script:ConfigPath;return @($config.sources|Where-Object{$_.sourceKey -eq $Key})[0]}
$script:PoolAliases=@{CRH='CRYSTAL RUN';CMM='CAREMOUNT';PHNY='PROHEALTH';RIV='RIVERSIDE'}
function ConvertTo-RiskPoolName([string]$Value){
 $v=([string]$Value).Trim().ToUpperInvariant()
 if($script:PoolAliases.ContainsKey($v)){return $script:PoolAliases[$v]}
 return $v
}
function Find-IdentityHeader($Source,$Worksheet){
 # Returns @{row=<header row>;map=@{header=column}} for the first row (within the first 25) that carries the provider/NPI/pool identity columns, else $null.
 if($null -eq $Worksheet.Dimension){return $null}
 $limit=[Math]::Min(25,$Worksheet.Dimension.End.Row);$maxCol=$Worksheet.Dimension.End.Column
 $required=@($Source.providerColumns);if($Source.npiColumn){$required+=@([string]$Source.npiColumn)}
 $needsPool=([string]$Source.sourceKey -notlike 'HR-*')
 for($hr=1;$hr -le $limit;$hr++){
  $map=@{}
  for($c=1;$c -le $maxCol;$c++){$t=([string]$Worksheet.Cells[$hr,$c].Text).Trim();if($t -and !$map.ContainsKey($t)){$map[$t]=$c}}
  if($map.Count -eq 0){continue}
  $missing=@($required|Where-Object{!$map.ContainsKey([string]$_)})
  if($missing.Count -gt 0){continue}
  if($needsPool -and !($map.ContainsKey('Risk Pool') -or $map.ContainsKey('Cdo'))){continue}
  return @{row=$hr;map=$map}
 }
 return $null
}
function Get-ColumnText($Worksheet,[int]$Column,[int]$FromRow,[int]$ToRow){
 # One EPPlus range read per column instead of one indexer call per cell; returns a trimmed string per row (index 0 = FromRow).
 $count=$ToRow-$FromRow+1;if($count -le 0){return ,@()}
 $out=New-Object string[] $count;$inv=[Globalization.CultureInfo]::InvariantCulture
 $raw=$Worksheet.Cells[$FromRow,$Column,$ToRow,$Column].Value
 if($raw -is [Array] -and $raw.Rank -eq 2){
  for($i=0;$i -lt $count;$i++){
   $v=$raw[$i,0]
   if($null -eq $v){$out[$i]=''}
   elseif($v -is [string]){$out[$i]=$v.Trim()}
   elseif($v -is [DateTime]){$out[$i]=$v.ToString('o')}
   elseif($v -is [double]){$out[$i]=$v.ToString('0.###############',$inv)}
   else{$out[$i]=([string]$v).Trim()}
  }
 }else{
  $v=$raw
  if($null -eq $v){$out[0]=''}elseif($v -is [DateTime]){$out[0]=$v.ToString('o')}elseif($v -is [double]){$out[0]=$v.ToString('0.###############',$inv)}else{$out[0]=([string]$v).Trim()}
 }
 return ,$out
}
function Get-ProviderIndex([string]$ProviderSourceKey,[switch]$Force){
 $source=Get-SourceConfig $ProviderSourceKey
 if($null -eq $source){throw 'Unknown source key.'}
 $path=Join-Path $script:Paths.CanonicalCurrent $source.canonicalFileName
 $file=Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
 if($null -eq $file){throw ('Canonical file is missing: '+$source.canonicalFileName+'. Import it from the dashboard first.')}
 $stamp=[string]$file.Length+'|'+$file.LastWriteTimeUtc.Ticks
 if(!$Force -and $script:IndexMemo.ContainsKey($ProviderSourceKey) -and $script:IndexMemo[$ProviderSourceKey].stamp -eq $stamp){return $script:IndexMemo[$ProviderSourceKey].items}
 $hash=Get-FileHash256 $path
 $columnSignature=(@($source.providerColumns)-join '|')+'|NPI='+[string]$source.npiColumn+'|POOL=STRICT_FORWARD_FILL|IDENTITY_SCHEMA=1'
 $cache=Join-Path $script:Paths.State ('provider-index-'+$ProviderSourceKey+'.json')
 if(!$Force){
  $cached=Json $cache
  if($cached){
   $p=$cached.PSObject.Properties
   $cacheItems=@($(if($p['items']){$p['items'].Value}else{@()}))
   $valid=($p['sha256'] -and [string]$p['sha256'].Value -eq $hash) -and ($p['indexVersion'] -and [int]$p['indexVersion'].Value -eq $script:IndexVersion) -and ($p['columnSignature'] -and [string]$p['columnSignature'].Value -eq $columnSignature) -and ($p['count'] -and [int]$p['count'].Value -eq $cacheItems.Count)
   if($valid){
    $script:IndexMemo[$ProviderSourceKey]=@{stamp=$stamp;hash=$hash;items=$cacheItems;count=$cacheItems.Count;builtUtc=$(if($p['builtUtc']){[string]$p['builtUtc'].Value}else{''})}
    return $cacheItems
   }
  }
 }
 $sw=[Diagnostics.Stopwatch]::StartNew();$package=$null;$items=@{};$found=$false
 try{
  $package=Open-ExcelPackage -Path $path -ErrorAction Stop
  foreach($worksheet in $package.Workbook.Worksheets){
   $header=Find-IdentityHeader $source $worksheet
   if($null -eq $header){continue}
   $found=$true;$map=$header.map;$first=$header.row+1;$last=$worksheet.Dimension.End.Row
   if($last -lt $first){break}
   $count=$last-$first+1
   $providerValues=@()
   foreach($providerHeader in @($source.providerColumns)){if($map.ContainsKey([string]$providerHeader)){$providerValues+=,(Get-ColumnText $worksheet $map[[string]$providerHeader] $first $last)}}
   $npiValues=$null
   if($source.npiColumn -and $map.ContainsKey([string]$source.npiColumn)){$npiValues=Get-ColumnText $worksheet $map[[string]$source.npiColumn] $first $last}
   $poolValues=$null
   foreach($poolHeader in @('Risk Pool','Cdo')){if($map.ContainsKey($poolHeader)){$poolValues=Get-ColumnText $worksheet $map[$poolHeader] $first $last;break}}
   $currentRiskPool='';$poolAliases=$script:PoolAliases
   for($i=0;$i -lt $count;$i++){
    $npi=$(if($null -ne $npiValues){$npiValues[$i]}else{''})
    if($null -ne $poolValues){$rowRiskPool=$poolValues[$i].ToUpperInvariant();if($poolAliases.ContainsKey($rowRiskPool)){$rowRiskPool=$poolAliases[$rowRiskPool]};if($rowRiskPool){$currentRiskPool=$rowRiskPool}}
    foreach($values in $providerValues){
     $providerName=$values[$i]
     if(!$providerName){continue}
     if(!$items.ContainsKey($providerName)){$items[$providerName]=@{name=$providerName;npi=$npi;pools=New-Object Collections.Generic.List[string]}}
     $entry=$items[$providerName]
     if(!$entry.npi -and $npi){$entry.npi=$npi}
     if($currentRiskPool -and !$entry.pools.Contains($currentRiskPool)){$entry.pools.Add($currentRiskPool)}
    }
   }
   break
  }
  if(!$found){throw ('Provider identity schema was not found for '+$source.displayName+'. Required provider/NPI and risk-pool columns are missing from the canonical file. Re-import that source.')}
 }finally{
  if($package){Close-ExcelPackage $package -NoSave}
 }
 $sorted=@($items.Values|ForEach-Object{[pscustomobject][ordered]@{name=[string]$_.name;npi=[string]$_.npi;riskPools=@($_.pools.ToArray())}}|Sort-Object -Property name)
 $builtUtc=[DateTime]::UtcNow.ToString('o')
 Save-JsonAtomic $cache ([ordered]@{indexVersion=$script:IndexVersion;columnSignature=$columnSignature;sourceKey=$ProviderSourceKey;sha256=$hash;builtUtc=$builtUtc;count=$sorted.Count;items=$sorted})
 $script:IndexMemo[$ProviderSourceKey]=@{stamp=$stamp;hash=$hash;items=$sorted;count=$sorted.Count;builtUtc=$builtUtc}
 Log 'PROVIDER_INDEX_BUILT' 'OK' ($sorted.Count.ToString()+' providers in '+[Math]::Round($sw.Elapsed.TotalSeconds,1)+'s') $ProviderSourceKey
 return $sorted
}
function Get-ProviderIndexForPool([string]$ProviderSourceKey,[string]$RiskPool){
 $all=@(Get-ProviderIndex $ProviderSourceKey)
 if(!$RiskPool -or $ProviderSourceKey -like 'HR-*'){return $all}
 $normalized=ConvertTo-RiskPoolName $RiskPool
 return @($all|Where-Object{$p=$_.PSObject.Properties['riskPools'];$null -ne $p -and @($p.Value) -contains $normalized})
}
function ConvertTo-NormalProvider([string]$Name){return (($Name.ToUpperInvariant() -replace '[^A-Z0-9 ]',' ' -replace '\s+',' ').Trim())}
$script:CredentialTokens=@('MD','DO','NP','PA','PAC','APRN','APN','FNP','ANP','AGNP','AGPCNP','CNP','DNP','ARNP','CRNP','NPC','PHD','DPM','MBBS','RN','RPA','DR','JR','SR','II','III','FACP','FAAFP','FACOG','BC')
function Get-ProviderTokens([string]$Name){
 if($script:TokenCache.ContainsKey($Name)){return ,$script:TokenCache[$Name]}
 $out=New-Object Collections.Generic.List[string];$previousWasCredential=$false
 foreach($t in (ConvertTo-NormalProvider $Name).Split(' ')){
  if(!$t){continue}
  $isCredential=($script:CredentialTokens -contains $t) -or ($t -eq 'C' -and $previousWasCredential)   # FNP-C / PA-C / NP-C
  if(!$isCredential){$out.Add($t)}
  $previousWasCredential=$isCredential
 }
 $arr=$out.ToArray();$script:TokenCache[$Name]=$arr;return ,$arr
}
function Get-TokenScore([string[]]$At,[string[]]$Bt){
 # Token overlap after credential stripping: full-token matches count 1, an initial matching the first letter of an unmatched token counts 0.5. 100 = same name tokens in any order.
 if($At.Count -eq 0 -or $Bt.Count -eq 0){return 0}
 $sa=[string[]]$At.Clone();$sb=[string[]]$Bt.Clone();[Array]::Sort($sa);[Array]::Sort($sb)
 if(($sa -join ' ') -eq ($sb -join ' ')){return 100}
 $credit=0.0;$usedA=New-Object Collections.Generic.HashSet[int];$usedB=New-Object Collections.Generic.HashSet[int]
 for($x=0;$x -lt $At.Count;$x++){if($At[$x].Length -le 1){continue};for($y=0;$y -lt $Bt.Count;$y++){if($Bt[$y].Length -gt 1 -and !$usedB.Contains($y) -and $Bt[$y] -eq $At[$x]){$credit+=1;[void]$usedA.Add($x);[void]$usedB.Add($y);break}}}
 for($x=0;$x -lt $At.Count;$x++){if($At[$x].Length -ne 1){continue};for($y=0;$y -lt $Bt.Count;$y++){if($Bt[$y].Length -gt 1 -and !$usedB.Contains($y) -and $Bt[$y].StartsWith($At[$x])){$credit+=0.5;[void]$usedB.Add($y);break}}}
 for($y=0;$y -lt $Bt.Count;$y++){if($Bt[$y].Length -ne 1){continue};for($x=0;$x -lt $At.Count;$x++){if($At[$x].Length -gt 1 -and !$usedA.Contains($x) -and $At[$x].StartsWith($Bt[$y])){$credit+=0.5;[void]$usedA.Add($x);break}}}
 return [Math]::Round((200*$credit/[Math]::Max(1,$At.Count+$Bt.Count)),1)
}
function Get-ProviderScore([string]$A,[string]$B){return Get-TokenScore (Get-ProviderTokens $A) (Get-ProviderTokens $B)}
function Get-ProviderSuggestions([string]$ProviderSourceKey,[string[]]$Names,[string]$RiskPool){
 # Each candidate is scored against every seed name (aliases already confirmed in other sources) and keeps its best score.
 $seeds=@();foreach($n in @($Names)){if(!$n){continue};$t=Get-ProviderTokens ([string]$n);if($t.Count -gt 0){$seeds+=,$t}}
 if($seeds.Count -eq 0){return @()}
 $scored=New-Object Collections.Generic.List[object]
 foreach($item in @(Get-ProviderIndexForPool $ProviderSourceKey $RiskPool)){
  $candidate=Get-ProviderTokens ([string]$item.name);if($candidate.Count -eq 0){continue}
  $best=0.0;foreach($seed in $seeds){$s=Get-TokenScore $seed $candidate;if($s -gt $best){$best=$s}}
  if($best -gt 0){$scored.Add([ordered]@{name=[string]$item.name;npi=[string]$item.npi;score=$best})}
 }
 return @($scored|Sort-Object -Property @{Expression='score';Descending=$true},@{Expression='name';Ascending=$true}|Select-Object -First 10)
}
function Get-NpiForAlias([string]$ProviderSourceKey,[string]$Alias){if(!$Alias){return $null};$item=@(Get-ProviderIndex $ProviderSourceKey|Where-Object{$_.name -eq $Alias})[0];if($item){return [string]$item.npi};return $null}
function Save-ProviderProfile($Body){if($null -eq $Body -or $null -eq $Body.aliases){throw 'Profile aliases are required.'};$exportNpi=Get-NpiForAlias 'Export' ([string]$Body.aliases.Export);$qualityNpi=Get-NpiForAlias 'PtListQuality' ([string]$Body.aliases.PtListQuality);if(!$exportNpi -or !$qualityNpi){throw 'Export and PtListQuality selections must both contain an NPI.'};if($exportNpi -ne $qualityNpi){throw ('NPI conflict: Export '+$exportNpi+' versus PtListQuality '+$qualityNpi)};$path=Join-Path $script:Paths.Profiles ($exportNpi+'.json');$old=Json $path;$providerProfile=[ordered]@{profileVersion=1;profileId=$exportNpi;npi=$exportNpi;displayName=$(if($Body.displayName){[string]$Body.displayName}else{[string]$Body.aliases.Export});riskPool=[string]$Body.riskPool;location=([string](Get-P $Body 'location' '')).Trim();aliases=$Body.aliases;createdUtc=$(if($old){$old.createdUtc}else{[DateTime]::UtcNow.ToString('o')});updatedUtc=[DateTime]::UtcNow.ToString('o')};if(Test-Path $path){Copy-Item $path ($path+'.'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.bak')};Save-JsonAtomic $path $providerProfile;Add-KnownLocation $providerProfile.riskPool $providerProfile.location;Log 'PROFILE_SAVED' 'OK' ('NPI '+$exportNpi);return $providerProfile}
function Get-ProviderProfiles{return @(Get-ChildItem $script:Paths.Profiles -Filter '*.json' -File -ErrorAction SilentlyContinue|Where-Object{$_.Name -notlike '*.bak'}|ForEach-Object{Json $_.FullName}|Sort-Object displayName)}
function New-PreparedJob($Body){if($null -eq $Body.aliases){throw 'Confirmed aliases are required.'};$exportNpi=Get-NpiForAlias 'Export' ([string]$Body.aliases.Export);$qualityNpi=Get-NpiForAlias 'PtListQuality' ([string]$Body.aliases.PtListQuality);if(!$exportNpi -or !$qualityNpi -or $exportNpi -ne $qualityNpi){throw 'Export/PtListQuality NPI is missing or conflicting; job blocked.'};$job=[ordered]@{jobId=[Guid]::NewGuid().ToString('N');providerKey=$exportNpi;displayName=$(if($Body.displayName){[string]$Body.displayName}else{[string]$Body.aliases.Export});riskPool=[string]$Body.riskPool;location=([string](Get-P $Body 'location' '')).Trim();aliases=$Body.aliases;queuedUtc=[DateTime]::UtcNow.ToString('o');state='Prepared';percent=0;stage='Queued for Draft 4 analysis'};$jobPath=Join-Path $script:Paths.State ('job-'+$job.jobId+'.json');Save-JsonAtomic $jobPath $job;Log 'JOB_PREPARED' 'OK' ('NPI '+$exportNpi);return $job}
function Get-Jobs{return @(Get-ChildItem $script:Paths.State -Filter 'job-*.json' -File -ErrorAction SilentlyContinue|ForEach-Object{Json $_.FullName}|Sort-Object queuedUtc -Descending)}
$script:TabCss='.tabs{display:flex;gap:4px;background:#17365d;padding:0 22px}.tabs a{color:#cfe0f5;text-decoration:none;padding:10px 18px;border-radius:8px 8px 0 0;font-weight:600}.tabs a:hover{background:#274b7a;color:#fff}.tabs a.active{background:#f4f7fb;color:#17365d}'
function Get-NavHtml([string]$Active){
 $tabs=@(@('sources','/','Sources'),@('wizard','/providers','Provider Wizard'),@('index','/provider-index','Provider Index'),@('communication','/communication','Communication'))
 $links=foreach($t in $tabs){'<a href="'+$t[1]+'"'+$(if($t[0] -eq $Active){' class="active"'}else{''})+'>'+$t[2]+'</a>'}
 return '<nav class="tabs">'+($links -join '')+'</nav>'
}
function Get-KnownLocations([string]$RiskPool){
 # Union of the per-pool location store and every saved profile's location, de-duplicated case-insensitively.
 $pool=ConvertTo-RiskPoolName $RiskPool;if(!$pool){return @()}
 $values=@();$store=Json (Join-Path $script:Paths.State 'locations.json')
 if($store){$p=$store.PSObject.Properties['pools'];if($p -and $p.Value){$q=$p.Value.PSObject.Properties[$pool];if($q){$values+=@($q.Value)}}}
 foreach($profile in @(Get-ProviderProfiles)){if((ConvertTo-RiskPoolName ([string](Get-P $profile 'riskPool' ''))) -eq $pool){$values+=[string](Get-P $profile 'location' '')}}
 $seen=@{};$out=New-Object Collections.Generic.List[string]
 foreach($v in $values){$t=([string]$v).Trim();if(!$t){continue};$k=$t.ToUpperInvariant();if($seen.ContainsKey($k)){continue};$seen[$k]=$true;$out.Add($t)}
 return @($out|Sort-Object)
}
function Add-KnownLocation([string]$RiskPool,[string]$Location){
 $pool=ConvertTo-RiskPoolName $RiskPool;$loc=([string]$Location).Trim();if(!$pool -or !$loc){return}
 $path=Join-Path $script:Paths.State 'locations.json';$store=Json $path;$pools=[ordered]@{}
 if($store){$p=$store.PSObject.Properties['pools'];if($p -and $p.Value){foreach($prop in $p.Value.PSObject.Properties){$pools[$prop.Name]=@($prop.Value)}}}
 $current=@($(if($pools.Contains($pool)){$pools[$pool]}else{@()}))
 if(@($current|Where-Object{[string]$_ -ieq $loc}).Count -eq 0){$current+=$loc}
 $pools[$pool]=@($current|Sort-Object)
 Save-JsonAtomic $path ([ordered]@{version=1;updatedUtc=[DateTime]::UtcNow.ToString('o');pools=$pools})
}
function New-JobFromProfile([string]$Npi){
 if($Npi -notmatch '^\d{10}$'){throw 'Invalid NPI.'}
 $profile=Json (Join-Path $script:Paths.Profiles ($Npi+'.json'));if($null -eq $profile){throw 'Profile not found.'}
 return New-PreparedJob $profile
}
function Resolve-JobOutputPath($Job,[string]$Type){
 # Outputs move from providers-current to providers-archive when a newer run is published; follow them there.
 $path=[string](Get-P $Job ($Type+'Path') '');if(!$path){return ''}
 if(Test-Path -LiteralPath $path){return $path}
 $archiveDir=Join-Path $script:Paths.ProvidersArchive (ConvertTo-ProviderSlug ([string](Get-P $Job 'displayName' '')))
 if(!(Test-Path -LiteralPath $archiveDir)){return ''}
 $name=[IO.Path]::GetFileName($path);$base=[IO.Path]::GetFileNameWithoutExtension($name);$ext=[IO.Path]::GetExtension($name)
 $candidates=@(Get-ChildItem -LiteralPath $archiveDir -File -Filter ($base+'*'+$ext) -ErrorAction SilentlyContinue|Sort-Object LastWriteTimeUtc -Descending)
 if($candidates.Count -gt 0){return $candidates[0].FullName}
 return ''
}
function Get-ProviderIndexModel{
 $jobs=@(Get-Jobs);$groups=@{}
 foreach($profile in @(Get-ProviderProfiles)){
  $npi=[string](Get-P $profile 'npi' '');if(!$npi){continue}
  $pool=ConvertTo-RiskPoolName ([string](Get-P $profile 'riskPool' ''));if(!$pool){$pool='UNASSIGNED'}
  if(!$groups.ContainsKey($pool)){$groups[$pool]=New-Object Collections.Generic.List[object]}
  $reports=New-Object Collections.Generic.List[object];$active=$null
  foreach($job in @($jobs|Where-Object{[string](Get-P $_ 'providerKey' '') -eq $npi})){
   $state=[string](Get-P $job 'state' '')
   $entry=[ordered]@{jobId=[string]$job.jobId;state=$state;percent=(Get-P $job 'percent' 0);stage=[string](Get-P $job 'stage' '');queuedUtc=[string](Get-P $job 'queuedUtc' '');completedUtc=[string](Get-P $job 'completedUtc' '');errorSummary=[string](Get-P $job 'errorSummary' '');location=[string](Get-P $job 'location' '');htmlUrl=$null;pdfUrl=$null;flagUrl=$null}
   if($state -eq 'Completed'){if(Resolve-JobOutputPath $job 'html'){$entry.htmlUrl='/html?jobId='+$job.jobId};if(Resolve-JobOutputPath $job 'pdf'){$entry.pdfUrl='/pdf?jobId='+$job.jobId};if(Test-Path -LiteralPath (Join-Path $script:Paths.State ('analysis-'+$job.jobId+'.json'))){$entry.flagUrl='/flag?jobId='+$job.jobId};$reports.Add($entry)}
   elseif($state -eq 'Failed'){$reports.Add($entry)}
   elseif($null -eq $active){$active=$entry}
  }
  $groups[$pool].Add([ordered]@{npi=$npi;displayName=[string](Get-P $profile 'displayName' $npi);location=[string](Get-P $profile 'location' '');riskPool=$pool;updatedUtc=[string](Get-P $profile 'updatedUtc' '');reports=@($reports.ToArray());active=$active})
 }
 $out=@()
 foreach($pool in @($groups.Keys|Sort-Object)){$out+=[ordered]@{riskPool=$pool;providers=@($groups[$pool].ToArray()|Sort-Object -Property @{Expression={[string]$_.displayName}})}}
 return $out
}
function ProviderPage{
 $html=@'
<!doctype html><html><head><meta charset="utf-8"><title>Provider Wizard</title>
<style>
body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:white;padding:22px 22px 12px}main{padding:22px;max-width:960px}
__TABCSS__
.card{background:white;border:1px solid #dce4ef;border-radius:8px;padding:18px;margin:14px 0}
select,input[type=text],input:not([type]){padding:8px;width:100%;box-sizing:border-box;margin:6px 0}select[size]{height:auto}
button{padding:8px 12px;margin:6px 6px 6px 0;cursor:pointer}button:disabled{opacity:.5;cursor:default}
.map{display:grid;grid-template-columns:220px 1fr auto;gap:5px 12px;align-items:center;margin-top:12px}.map .current{font-weight:600}
.muted{color:#667085}.error{color:#a61b1b}label{display:block;margin-top:8px;font-weight:600}label.inline{display:inline;font-weight:normal}
.progress{height:10px;background:#e4e9f0;border-radius:6px;overflow:hidden;margin:7px 0}.progress span{display:block;height:100%;background:#1769aa;transition:width .3s}
.job{padding:8px 0;border-bottom:1px solid #e5e9f0}
</style></head><body>
<header><h1>Provider Analysis Wizard</h1><p>Draft 4.7 - map a provider across sources, save the profile, and prepare a report</p></header>__NAV__
<main>
<div id="startCard" class="card"><h2>Select risk pool and location</h2>
<label>Risk pool</label><select id="riskPoolSelect"><option value="">Choose risk pool...</option><option>CRYSTAL RUN</option><option>PROHEALTH</option><option>CAREMOUNT</option><option>RIVERSIDE</option></select>
<label>Location</label><input id="locationInput" type="text" list="locationList" placeholder="Choose a known location for this risk pool, or type a new one" autocomplete="off" disabled><datalist id="locationList"></datalist>
<div id="locationNote" class="muted">Choose a risk pool first.</div>
<button id="startWizard" disabled>Start provider mapping</button>
<p class="muted">Provider lists will be limited to the selected risk pool. Crystal Run uses HR-CRH; the other pools use HR-RIVPHNYCMM. Locations you enter are remembered for that risk pool and offered next time.</p></div>
<div id="wizardCard" class="card" style="display:none">
<div id="context" class="muted"></div>
<label>Saved profile</label><select id="profile"><option value="">New mapping</option></select>
<h2 id="step">Loading sources...</h2><div id="fieldNote" class="muted"></div>
<div id="pickCard">
<label>Filter names</label><input id="filter" type="text" placeholder="Type part of a name to narrow the list..." autocomplete="off">
<label>Provider name from this source <span id="nameCount" class="muted"></span></label><select id="names" size="10"></select>
<label>Suggested matches (optional; selecting one changes the provider above)</label><select id="suggest"><option value="">Choose a suggestion...</option></select>
<p class="muted">Only configured provider fields are listed; patient-name fields are excluded. Suggestions are scored against the names you have already confirmed for other sources. Nothing is confirmed until you click the button, although an exact match is preselected for you.</p>
<button id="ok">Confirm selected provider for this source</button><button id="blank">Blank / skip</button><button id="retry" style="display:none">Retry loading</button>
</div>
<div id="mapping" class="map"></div>
</div>
<div id="finish" class="card" style="display:none"><label>Display name</label><input id="displayName" type="text"><label>Location</label><input id="finishLocation" type="text" list="locationList" autocomplete="off"><label class="inline"><input id="save" type="checkbox" style="width:auto"> Save/update provider profile</label><br><button id="queue">Validate mapping and prepare job</button><button id="saveOnly">Save profile only</button><button id="another" style="display:none">Select another provider</button><div id="result"></div></div>
<div class="card"><h2>Prepared jobs</h2><div id="jobs"></div></div>
</main>
<script>
let allSources=[],sources=[],riskPool='',providerLocation='',i=0,aliases={},items=[],profiles=[],runner=false,jobsBusy=false,loadToken=0,complete=false;
const el=id=>document.getElementById(id);
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
async function api(u,o){const r=await fetch(u,o);const text=await r.text();let j=null;try{j=text?JSON.parse(text):null}catch(e){throw Error('Server returned invalid JSON for '+u+': '+text.slice(0,200))}if(!r.ok)throw Error((j&&j.error)||('Request failed ('+r.status+')'));if(j===null)throw Error('Server returned an empty response for '+u);return j}
function poolSources(){return allSources.filter(s=>riskPool==='CRYSTAL RUN'?s.sourceKey!=='HR-RIVPHNYCMM':s.sourceKey!=='HR-CRH')}
function otherHrKey(){return riskPool==='CRYSTAL RUN'?'HR-RIVPHNYCMM':'HR-CRH'}
function setNote(t,isError){const n=el('fieldNote');n.textContent=t;n.className=isError?'error':'muted'}
function fillNames(){const q=el('filter').value.trim().toLowerCase();const names=el('names');const keep=names.value;names.innerHTML='';let shown=0;items.forEach(x=>{if(!q||x.name.toLowerCase().includes(q)){names.add(new Option(x.name+(x.npi?'   ('+x.npi+')':''),x.name));shown++}});if(keep&&[...names.options].some(o=>o.value===keep))names.value=keep;el('nameCount').textContent=items.length?'('+shown+' of '+items.length+')':''}
function selectName(v){el('filter').value='';fillNames();el('names').value=v}
function startEnabled(){el('startWizard').disabled=!(el('riskPoolSelect').value&&el('locationInput').value.trim())}
async function loadLocations(pool){const list=el('locationList');list.innerHTML='';if(!pool)return[];try{const locs=await api('/api/locations?riskPool='+encodeURIComponent(pool));locs.forEach(l=>list.appendChild(new Option(l)));el('locationNote').textContent=locs.length?locs.length+' known location'+(locs.length===1?'':'s')+' for '+pool+': '+locs.join(', ')+'. Pick one or type a new one.':'No locations saved for '+pool+' yet; type the first one.';return locs}catch(e){el('locationNote').textContent='Could not load locations: '+e.message;return[]}}
function showContext(){el('context').textContent='Risk pool: '+riskPool+'  |  Location: '+(providerLocation||'(none)')}
function openWizard(){sources=poolSources();if(complete)i=sources.length;el('startCard').style.display='none';el('wizardCard').style.display='block';el('finishLocation').value=providerLocation;showContext();load()}
async function init(){
 allSources=await api('/api/provider-sources');profiles=await api('/api/profiles');
 const profile=el('profile');profile.innerHTML='<option value="">New mapping</option>'+profiles.map((p,n)=>'<option value="'+n+'">'+esc(p.displayName)+' ('+esc(p.npi)+')'+(p.location?' - '+esc(p.location):'')+'</option>').join('');
 profile.onchange=()=>{if(profile.value!==''){applyProfile(profiles[+profile.value])}load()};
 el('riskPoolSelect').onchange=async()=>{const pool=el('riskPoolSelect').value;el('locationInput').disabled=!pool;el('locationInput').value='';startEnabled();await loadLocations(pool)};
 el('locationInput').oninput=startEnabled;
 el('startWizard').onclick=()=>{riskPool=el('riskPoolSelect').value;providerLocation=el('locationInput').value.trim();if(!riskPool||!providerLocation)return;aliases={};aliases[otherHrKey()]='';i=0;complete=false;openWizard()};
 el('filter').oninput=fillNames;
 el('names').onchange=()=>refreshSuggestions(false);
 el('suggest').onchange=()=>{const s=el('suggest');if(s.value){selectName(s.value);s.value=''}};
 el('ok').onclick=()=>{const v=el('names').value;if(!v)return;aliases[sources[i].sourceKey]=v;advance()};
 el('blank').onclick=()=>{aliases[sources[i].sourceKey]='';advance()};
 el('retry').onclick=()=>load();
 el('queue').onclick=queue;el('saveOnly').onclick=saveProfileOnly;el('another').onclick=reset;
 const wanted=new URLSearchParams(window.location.search).get('profile');
 if(wanted){const n=profiles.findIndex(p=>p.npi===wanted);if(n>=0){profile.value=String(n);applyProfile(profiles[n]);aliases[otherHrKey()]=aliases[otherHrKey()]||'';openWizard()}}
}
function applyProfile(p){if(p.riskPool&&p.riskPool!==riskPool){riskPool=p.riskPool;el('riskPoolSelect').value=riskPool;sources=poolSources();loadLocations(riskPool)}if(p.location){providerLocation=p.location;el('finishLocation').value=providerLocation}aliases=Object.assign({},p.aliases||{});el('displayName').value=p.displayName||'';el('save').checked=true;complete=true;i=sources.length;showContext()}
function advance(){loadToken++;if(complete){i=sources.length}else{i++}load()}
async function load(){
 const my=++loadToken;
 if(i>=sources.length){complete=true;el('finish').style.display='block';el('pickCard').style.display='none';el('step').textContent=el('profile').value!==''?'Saved profile loaded':'Mapping complete';setNote(el('profile').value!==''?'Every source is already confirmed from the saved profile. Use edit to change one, update the display name or location, then save the profile (with or without running a report).':'Review the mapping below (use edit to change a source), confirm the location, then validate and prepare the job.');if(!el('displayName').value)el('displayName').value=aliases.Export||'';if(!el('finishLocation').value)el('finishLocation').value=providerLocation;render();return}
 el('finish').style.display=complete?'block':'none';el('pickCard').style.display='';
 const s=sources[i];el('step').textContent=(i+1)+' of '+sources.length+': '+s.displayName;
 const fields='Provider field'+(s.providerColumns.length>1?'s':'')+': '+s.providerColumns.join(' / ');
 setNote(fields+' - loading distinct names (the first load after importing a large file can take a minute)...');
 items=[];el('filter').value='';fillNames();el('suggest').innerHTML='<option value="">Choose a suggestion...</option>';el('ok').disabled=true;el('retry').style.display='none';render();
 try{const list=await api('/api/providers?sourceKey='+encodeURIComponent(s.sourceKey)+'&riskPool='+encodeURIComponent(riskPool));if(my!==loadToken)return;items=list}
 catch(e){if(my!==loadToken)return;setNote('Could not load provider names for '+s.displayName+': '+e.message+'  You can retry, or skip this source with Blank / skip.',true);el('retry').style.display='inline-block';return}
 items.sort((a,b)=>a.name.localeCompare(b.name,undefined,{sensitivity:'base'}));
 el('ok').disabled=false;
 if(!items.length){setNote(fields+' - no provider names found'+(riskPool&&!s.sourceKey.startsWith('HR-')?' for '+riskPool:'')+' in this source. Skip it, or check the imported file on the Sources tab.',true)}
 else{setNote(fields+' - '+items.length+' distinct provider names'+(riskPool&&!s.sourceKey.startsWith('HR-')?' in '+riskPool:''))}
 fillNames();if(aliases[s.sourceKey])selectName(aliases[s.sourceKey]);
 await refreshSuggestions(true);
}
async function refreshSuggestions(preselect){
 const s=sources[i];if(!s||!items.length)return;const my=loadToken;const sug=el('suggest');sug.innerHTML='<option value="">Choose a suggestion...</option>';
 const seeds=Object.values(aliases).filter(v=>v);const dn=el('displayName').value.trim();if(dn)seeds.push(dn);if(!seeds.length&&el('names').value)seeds.push(el('names').value);if(!seeds.length)return;
 try{
  const q=await api('/api/suggest?sourceKey='+encodeURIComponent(s.sourceKey)+'&riskPool='+encodeURIComponent(riskPool)+seeds.map(x=>'&name='+encodeURIComponent(x)).join(''));
  if(my!==loadToken)return;
  q.slice(0,8).forEach(x=>sug.add(new Option(x.name+' - match '+x.score,x.name)));
  if(preselect&&!el('names').value&&q.length&&q[0].score>=100){selectName(q[0].name);setNote(el('fieldNote').textContent+' - exact match preselected; confirm it or choose another.',false)}
 }catch(e){}
}
function render(){
 const m=el('mapping');
 m.innerHTML=sources.map((s,n)=>{const v=aliases[s.sourceKey];const txt=v?esc(v):(s.sourceKey in aliases?'<em>Blank</em>':'<em class="muted">Pending</em>');return '<b'+(n===i&&!complete?' class="current"':'')+'>'+esc(s.displayName)+'</b><span>'+txt+'</span><a href="#" data-n="'+n+'">edit</a>'}).join('');
 [...m.querySelectorAll('a[data-n]')].forEach(a=>a.onclick=e=>{e.preventDefault();i=+a.dataset.n;load()});
}
async function queue(){
 const result=el('result');
 try{
  providerLocation=el('finishLocation').value.trim();if(!providerLocation){result.textContent='Error: a location is required.';return}showContext();
  const body={displayName:el('displayName').value,aliases,riskPool,location:providerLocation};
  if(el('save').checked&&el('profile').value!==''&&!confirm('Overwrite this saved provider profile?'))return;
  el('queue').disabled=true;result.textContent='Validating mapping...';
  if(el('save').checked)await api('/api/profile',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  const j=await api('/api/job',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  result.textContent='Job '+j.state+': '+j.displayName+'. The analysis runs in the background; progress appears under Prepared jobs and on the Provider Index tab.';
  el('another').style.display='inline-block';await jobs();
 }catch(e){result.textContent='Error: '+e.message}
 finally{el('queue').disabled=false}
}
async function reloadProfiles(selectNpi){profiles=await api('/api/profiles');const profile=el('profile');profile.innerHTML='<option value="">New mapping</option>'+profiles.map((p,n)=>'<option value="'+n+'">'+esc(p.displayName)+' ('+esc(p.npi)+')'+(p.location?' - '+esc(p.location):'')+'</option>').join('');const n=profiles.findIndex(p=>p.npi===selectNpi);profile.value=n>=0?String(n):''}
async function saveProfileOnly(){
 const result=el('result');
 try{
  providerLocation=el('finishLocation').value.trim();if(!providerLocation){result.textContent='Error: a location is required.';return}showContext();
  const body={displayName:el('displayName').value,aliases,riskPool,location:providerLocation};
  if(el('profile').value!==''&&!confirm('Overwrite this saved provider profile?'))return;
  el('saveOnly').disabled=true;result.textContent='Saving profile...';
  const p=await api('/api/profile',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  await reloadProfiles(p.npi);result.textContent='Profile saved: '+p.displayName+' (NPI '+p.npi+') at '+p.location+'. No report was generated; use Generate report on the Provider Index tab when you want one.';
  el('another').style.display='inline-block';
 }catch(e){result.textContent='Error: '+e.message}
 finally{el('saveOnly').disabled=false}
}
function reset(){loadToken++;i=0;aliases={};items=[];sources=[];riskPool='';providerLocation='';complete=false;el('riskPoolSelect').value='';el('locationInput').value='';el('locationInput').disabled=true;el('locationNote').textContent='Choose a risk pool first.';startEnabled();el('profile').value='';el('displayName').value='';el('finishLocation').value='';el('save').checked=false;el('finish').style.display='none';el('pickCard').style.display='';el('another').style.display='none';el('result').textContent='';el('mapping').innerHTML='';el('wizardCard').style.display='none';el('startCard').style.display='block'}
async function jobs(){
 if(jobsBusy)return;jobsBusy=true;
 try{
  const j=await api('/api/jobs');
  el('jobs').innerHTML=j.map(x=>'<div class="job"><b>'+esc(x.displayName)+'</b>'+(x.location?' <span class="muted">'+esc(x.location)+'</span>':'')+' - '+esc(x.state)+' '+(x.percent||0)+'%<div class="progress"><span style="width:'+Math.max(0,Math.min(100,x.percent||0))+'%"></span></div><span class="muted">'+esc(x.stage)+'</span>'+(x.state==='Completed'?' &middot; <a target="_blank" href="/html?jobId='+esc(x.jobId)+'">Open HTML</a> &middot; <a target="_blank" href="/pdf?jobId='+esc(x.jobId)+'">Open PDF</a>'+(x.flagUrl?' &middot; <a target="_blank" href="'+esc(x.flagUrl)+'">Flag</a>':''):'')+(x.errorSummary?'<br><span class="error">'+esc(x.errorSummary)+'</span>':'')+(x.warnings&&x.warnings.length?'<br><span class="muted">Warning: '+esc([].concat(x.warnings).join(' | '))+'</span>':'')+'</div>').join('')||'No jobs yet.';
  if(!runner&&j.some(x=>x.state==='Prepared'||x.state==='Queued')){runner=true;try{await fetch('/api/run-next',{method:'POST'})}finally{runner=false}}
 }catch(e){}
 finally{jobsBusy=false}
}
init().then(()=>{jobs();setInterval(jobs,2000)}).catch(e=>{document.body.insertAdjacentHTML('beforeend','<p class="error" style="padding:22px">Error: '+esc(e.message)+'</p>')});
</script></body></html>
'@
 return $html.Replace('__NAV__',(Get-NavHtml 'wizard')).Replace('__TABCSS__',$script:TabCss)
}
function ProviderIndexPage{
 $html=@'
<!doctype html><html><head><meta charset="utf-8"><title>Provider Index</title>
<style>
body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:white;padding:22px 22px 12px}main{padding:22px;max-width:1100px}
__TABCSS__
.card{background:white;border:1px solid #dce4ef;border-radius:8px;padding:18px;margin:14px 0}.muted{color:#667085}.error{color:#a61b1b}
details.pool{border:1px solid #dce4ef;border-radius:8px;margin:10px 0;background:#fff}details.pool>summary{cursor:pointer;padding:12px 16px;font-size:16px;background:#eaf1f8;border-radius:8px}details.pool[open]>summary{border-radius:8px 8px 0 0}
details.loc{margin:8px 12px;border:1px solid #e5e9f0;border-radius:6px;background:#fff}details.loc>summary{cursor:pointer;padding:8px 12px;font-weight:600;color:#17365d;background:#f7f9fc;border-radius:6px}details.loc[open]>summary{border-radius:6px 6px 0 0}
.provider{padding:12px 16px;border-top:1px solid #e5e9f0}.row{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
button,a.btn{padding:7px 12px;cursor:pointer;background:#1769aa;color:#fff;border:0;border-radius:5px;text-decoration:none;font:inherit}button:disabled{opacity:.5;cursor:default}a.btn{background:#e4e9f0;color:#172033}
details.reports{margin-top:8px}details.reports>summary{cursor:pointer;color:#17365d}.report{padding:6px 0 6px 14px;border-left:3px solid #dce4ef;margin:4px 0}
.progress{height:10px;background:#e4e9f0;border-radius:6px;overflow:hidden;margin:7px 0}.progress span{display:block;height:100%;background:#1769aa;transition:width .3s}
</style></head><body>
<header><h1>Provider Index</h1><p>Draft 5.1 - saved provider profiles by risk pool, with every generated report and patient flags</p></header>__NAV__
<main><div class="card"><p class="muted">Every provider with a saved profile, grouped by risk pool. <b>Generate report</b> runs a fresh analysis with the saved mapping and location; <b>Edit profile</b> opens it in the wizard; <b>Flag</b> opens a report in flag mode to exclude patients who are not current or whose data is wrong. Reports are listed newest first.</p><div id="index">Loading...</div></div></main>
<script>
const el=id=>document.getElementById(id);let runner=false,busy=false,lastJson='',openState={};
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
async function api(u,o){const r=await fetch(u,o);const text=await r.text();let j=null;try{j=text?JSON.parse(text):null}catch(e){throw Error('Server returned invalid JSON: '+text.slice(0,200))}if(!r.ok)throw Error((j&&j.error)||('Request failed ('+r.status+')'));return j}
function when(iso){if(!iso)return '';const d=new Date(iso);return isNaN(d)?iso:d.toLocaleString()}
function rememberOpen(){[...el('index').querySelectorAll('details[data-key]')].forEach(d=>openState[d.dataset.key]=d.open)}
function isOpen(k,dflt){return openState[k]===undefined?dflt:openState[k]}
function reportHtml(r){
 if(r.state==='Completed')return '<div class="report"><span>'+esc(when(r.completedUtc))+'</span>'+(r.location?' <span class="muted">'+esc(r.location)+'</span>':'')+(r.htmlUrl?' &middot; <a target="_blank" href="'+r.htmlUrl+'">Open HTML</a>':'')+(r.pdfUrl?' &middot; <a target="_blank" href="'+r.pdfUrl+'">Open PDF</a>':'')+(r.flagUrl?' &middot; <a target="_blank" href="'+r.flagUrl+'">Flag</a>':'')+(!r.htmlUrl&&!r.pdfUrl?' <span class="muted">(files no longer available)</span>':'')+'</div>';
 return '<div class="report"><span>'+esc(when(r.completedUtc||r.queuedUtc))+'</span> <span class="error">'+esc(r.state)+(r.errorSummary?': '+esc(r.errorSummary):'')+'</span></div>';
}
function providerHtml(p){
 const k='prov:'+p.npi;const done=p.reports.filter(r=>r.state==='Completed').length;
 const active=p.active?'<div class="progress"><span style="width:'+Math.max(0,Math.min(100,p.active.percent||0))+'%"></span></div><span class="muted">'+esc(p.active.state)+' '+(p.active.percent||0)+'% - '+esc(p.active.stage)+'</span>':'';
 return '<div class="provider"><div class="row"><div><b>'+esc(p.displayName)+'</b> <span class="muted">NPI '+esc(p.npi)+' &middot; '+(p.location?esc(p.location):'<em>no location saved</em>')+'</span></div><div><button data-npi="'+esc(p.npi)+'"'+(p.active?' disabled':'')+'>'+(p.active?'Running...':'Generate report')+'</button> <a class="btn" href="/providers?profile='+encodeURIComponent(p.npi)+'">Edit profile</a></div></div>'+active+'<details class="reports" data-key="'+esc(k)+'"'+(isOpen(k,false)?' open':'')+'><summary>Reports <span class="muted">('+done+' completed'+(p.reports.length-done?', '+(p.reports.length-done)+' failed':'')+')</span></summary>'+(p.reports.length?p.reports.map(reportHtml).join(''):'<div class="muted report">No reports yet.</div>')+'</details></div>';
}
function locationGroups(g){
 const byKey={};
 g.providers.forEach(p=>{const name=(p.location||'').trim();const key=name?name.toLowerCase():'~none';if(!byKey[key])byKey[key]={key:key,name:name||'No location saved',providers:[]};byKey[key].providers.push(p)});
 return Object.keys(byKey).sort((a,b)=>a==='~none'?1:b==='~none'?-1:a.localeCompare(b)).map(k=>byKey[k]);
}
function render(groups){
 rememberOpen();
 if(!groups.length){el('index').innerHTML='<p class="muted">No saved profiles yet. Map a provider in the Provider Wizard and tick "Save/update provider profile".</p>';return}
 el('index').innerHTML=groups.map(g=>{const k='pool:'+g.riskPool;const locs=locationGroups(g);return '<details class="pool" data-key="'+esc(k)+'"'+(isOpen(k,true)?' open':'')+'><summary><b>'+esc(g.riskPool)+'</b> <span class="muted">'+g.providers.length+' provider'+(g.providers.length===1?'':'s')+' &middot; '+locs.length+' location'+(locs.length===1?'':'s')+'</span></summary>'+locs.map(loc=>{const lk='loc:'+g.riskPool+'|'+loc.key;return '<details class="loc" data-key="'+esc(lk)+'"'+(isOpen(lk,true)?' open':'')+'><summary>'+esc(loc.name)+' <span class="muted">'+loc.providers.length+' provider'+(loc.providers.length===1?'':'s')+'</span></summary>'+loc.providers.map(providerHtml).join('')+'</details>'}).join('')+'</details>'}).join('');
 [...el('index').querySelectorAll('button[data-npi]')].forEach(b=>b.onclick=()=>generate(b.dataset.npi,b));
}
async function generate(npi,btn){btn.disabled=true;btn.textContent='Queuing...';try{await api('/api/profile-job?npi='+encodeURIComponent(npi),{method:'POST'});lastJson='';await refresh()}catch(e){alert('Could not queue the report: '+e.message);btn.disabled=false;btn.textContent='Generate report'}}
async function refresh(){
 if(busy)return;busy=true;
 try{const groups=await api('/api/provider-index');const txt=JSON.stringify(groups);if(txt!==lastJson){lastJson=txt;render(groups)}
  const all=groups.flatMap(g=>g.providers.flatMap(p=>p.reports.concat(p.active?[p.active]:[])));
  if(!runner&&all.some(r=>r.state==='Prepared'||r.state==='Queued')){runner=true;try{await fetch('/api/run-next',{method:'POST'})}finally{runner=false}}
 }catch(e){el('index').innerHTML='<p class="error">'+esc(e.message)+'</p>'}
 finally{busy=false}
}
refresh();setInterval(refresh,3000);
</script></body></html>
'@
 return $html.Replace('__NAV__',(Get-NavHtml 'index')).Replace('__TABCSS__',$script:TabCss)
}
# --- Draft 5.2: Communication tab - provider contact list, per-pool email campaigns with fresh report PDFs saved to Outlook Drafts ---
$script:ContactFields=[ordered]@{lastName=@('lastname','last','providerlastname','surname');firstName=@('firstname','first','providerfirstname');degree=@('degree','degrees','credential','credentials','title');location=@('location','site','office','practice','practicelocation');specialty=@('specialty','speciality','specialtydescription');phone=@('phone','officephone','workphone','phonenumber','businessphone');email=@('email','emailaddress','provideremail','mail');npi=@('npi','npinumber','providernpi');cellPhone=@('pvtcellphone','pvtcell','cellphone','cell','mobile','mobilephone','privatecell','privatecellphone');homePhone=@('homephone','home')}
$script:ContactFieldLabels=[ordered]@{lastName='Last name';firstName='First name';degree='Degree';location='Location';specialty='Specialty';phone='Phone';email='Email';npi='NPI';cellPhone='Pvt cell phone';homePhone='Home phone'}
$script:ContactRequiredFields=@('npi','email')
$script:ContactCache=@{}
function ConvertTo-HeaderToken([string]$Text){return (([string]$Text).ToLowerInvariant() -replace '[^a-z0-9]','')}
function Get-CellText($Worksheet,[int]$Row,[int]$Column){
 # Numbers come back as digits (an NPI or phone stored as a number must not be rendered in scientific notation).
 $v=$Worksheet.Cells[$Row,$Column].Value
 if($v -is [double] -or $v -is [decimal] -or $v -is [int] -or $v -is [long]){return ([decimal]$v).ToString('0.############')}
 if($v -is [DateTime]){return ([DateTime]$v).ToString('M/d/yyyy')}
 return ([string]$Worksheet.Cells[$Row,$Column].Text).Trim()
}
function Find-ContactHeader($Worksheet){
 # Fingerprint by column titles: the first 25 rows are scanned for the row that names the most contact fields (NPI and Email are required).
 if($null -eq $Worksheet.Dimension){return $null}
 $maxCol=$Worksheet.Dimension.End.Column;$limit=[Math]::Min(25,$Worksheet.Dimension.End.Row);$best=$null
 for($row=1;$row -le $limit;$row++){
  $map=[ordered]@{};$headers=@()
  for($col=1;$col -le $maxCol;$col++){
   $text=([string]$Worksheet.Cells[$row,$col].Text).Trim();if(!$text){continue};$headers+=$text;$token=ConvertTo-HeaderToken $text
   foreach($field in @($script:ContactFields.Keys)){if($map.Contains($field)){continue};if(@($script:ContactFields[$field]) -contains $token){$map[$field]=$col;break}}
  }
  if(@($script:ContactRequiredFields|Where-Object{!$map.Contains($_)}).Count -gt 0){continue}
  if($null -eq $best -or $map.Count -gt $best.map.Count){$best=[ordered]@{row=$row;map=$map;headers=$headers}}
 }
 return $best
}
function Get-ContactDisplayName($c){$n=(([string]$c.firstName+' '+[string]$c.lastName).Trim());if($c.degree){$n=(($n+', '+[string]$c.degree).Trim(',',' '))};return $n}
function Read-ContactList([string]$Path){
 if(!(Test-Path -LiteralPath $Path)){throw ('Provider contact list was not found: '+$Path)}
 $item=Get-Item -LiteralPath $Path;$cacheKey=$item.FullName+'|'+$item.LastWriteTimeUtc.Ticks+'|'+$item.Length
 if($script:ContactCache.ContainsKey($cacheKey)){return $script:ContactCache[$cacheKey]}
 $package=$null;$best=$null;$bestWs=$null
 try{
  $package=Open-ExcelPackage -Path $item.FullName -ErrorAction Stop
  foreach($ws in $package.Workbook.Worksheets){$h=Find-ContactHeader $ws;if($h -and ($null -eq $best -or $h.map.Count -gt $best.map.Count)){$best=$h;$bestWs=$ws}}
  if($null -eq $best){throw ('No worksheet in '+$item.Name+' has a header row with at least NPI and Email columns (the first 25 rows of every sheet were checked).')}
  $contacts=New-Object Collections.Generic.List[object];$byNpi=@{};$duplicates=@();$noNpi=0;$last=$bestWs.Dimension.End.Row
  for($r=$best.row+1;$r -le $last;$r++){
   $c=[ordered]@{};foreach($field in @($script:ContactFields.Keys)){$c[$field]=$(if($best.map.Contains($field)){Get-CellText $bestWs $r $best.map[$field]}else{''})}
   $npi=([string]$c.npi -replace '[^0-9]','');if($npi.Length -ne 10){if(([string]$c.lastName+[string]$c.firstName+[string]$c.email).Trim() -ne ''){$noNpi++};continue}
   $c.npi=$npi;$c.name=Get-ContactDisplayName $c;$c.row=$r
   if($byNpi.ContainsKey($npi)){$duplicates+=$npi;continue}
   $byNpi[$npi]=$c;$contacts.Add($c)
  }
  $result=[ordered]@{path=$item.FullName;fileName=$item.Name;lastWriteUtc=$item.LastWriteTimeUtc.ToString('o');sha256=(Get-FileHash256 $item.FullName);worksheet=[string]$bestWs.Name;headerRow=[int]$best.row;headers=@($best.headers);recognized=@(foreach($k in @($best.map.Keys)){[string]$script:ContactFieldLabels[$k]});missing=@(foreach($k in @($script:ContactFields.Keys)){if(!$best.map.Contains($k)){[string]$script:ContactFieldLabels[$k]}});contactCount=$contacts.Count;duplicateNpis=@($duplicates|Select-Object -Unique);rowsWithoutNpi=$noNpi;contacts=@($contacts.ToArray());byNpi=$byNpi}
 }finally{if($package){Close-ExcelPackage $package -NoSave}}
 $script:ContactCache.Clear();$script:ContactCache[$cacheKey]=$result;return $result
}
function Get-ContactListSummary($List){if($null -eq $List){return $null};$o=[ordered]@{};foreach($k in @($List.Keys)){if($k -in @('contacts','byNpi')){continue};$o[$k]=$List[$k]};return $o}
function Get-CommunicationSettings{$s=Json (Join-Path $script:Paths.State 'communication.json');return [ordered]@{version=1;contactListPath=$(if($s){[string](Get-P $s 'contactListPath' '')}else{''});updatedUtc=$(if($s){[string](Get-P $s 'updatedUtc' '')}else{''})}}
function Save-CommunicationSettings($Settings){$Settings.updatedUtc=[DateTime]::UtcNow.ToString('o');Save-JsonAtomic (Join-Path $script:Paths.State 'communication.json') $Settings}
function Get-ContactListCandidates{
 # XLSX files in the application root and the contacts folder; canonical source exports live in subfolders and are never offered.
 $out=@()
 foreach($dir in @($script:Root,$script:Paths.Contacts)){foreach($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.xlsx' -ErrorAction SilentlyContinue|Where-Object{$_.Name -notlike '~$*'})){$out+=[ordered]@{path=$f.FullName;fileName=$f.Name;folder=$(if($dir -eq $script:Root){'root'}else{'contacts'});byteLength=$f.Length;lastWriteUtc=$f.LastWriteTimeUtc.ToString('o')}}}
 return @($out|Sort-Object -Property @{Expression={[string]$_.fileName}})
}
function Resolve-ContactListPath([string]$Path){
 if([string]::IsNullOrWhiteSpace($Path)){throw 'Choose a provider contact list.'}
 $candidate=$(if([IO.Path]::IsPathRooted($Path)){$Path}else{Join-Path $script:Root $Path})
 if(!(Test-Path -LiteralPath $candidate -PathType Leaf)){throw ('Provider contact list was not found: '+$Path)}
 $full=(Get-Item -LiteralPath $candidate).FullName
 if([IO.Path]::GetExtension($full) -ine '.xlsx'){throw 'The contact list must be an XLSX file.'}
 $dir=(Split-Path -Parent $full).TrimEnd('\','/');$allowed=@(foreach($d in @($script:Root,$script:Paths.Contacts)){([string]$d).TrimEnd('\','/')})
 if(@($allowed|Where-Object{$_ -ieq $dir}).Count -eq 0){throw 'The contact list must be in the application folder or its contacts folder.'}
 return $full
}
function Select-ContactList([string]$Path){$full=Resolve-ContactListPath $Path;$list=Read-ContactList $full;$settings=Get-CommunicationSettings;$settings.contactListPath=$full;Save-CommunicationSettings $settings;Log 'CONTACT_LIST_SELECTED' 'OK' ($list.fileName+'; '+$list.contactCount+' contacts');return (Get-ContactListSummary $list)}
function Get-CurrentContactList{$settings=Get-CommunicationSettings;if(!$settings.contactListPath){return $null};if(!(Test-Path -LiteralPath $settings.contactListPath)){return $null};return (Read-ContactList $settings.contactListPath)}
function Receive-ContactListUpload($Context){
 $name=[Net.WebUtility]::UrlDecode([string]$Context.Request.Headers['X-File-Name']);if([string]::IsNullOrWhiteSpace($name)){throw 'Missing file name.'};if([IO.Path]::GetExtension($name) -ine '.xlsx'){throw 'Only XLSX files are accepted.'}
 $safe=([IO.Path]::GetFileName($name) -replace '[^A-Za-z0-9 ._-]','_');$target=Join-Path $script:Paths.Contacts $safe;$tmp=Join-Path $script:Paths.Staging ([Guid]::NewGuid().ToString('N')+'.xlsx')
 $stream=[IO.File]::Create($tmp);try{$buffer=New-Object byte[] 1048576;$total=0;while(($read=$Context.Request.InputStream.Read($buffer,0,$buffer.Length)) -gt 0){$total+=$read;if($total -gt 52428800){throw 'Contact list exceeds the 50 MB limit.'};$stream.Write($buffer,0,$read)}}finally{$stream.Dispose()}
 try{$null=Read-ContactList $tmp}catch{Remove-Item $tmp -Force -ErrorAction SilentlyContinue;throw}
 if(Test-Path -LiteralPath $target){Copy-Item -LiteralPath $target -Destination ($target+'.'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.bak') -Force}
 Move-Item -LiteralPath $tmp -Destination $target -Force;$script:ContactCache.Clear();Log 'CONTACT_LIST_UPLOADED' 'OK' $safe
 return (Select-ContactList $target)
}
function Get-CommunicationPools{return @(Get-ProviderProfiles|ForEach-Object{ConvertTo-RiskPoolName ([string](Get-P $_ 'riskPool' ''))}|Where-Object{$_}|Select-Object -Unique|Sort-Object)}
function Get-CommunicationModel{$list=$null;$listError='';try{$list=Get-CurrentContactList}catch{$listError=$_.Exception.Message};return [ordered]@{pools=@(Get-CommunicationPools);contactList=(Get-ContactListSummary $list);contactListPath=(Get-CommunicationSettings).contactListPath;contactListError=$listError;candidates=@(Get-ContactListCandidates)}}
function Get-LatestCompletedJob([string]$Npi){foreach($j in @(Get-Jobs|Where-Object{[string](Get-P $_ 'providerKey' '') -eq $Npi -and [string](Get-P $_ 'state' '') -eq 'Completed'}|Sort-Object -Property @{Expression={ConvertTo-IsoText (Get-P $_ 'completedUtc' '')}} -Descending)){if(Resolve-JobOutputPath $j 'pdf'){return $j}};return $null}
function Get-CommunicationRecipients([string]$RiskPool){
 $pool=ConvertTo-RiskPoolName $RiskPool;if(!$pool){throw 'Choose a risk pool.'}
 $list=Get-CurrentContactList;$out=@()
 foreach($profile in @(Get-ProviderProfiles)){
  if((ConvertTo-RiskPoolName ([string](Get-P $profile 'riskPool' ''))) -ne $pool){continue}
  $npi=[string](Get-P $profile 'npi' '');if(!$npi){continue}
  $contact=$(if($list -and $list.byNpi.ContainsKey($npi)){$list.byNpi[$npi]}else{$null})
  $latest=Get-LatestCompletedJob $npi
  $out+=[ordered]@{npi=$npi;displayName=[string](Get-P $profile 'displayName' $npi);location=[string](Get-P $profile 'location' '');matched=[bool]($null -ne $contact);contactName=$(if($contact){[string]$contact.name}else{''});contactLocation=$(if($contact){[string]$contact.location}else{''});contactSpecialty=$(if($contact){[string]$contact.specialty}else{''});email=$(if($contact){[string]$contact.email}else{''});latestReportUtc=$(if($latest){ConvertTo-IsoText (Get-P $latest 'completedUtc' '')}else{''})}
 }
 return [ordered]@{riskPool=$pool;contactList=(Get-ContactListSummary $list);recipients=@($out|Sort-Object -Property @{Expression={[string]$_.displayName}})}
}
$script:EmailAllowedTags=@('p','br','b','strong','i','em','u','ul','ol','li','a','div','span','blockquote','h1','h2','h3','h4','hr','sub','sup','s','strike')
function ConvertTo-SafeEmailHtml([string]$Html){
 # Keeps ordinary formatting from the composer (paragraphs, bold/italic/underline, lists, links) and drops everything else, including every attribute except a safe href.
 if([string]::IsNullOrWhiteSpace($Html)){return ''}
 $h=[regex]::Replace($Html,'(?is)<(script|style|iframe|object|embed|form|textarea|select|button|title|head)\b[^>]*>.*?</\1\s*>','')
 $h=[regex]::Replace($h,'(?is)<(script|style|iframe|object|embed|form|input|meta|link|img)\b[^>]*/?>','')
 $h=[regex]::Replace($h,'(?s)<!--.*?-->','')
 $h=[regex]::Replace($h,'(?is)<\s*(/?)\s*([a-z0-9]+)([^>]*)>',{param($m)$close=$m.Groups[1].Value;$tag=$m.Groups[2].Value.ToLowerInvariant();$attrs=$m.Groups[3].Value;if($script:EmailAllowedTags -notcontains $tag){return ''};if($close){return ('</'+$tag+'>')};if($tag -eq 'br' -or $tag -eq 'hr'){return ('<'+$tag+'>')};$keep='';if($tag -eq 'a'){$hm=[regex]::Match($attrs,'(?i)href\s*=\s*("([^"]*)"|''([^'']*)''|([^\s>]+))');if($hm.Success){$href=$(if($hm.Groups[2].Success){$hm.Groups[2].Value}elseif($hm.Groups[3].Success){$hm.Groups[3].Value}else{$hm.Groups[4].Value}).Trim();if($href -match '^(https?:|mailto:)'){$keep=' href="'+[Net.WebUtility]::HtmlEncode($href)+'"'}}};return ('<'+$tag+$keep+'>')})
 return $h.Trim()
}
function Get-EmailPlainText([string]$Html){return (([Net.WebUtility]::HtmlDecode(($Html -replace '<[^>]+>',' '))) -replace '\s+',' ').Trim()}
function New-EmailHtml([string]$Greeting,[string]$BodyHtml){return ('<div style="font-family:Calibri,Segoe UI,Arial,sans-serif;font-size:11pt"><p>Dear '+(ConvertTo-HtmlEncoded $Greeting)+',</p>'+$BodyHtml+'</div>')}
$script:Outlook=$null
function Get-OutlookApplication{
 if($script:Outlook){try{$null=$script:Outlook.Name;return $script:Outlook}catch{$script:Outlook=$null}}
 try{$script:Outlook=[Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')}catch{$script:Outlook=$null}
 if(!$script:Outlook){try{$script:Outlook=New-Object -ComObject Outlook.Application}catch{$script:Outlook=$null;throw ('Outlook could not be started for drafting: '+$_.Exception.Message)}}
 return $script:Outlook
}
$script:DraftDateTags=@('0x00390040','0x0E060040')   # PR_CLIENT_SUBMIT_TIME (the Drafts folder's Sent/Date column) and PR_MESSAGE_DELIVERY_TIME
function Set-OutlookItemDate($Mail){
 # A draft saved through COM carries no date, so Outlook sorts it to the bottom of Drafts; stamping the MAPI date properties with "now" puts it at the top like a hand-written draft.
 $now=Get-Date
 foreach($tag in $script:DraftDateTags){
  $name='http://schemas.microsoft.com/mapi/proptag/'+$tag
  try{$utc=$null;try{$utc=$Mail.PropertyAccessor.LocalTimeToUTC($name,$now)}catch{$utc=$now.ToUniversalTime()};$Mail.PropertyAccessor.SetProperty($name,$utc)}
  catch{Log 'DRAFT_DATE' 'WARN' ($tag+': '+$_.Exception.Message)}
 }
}
function New-OutlookDraft($Spec){
 # $Spec keys: to, subject, html, attachmentPath, attachmentName, includeSignature. Returns the saved draft's Outlook EntryID.
 # Test hook: when PA_MAIL_SINK names a folder, each draft is written there as JSON (plus a copy of the attachment) instead of going to Outlook.
 $sink=[string]$env:PA_MAIL_SINK
 if($sink){if(!(Test-Path -LiteralPath $sink)){New-Item -ItemType Directory -Path $sink -Force|Out-Null};$id='SINK-'+[Guid]::NewGuid().ToString('N');$copy='';if($Spec.attachmentPath){Copy-Item -LiteralPath $Spec.attachmentPath -Destination (Join-Path $sink ([string]$Spec.attachmentName)) -Force;$copy=[string]$Spec.attachmentName};Save-JsonAtomic (Join-Path $sink ($id+'.json')) ([ordered]@{entryId=$id;to=[string]$Spec.to;subject=[string]$Spec.subject;html=[string]$Spec.html;attachment=$copy;includeSignature=[bool]$Spec.includeSignature;draftedUtc=[DateTime]::UtcNow.ToString('o')});return $id}
 $ol=Get-OutlookApplication;$mail=$ol.CreateItem(0);$tempDir=''
 try{
  $mail.Subject=[string]$Spec.subject;$mail.To=[string]$Spec.to;$mail.BodyFormat=2
  $signature='';if($Spec.includeSignature){try{$null=$mail.GetInspector;$signature=[string]$mail.HTMLBody}catch{$signature=''}}
  # The default signature arrives as a full HTML document once the inspector exists; the message goes in right after its <body> tag so the signature keeps its styles.
  if($signature -match '(?is)<body[^>]*>'){$tag=$matches[0];$pos=$signature.IndexOf($tag);$mail.HTMLBody=$signature.Substring(0,$pos+$tag.Length)+[string]$Spec.html+$signature.Substring($pos+$tag.Length)}else{$mail.HTMLBody=[string]$Spec.html}
  # HTML mail shows the attached file's own name, so the PDF is attached from a temporary copy that already carries the friendly name.
  if($Spec.attachmentPath){$tempDir=Join-Path $script:Paths.Staging ('mail-'+[Guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $tempDir -Force|Out-Null;$copy=Join-Path $tempDir ([string]$Spec.attachmentName);Copy-Item -LiteralPath ([string]$Spec.attachmentPath) -Destination $copy -Force;$null=$mail.Attachments.Add($copy,1,1,[string]$Spec.attachmentName)}
  Set-OutlookItemDate $mail
  $mail.Save();$id=[string]$mail.EntryID;try{$mail.Close(0)}catch{};return $id
 }finally{try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($mail)}catch{};if($tempDir){Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue}}
}
function Get-CampaignPath([string]$Id){if($Id -notmatch '^[a-f0-9]{32}$'){throw 'Invalid campaign ID.'};return (Join-Path $script:Paths.State ('campaign-'+$Id+'.json'))}
function Get-Campaign([string]$Id){$c=Json (Get-CampaignPath $Id);if($null -eq $c){throw 'Campaign not found.'};return $c}
function Get-Campaigns{return @(Get-ChildItem $script:Paths.State -Filter 'campaign-*.json' -File -ErrorAction SilentlyContinue|ForEach-Object{Json $_.FullName}|Where-Object{$_}|Sort-Object -Property @{Expression={ConvertTo-IsoText (Get-P $_ 'createdUtc' '')}} -Descending)}
function Update-CampaignState($Campaign){
 $ps=@($Campaign.providers)
 $reportsPending=@($ps|Where-Object{$_.draftState -eq 'Pending' -and $_.reportState -notin @('Completed','Failed')}).Count
 $draftsPending=@($ps|Where-Object{$_.draftState -eq 'Pending' -and $_.reportState -eq 'Completed'}).Count
 $Campaign.state=$(if($reportsPending -gt 0){'Running'}elseif($draftsPending -gt 0){'Drafting'}else{'Completed'})
 Set-P $Campaign 'summary' ([ordered]@{total=$ps.Count;reportsPending=$reportsPending;draftsPending=$draftsPending;draftsCreated=@($ps|Where-Object{$_.draftState -eq 'Created'}).Count;draftsFailed=@($ps|Where-Object{$_.draftState -eq 'Failed'}).Count;skipped=@($ps|Where-Object{$_.draftState -eq 'Skipped'}).Count})
}
function New-Campaign($Body){
 if($null -eq $Body){throw 'Email details are required.'}
 $pool=ConvertTo-RiskPoolName ([string](Get-P $Body 'riskPool' ''));if(!$pool){throw 'Choose a risk pool.'}
 $subject=([string](Get-P $Body 'subject' '')).Trim();if(!$subject){throw 'A subject line is required.'}
 $bodyHtml=ConvertTo-SafeEmailHtml ([string](Get-P $Body 'bodyHtml' ''));if(!(Get-EmailPlainText $bodyHtml)){throw 'An email body is required.'}
 $fresh=[bool](Get-P $Body 'freshReports' $true);$sig=[bool](Get-P $Body 'includeSignature' $true);$settings=Get-CommunicationSettings
 $providers=@();$seen=@{}
 foreach($r in @(Get-P $Body 'providers' @())){
  if($null -eq $r){continue};$npi=([string](Get-P $r 'npi' '') -replace '[^0-9]','');if($npi -notmatch '^\d{10}$'){throw ('Invalid NPI in the recipient list: '+[string](Get-P $r 'npi' ''))};if($seen.ContainsKey($npi)){continue};$seen[$npi]=$true
  $profile=Json (Join-Path $script:Paths.Profiles ($npi+'.json'));if($null -eq $profile){throw ('No saved profile for NPI '+$npi+'.')}
  $name=[string](Get-P $profile 'displayName' $npi)
  if((ConvertTo-RiskPoolName ([string](Get-P $profile 'riskPool' ''))) -ne $pool){throw ($name+' is not in risk pool '+$pool+'.')}
  $email=([string](Get-P $r 'email' '')).Trim();if($email -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$'){throw ('A valid email address is required for '+$name+'.')}
  $greeting=([string](Get-P $r 'greeting' '')).Trim();if(!$greeting){$greeting=$name}
  $providers+=[ordered]@{npi=$npi;displayName=$name;location=[string](Get-P $profile 'location' '');greeting=$greeting;email=$email;jobId='';reportState='';percent=0;stage='';pdfPath='';draftState='Pending';draftUtc='';entryId='';error=''}
 }
 if($providers.Count -eq 0){throw 'Select at least one recipient with an email address.'}
 $campaign=[ordered]@{campaignVersion=1;campaignId=[Guid]::NewGuid().ToString('N');riskPool=$pool;subject=$subject;bodyHtml=$bodyHtml;includeSignature=$sig;freshReports=$fresh;contactListPath=[string]$settings.contactListPath;createdUtc=[DateTime]::UtcNow.ToString('o');state='Running';providers=$providers}
 foreach($p in $campaign.providers){
  if($fresh){$job=New-JobFromProfile $p.npi;$job['campaignId']=$campaign.campaignId;Save-JsonAtomic (Join-Path $script:Paths.State ('job-'+$job.jobId+'.json')) $job;$p.jobId=[string]$job.jobId;$p.reportState='Prepared';$p.stage=[string]$job.stage}
  else{$latest=Get-LatestCompletedJob $p.npi;if($latest){$p.jobId=[string]$latest.jobId;$p.reportState='Completed';$p.percent=100;$p.stage='Existing report'}else{$p.reportState='Failed';$p.draftState='Skipped';$p.error='No existing report PDF to attach; generate a fresh report instead.'}}
 }
 Save-JsonAtomic (Get-CampaignPath $campaign.campaignId) $campaign
 Log 'CAMPAIGN_CREATED' 'OK' ($pool+'; '+$providers.Count+' recipients; fresh='+$fresh)
 return (Get-CampaignModel $campaign.campaignId)
}
function Get-CampaignModel([string]$Id){
 # Merges live job progress into the campaign record; a failed report marks the draft as skipped.
 $c=Get-Campaign $Id;$changed=$false
 foreach($p in @($c.providers)){
  if(!$p.jobId -or $p.draftState -ne 'Pending' -or $p.reportState -in @('Completed','Failed')){continue}
  $job=Json (Join-Path $script:Paths.State ('job-'+$p.jobId+'.json'))
  if(!$job){$p.reportState='Failed';$p.draftState='Skipped';$p.error='Report job record not found.';$changed=$true;continue}
  $s=[string](Get-P $job 'state' '');$pct=[int](Get-P $job 'percent' 0);$stage=[string](Get-P $job 'stage' '')
  if($s -ne [string]$p.reportState -or $pct -ne [int]$p.percent -or $stage -ne [string]$p.stage){$p.reportState=$s;$p.percent=$pct;$p.stage=$stage;$changed=$true}
  if($s -eq 'Failed'){$p.draftState='Skipped';$p.error='Report failed: '+[string](Get-P $job 'errorSummary' '');$changed=$true}
 }
 $prev=[string]$c.state;Update-CampaignState $c
 if($changed -or $prev -ne [string]$c.state){Save-JsonAtomic (Get-CampaignPath $Id) $c}
 return $c
}
function Invoke-CampaignDrafts([string]$Id,[int]$Limit=5){
 # Saves an Outlook draft for every provider whose report is complete and whose draft is still pending (at most $Limit per call so a poll never blocks for long).
 $c=Get-CampaignModel $Id;$created=0
 foreach($p in @($c.providers)){
  if($created -ge $Limit){break}
  if($p.draftState -ne 'Pending' -or $p.reportState -ne 'Completed'){continue}
  try{
   $job=Json (Join-Path $script:Paths.State ('job-'+$p.jobId+'.json'));if(!$job){throw 'Report job record not found.'}
   $pdf=Resolve-JobOutputPath $job 'pdf';if(!$pdf){throw 'Report PDF is not available.'}
   $when=ConvertTo-DateValue ([string](Get-P $job 'completedUtc' ''));$stamp=$(if($when){$when.ToString('yyyy-MM-dd')}else{Get-Date -Format 'yyyy-MM-dd'})
   $attachmentName=((('Provider Patient Dashboard - '+[string]$p.displayName+' - '+$stamp) -replace '[\\/:*?"<>|]','-')+'.pdf')
   $entry=New-OutlookDraft ([ordered]@{to=[string]$p.email;subject=[string]$c.subject;html=(New-EmailHtml ([string]$p.greeting) ([string]$c.bodyHtml));attachmentPath=$pdf;attachmentName=$attachmentName;includeSignature=[bool]$c.includeSignature})
   $p.pdfPath=$pdf;$p.draftState='Created';$p.draftUtc=[DateTime]::UtcNow.ToString('o');$p.entryId=[string]$entry;$p.error='';$created++
  }catch{$p.draftState='Failed';$p.error=$_.Exception.Message;Log 'CAMPAIGN_DRAFT' 'FAILED' ([string]$p.npi+' '+$_.Exception.Message)}
 }
 Update-CampaignState $c;Save-JsonAtomic (Get-CampaignPath $Id) $c
 if($created -gt 0){Log 'CAMPAIGN_DRAFTS' 'OK' ($created.ToString()+' drafts saved for campaign '+$Id)}
 return $c
}
function Resume-Campaign([string]$Id){
 # Retries failed drafts and, for a fresh-report campaign, queues a new report for every provider whose report failed; existing-report campaigns look for a newer PDF instead.
 $c=Get-CampaignModel $Id;$fresh=[bool](Get-P $c 'freshReports' $true);$requeued=0;$retried=0
 foreach($p in @($c.providers)){
  if($p.draftState -eq 'Failed'){$p.draftState='Pending';$p.error='';$retried++;continue}
  if($p.draftState -ne 'Skipped' -or $p.reportState -ne 'Failed'){continue}
  if($fresh){$job=New-JobFromProfile ([string]$p.npi);$job['campaignId']=$Id;Save-JsonAtomic (Join-Path $script:Paths.State ('job-'+$job.jobId+'.json')) $job;$p.jobId=[string]$job.jobId;$p.reportState='Prepared';$p.percent=0;$p.stage=[string]$job.stage;$p.draftState='Pending';$p.error='';$requeued++}
  else{$latest=Get-LatestCompletedJob ([string]$p.npi);if($latest){$p.jobId=[string]$latest.jobId;$p.reportState='Completed';$p.percent=100;$p.stage='Existing report';$p.draftState='Pending';$p.error='';$retried++}}
 }
 Update-CampaignState $c;Save-JsonAtomic (Get-CampaignPath $Id) $c;Log 'CAMPAIGN_RESUMED' 'OK' ($Id+'; reports requeued '+$requeued+'; drafts retried '+$retried)
 return (Invoke-CampaignDrafts $Id)
}
function Get-CampaignList([int]$Take=10){$out=@();foreach($c in @(Get-Campaigns|Select-Object -First $Take)){$m=Get-CampaignModel ([string]$c.campaignId);$m.createdUtc=(ConvertTo-IsoText $m.createdUtc);$out+=$m};return $out}
function CommunicationPage{
 $html=@'
<!doctype html><html><head><meta charset="utf-8"><title>Communication</title>
<style>
body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:white;padding:22px 22px 12px}main{padding:22px;max-width:1200px}
__TABCSS__
.card{background:white;border:1px solid #dce4ef;border-radius:8px;padding:18px;margin:14px 0}.card h2{margin:0 0 8px;font-size:17px;color:#17365d}.muted{color:#667085}.error{color:#a61b1b}.warn{color:#9a5b00}.ok{color:#1d7a3a;font-weight:600}
.row{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:8px 0}label{font-weight:600}select,input[type=text],input:not([type]){padding:7px 9px;border:1px solid #c8d2e0;border-radius:5px;font:inherit;background:#fff}
button{padding:7px 12px;cursor:pointer;background:#1769aa;color:#fff;border:0;border-radius:5px;font:inherit}button:disabled{opacity:.5;cursor:default}button.alt{background:#e4e9f0;color:#172033}button.go{background:#1d7a3a;font-weight:600;padding:10px 16px}
table{border-collapse:collapse;width:100%;margin-top:8px}th,td{padding:7px 8px;border-bottom:1px solid #dce4ef;text-align:left;vertical-align:top}th{background:#eaf1f8}tr.unmatched td{background:#fff8ec}td input{width:100%;box-sizing:border-box}
.greet{font:14px Calibri,Segoe UI,Arial;padding:8px 10px 0;color:#172033}.toolbar{display:flex;gap:4px;margin:6px 0 4px}.toolbar button{background:#e4e9f0;color:#172033;min-width:34px}
.editor{min-height:200px;border:1px solid #c8d2e0;border-radius:6px;padding:10px;background:#fff;font:14px Calibri,Segoe UI,Arial;line-height:1.45;outline:none}.editor:focus{border-color:#1769aa}.editor p{margin:0 0 10px}
.progress{height:10px;background:#e4e9f0;border-radius:6px;overflow:hidden;margin:6px 0}.progress span{display:block;height:100%;background:#1769aa;transition:width .3s}
.subject{width:100%;box-sizing:border-box}.check{font-weight:400;display:block;margin:6px 0}.queue{background:#fff8ec;border:1px solid #f0d9a8;border-radius:6px;padding:8px 12px;margin:6px 0}.queue.stuck{background:#fdecec;border-color:#e6b3b3}.camp{border-top:1px solid #dce4ef;padding-top:10px;margin-top:14px}.camp:first-of-type{border-top:0;margin-top:0;padding-top:0}
</style></head><body>
<header><h1>Communication</h1><p>Draft 5.3 - email every provider in a risk pool a fresh report, saved to your Outlook Drafts folder for review before sending</p></header>__NAV__
<main>
<div class="card"><h2>1. Provider contact list</h2>
<p class="muted">Choose the XLSX that holds provider emails. Its columns are recognized by title (last name, first name, degree, location, specialty, phone, email, NPI, pvt cell phone, home phone) and providers are matched to saved profiles by NPI. Files in the application folder and its <b>contacts</b> folder are listed; uploads are copied into <b>contacts</b>.</p>
<div class="row"><select id="listSelect"><option value="">Choose a contact list...</option></select><button id="useList">Use this list</button><button class="alt" id="refreshLists">Refresh files</button><button class="alt" onclick="pickFile()">Upload XLSX...</button><input id="file" type="file" accept=".xlsx" hidden></div>
<div id="listInfo" class="muted">No contact list selected.</div></div>
<div class="card"><h2>2. Recipients</h2>
<div class="row"><label for="pool">Risk pool</label><select id="pool"><option value="">Choose risk pool...</option></select><button class="alt" id="selAll">Select all with an email</button><button class="alt" id="selNone">Clear selection</button><span id="recipNote" class="muted"></span></div>
<div id="recipients" class="muted">Choose a risk pool to list its saved provider profiles.</div></div>
<div class="card"><h2>3. Email</h2>
<label for="subject">Subject</label><div class="row"><input id="subject" class="subject" placeholder="Subject line"></div>
<label>Body</label><div class="greet">Dear [Provider Name],</div>
<div class="toolbar"><button type="button" data-cmd="bold" title="Bold"><b>B</b></button><button type="button" data-cmd="italic" title="Italic"><i>I</i></button><button type="button" data-cmd="underline" title="Underline"><u>U</u></button><button type="button" data-cmd="insertUnorderedList" title="Bulleted list">&bull; List</button><button type="button" data-cmd="insertOrderedList" title="Numbered list">1. List</button><button type="button" data-cmd="link" title="Insert link">Link</button><button type="button" data-cmd="removeFormat" title="Clear formatting">Clear</button><button type="button" data-cmd="undo" title="Undo">Undo</button></div>
<div id="body" class="editor" contenteditable="true"></div>
<label class="check"><input type="checkbox" id="sig" checked> Include my default Outlook signature</label>
<label class="check"><input type="checkbox" id="fresh" checked> Generate a fresh report for each provider (uncheck to attach each provider's latest existing PDF)</label>
<div class="row"><button class="go" id="go" disabled>Generate reports and save Outlook drafts</button><span id="goNote" class="muted"></span></div></div>
<div class="card" id="progressCard" style="display:none"><h2>4. Progress</h2><div id="campaign"></div></div>
</main>
<script>
const el=id=>document.getElementById(id);let recipients=[],pollTimer=null,pollBusy=false,draftBusy=false,runner=false,lastCampaignJson='';
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
async function api(u,o){const r=await fetch(u,o);const text=await r.text();let j=null;try{j=text?JSON.parse(text):null}catch(e){throw Error('Server returned invalid JSON: '+text.slice(0,200))}if(!r.ok)throw Error((j&&j.error)||('Request failed ('+r.status+')'));return j}
function when(iso){if(!iso)return '';const d=new Date(iso);return isNaN(d)?iso:d.toLocaleString()}
function note(id,t,err){const n=el(id);n.textContent=t;n.className=err?'error':'muted'}
function saveDraft(){try{localStorage.setItem('pa-communication',JSON.stringify({pool:el('pool').value,subject:el('subject').value,body:el('body').innerHTML,sig:el('sig').checked,fresh:el('fresh').checked}))}catch(e){}}
function loadDraft(){try{const d=JSON.parse(localStorage.getItem('pa-communication')||'null');if(!d)return null;el('subject').value=d.subject||'';el('body').innerHTML=d.body||'';el('sig').checked=d.sig!==false;el('fresh').checked=d.fresh!==false;return d}catch(e){return null}}
function listInfoHtml(l,path,err){
 if(err)return '<span class="error">'+esc(err)+'</span>';
 if(!l)return '<span class="muted">No contact list selected'+(path?' (the saved list '+esc(path)+' is no longer available)':'')+'.</span>';
 const d=l.duplicateNpis.length,n=l.rowsWithoutNpi;
 return '<b>'+esc(l.fileName)+'</b> <span class="muted">sheet '+esc(l.worksheet)+', header row '+l.headerRow+', '+l.contactCount+' contact'+(l.contactCount===1?'':'s')+' with an NPI'+(d?', '+d+' duplicate NPI'+(d===1?'':'s')+' ignored':'')+(n?', '+n+' row'+(n===1?'':'s')+' without a valid NPI skipped':'')+'</span><br>Recognized columns: '+esc(l.recognized.join(', '))+(l.missing.length?'<br><span class="warn">Not found: '+esc(l.missing.join(', '))+'</span>':'');
}
function fillLists(m){const s=el('listSelect');s.innerHTML='<option value="">Choose a contact list...</option>';m.candidates.forEach(c=>s.add(new Option(c.fileName+(c.folder==='contacts'?'  (contacts folder)':'')+'  -  '+when(c.lastWriteUtc),c.path)));if(m.contactListPath&&[...s.options].some(o=>o.value===m.contactListPath))s.value=m.contactListPath;el('listInfo').innerHTML=listInfoHtml(m.contactList,m.contactListPath,m.contactListError)}
async function refreshLists(){try{const m=await api('/api/communication');fillLists(m);fillPools(m.pools)}catch(e){el('listInfo').innerHTML='<span class="error">'+esc(e.message)+'</span>'}}
function fillPools(pools){const s=el('pool');const keep=s.value;s.innerHTML='<option value="">Choose risk pool...</option>';pools.forEach(p=>s.add(new Option(p,p)));if(keep&&pools.includes(keep))s.value=keep}
async function useList(){const p=el('listSelect').value;if(!p){note('goNote','Choose a contact list first.',true);return}el('listInfo').innerHTML='<span class="muted">Reading '+esc(p)+' ...</span>';try{const l=await api('/api/contact-list?path='+encodeURIComponent(p),{method:'POST'});el('listInfo').innerHTML=listInfoHtml(l,p,'');if(el('pool').value)await loadRecipients()}catch(e){el('listInfo').innerHTML='<span class="error">'+esc(e.message)+'</span>'}}
function pickFile(){const f=el('file');f.value='';f.click()}
el('file').onchange=async()=>{const f=el('file');if(!f.files.length)return;const file=f.files[0];el('listInfo').innerHTML='<span class="muted">Uploading and reading '+esc(file.name)+' ...</span>';try{const l=await api('/api/contact-list-upload',{method:'POST',headers:{'X-File-Name':encodeURIComponent(file.name)},body:file});await refreshLists();el('listInfo').innerHTML=listInfoHtml(l,l.path,'');if(el('pool').value)await loadRecipients()}catch(e){el('listInfo').innerHTML='<span class="error">'+esc(e.message)+'</span>'}};
function recipientsHtml(){
 if(!recipients.length)return '<span class="muted">No saved profiles in this risk pool. Map providers in the Provider Wizard and save their profiles first.</span>';
 const matched=recipients.filter(r=>r.matched).length;
 return '<p class="muted">'+recipients.length+' saved profile'+(recipients.length===1?'':'s')+' in this pool, '+matched+' matched to the contact list by NPI. Rows without a match are highlighted; type an email to include one anyway. The greeting name is prefilled from the profile.</p><table><thead><tr><th></th><th>Provider (profile)</th><th>NPI</th><th>Location</th><th>Contact list</th><th>Email</th><th>Greeting name</th><th>Latest report</th></tr></thead><tbody>'+recipients.map((r,i)=>'<tr'+(r.matched?'':' class="unmatched"')+'><td><input type="checkbox" data-i="'+i+'"'+(r.matched&&r.email?' checked':'')+'></td><td><b>'+esc(r.displayName)+'</b></td><td>'+esc(r.npi)+'</td><td>'+esc(r.location)+'</td><td>'+(r.matched?esc(r.contactName)+(r.contactSpecialty?'<br><span class="muted">'+esc(r.contactSpecialty)+'</span>':''):'<span class="warn">Not in contact list</span>')+'</td><td><input data-email="'+i+'" value="'+esc(r.email)+'" placeholder="email address"></td><td><input data-greet="'+i+'" value="'+esc(r.displayName)+'"></td><td class="muted">'+(r.latestReportUtc?esc(when(r.latestReportUtc)):'none yet')+'</td></tr>').join('')+'</tbody></table>';
}
async function loadRecipients(){const pool=el('pool').value;recipients=[];if(!pool){el('recipients').innerHTML='<span class="muted">Choose a risk pool to list its saved provider profiles.</span>';updateGo();return}el('recipients').innerHTML='<span class="muted">Loading...</span>';try{const m=await api('/api/recipients?riskPool='+encodeURIComponent(pool));recipients=m.recipients;el('recipients').innerHTML=recipientsHtml();[...el('recipients').querySelectorAll('input')].forEach(x=>{x.oninput=updateGo;x.onchange=updateGo})}catch(e){el('recipients').innerHTML='<span class="error">'+esc(e.message)+'</span>'}updateGo()}
function selected(){return [...el('recipients').querySelectorAll('input[type=checkbox]:checked')].map(cb=>{const i=+cb.dataset.i;return {npi:recipients[i].npi,displayName:recipients[i].displayName,email:el('recipients').querySelector('input[data-email="'+i+'"]').value.trim(),greeting:el('recipients').querySelector('input[data-greet="'+i+'"]').value.trim()}})}
function bodyText(){return el('body').innerText.replace(/\s+/g,' ').trim()}
function updateGo(){
 const sel=selected();const noEmail=sel.filter(s=>!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s.email));const problems=[];
 if(!el('pool').value)problems.push('choose a risk pool');
 if(!sel.length)problems.push('tick at least one recipient');
 if(noEmail.length)problems.push(noEmail.length+' selected recipient'+(noEmail.length===1?' has':'s have')+' no valid email');
 if(!el('subject').value.trim())problems.push('enter a subject');
 if(!bodyText())problems.push('write the email body');
 el('go').disabled=problems.length>0;
 note('goNote',problems.length?('To continue: '+problems.join('; ')+'.'):(sel.length+' email'+(sel.length===1?'':'s')+' will be drafted'+(el('fresh').checked?' after fresh reports are generated':' with each provider\'s latest existing PDF')+'.'),false);
}
el('selAll').onclick=()=>{[...el('recipients').querySelectorAll('input[type=checkbox]')].forEach(cb=>{const i=+cb.dataset.i;cb.checked=!!el('recipients').querySelector('input[data-email="'+i+'"]').value.trim()});updateGo()};
el('selNone').onclick=()=>{[...el('recipients').querySelectorAll('input[type=checkbox]')].forEach(cb=>cb.checked=false);updateGo()};
[...document.querySelectorAll('.toolbar button')].forEach(b=>{b.onmousedown=e=>e.preventDefault();b.onclick=()=>{el('body').focus();const c=b.dataset.cmd;if(c==='link'){const u=prompt('Link address (https://...)');if(u)document.execCommand('createLink',false,u)}else{document.execCommand(c,false,null)}saveDraft();updateGo()}});
['subject','body','sig','fresh'].forEach(id=>{el(id).addEventListener('input',()=>{saveDraft();updateGo()});el(id).addEventListener('change',()=>{saveDraft();updateGo()})});
el('pool').onchange=async()=>{saveDraft();await loadRecipients()};
el('useList').onclick=useList;el('refreshLists').onclick=refreshLists;
el('go').onclick=async()=>{
 updateGo();if(el('go').disabled)return;const sel=selected();
 if(!confirm('Generate '+(el('fresh').checked?'fresh reports and ':'')+'Outlook drafts for '+sel.length+' provider'+(sel.length===1?'':'s')+' in '+el('pool').value+'? Nothing is sent; drafts wait in Outlook for your review.'))return;
 el('go').disabled=true;note('goNote','Starting...',false);
 try{const c=await api('/api/campaign',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({riskPool:el('pool').value,subject:el('subject').value,bodyHtml:el('body').innerHTML,includeSignature:el('sig').checked,freshReports:el('fresh').checked,providers:sel})});startCampaign(c);note('goNote','Started. Progress is shown below; you can leave this tab open while reports run.',false)}
 catch(e){note('goNote','Could not start: '+e.message,true);updateGo()}
};
function queueHtml(q){
 if(!q)return '';
 const parts=q.running.map(r=>{const stuck=r.minutesSinceProgress>=5||!r.workerAlive;return '<div class="queue'+(stuck?' stuck':'')+'"><b>Report running:</b> '+esc(r.displayName)+' &middot; '+r.percent+'% '+esc(r.stage)+' &middot; last progress '+esc(when(r.lastProgressUtc))+(r.minutesSinceProgress>=2?' ('+r.minutesSinceProgress+' min ago)':'')+(r.workerAlive?'':' &middot; <span class="error">worker process not found</span>')+(r.attempts>1?' &middot; attempt '+r.attempts:'')+' <button class="alt" data-reset="'+esc(r.jobId)+'">Stop and requeue</button>'+(stuck?'<br><span class="muted">This report is holding up the queue. The server stops and retries it on its own after about 15 minutes without progress; press Stop and requeue to do that now.</span>':'')+'</div>'});
 return parts.join('')+(q.queued?'<p class="muted">'+q.queued+' report'+(q.queued===1?'':'s')+' waiting in the queue; reports run one at a time.</p>':'');
}
function campaignHtml(c,newest){
 const s=c.summary||{};
 const rows=c.providers.map(p=>{const rep=p.reportState==='Completed'?'<span class="ok">Completed</span>'+(p.stage==='Existing report'?' <span class="muted">(existing)</span>':''):p.reportState==='Failed'?'<span class="error">Failed</span>':'<div class="progress"><span style="width:'+Math.max(0,Math.min(100,p.percent||0))+'%"></span></div><span class="muted">'+esc(p.reportState||'Queued')+' '+(p.percent||0)+'% '+esc(p.stage||'')+'</span>';
  const dr=p.draftState==='Created'?'<span class="ok">Saved to Drafts</span> <span class="muted">'+esc(when(p.draftUtc))+'</span>':p.draftState==='Failed'?'<span class="error">Failed</span>':p.draftState==='Skipped'?'<span class="warn">Skipped</span>':p.reportState==='Completed'?'<span class="muted">Saving...</span>':'<span class="muted">Waiting for the report</span>';
  return '<tr><td><b>'+esc(p.displayName)+'</b><br><span class="muted">'+esc(p.email)+' &middot; Dear '+esc(p.greeting)+',</span></td><td>'+rep+'</td><td>'+dr+(p.error?'<br><span class="error">'+esc(p.error)+'</span>':'')+'</td></tr>'}).join('');
 const head='<p><b>'+esc(c.subject)+'</b> <span class="muted">'+esc(c.riskPool)+' &middot; started '+esc(when(c.createdUtc))+' &middot; '+esc(c.state)+'</span></p>';
 const sum='<p>'+(s.draftsCreated||0)+' of '+(s.total||0)+' drafts saved to Outlook'+(s.draftsFailed?', <span class="error">'+s.draftsFailed+' failed</span>':'')+(s.skipped?', '+s.skipped+' skipped':'')+(s.reportsPending?', '+s.reportsPending+' report'+(s.reportsPending===1?'':'s')+' still to run':'')+'.'+(c.state==='Completed'?' Open Outlook, review each message in the Drafts folder, then send.':' Reports and drafts keep going in the background even if this page is closed; the computer is kept from idle sleep while work is queued (closing the lid still sleeps it). An interrupted report is retried once automatically.')+'</p><div class="row">'+((s.draftsFailed||s.skipped)?'<button data-retry="'+esc(c.campaignId)+'">Retry failed items</button>':'')+(c.state!=='Completed'?'<button class="alt" data-cancel="'+esc(c.campaignId)+'">Cancel remaining</button>':'')+(newest&&c.state==='Completed'?'<button class="alt" id="another">Start another email</button>':'')+'</div>';
 return '<div class="camp">'+head+'<table><thead><tr><th>Provider</th><th>Report</th><th>Outlook draft</th></tr></thead><tbody>'+rows+'</tbody></table>'+sum+'</div>';
}
function shownCampaigns(list){return list.filter((c,i)=>i===0||c.state!=='Completed')}
function renderAll(data){
 const list=shownCampaigns(data.campaigns);const txt=JSON.stringify([data.queue,list]);if(txt===lastCampaignJson)return;lastCampaignJson=txt;
 if(!list.length){el('progressCard').style.display='none';return}
 el('progressCard').style.display='';el('campaign').innerHTML=queueHtml(data.queue)+list.map((c,i)=>campaignHtml(c,i===0)).join('');
 [...el('campaign').querySelectorAll('button[data-retry]')].forEach(b=>b.onclick=()=>act('/api/campaign-resume?id='+b.dataset.retry,b,'Retry failed'));
 [...el('campaign').querySelectorAll('button[data-cancel]')].forEach(b=>b.onclick=()=>{if(confirm('Cancel the remaining reports and drafts of this email? Drafts already saved stay in Outlook.'))act('/api/campaign-cancel?id='+b.dataset.cancel,b,'Cancel')});
 [...el('campaign').querySelectorAll('button[data-reset]')].forEach(b=>b.onclick=()=>{if(confirm('Stop this report and put it back in the queue?'))act('/api/job-reset?jobId='+b.dataset.reset,b,'Stop and requeue')});
 const a=el('another');if(a)a.onclick=()=>{updateGo();window.scrollTo(0,0)};
}
async function act(url,btn,label){btn.disabled=true;try{await api(url,{method:'POST'});lastCampaignJson='';startPoll()}catch(e){alert(label+' failed: '+e.message);btn.disabled=false}}
function startPoll(){if(!pollTimer)pollTimer=setInterval(poll,2500);poll()}
function startCampaign(c){lastCampaignJson='';startPoll()}
function stopPoll(){if(pollTimer){clearInterval(pollTimer);pollTimer=null}}
async function poll(){
 if(pollBusy)return;pollBusy=true;
 try{const data=await api('/api/campaigns');renderAll(data);const list=shownCampaigns(data.campaigns);
  if(!runner&&data.queue.queued>0&&data.queue.running.length===0){runner=true;try{await fetch('/api/run-next',{method:'POST'})}finally{runner=false}}
  const drafting=list.find(c=>c.summary&&c.summary.draftsPending>0);
  if(drafting&&!draftBusy){draftBusy=true;try{await api('/api/campaign-drafts?id='+drafting.campaignId,{method:'POST'});lastCampaignJson=''}catch(e){el('campaign').insertAdjacentHTML('beforeend','<p class="error">Drafting failed: '+esc(e.message)+'</p>')}finally{draftBusy=false}}
  if(list.every(c=>c.state==='Completed')&&data.queue.running.length===0)stopPoll();
 }catch(e){el('campaign').innerHTML='<p class="error">'+esc(e.message)+'</p>';stopPoll()}
 finally{pollBusy=false}
}
async function init(){
 const m=await api('/api/communication');fillLists(m);fillPools(m.pools);
 const d=loadDraft();if(d&&d.pool&&m.pools.includes(d.pool)){el('pool').value=d.pool;await loadRecipients()}else{updateGo()}
 try{startPoll()}catch(e){}
}
init().catch(e=>{document.body.insertAdjacentHTML('beforeend','<p class="error" style="padding:22px">Error: '+esc(e.message)+'</p>')});
</script></body></html>
'@
 return $html.Replace('__NAV__',(Get-NavHtml 'communication')).Replace('__TABCSS__',$script:TabCss)
}
# --- Draft 4: analysis transformation, HTML/PDF publishing, and persisted execution ---
function ConvertTo-HtmlEncoded([object]$Value){return [Net.WebUtility]::HtmlEncode([string]$Value)}
function ConvertTo-ProviderSlug([string]$Name){$s=($Name.ToLowerInvariant() -replace '[^a-z0-9]+','-').Trim('-');if(!$s){$s='provider'};return $s}
function ConvertTo-DateValue($Value){if($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)){return $null};if($Value -is [DateTime]){return [DateTime]$Value};$n=0.0;if([double]::TryParse(([string]$Value),[ref]$n) -and $n -gt 20000 -and $n -lt 80000){return [DateTime]::FromOADate($n)};$d=[DateTime]::MinValue;if([DateTime]::TryParse(([string]$Value),[ref]$d)){return $d};return $null}
function ConvertTo-DateText($Value){$d=ConvertTo-DateValue $Value;if($d){return $d.ToString('M/d/yyyy')};return ''}
function Get-Alias($Aliases,[string]$Key){if($null -eq $Aliases){return ''};$p=$Aliases.PSObject.Properties[$Key];if($p){return [string]$p.Value};return ''};function Get-P($Object,[string]$Name,$Default){$p=$Object.PSObject.Properties[$Name];if($p){return $p.Value};return $Default};function Set-P($Object,[string]$Name,$Value){$p=$Object.PSObject.Properties[$Name];if($p){$p.Value=$Value}else{$Object|Add-Member -NotePropertyName $Name -NotePropertyValue $Value}}
function Test-TrueValue($Value){return ([string]$Value).Trim() -match '^(1|Y|YES|TRUE)$'}
function ConvertTo-ProviderKey([string]$Name){if(!$Name){return ''};$n=($Name.ToUpperInvariant() -replace '[^A-Z0-9, ]',' ' -replace '\s+',' ').Trim();if($n -match '^([^,]+),\s*([^ ]+)'){return (($matches[1] -replace '[- ]','')+'|'+($matches[2] -replace '[- ]',''))};return ($n -replace '[- ,]','')}
function ConvertTo-NumberValue($Value){$n=0.0;$s=([string]$Value).Trim().TrimEnd('%');if([double]::TryParse($s,[ref]$n)){return $n};return 0};function Set-JobProgress($Job,[int]$Percent,[string]$Stage){$Job.state='Running';$Job.percent=$Percent;$Job.stage=$Stage;Save-JsonAtomic (Join-Path $script:Paths.State ('job-'+$Job.jobId+'.json')) $Job}
function Get-SourceRows([string]$Key,[string[]]$Wanted,[string]$FilterHeader,[string]$FilterValue){
 # The filter column (and the forward-filled Cdo/Provider columns) are read in bulk; the remaining fields are read only for rows that pass the filter.
 $source=Get-SourceConfig $Key;$path=Join-Path $script:Paths.CanonicalCurrent $source.canonicalFileName;if(!(Test-Path $path)){throw ('Canonical source missing: '+$source.displayName)}
 $carryHeaders=@('Cdo','Provider');$package=$null;$rows=New-Object Collections.Generic.List[object];$found=$false
 try{
  $package=Open-ExcelPackage -Path $path -ErrorAction Stop
  foreach($sheet in $package.Workbook.Worksheets){
   $header=Find-IdentityHeader $source $sheet;if($null -eq $header){continue}
   $found=$true;$map=$header.map;$first=$header.row+1;$last=$sheet.Dimension.End.Row
   if($last -lt $first){break}
   $count=$last-$first+1;$bulk=@{}
   $bulkHeaders=@($Wanted|Where-Object{$carryHeaders -contains $_});if($FilterHeader){$bulkHeaders+=$FilterHeader}
   foreach($hdr in @($bulkHeaders|Select-Object -Unique)){
    if(!$map.ContainsKey($hdr)){continue}
    $vals=Get-ColumnText $sheet $map[$hdr] $first $last
    if($carryHeaders -contains $hdr){$carry='';for($i=0;$i -lt $count;$i++){if($vals[$i]){$carry=$vals[$i]}elseif($carry){$vals[$i]=$carry}}}
    $bulk[$hdr]=$vals
   }
   $filtering=[bool]($FilterHeader -and $FilterValue);$target=$(if($filtering){$FilterValue.Trim()}else{''})
   $filterVals=$(if($filtering -and $bulk.ContainsKey($FilterHeader)){$bulk[$FilterHeader]}else{$null})
   for($i=0;$i -lt $count;$i++){
    if($filtering){$fv=$(if($null -ne $filterVals){$filterVals[$i]}else{''});if($fv -ne $target){continue}}
    $r=$first+$i;$o=[ordered]@{}
    foreach($name in $Wanted){
     if($bulk.ContainsKey($name)){$o[$name]=$bulk[$name][$i];continue}
     $v='';if($map.ContainsKey($name)){$cell=$sheet.Cells[$r,$map[$name]];$v=$(if($cell.Value -is [DateTime]){([DateTime]$cell.Value).ToString('o')}else{[string]$cell.Text})}
     $o[$name]=$v
    }
    $rows.Add([pscustomobject]$o)
   }
   break
  }
  if(!$found){throw ('Provider identity schema was not found for '+$source.displayName)}
 }finally{if($package){Close-ExcelPackage $package -NoSave}}
 return @($rows.ToArray())
}
function New-UniqueIndex([object[]]$Rows,[string]$Key){$counts=@{};$first=@{};foreach($r in $Rows){$id=([string]$r.$Key).Trim();if(!$id){continue};$counts[$id]=1+[int]$counts[$id];if(!$first.ContainsKey($id)){$first[$id]=$r}};$out=@{};foreach($id in $first.Keys){if($counts[$id] -eq 1){$out[$id]=$first[$id]}};return $out}
function Get-LatestDateFromText([string]$Text){$best=$null;foreach($part in @($Text -split ',')){$d=ConvertTo-DateValue $part;if($d -and $d -le (Get-Date) -and (!$best -or $d -gt $best)){$best=$d}};return $best}
function Get-Field($Object,[string]$Name,$Default){if($null -eq $Object){return $Default};if($Object -is [Collections.IDictionary]){if($Object.Contains($Name)){return $Object[$Name]};return $Default};return (Get-P $Object $Name $Default)}
function Get-ProviderPatientFacts($Job){
 # Reads the seven sources for one provider and returns plain patient facts (no scoring, no tables) so a report can be rebuilt later without Excel.
 $a=$Job.aliases;$exportAlias=Get-Alias $a 'Export';$qualityAlias=Get-Alias $a 'PtListQuality';$serialAlias=Get-Alias $a 'SerialScheduling';$dscAlias=Get-Alias $a 'DiabetesScorecard';$rpoAlias=Get-Alias $a 'RiskPopulationOutreach';$hrKey=$(if(Get-Alias $a 'HR-CRH'){'HR-CRH'}elseif(Get-Alias $a 'HR-RIVPHNYCMM'){'HR-RIVPHNYCMM'}else{''});$hrAlias=$(if($hrKey){Get-Alias $a $hrKey}else{''})
 $exportFields=@('Provider Name','Risk Pool','NPI','MemberID','First Name','Last Name','Date of Birth','Payer','Last QEM date with non-PCP','Last ACV Date','Last QEM Visit Date with any PCP in assigned TIN','Completed Attestations','Incompleted Attestations','Open ICDs','New Patient','Total Care Gaps','TCM')
 $qualityFields=@('Provider Name','Risk Pool','NPI','MemberID','First Name','Last Name','DOB','BCS','COLO','EED','GSD','CBP','OMW','KED','SPC','MAD','MAC','MAH','SUPD','COB','POLY','OMW Critical Due Date')
 $dscFields=@('Provider','Cdo','Patient','Member ID','Mrn','KED','EED','Eye Exam Gap Status','Eye Exam Date','Next Appt Date','Next Appt Specialty','Next Appt Location','Risk','GSD','Med Adherence DM','MAD Days Supply','Dx Date','Avg Last A1c','% eGFR last 12 mo.','% uACR last 12 mo.')
 $serialFields=@('Provider Name','Risk Pool','Member ID','MRN','Future PCP Visits 2026','PCP Visit Dates','A1c Date');$hrFields=@('PCP Name','Patient First Name','Patient Last Name','DOB','Patient Insurance ID');$rpoFields=@('Epic Pcp','Cdo','Member ID','Mrn','Next Acv')
 $exports=@(Get-SourceRows 'Export' $exportFields 'Provider Name' $exportAlias);Set-JobProgress $Job 15 'Loaded Export source';$quality=@(Get-SourceRows 'PtListQuality' $qualityFields 'Provider Name' $qualityAlias);Set-JobProgress $Job 28 'Loaded PtListQuality source';$dsc=@();if($dscAlias){$dsc=@(Get-SourceRows 'DiabetesScorecard' $dscFields 'Provider' $dscAlias)};Set-JobProgress $Job 40 'Loaded Diabetes Scorecard source';$serial=@();if($serialAlias){$serial=@(Get-SourceRows 'SerialScheduling' $serialFields 'Provider Name' $serialAlias)};Set-JobProgress $Job 50 'Loaded Serial Scheduling source';$rpo=@();if($rpoAlias){$rpo=@(Get-SourceRows 'RiskPopulationOutreach' $rpoFields 'Epic Pcp' $rpoAlias)};Set-JobProgress $Job 54 'Loaded Risk Population Outreach source';$hr=@();if($hrKey -and $hrAlias){$hr=@(Get-SourceRows $hrKey $hrFields 'PCP Name' $hrAlias)};Set-JobProgress $Job 58 'Loaded HR source'
 $qIndex=New-UniqueIndex $quality 'MemberID';$dIndex=New-UniqueIndex $dsc 'Member ID';$sIndex=New-UniqueIndex $serial 'Member ID';$rpoIndex=New-UniqueIndex $rpo 'Member ID';$hrIds=@{};foreach($x in $hr){$id=([string]$x.'Patient Insurance ID').Trim();if($id){$hrIds[$id]=$x}};$exportIds=New-Object Collections.Generic.List[string];foreach($x in $exports){$exportIds.Add(([string]$x.MemberID).Trim())};$highIds=@{};$unmatchedHr=@{};foreach($hid in $hrIds.Keys){$matchCount=0;$matchId='';foreach($eid in $exportIds){if($eid -eq $hid -or $eid.StartsWith($hid)){$matchCount++;if($matchCount -eq 1){$matchId=$eid}}};if($matchCount -eq 1){$highIds[$matchId]=$true}elseif($matchCount -eq 0){$unmatchedHr[$hid]=$hrIds[$hid]}}
 $hedis=@('BCS','COLO','EED','GSD','CBP','OMW','KED','SPC','MAD','MAC','MAH','SUPD','COB','POLY');$patients=@();$seen=@{}
 foreach($e in $exports){$id=([string]$e.MemberID).Trim();if(!$id){continue};$seen[$id]=$true;$q=$(if($qIndex.ContainsKey($id)){$qIndex[$id]}else{$null});$d=$(if($dIndex.ContainsKey($id)){$dIndex[$id]}else{$null});$s=$(if($sIndex.ContainsKey($id)){$sIndex[$id]}else{$null});$rp=$(if($rpoIndex.ContainsKey($id)){$rpoIndex[$id]}else{$null})
  $open=@();$closed=@();if($q){foreach($m in $hedis){$v=([string]$q.$m).Trim();if($v -in @('Needed','Non-compliant')){$open+=$m}elseif($v -in @('Completed','Compliant')){$closed+=$m}}}
  if($d){$overrides=@{'KED'=[string]($d.KED);'EED'=[string]($d.EED);'GSD'=[string]($d.GSD);'MAD'=[string]($d.'Med Adherence DM')};foreach($m in $overrides.Keys){$v=$overrides[$m];if($v -in @('Needed','Non-compliant')){if($open -notcontains $m){$open+=$m};$closed=@($closed|Where-Object{$_ -ne $m})}elseif($v -in @('Completed','Compliant')){$open=@($open|Where-Object{$_ -ne $m});if($closed -notcontains $m){$closed+=$m}}};$egfrComplete=(ConvertTo-NumberValue $d.'% eGFR last 12 mo.') -ge 100;$uacrComplete=(ConvertTo-NumberValue $d.'% uACR last 12 mo.') -ge 100;if($egfrComplete -and $uacrComplete){$open=@($open|Where-Object{$_ -ne 'KED'});if($closed -notcontains 'KED'){$closed+='KED'}}}
  $high=$highIds.ContainsKey($id);$lastAcv=(ConvertTo-DateValue $e.'Last ACV Date');$acor=([string]$e.Payer).Trim() -eq 'ACOR';$due=$false;if($lastAcv){$due=$(if($acor){$lastAcv -le (Get-Date).AddDays(-366)}else{$lastAcv -lt (Get-Date -Day 1 -Month 1)})};$nextAcv=$null;if($rp){$nextAcv=ConvertTo-DateValue $rp.'Next Acv'};if($nextAcv -and $nextAcv.Date -ge (Get-Date).Date){$due=$false}
  $vis=@((ConvertTo-DateValue $e.'Last QEM date with non-PCP'),(ConvertTo-DateValue $e.'Last ACV Date'),(ConvertTo-DateValue $e.'Last QEM Visit Date with any PCP in assigned TIN'))|Where-Object{$_};$last=$(if($vis){$vis|Sort-Object -Descending|Select-Object -First 1}else{$null});$serialLast=$(if($s){Get-LatestDateFromText ([string]$s.'PCP Visit Dates')}else{$null})
  $mrn='';foreach($cand in @($(if($d){$d.Mrn}else{''}),$(if($s){$s.MRN}else{''}),$(if($rp){$rp.Mrn}else{''}))){$c=([string]$cand).Trim();if($c -and $c -ne '0'){$mrn=$c;break}}
  $patients+=[pscustomobject][ordered]@{MemberID=$id;MRN=$mrn;First=[string]$e.'First Name';Last=[string]$e.'Last Name';DOB=$(ConvertTo-DateText $e.'Date of Birth');ACOR=$acor;ACVDue=$due;Attest=((ConvertTo-NumberValue $e.'Completed Attestations')+(ConvertTo-NumberValue $e.'Incompleted Attestations')) -gt 0;OpenICD=[string]$e.'Open ICDs';NewPatient=(Test-TrueValue $e.'New Patient');OpenHedis=($open -join ', ');ClosedHedis=($closed -join ', ');HighRisk=$high;LastVisit=$last;SerialLast=$serialLast;A1cDate=$(if($s){ConvertTo-DateValue $s.'A1c Date'}else{$null});FutureVisits=$(if($s){[string]$s.'Future PCP Visits 2026'}else{''});NextAppt=$(if($d){ConvertTo-DateValue $d.'Next Appt Date'}else{$null});NextApptSpecialty=$(if($d){([string]$d.'Next Appt Specialty').Trim()}else{''});A1c=$(if($d){ConvertTo-NumberOrNull $d.'Avg Last A1c'}else{$null});A1cText=$(if($d){([string]$d.'Avg Last A1c').Trim()}else{''});Diabetic=($null -ne $d);EyeOpen=[bool]($d -and (([string]$d.'Eye Exam Gap Status').Trim() -eq 'Open'));EyeDate=$(if($d){ConvertTo-DateValue $d.'Eye Exam Date'}else{$null});EgfrNeeded=[bool]($d -and ((ConvertTo-NumberValue $d.'% eGFR last 12 mo.') -lt 100));UacrNeeded=[bool]($d -and ((ConvertTo-NumberValue $d.'% uACR last 12 mo.') -lt 100));MedAdhDM=$(if($d){([string]$d.'Med Adherence DM').Trim()}else{''});OmwDue=$(if($q){ConvertTo-DateText $q.'OMW Critical Due Date'}else{''});TotalGaps=(ConvertTo-NumberValue $e.'Total Care Gaps');IncompleteAttest=(ConvertTo-NumberValue $e.'Incompleted Attestations');OpenIcdCount=@(([string]$e.'Open ICDs') -split ','|Where-Object{$_.Trim() -and $_.Trim() -ne '0'}).Count;Tcm=[bool]((Test-TrueValue $e.TCM) -or ((ConvertTo-NumberValue $e.TCM) -gt 0));RiskContract=$true;DSC=$d}
 }
 foreach($hid in $unmatchedHr.Keys){$x=$unmatchedHr[$hid];$patients+=[pscustomobject][ordered]@{MemberID=$hid;MRN='';First=[string]$x.'Patient First Name';Last=[string]$x.'Patient Last Name';DOB=$(ConvertTo-DateText $x.DOB);ACOR=$null;ACVDue=$false;Attest=$false;OpenICD='';NewPatient=$false;OpenHedis='';ClosedHedis='';HighRisk=$true;LastVisit=$null;SerialLast=$null;A1cDate=$null;FutureVisits='';NextAppt=$null;NextApptSpecialty='';A1c=$null;A1cText='';Diabetic=$false;EyeOpen=$false;EyeDate=$null;EgfrNeeded=$false;UacrNeeded=$false;MedAdhDM='';OmwDue='';TotalGaps=0;IncompleteAttest=0;OpenIcdCount=0;Tcm=$false;RiskContract=$true;DSC=$null}}
 return [ordered]@{asOf=(Get-Date).ToString('o');riskPool=$(if($exports.Count){[string]$exports[0].'Risk Pool'}else{''});patients=@($patients);scorecardNonRisk=@($dsc|Where-Object{([string]$_.Risk) -match 'Non-Risk Population'})}
}
$script:PatientDateFields=@('LastVisit','SerialLast','A1cDate','NextAppt','EyeDate')
function ConvertFrom-PatientFacts($Stored){
 # Rebuilds patient facts from analysis-<jobId>.json (dates come back as strings or DateTime depending on the PowerShell version).
 if($null -eq $Stored){throw 'This report was generated before flag support was added; generate a new report for this provider first.'}
 $patients=@();foreach($p in @(Get-Field $Stored 'patients' @())){if($null -eq $p){continue};foreach($f in $script:PatientDateFields){Set-P $p $f (ConvertTo-DateValue (Get-P $p $f $null))};if($null -eq (Get-P $p 'MRN' $null)){Set-P $p 'MRN' ''};$patients+=$p}
 return [ordered]@{asOf=[string](Get-Field $Stored 'asOf' '');riskPool=[string](Get-Field $Stored 'riskPool' '');patients=@($patients);scorecardNonRisk=@(Get-Field $Stored 'scorecardNonRisk' @()|Where-Object{$null -ne $_})}
}
function New-ProviderAnalysis($Job){
 $facts=Get-ProviderPatientFacts $Job
 Set-JobProgress $Job 62 'Scoring outreach needs'
 $flags=@(Get-ProviderFlags ([string]$Job.providerKey))
 $provider=[ordered]@{displayName=$Job.displayName;npi=$Job.providerKey;riskPool=$(Get-P $Job 'riskPool' $facts.riskPool);location=[string](Get-P $Job 'location' '');aliases=$Job.aliases}
 return (New-AnalysisModel $provider $facts $flags)
}
function New-ReferenceTables([object[]]$Patients,[object[]]$NonRiskDsc){
 $now=Get-AsOfDate
 $openRows=@();$openIds=@();$openNames=@();foreach($p in @($Patients|Where-Object{$_.OpenHedis}|Sort-Object Last,First)){$openRows+=,@($p.First,$p.Last,$p.DOB,$p.OpenHedis,$(ConvertTo-DateText $p.SerialLast),$(if($p.OpenICD){$p.OpenICD}else{'0'}),$(if($p.ACOR -eq $true -and $p.ACVDue){'Yes'}else{'No'}),$(if($p.ACOR -eq $false -and $p.ACVDue){'Yes'}else{'No'}));$openIds+=[string]$p.MemberID;$openNames+=(Get-PatientDisplayName $p)}
 $highRows=@();$highIds=@();$highNames=@();foreach($p in @($Patients|Where-Object{$_.HighRisk}|Sort-Object Last,First)){$highRows+=,@($p.First,$p.Last,$p.DOB,$p.OpenHedis,$(ConvertTo-DateText $p.LastVisit),$(if($p.OpenICD){$p.OpenICD}else{'0'}),$(if($p.ACOR -eq $true -and $p.ACVDue){'Yes'}else{'No'}),$(if($p.ACOR -eq $false -and $p.ACVDue){'Yes'}else{'No'}));$highIds+=[string]$p.MemberID;$highNames+=(Get-PatientDisplayName $p)}
 $diabRows=@();$diabIds=@();$diabNames=@();foreach($p in $Patients){$d=$p.DSC;if(!$d){continue};$metric=([string]$d.KED -eq 'Needed') -or ([string]$d.EED -eq 'Needed') -or ([string]$d.GSD -eq 'Needed') -or ([string]$d.'Med Adherence DM' -eq 'Non-compliant');$appt=(ConvertTo-DateValue $d.'Next Appt Date');if($metric -and (!$appt -or $appt -gt $now.AddMonths(3))){$diabRows+=,@($p.First,$p.Last,$p.DOB,$p.OpenHedis,$(if($appt){ConvertTo-DateText $appt}else{'Needed'}),[string]$d.'Next Appt Specialty',[string]$d.'Avg Last A1c',$(ConvertTo-DateText $p.A1cDate),$(Get-KedNeedsText $d));$diabIds+=[string]$p.MemberID;$diabNames+=(Get-PatientDisplayName $p)}}
 $nonRisk=@();$nonIds=@();$nonNames=@();foreach($d in $NonRiskDsc){$eyeDate=(ConvertTo-DateValue $d.'Eye Exam Date');$eye=([string]$d.'Eye Exam Gap Status' -eq 'Open') -and (!$eyeDate -or $eyeDate -le $now);$appt=(ConvertTo-DateValue $d.'Next Appt Date');$apptFlag=!$appt -or $appt -gt $now.AddMonths(3);$a1c=(ConvertTo-NumberValue $d.'Avg Last A1c') -ge 9;$egfr=(ConvertTo-NumberValue $d.'% eGFR last 12 mo.') -lt 100;$uacr=(ConvertTo-NumberValue $d.'% uACR last 12 mo.') -lt 100;if($a1c -and ($eye -or $apptFlag -or $egfr -or $uacr)){$reason=@();if($eye){$reason+='Eye exam'};if($apptFlag){$reason+='Appt needed/>3 mo'};if($a1c){$reason+='A1c >=9'};if($egfr){$reason+='eGFR needed'};if($uacr){$reason+='uACR needed'};$nonRisk+=,@([string]$d.Patient,[string]$d.'Eye Exam Gap Status',$(ConvertTo-DateText $eyeDate),$(if($appt){ConvertTo-DateText $appt}else{'Needed'}),[string]$d.'Next Appt Specialty',[string]$d.'Avg Last A1c',$(if(!$egfr){'Completed'}else{'Needed'}),$(if(!$uacr){'Completed'}else{'Needed'}),($reason -join '; '));$nonIds+=([string](Get-P $d 'Member ID' '')).Trim();$nonNames+=[string]$d.Patient}}
 $panelColumns=@('First Name','Last Name','DOB','Open HEDIS','Last Visit','Open ICDs','PCP AWV Due','NP AWV Due')
 return [ordered]@{openHedis=[ordered]@{columns=$panelColumns;rows=$openRows;memberIds=$openIds;memberNames=$openNames};highRisk=[ordered]@{columns=$panelColumns;rows=$highRows;memberIds=$highIds;memberNames=$highNames};diabetes=[ordered]@{columns=@('First Name','Last Name','DOB','Open HEDIS','Next Appt Date','Next Appt Specialty','Avg Last A1c','Last A1c Date','KED Needs');rows=$diabRows;memberIds=$diabIds;memberNames=$diabNames};nonRiskDiabetes=[ordered]@{columns=@('Patient (MRN)','Eye Exam Gap','Eye Exam Date','Next Appt Date','Next Appt Specialty','Avg Last A1c','eGFR','uACR','Reason Included');rows=$nonRisk;memberIds=$nonIds;memberNames=$nonNames}}
}
$script:ReferencePanelTitles=[ordered]@{openHedis='Open HEDIS List';highRisk='High Risk / Tuck-In Patient Panel';diabetes='Diabetes Care Gap Patient Panel';nonRiskDiabetes='Non-Risk Diabetes Scorecard Follow-Up Panel'}
function New-AnalysisModel($Provider,$Facts,[object[]]$Flags){
 # Scores every patient as of the facts date, removes flagged patients from every list and panel, and hangs them off the lists they would have reached.
 $prevAsOf=$script:AsOf;$script:AsOf=$(ConvertTo-DateValue (Get-Field $Facts 'asOf' $null))
 try{
  $patients=@(Get-Field $Facts 'patients' @()|Where-Object{$null -ne $_});$nonRiskDsc=@(Get-Field $Facts 'scorecardNonRisk' @()|Where-Object{$null -ne $_})
  foreach($p in $patients){Set-P $p 'Items' (Get-PatientItems $p);Set-P $p 'Needs' @(Get-PatientNeeds $p);Set-P $p 'Urgency' (Get-PatientUrgency $p)}
  $resolved=Resolve-ProviderFlags $Flags $patients $nonRiskDsc;$applied=$resolved.applied
  $active=@($patients|Where-Object{!$applied.ContainsKey([string]$_.MemberID)});$activeNonRisk=@($nonRiskDsc|Where-Object{!$applied.ContainsKey(([string](Get-P $_ 'Member ID' '')).Trim())})
  $outreach=@(New-OutreachTables $active);$tables=New-ReferenceTables $active $activeNonRisk
  if($applied.Count -gt 0){
   # Default placement: where would each flagged patient have landed with nobody excluded?
   $defaultOutreach=@(New-OutreachTables $patients);$defaultTables=New-ReferenceTables $patients $nonRiskDsc
   foreach($t in $defaultOutreach){foreach($id in @($t.memberIds)){if($applied.ContainsKey($id) -and !$applied[$id].defaultList){$applied[$id].defaultList=[string]$t.key;$applied[$id].defaultTitle=[string]$t.title}}}
   foreach($k in @($defaultTables.Keys)){foreach($id in @($defaultTables[$k].memberIds)){if($applied.ContainsKey($id)){$applied[$id].panels+=$k}}}
   $activeKeys=@($outreach|ForEach-Object{[string]$_.key})
   foreach($e in @($applied.Values)){
    $e.panelTitles=@(foreach($k in @($e.panels)){[string]$script:ReferencePanelTitles[$k]});$e.removedFrom=@($(if($e.defaultTitle){@($e.defaultTitle)}else{@()}))+@($e.panelTitles)
    $e.targetList=$(if($activeKeys -contains $e.defaultList){$e.defaultList}else{''})
    if(!$e.targetList -and $e.defaultList){
     if($e.defaultList -like 'outreach-awv*'){if($activeKeys -contains 'outreach-awv'){$e.targetList='outreach-awv'}elseif($e.awv -eq 'pcp' -and $activeKeys -contains 'outreach-awv-pcp'){$e.targetList='outreach-awv-pcp'}elseif($e.awv -eq 'np' -and $activeKeys -contains 'outreach-awv-np'){$e.targetList='outreach-awv-np'}}
     # Otherwise the live list sharing the most of the patient's items (a full match outranks a partial one); the note names both lists.
     if(!$e.targetList){$bestScore=0;foreach($t in $outreach){$need=@($t.items);if($need.Count -eq 0){continue};$shared=@($need|Where-Object{$e.items -contains $_}).Count;if($shared -eq 0){continue};$score=$shared*10+$(if($shared -eq $need.Count){5}else{0});if($score -gt $bestScore){$bestScore=$score;$e.targetList=[string]$t.key}}}
     $live=@($outreach|Where-Object{[string]$_.key -eq $e.targetList})[0]
     if($live){$e.targetTitle=[string]$live.title;$e.notes+=($script:FlagText.MovedList.Replace('{default}',$e.defaultTitle).Replace('{live}',$e.targetTitle))}else{$e.notes+=($script:FlagText.LostList.Replace('{default}',$e.defaultTitle))}
    }
   }
   foreach($t in $outreach){$members=@($applied.Values|Where-Object{$_.targetList -eq [string]$t.key});if($members.Count -gt 0){$t.flagged=New-FlaggedSubList $members}}
   foreach($k in @($tables.Keys)){$members=@($applied.Values|Where-Object{$_.panels -contains $k});if($members.Count -gt 0){$tables[$k].flagged=New-FlaggedSubList $members}}
  }
  $total=$active.Count;$acorCount=@($active|Where-Object{$_.ACOR -eq $true}).Count;$non=@($active|Where-Object{$null -ne $_.ACOR -and $_.ACOR -eq $false}).Count
  $kpi=[ordered]@{total=$total;acor=$acorCount;nonAcor=$non;acorAcvDue=@($active|Where-Object{$_.ACOR -eq $true -and $_.ACVDue}).Count;nonAcorAcvDue=@($active|Where-Object{$_.ACOR -eq $false -and $_.ACVDue}).Count;newPatients=@($active|Where-Object{$_.NewPatient}).Count;hasAttestations=@($active|Where-Object{$_.Attest}).Count;openIcd=@($active|Where-Object{$_.OpenICD}).Count;hasHedis=@($active|Where-Object{$_.OpenHedis -or $_.ClosedHedis}).Count;openHedis=@($active|Where-Object{$_.OpenHedis}).Count;highRisk=@($active|Where-Object{$_.HighRisk}).Count;flagged=$applied.Count}
  $narrative=([string]$Provider.displayName+' has '+$total+' patients: '+$acorCount+' ACOR and '+$non+' non-ACOR. New patients total '+$kpi.newPatients+'; '+$kpi.hasAttestations+' have attestations and '+$kpi.openIcd+' have Open ICD content.'+$(if($applied.Count -gt 0){' '+$applied.Count+' flagged patient(s) are excluded from every list, panel, and count above and listed under Flagged patients.'}else{''}))
  $storedPatients=@($patients|Select-Object -Property * -ExcludeProperty Items,Needs,Urgency)
  return [ordered]@{provider=$Provider;generatedUtc=[DateTime]::UtcNow.ToString('o');asOf=[string](Get-Field $Facts 'asOf' '');kpi=$kpi;narrative=$narrative;tables=$tables;outreach=@($outreach);flags=(New-FlagSummary $resolved $outreach);facts=[ordered]@{asOf=[string](Get-Field $Facts 'asOf' '');riskPool=[string](Get-Field $Facts 'riskPool' '');patients=$storedPatients;scorecardNonRisk=@($nonRiskDsc)};warnings=@()}
 }finally{$script:AsOf=$prevAsOf}
}
# --- Draft 5.1: per-provider patient flags (exceptions) ---
$script:FlagKinds=[ordered]@{'data-incorrect'=@{label='Data incorrect';order=0};'not-current'=@{label='Not a current patient';order=1}}
$script:FlagText=[ordered]@{
 SubListTitle='Flagged exceptions ({count})'
 SubListNote='Removed from this list by a saved flag; shown here so the exclusion stays visible. Data-incorrect flags are listed before not-a-current-patient flags, each by urgency.'
 SectionTitle='Flagged patients'
 SectionNote='Every patient excluded from the lists and panels above by a flag saved on this provider profile. Flags stay in force on every future report until they are cleared by hand; a note appears when the source data for a flagged patient has changed since the flag was saved.'
 NoDefault='Would not have reached an outreach list'
 MovedList='Default list "{default}" changed after exclusions; shown under "{live}"'
 LostList='Default list "{default}" no longer forms after exclusions; listed here only'
 Review='Needs review'
 Dormant='Inactive (patient not in any current source; the flag re-applies if the patient returns)'
 Empty='No patients are flagged for this provider.'
}
function Get-PatientDisplayName($p){$last=([string]$p.Last).Trim();$first=([string]$p.First).Trim();if($last -and $first){return ($last+', '+$first)};return ($last+$first)}
function Get-FlagNameKey([string]$Last,[string]$First,[string]$Dob){$l=($Last -replace '[^A-Za-z0-9]','').ToUpperInvariant();$f=($First -replace '[^A-Za-z0-9]','').ToUpperInvariant();if(!$l -or !$f){return ''};$d=ConvertTo-DateText $Dob;if(!$d){$d=([string]$Dob).Trim()};return ($l+'|'+$f+'|'+$d)}
function New-FlagSnapshot($p){if($null -eq $p){return $null};return [ordered]@{acvDue=[bool]$p.ACVDue;openHedis=[string]$p.OpenHedis;openIcdCount=[int]$p.OpenIcdCount;lastVisit=(ConvertTo-DateText (Get-PatientLastVisit $p));highRisk=[bool]$p.HighRisk}}
function Get-FlagChangeNote($Old,$New){
 if($null -eq $Old -or $null -eq $New){return ''}
 $changed=@();$labels=[ordered]@{acvDue='AWV due';openHedis='open HEDIS';openIcdCount='open ICD count';lastVisit='last visit';highRisk='high-risk status'}
 foreach($k in $labels.Keys){if([string](Get-Field $Old $k '') -ne [string](Get-Field $New $k '')){$changed+=$labels[$k]}}
 if($changed.Count -eq 0){return ''};return ('Source changed since flagged: '+($changed -join ', '))
}
function ConvertTo-IsoText($Value){if($null -eq $Value){return ''};if($Value -is [DateTime]){return ([DateTime]$Value).ToUniversalTime().ToString('o')};return [string]$Value}
function Get-FlagStorePath([string]$Npi){if($Npi -notmatch '^\d{10}$'){throw 'Invalid NPI.'};return (Join-Path $script:Paths.State ('flags-'+$Npi+'.json'))}
function Get-ProviderFlags([string]$Npi){if($Npi -notmatch '^\d{10}$'){return @()};$store=Json (Get-FlagStorePath $Npi);if($null -eq $store){return @()};return @(Get-P $store 'flags' @()|Where-Object{$null -ne $_})}
function Set-ProviderFlag([string]$Npi,[string]$MemberId,[string]$Kind,[string]$Note,$Identity,[string]$PreviousMemberId=''){
 # One patient per call; the server merges by member ID so two open pages never overwrite each other's other patients. Previous states are kept in history.
 $MemberId=([string]$MemberId).Trim();if(!$MemberId){throw 'Member ID is required.'};if($Kind -ne 'cleared' -and !$script:FlagKinds.Contains($Kind)){throw ('Unknown flag kind: '+$Kind)}
 $Note=([string]$Note -replace '[\r\n\t]+',' ').Trim();if($Note.Length -gt 1000){$Note=$Note.Substring(0,1000)}
 $path=Get-FlagStorePath $Npi;$store=Json $path;$now=[DateTime]::UtcNow.ToString('o');$flags=@();if($store){$flags=@(Get-P $store 'flags' @()|Where-Object{$null -ne $_})}
 $PreviousMemberId=([string]$PreviousMemberId).Trim();if(!$PreviousMemberId){$PreviousMemberId=$MemberId}
 $existing=@($flags|Where-Object{[string](Get-P $_ 'memberId' '') -eq $MemberId})[0];if(!$existing -and $PreviousMemberId -ne $MemberId){$existing=@($flags|Where-Object{[string](Get-P $_ 'memberId' '') -eq $PreviousMemberId})[0]}
 $history=@();if($existing){$history=@(Get-P $existing 'history' @()|Where-Object{$null -ne $_});$history=@([ordered]@{kind=[string](Get-P $existing 'kind' '');note=[string](Get-P $existing 'note' '');updatedUtc=(ConvertTo-IsoText (Get-P $existing 'updatedUtc' ''))})+$history;if($history.Count -gt 20){$history=@($history[0..19])}}
 $pick={param($name);$v=$null;if($null -ne $Identity){$v=Get-Field $Identity $name $null};if(($null -eq $v -or [string]$v -eq '') -and $existing){$v=Get-P $existing $name $null};return $v}
 $record=[ordered]@{memberId=$MemberId;first=[string](& $pick 'first');last=[string](& $pick 'last');dob=[string](& $pick 'dob');mrn=[string](& $pick 'mrn');kind=$Kind;note=$Note;createdUtc=$(if($existing){ConvertTo-IsoText (Get-P $existing 'createdUtc' $now)}else{$now});updatedUtc=$now;snapshot=(& $pick 'snapshot');history=@($history)}
 $out=@($flags|Where-Object{$mid=[string](Get-P $_ 'memberId' '');$mid -ne $MemberId -and $mid -ne $PreviousMemberId})+@($record)
 Save-JsonAtomic $path ([ordered]@{flagsVersion=1;npi=$Npi;updatedUtc=$now;flags=@($out)})
 Log 'FLAG_SAVED' 'OK' ('NPI '+$Npi+' member '+$MemberId+' '+$Kind)
 return $record
}
function New-FlagEntry($Flag,$Target,[string]$Match,[string]$StatusNote){
 $kind=[string](Get-P $Flag 'kind' '');$def=$script:FlagKinds[$kind];$notes=@();if($StatusNote){$notes+=$StatusNote}
 $name='';$dob='';$mrn='';$urgency=0;$snapshot=$null
 $awv='';$itemKeys=@();if($Target){$name=[string]$Target.name;$dob=[string]$Target.dob;$mrn=[string]$Target.mrn;$urgency=[int]$Target.urgency;$snapshot=$Target.snapshot;$awv=[string]$Target.awv;$itemKeys=@($Target.items);if($Match -eq 'name'){$notes+=('Matched by name and DOB (member ID changed from '+[string](Get-P $Flag 'memberId' '')+' to '+[string]$Target.id+')')};$change=Get-FlagChangeNote (Get-P $Flag 'snapshot' $null) $snapshot;if($change){$notes+=$change}}
 else{$name=Get-PatientDisplayName ([pscustomobject]@{Last=[string](Get-P $Flag 'last' '');First=[string](Get-P $Flag 'first' '')});$dob=[string](Get-P $Flag 'dob' '');$mrn=[string](Get-P $Flag 'mrn' '')}
 if(!$mrn){$mrn=[string](Get-P $Flag 'mrn' '')}
 return @{memberId=$(if($Target){[string]$Target.id}else{[string](Get-P $Flag 'memberId' '')});flagMemberId=[string](Get-P $Flag 'memberId' '');name=$name;dob=$dob;mrn=$mrn;kind=$kind;kindLabel=$(if($def){[string]$def.label}else{$kind});order=$(if($def){[int]$def.order}else{9});note=[string](Get-P $Flag 'note' '');urgency=$urgency;match=$Match;notes=@($notes);updatedUtc=(ConvertTo-IsoText (Get-P $Flag 'updatedUtc' ''));defaultList='';defaultTitle='';targetList='';targetTitle='';awv=$awv;items=$itemKeys;panels=@();panelTitles=@();removedFrom=@()}
}
function Resolve-ProviderFlags([object[]]$Flags,[object[]]$Patients,[object[]]$NonRiskDsc){
 # Member ID first; then an exact last name + first name + DOB match; more than one name match means nobody is excluded and the flag is listed for review.
 $byId=@{};$byName=@{}
 foreach($p in @($Patients)){$id=[string]$p.MemberID;$byId[$id]=@{id=$id;name=(Get-PatientDisplayName $p);dob=[string]$p.DOB;mrn=[string](Get-P $p 'MRN' '');urgency=[int](Get-P $p 'Urgency' 0);snapshot=(New-FlagSnapshot $p);awv=$(if($p.Items.Contains('awv-pcp')){'pcp'}elseif($p.Items.Contains('awv-np')){'np'}else{''});items=@($p.Items.Keys)};$nk=Get-FlagNameKey ([string]$p.Last) ([string]$p.First) ([string]$p.DOB);if($nk){if(!$byName.ContainsKey($nk)){$byName[$nk]=@()};$byName[$nk]+=$id}}
 foreach($d in @($NonRiskDsc)){$id=([string](Get-P $d 'Member ID' '')).Trim();if(!$id -or $byId.ContainsKey($id)){continue};$byId[$id]=@{id=$id;name=[string](Get-P $d 'Patient' '');dob='';mrn=([string](Get-P $d 'Mrn' '')).Trim();urgency=0;snapshot=$null;awv='';items=@()}}
 $applied=@{};$review=@();$dormant=@()
 foreach($f in @($Flags)){
  if($null -eq $f){continue};$kind=[string](Get-P $f 'kind' '');if(!$script:FlagKinds.Contains($kind)){continue}
  $fid=([string](Get-P $f 'memberId' '')).Trim();$target=$null;$match='id'
  if($fid -and $byId.ContainsKey($fid)){$target=$fid}
  else{$nk=Get-FlagNameKey ([string](Get-P $f 'last' '')) ([string](Get-P $f 'first' '')) ([string](Get-P $f 'dob' ''));if($nk -and $byName.ContainsKey($nk)){$ids=@($byName[$nk]);if($ids.Count -eq 1){$target=$ids[0];$match='name'}else{$review+=New-FlagEntry $f $null 'ambiguous' ('Name and DOB match '+$ids.Count+' patients; nobody was excluded');continue}}}
  if($null -eq $target){$dormant+=New-FlagEntry $f $null 'dormant' '';continue}
  if($applied.ContainsKey($target)){$review+=New-FlagEntry $f $null 'ambiguous' ('Another flag already applies to '+$byId[$target].name+' (member ID '+$target+')');continue}
  $applied[$target]=New-FlagEntry $f $byId[$target] $match ''
 }
 return @{applied=$applied;review=@($review);dormant=@($dormant)}
}
function Sort-FlagEntries([object[]]$Entries){return @($Entries|Sort-Object -Property @{Expression={[int]$_.order}},@{Expression={-[int]$_.urgency}},@{Expression={[string]$_.name}})}
function Get-FlagReasonText($e){return ('['+[string]$e.kindLabel+']'+$(if($e.note){' '+[string]$e.note}else{''}))}
function New-FlaggedSubList([object[]]$Entries,[string[]]$ExtraColumns){
 $sorted=@(Sort-FlagEntries $Entries);$extra=@($ExtraColumns|Where-Object{$_});$rows=@();$ids=@();$names=@();$anyNote=$false
 foreach($e in $sorted){$row=@([string]$e.name,[string]$e.dob,[string]$e.mrn,(Get-FlagReasonText $e));foreach($c in $extra){$row+=$(switch($c){'Removed From'{$(if(@($e.removedFrom).Count -gt 0){(@($e.removedFrom) -join '; ')}else{$script:FlagText.NoDefault})}'Status'{[string](@($e.notes)[0])}default{''}})};$noteText=(@($e.notes) -join '; ');if($noteText){$anyNote=$true};$row+=$noteText;$rows+=,$row;$ids+=[string]$e.memberId;$names+=[string]$e.name}
 $columns=@('Patient','DOB','MRN','Reason')+$extra+@('Notes')
 if(!$anyNote){$rows=@(foreach($r in $rows){,@($r[0..($r.Count-2)])});$columns=@($columns[0..($columns.Count-2)])}
 return [ordered]@{columns=$columns;rows=@($rows);memberIds=@($ids);memberNames=@($names);count=$sorted.Count}
}
function New-FlagSummary($Resolved,[object[]]$Outreach){
 $applied=@($Resolved.applied.Values);$review=@($Resolved.review);$dormant=@($Resolved.dormant)
 $out=[ordered]@{count=$applied.Count;reviewCount=$review.Count;dormantCount=$dormant.Count;active=$null;review=$null;dormant=$null}
 if($applied.Count -gt 0){$out.active=New-FlaggedSubList $applied @('Removed From')}
 if($review.Count -gt 0){$rv=New-FlaggedSubList $review;$out.review=$rv}
 if($dormant.Count -gt 0){$out.dormant=New-FlaggedSubList $dormant}
 $out.entries=@(foreach($e in @($applied+$review+$dormant)){[ordered]@{memberId=[string]$e.memberId;flagMemberId=[string]$e.flagMemberId;kind=[string]$e.kind;note=[string]$e.note;match=[string]$e.match}})
 return $out
}
function Get-FlagReportModel([string]$JobId){
 # Rebuilds the report for a finished job from its stored patient facts plus the provider's current flags (no Excel access needed).
 if($JobId -notmatch '^[a-f0-9]{32}$'){throw 'Invalid job ID.'}
 $job=Json (Join-Path $script:Paths.State ('job-'+$JobId+'.json'));if(!$job){throw 'Job not found.'}
 $stored=Json (Join-Path $script:Paths.State ('analysis-'+$JobId+'.json'));if(!$stored){throw 'The analysis data for this report is no longer available; generate a new report.'}
 $facts=ConvertFrom-PatientFacts (Get-P $stored 'facts' $null)
 $sp=Get-P $stored 'provider' $null;$provider=[ordered]@{displayName=[string](Get-P $sp 'displayName' (Get-P $job 'displayName' ''));npi=[string](Get-P $sp 'npi' (Get-P $job 'providerKey' ''));riskPool=[string](Get-P $sp 'riskPool' '');location=[string](Get-P $sp 'location' '');aliases=(Get-P $sp 'aliases' $null)}
 $model=New-AnalysisModel $provider $facts @(Get-ProviderFlags ([string]$provider.npi))
 $model.generatedUtc=(ConvertTo-IsoText (Get-P $stored 'generatedUtc' $model.generatedUtc));$model.jobId=$JobId
 return $model
}
function Save-FlagFromRequest($Body){
 if($null -eq $Body){throw 'A JSON body is required.'}
 $jobId=[string](Get-P $Body 'jobId' '');if($jobId -notmatch '^[a-f0-9]{32}$'){throw 'Invalid job ID.'}
 $job=Json (Join-Path $script:Paths.State ('job-'+$jobId+'.json'));if(!$job){throw 'Job not found.'}
 $npi=[string](Get-P $job 'providerKey' '');$memberId=([string](Get-P $Body 'memberId' '')).Trim();$kind=[string](Get-P $Body 'kind' '');$note=[string](Get-P $Body 'note' '')
 $identity=$null;$stored=Json (Join-Path $script:Paths.State ('analysis-'+$jobId+'.json'))
 if($stored){$facts=ConvertFrom-PatientFacts (Get-P $stored 'facts' $null)
  $p=@($facts.patients|Where-Object{[string]$_.MemberID -eq $memberId})[0]
  if($p){$identity=@{first=[string]$p.First;last=[string]$p.Last;dob=[string]$p.DOB;mrn=[string](Get-P $p 'MRN' '');snapshot=(New-FlagSnapshot $p)}}
  else{$d=@($facts.scorecardNonRisk|Where-Object{([string](Get-P $_ 'Member ID' '')).Trim() -eq $memberId})[0];if($d){$identity=@{first='';last=[string](Get-P $d 'Patient' '');dob='';mrn=([string](Get-P $d 'Mrn' '')).Trim();snapshot=$null}}}}
 $record=Set-ProviderFlag $npi $memberId $kind $note $identity ([string](Get-P $Body 'previousMemberId' ''))
 return [ordered]@{npi=$npi;flag=$record;flagCount=@(Get-ProviderFlags $npi|Where-Object{$script:FlagKinds.Contains([string](Get-P $_ 'kind' ''))}).Count}
}
$script:ListColumns=@{'Removed From'=';';'Open ICDs'=',';'Open HEDIS'=',';'Reason Included'=';';'KED Needs'=',';'Open Measures'=',';'Why Ranked Here'=';';'Diabetes Measures'=',';'Cardio Measures'=',';'Screenings Due'=',';'Med Safety Measures'=',';'Also Needs'=';'}
function Get-KedNeedsText($Row){$needs=@();if((ConvertTo-NumberValue $Row.'% eGFR last 12 mo.') -lt 100){$needs+='eGFR'};if((ConvertTo-NumberValue $Row.'% uACR last 12 mo.') -lt 100){$needs+='uACR'};if($needs.Count -eq 0){return 'None'};return ($needs -join ', ')}
function ConvertTo-ListCellHtml([object]$Value,[string]$Separator){
 # Comma/semicolon lists wrap only between items: each item is a nowrap span inside a max-width block.
 $text=[string]$Value;$parts=@($text.Split($Separator)|ForEach-Object{$_.Trim()}|Where-Object{$_})
 if($parts.Count -le 1){return '<td>'+(ConvertTo-HtmlEncoded $text)+'</td>'}
 $spans=@();for($i=0;$i -lt $parts.Count;$i++){$suffix=$(if($i -lt $parts.Count-1){$Separator}else{''});$spans+='<span>'+(ConvertTo-HtmlEncoded ($parts[$i]+$suffix))+'</span>'}
 return '<td><div class="list">'+($spans -join ' ')+'</div></td>'
}
# --- Draft 5.0: ranked outreach lists (items shared by every patient on a list -> short dynamic tables with a preloaded-text reason builder) ---
$script:AsOf=$null
function Get-AsOfDate{if($script:AsOf){return [DateTime]$script:AsOf};return (Get-Date)}
$script:OutreachRules=[ordered]@{NoVisitMonths=3;NoVisitLongMonths=12;NoApptMonths=3;UrgentApptMonths=1;A1cHigh=9.0;A1cVeryHigh=10.0;A1cStaleMonths=12;ManyGaps=3;UrgentScore=6;SoonScore=3;PreferredRowsMin=8;PreferredRowsMax=12;MaxRows=25;MinRowsPerList=3;MaxLists=6;MaxItemsPerList=3;RiskDiabetesMultiplier=3}
$script:HedisMeasures=[ordered]@{
 EED=@{label='Eye exam (EED)';group='diabetes'};KED=@{label='Kidney evaluation (KED)';group='diabetes'};GSD=@{label='A1c control (GSD)';group='diabetes'};MAD=@{label='Diabetes med adherence (MAD)';group='diabetes'};SUPD=@{label='Statin in diabetes (SUPD)';group='diabetes'}
 CBP=@{label='Blood pressure control (CBP)';group='cardio'};MAH=@{label='Hypertension med adherence (MAH)';group='cardio'};MAC=@{label='Cholesterol med adherence (MAC)';group='cardio'};SPC=@{label='Statin therapy (SPC)';group='cardio'}
 BCS=@{label='Breast cancer screening (BCS)';group='screening'};COLO=@{label='Colorectal cancer screening (COLO)';group='screening'};OMW=@{label='Osteoporosis management after fracture (OMW)';group='screening'}
 COB=@{label='Concurrent opioid and benzodiazepine (COB)';group='medsafety'};POLY=@{label='Polypharmacy (POLY)';group='medsafety'}
}
# Items are the things a list is built around; every patient on a list shares all of that list's items (1 to MaxItemsPerList).
$script:OutreachItems=[ordered]@{
 'hedis-diabetes'=@{label='open diabetes measures (EED, KED, GSD, MAD, SUPD)';short='Diabetes HEDIS';ask='close the open diabetes measures';column='Diabetes Measures';group='diabetes';context=@('Avg Last A1c','Last A1c Date');sources=@('PtListQuality (HEDIS status)','Diabetes Scorecard (A1c, eye exam, eGFR/uACR, next appointment)','Serial Scheduling (A1c date, future visits)')}
 'hedis-cardio'=@{label='open cardiovascular measures (CBP, MAH, MAC, SPC)';short='Cardio HEDIS';ask='address blood pressure control and statin or antihypertensive adherence';column='Cardio Measures';group='cardio';context=@();sources=@('PtListQuality (HEDIS status)')}
 'hedis-screening'=@{label='open preventive screenings (BCS, COLO, OMW)';short='Screening HEDIS';ask='order the overdue screenings';column='Screenings Due';group='screening';context=@('OMW Critical Due');sources=@('PtListQuality (HEDIS status, OMW critical due date)')}
 'hedis-medsafety'=@{label='open medication-safety measures (COB, POLY)';short='Med Safety HEDIS';ask='reconcile medications for opioid and benzodiazepine overlap and polypharmacy';column='Med Safety Measures';group='medsafety';context=@();sources=@('PtListQuality (HEDIS status)')}
 'icds'=@{label='open ICDs to recapture';short='Open ICDs';ask='recapture the open ICDs';column='Open ICDs';group='';context=@();sources=@('Export (open ICDs)')}
 'awv-pcp'=@{label='an annual wellness visit due with the PCP (ACOR)';short='AWV due (PCP)';ask='complete the annual wellness visit with the PCP';column='AWV Due';group='';context=@();sources=@('Export (AWV due, payer)')}
 'awv-np'=@{label='an annual wellness visit due that an NP can complete (non-ACOR)';short='AWV due (NP)';ask='complete the annual wellness visit (NP-eligible)';column='AWV Due';group='';context=@();sources=@('Export (AWV due, payer)')}
}
foreach($itemKey in @($script:OutreachItems.Keys)){$script:OutreachItems[$itemKey].key=$itemKey}
# Boosters never qualify a patient for a list; they raise urgency and are explained in "Why Ranked Here". short = table cell text, label = reason-block text.
$script:NeedCatalog=[ordered]@{
 'a1c:very-high'=@{label='A1c 10 or higher';short='A1c 10+';weight=3;group='booster'}
 'a1c:high'=@{label='A1c 9 to 9.9';short='A1c 9-9.9';weight=2;group='booster'}
 'a1c:overdue'=@{label='A1c older than 12 months';short='A1c >12 mo';weight=1;group='booster'}
 'eye:open'=@{label='Eye exam gap open (scorecard)';short='Eye gap';weight=1;group='booster'}
 'new:unseen'=@{label='New patient, never seen';short='Never seen';weight=1;group='booster'}
 'attest:incomplete'=@{label='Incomplete attestations';short='Attestations';weight=1;group='booster'}
 'visit:none-record'=@{label='No visit on record in any source (assumed overdue)';short='No visit on record';weight=2;group='booster'}
 'visit:none-12mo'=@{label='No visit in 12 months';short='No visit 12 mo';weight=2;group='booster'}
 'visit:none-3mo'=@{label='No visit in 3 months';short='No visit 3 mo';weight=1;group='booster'}
 'appt:none-1mo'=@{label='No appointment within 1 month';short='No appt 1 mo';weight=2;group='booster'}   # weight 3 when nothing is booked within 3 months either (stated in the detail)
 'risk:high'=@{label='On the high-risk list';short='High risk';weight=2;group='booster'}
 'gaps:many'=@{label='3 or more total care gaps (gentle intensifier)';short='Care gaps 3+';weight=1;group='booster'}
 'tcm'=@{label='Had a transitional care (TCM) visit (gentle intensifier)';short='TCM';weight=1;group='booster'}
}
$script:OutreachText=[ordered]@{
 IncludedAll='{count} patient(s) share {items}; all are shown.'
 IncludedCut='{shown} of {qualifying} patients share {items}; showing those with urgency {threshold} or higher (lists aim for {min}-{prefmax} patients and never exceed {max}).'
 IncludedCapped='{shown} of {qualifying} patients share {items}; the {max} most urgent are shown (all have urgency {threshold}).'
 Ask='Every patient here needs an appointment. At that visit: {asks}.'
 Ranking='Urgency: 1 per open measure (diabetes measures count {mult} each because every panel patient is on a risk contract; KED counts 1 per missing lab), 1 for open ICDs, 1 for an AWV due{boosters}. Urgent = {urgent} or more, Soon = {soon} or more; ties go to the oldest last visit.'
 AwvNote='Annual wellness visits are high urgency for ACOR and non-ACOR patients alike, so they get their own list(s) first; {total} patient(s) are due{split}.'
 AwvSplit=', split into PCP (ACOR) and NP (non-ACOR) lists because more than {max} are due'
 Boosters=', plus {list}'
 Highlights='Highlights: {parts}.'
 Assignment='Each patient appears on one list only, the list where they share the most items; {leftover} qualifying patient(s) did not reach a list.'
 Omitted='Columns omitted because no patient on this list had an actionable value: {columns}.'
 Sources='Built from: {sources}.'
 Empty='No patient on this panel currently has open HEDIS measures, open ICDs, or an annual wellness visit due.'
}
$script:OutreachProtectedColumns=@('First Name','Last Name','DOB','Urgency','Diabetes Measures','Cardio Measures','Screenings Due','Med Safety Measures','Open ICDs','AWV Due')
$script:OutreachBlankValues=@('','No','None','0','Closed')
function ConvertTo-NumberOrNull($Value){$n=0.0;$s=([string]$Value).Trim().TrimEnd('%');if($s -and [double]::TryParse($s,[ref]$n)){return $n};return $null}
function New-Need([string]$Key,[string]$Detail,[int]$Weight=-1){
 $c=$script:NeedCatalog[$Key];if($null -eq $c){throw ('Unknown need key '+$Key)}
 return [ordered]@{key=$Key;label=[string]$c.label;short=[string]$c.short;weight=$(if($Weight -ge 0){$Weight}else{[int]$c.weight});group=[string]$c.group;detail=$Detail}
}
function Get-PatientLastVisit($p){$dates=@(@($p.LastVisit,$p.SerialLast)|Where-Object{$_});if($dates.Count -eq 0){return $null};return ($dates|Sort-Object -Descending|Select-Object -First 1)}
function Get-PatientItems($p){
 $r=$script:OutreachRules;$items=[ordered]@{}
 $open=@(([string]$p.OpenHedis) -split ',\s*'|Where-Object{$_})
 $mult=$(if($p.RiskContract){[int]$r.RiskDiabetesMultiplier}else{1})
 foreach($itemKey in @('hedis-diabetes','hedis-cardio','hedis-screening','hedis-medsafety')){
  $group=[string]$script:OutreachItems[$itemKey].group;$measures=@();$weight=0
  foreach($m in $open){
   $def=$script:HedisMeasures[$m];if($null -eq $def -or [string]$def.group -ne $group){continue}
   if($m -eq 'KED' -and $p.Diabetic){$labs=@();if($p.EgfrNeeded){$labs+='eGFR'};if($p.UacrNeeded){$labs+='uACR'};if($labs.Count -eq 0){continue};$measures+=('KED: '+($labs -join ', '));$weight+=$labs.Count}
   else{$measures+=$m;$weight+=1}
  }
  if($measures.Count -eq 0){continue}
  if($itemKey -eq 'hedis-diabetes'){$weight=$weight*$mult}
  $items[$itemKey]=[ordered]@{key=$itemKey;weight=$weight;detail=($measures -join ', ')}
 }
 if($p.OpenIcdCount -gt 0){$items['icds']=[ordered]@{key='icds';weight=1;detail=[string]$p.OpenICD}}
 if($p.ACVDue){if($p.ACOR -eq $true){$items['awv-pcp']=[ordered]@{key='awv-pcp';weight=1;detail='PCP (ACOR)'}}else{$items['awv-np']=[ordered]@{key='awv-np';weight=1;detail='NP-eligible (non-ACOR)'}}}
 return $items
}
function Get-PatientNeeds($p){
 $r=$script:OutreachRules;$today=(Get-AsOfDate).Date;$needs=New-Object Collections.Generic.List[object]
 $open=@(([string]$p.OpenHedis) -split ',\s*'|Where-Object{$_})
 if($null -ne $p.A1c){if($p.A1c -ge $r.A1cVeryHigh){$needs.Add((New-Need 'a1c:very-high' ('A1c '+$p.A1cText)))}elseif($p.A1c -ge $r.A1cHigh){$needs.Add((New-Need 'a1c:high' ('A1c '+$p.A1cText)))}}
 if($p.A1cDate -and $p.A1cDate -lt $today.AddMonths(-$r.A1cStaleMonths)){$needs.Add((New-Need 'a1c:overdue' ('last '+(ConvertTo-DateText $p.A1cDate))))}
 if($p.Diabetic -and $p.EyeOpen -and $open -notcontains 'EED'){$needs.Add((New-Need 'eye:open' ''))}
 $last=Get-PatientLastVisit $p
 if(!$last){$needs.Add((New-Need 'visit:none-record' ''))}
 elseif($last -lt $today.AddMonths(-$r.NoVisitLongMonths)){$needs.Add((New-Need 'visit:none-12mo' ('last '+(ConvertTo-DateText $last))))}
 elseif($last -lt $today.AddMonths(-$r.NoVisitMonths)){$needs.Add((New-Need 'visit:none-3mo' ('last '+(ConvertTo-DateText $last))))}
 $future=ConvertTo-NumberOrNull $p.FutureVisits
 # Appointment needs fire when a scheduling source exists, or when no visit exists anywhere (a blank everywhere is read as overdue, not unknown).
 if($p.Diabetic -or $null -ne $future -or !$last){
  $within1=($p.NextAppt -and $p.NextAppt -le $today.AddMonths($r.UrgentApptMonths));$within3=($p.NextAppt -and $p.NextAppt -le $today.AddMonths($r.NoApptMonths));$hasFuture=($null -ne $future -and $future -gt 0)
  if(!$within1 -and !$hasFuture){if(!$within3){$needs.Add((New-Need 'appt:none-1mo' 'none within 3 months' 3))}else{$needs.Add((New-Need 'appt:none-1mo' ''))}}
 }
 if($p.NewPatient -and !$last){$needs.Add((New-Need 'new:unseen' ''))}
 if($p.IncompleteAttest -gt 0){$needs.Add((New-Need 'attest:incomplete' ([string]$p.IncompleteAttest)))}
 if($p.HighRisk){$needs.Add((New-Need 'risk:high' ''))}
 if($p.TotalGaps -ge $r.ManyGaps){$needs.Add((New-Need 'gaps:many' ([string]$p.TotalGaps+' gaps')))}
 if($p.Tcm){$needs.Add((New-Need 'tcm' ''))}
 return $needs.ToArray()
}
function Get-PatientUrgency($p){$s=0;foreach($k in @($p.Items.Keys)){$s+=[int]$p.Items[$k].weight};foreach($n in @($p.Needs)){$s+=[int]$n.weight};return $s}
function Get-OutreachItemCombos{
 $keys=@($script:OutreachItems.Keys|Where-Object{$_ -notlike 'awv-*'});$max=[int]$script:OutreachRules.MaxItemsPerList;$out=@()
 for($i=0;$i -lt $keys.Count;$i++){
  $out+=,@($keys[$i])
  if($max -ge 2){for($j=$i+1;$j -lt $keys.Count;$j++){
   $out+=,@($keys[$i],$keys[$j])
   if($max -ge 3){for($k=$j+1;$k -lt $keys.Count;$k++){$out+=,@($keys[$i],$keys[$j],$keys[$k])}}
  }}
 }
 return $out
}
function Select-OutreachRows([object[]]$Sorted,[int]$ShowAllUpTo=-1){
 # $Sorted is ordered by score desc. Show everything up to ShowAllUpTo (default PreferredRowsMax); otherwise raise the urgency cutoff until the list reaches PreferredRowsMin without exceeding MaxRows.
 $r=$script:OutreachRules;$n=$Sorted.Count;if($ShowAllUpTo -lt 0){$ShowAllUpTo=[int]$r.PreferredRowsMax}
 if($n -le $ShowAllUpTo){return @{threshold=$(if($n){$Sorted[$n-1].score}else{0});rows=@($Sorted);mode='all'}}
 $thresholds=@($Sorted|ForEach-Object{$_.score}|Select-Object -Unique);$best=$null
 foreach($t in $thresholds){
  $count=@($Sorted|Where-Object{$_.score -ge $t}).Count
  if($count -gt $r.MaxRows){break}
  $best=@{threshold=$t;count=$count}
  if($count -ge $r.PreferredRowsMin){break}
 }
 if($null -eq $best){return @{threshold=$thresholds[0];rows=@($Sorted|Select-Object -First $r.MaxRows);mode='capped'}}
 return @{threshold=$best.threshold;rows=@($Sorted|Where-Object{$_.score -ge $best.threshold});mode='cut'}
}
function Get-OutreachCell([string]$Header,$Row){
 $p=$Row.p;$items=$Row.items
 switch($Header){
  'First Name'{return [string]$p.First}
  'Last Name'{return [string]$p.Last}
  'DOB'{return [string]$p.DOB}
  'Urgency'{return ($Row.tier+' ('+$Row.score+')')}
  'Diabetes Measures'{if($items.Contains('hedis-diabetes')){return [string]$items['hedis-diabetes'].detail};return ''}
  'Cardio Measures'{if($items.Contains('hedis-cardio')){return [string]$items['hedis-cardio'].detail};return ''}
  'Screenings Due'{if($items.Contains('hedis-screening')){return [string]$items['hedis-screening'].detail};return ''}
  'Med Safety Measures'{if($items.Contains('hedis-medsafety')){return [string]$items['hedis-medsafety'].detail};return ''}
  'Open ICDs'{if($items.Contains('icds')){return [string]$items['icds'].detail};return ''}
  'AWV Due'{if($items.Contains('awv-pcp')){return [string]$items['awv-pcp'].detail};if($items.Contains('awv-np')){return [string]$items['awv-np'].detail};return ''}
  'Also Needs'{$others=@();foreach($k in @($items.Keys)){if($Row.k -contains $k){continue};$def=$script:OutreachItems[$k];$others+=$(if($k -like 'hedis-*'){[string]$def.short+': '+$items[$k].detail}else{[string]$def.short})};return ($others -join '; ')}
  'Why Ranked Here'{return ((@($Row.boost|ForEach-Object{$(if($_.detail){$_.short+': '+$_.detail}else{$_.short})})) -join '; ')}
  'Avg Last A1c'{return [string]$p.A1cText}
  'Last A1c Date'{return (ConvertTo-DateText $p.A1cDate)}
  'KED Needs'{if(!$p.Diabetic){return ''};$n=@();if($p.EgfrNeeded){$n+='eGFR'};if($p.UacrNeeded){$n+='uACR'};if($n.Count -eq 0){return 'None'};return ($n -join ', ')}
  'Next Appt'{if($p.NextAppt){return ((ConvertTo-DateText $p.NextAppt)+$(if($p.NextApptSpecialty){' '+$p.NextApptSpecialty}else{''}))};$f=ConvertTo-NumberOrNull $p.FutureVisits;if($null -ne $f -and $f -gt 0){return ('PCP visit scheduled ('+[int]$f+')')};if($p.Diabetic -or $null -ne $f){return 'None scheduled'};return ''}
  'Last Visit'{$l=Get-PatientLastVisit $p;if($l){return (ConvertTo-DateText $l)};return 'None on record'}
  'OMW Critical Due'{return [string]$p.OmwDue}
  default{return ''}
 }
}
function Format-OutreachList([string[]]$Items,[string]$Word='and'){$u=@($Items|Select-Object -Unique);if($u.Count -eq 0){return ''};if($u.Count -eq 1){return $u[0]};return ((($u[0..($u.Count-2)]) -join ', ')+' '+$Word+' '+$u[$u.Count-1])}
function New-OutreachTable([string]$Key,[string]$Title,[object[]]$Members,[object[]]$ItemDefs,[string[]]$LeadColumns,[string]$ItemText,[string[]]$Asks,[string[]]$Sources,[int]$ShowAllUpTo,[string]$Note){
 $r=$script:OutreachRules;$t=$script:OutreachText
 $rows=@(foreach($p in $Members){$score=[int]$p.Urgency;@{p=$p;score=$score;tier=$(if($score -ge $r.UrgentScore){'Urgent'}elseif($score -ge $r.SoonScore){'Soon'}else{'Routine'});items=$p.Items;boost=@($p.Needs);k=@($ItemDefs|ForEach-Object{[string]$_.key});lastVisit=(Get-PatientLastVisit $p)}})
 $sorted=@($rows|Sort-Object -Property @{Expression={-$_.score}},@{Expression={if($_.lastVisit){[DateTime]$_.lastVisit}else{[DateTime]::MinValue}}},@{Expression={[string]$_.p.Last}},@{Expression={[string]$_.p.First}})
 $selection=Select-OutreachRows $sorted $ShowAllUpTo;$shown=@($selection.rows)
 $headers=@('First Name','Last Name','DOB','Urgency');foreach($c in $LeadColumns){if($headers -notcontains $c){$headers+=$c}}
 foreach($d in $ItemDefs){foreach($c in @($d.context)){if($headers -notcontains $c){$headers+=$c}}}
 $headers+=@('Also Needs','Last Visit','Next Appt','Why Ranked Here')
 $matrix=@(foreach($row in $shown){,@(foreach($h in $headers){Get-OutreachCell $h $row})})
 $keep=@();$omitted=@()
 for($c=0;$c -lt $headers.Count;$c++){
  $h=$headers[$c];$hasValue=$false
  foreach($m in $matrix){if($script:OutreachBlankValues -notcontains ([string]$m[$c]).Trim()){$hasValue=$true;break}}
  if($hasValue -or $script:OutreachProtectedColumns -contains $h){$keep+=$c}else{$omitted+=$h}
 }
 $columns=@($keep|ForEach-Object{$headers[$_]});$tableRows=@(foreach($m in $matrix){,@($keep|ForEach-Object{$m[$_]})})
 $fired=@{};foreach($row in $shown){foreach($n in @($row.boost)){$fired[$n.key+'|'+$n.weight]=$n}}
 $boosterText=$(if($fired.Count -gt 0){$t.Boosters.Replace('{list}',((@($fired.Values|Sort-Object -Property @{Expression={-$_.weight}},@{Expression={$_.label}}|ForEach-Object{$_.label+$(if($_.key -eq 'appt:none-1mo' -and $_.detail){' ('+$_.detail+')'}else{''})+' (+'+$_.weight+')'})) -join ', '))}else{''})
 $nUrgent=@($shown|Where-Object{$_.tier -eq 'Urgent'}).Count;$nSoon=@($shown|Where-Object{$_.tier -eq 'Soon'}).Count;$nRoutine=$shown.Count-$nUrgent-$nSoon
 $parts=@();if($nUrgent){$parts+=($nUrgent.ToString()+' urgent')};if($nSoon){$parts+=($nSoon.ToString()+' soon')};if($nRoutine){$parts+=($nRoutine.ToString()+' routine')}
 $nAcor=@($shown|Where-Object{$_.p.ACOR -eq $true}).Count;if($nAcor){$parts+=($nAcor.ToString()+' ACOR')}
 $countKeys=[ordered]@{'appt:none-1mo'='with no appointment within 1 month';'visit:none-record'='with no visit on record';'visit:none-12mo'='not seen in 12 months';'a1c:very-high'='with A1c 10 or higher';'a1c:high'='with A1c 9 to 9.9';'risk:high'='on the high-risk list'}
 foreach($ck in @($countKeys.Keys)){$n=@($shown|Where-Object{@($_.boost|Where-Object{$_.key -eq $ck}).Count -gt 0}).Count;if($n){$parts+=($n.ToString()+' '+$countKeys[$ck])}}
 $included=$(switch([string]$selection.mode){
  'all'{$t.IncludedAll.Replace('{count}',$shown.Count.ToString()).Replace('{items}',$ItemText)}
  'capped'{$t.IncludedCapped.Replace('{shown}',$shown.Count.ToString()).Replace('{qualifying}',$sorted.Count.ToString()).Replace('{items}',$ItemText).Replace('{max}',$r.MaxRows.ToString()).Replace('{threshold}',$selection.threshold.ToString())}
  default{$t.IncludedCut.Replace('{shown}',$shown.Count.ToString()).Replace('{qualifying}',$sorted.Count.ToString()).Replace('{items}',$ItemText).Replace('{threshold}',$selection.threshold.ToString()).Replace('{min}',$r.PreferredRowsMin.ToString()).Replace('{prefmax}',$r.PreferredRowsMax.ToString()).Replace('{max}',$r.MaxRows.ToString())}
 })
 $reason=@($included)
 if($Note){$reason+=$Note}
 $reason+=@(
  $t.Ask.Replace('{asks}',(Format-OutreachList $Asks)),
  $t.Ranking.Replace('{mult}',$r.RiskDiabetesMultiplier.ToString()).Replace('{boosters}',$boosterText).Replace('{urgent}',$r.UrgentScore.ToString()).Replace('{soon}',$r.SoonScore.ToString()),
  $t.Highlights.Replace('{parts}',($parts -join '; '))
 )
 if($omitted.Count -gt 0){$reason+=$t.Omitted.Replace('{columns}',($omitted -join ', '))}
 $reason+=$t.Sources.Replace('{sources}',(@($Sources|Select-Object -Unique) -join '; '))
 return [ordered]@{key=$Key;title=$Title;items=@($ItemDefs|ForEach-Object{[string]$_.key});ask=(Format-OutreachList $Asks);reason=@($reason);columns=$columns;rows=$tableRows;omittedColumns=@($omitted);counts=[ordered]@{qualifying=$sorted.Count;shown=$shown.Count;threshold=$selection.threshold;urgent=$nUrgent;soon=$nSoon;routine=$nRoutine};memberIds=@($shown|ForEach-Object{[string]$_.p.MemberID});memberNames=@($shown|ForEach-Object{Get-PatientDisplayName $_.p})}
}
function New-OutreachTables([object[]]$Patients){
 $r=$script:OutreachRules;$t=$script:OutreachText;$out=@()
 $pool=@($Patients|Where-Object{@($_.Items.Keys).Count -gt 0});$assigned=@{}
 $common=@('Export (visits, payer)','HR list')
 # AWV lists come first: one list when MaxRows or fewer are due, otherwise split by ACOR (PCP) and non-ACOR (NP).
 $awv=@($pool|Where-Object{$_.Items.Contains('awv-pcp') -or $_.Items.Contains('awv-np')})
 if($awv.Count -gt 0){
  $split=($awv.Count -gt [int]$r.MaxRows);$note=$t.AwvNote.Replace('{total}',$awv.Count.ToString()).Replace('{split}',$(if($split){$t.AwvSplit.Replace('{max}',$r.MaxRows.ToString())}else{''}))
  $awvGroups=@()
  if(!$split){$awvGroups+=@{key='outreach-awv';title='AWV Due (PCP and NP)';members=$awv;defs=@($script:OutreachItems['awv-pcp'],$script:OutreachItems['awv-np']);lead=@('AWV Due');itemText='an annual wellness visit due';asks=@('complete the annual wellness visit (with the PCP for ACOR patients; NP-eligible for non-ACOR)')}}
  else{
   $pcp=@($awv|Where-Object{$_.Items.Contains('awv-pcp')});$np=@($awv|Where-Object{$_.Items.Contains('awv-np')})
   if($pcp.Count -gt 0){$awvGroups+=@{key='outreach-awv-pcp';title='AWV Due - PCP (ACOR)';members=$pcp;defs=@($script:OutreachItems['awv-pcp']);lead=@();itemText=[string]$script:OutreachItems['awv-pcp'].label;asks=@([string]$script:OutreachItems['awv-pcp'].ask)}}
   if($np.Count -gt 0){$awvGroups+=@{key='outreach-awv-np';title='AWV Due - NP (non-ACOR)';members=$np;defs=@($script:OutreachItems['awv-np']);lead=@();itemText=[string]$script:OutreachItems['awv-np'].label;asks=@([string]$script:OutreachItems['awv-np'].ask)}}
  }
  foreach($g in $awvGroups){
   $table=New-OutreachTable $g.key $g.title $g.members $g.defs $g.lead $g.itemText $g.asks (@($g.defs|ForEach-Object{@($_.sources)})+$common) ([int]$r.MaxRows) $note
   foreach($id in $table.memberIds){$assigned[$id]=$true};$out+=$table
  }
 }
 # Then greedy combination lists over the remaining patients: the combination with the most patients times shared items wins each round.
 $usedCombos=@{};$combos=@(Get-OutreachItemCombos)
 for($round=0;$round -lt [int]$r.MaxLists;$round++){
  $best=$null
  foreach($combo in $combos){
   $comboKey=($combo -join '+');if($usedCombos.ContainsKey($comboKey)){continue}
   $members=@($pool|Where-Object{$pt=$_;if($assigned.ContainsKey([string]$pt.MemberID)){$false}else{$ok=$true;foreach($k in $combo){if(!$pt.Items.Contains($k)){$ok=$false;break}};$ok}})
   if($members.Count -lt [int]$r.MinRowsPerList){continue}
   $value=[Math]::Min($members.Count,[int]$r.MaxRows)*$combo.Count;$urgency=0;foreach($m in $members){$urgency+=[int]$m.Urgency}
   if($null -eq $best -or $value -gt $best.value -or ($value -eq $best.value -and $urgency -gt $best.urgency)){$best=@{combo=$combo;key=$comboKey;members=$members;value=$value;urgency=$urgency}}
  }
  if($null -eq $best){break}
  $usedCombos[$best.key]=$true;$combo=@($best.combo);$defs=@($combo|ForEach-Object{$script:OutreachItems[$_]})
  $table=New-OutreachTable ('outreach-'+($combo -join '-')) (@($defs|ForEach-Object{[string]$_.short}) -join ' + ') $best.members $defs @($defs|ForEach-Object{[string]$_.column}) (Format-OutreachList @($defs|ForEach-Object{[string]$_.label})) @($defs|ForEach-Object{[string]$_.ask}) (@($defs|ForEach-Object{@($_.sources)})+$common) -1 ''
  foreach($id in $table.memberIds){$assigned[$id]=$true};$out+=$table
 }
 $leftover=@($pool|Where-Object{!$assigned.ContainsKey([string]$_.MemberID)}).Count
 foreach($table in $out){$table.reason=@($table.reason)+@($t.Assignment.Replace('{leftover}',$leftover.ToString()))}
 return $out
}
$script:RenderInteractive=$false
function ConvertTo-RowsHtml($Table,[string[]]$Columns){
 # In flag mode every row carries its member ID and display name so the page can open an editor under it.
 $ids=@(Get-Field $Table 'memberIds' @());$names=@(Get-Field $Table 'memberNames' @());$body='';$n=0
 foreach($row in @(Get-Field $Table 'rows' @())){
  $values=@($row);$cells=''
  for($i=0;$i -lt $values.Count;$i++){
   $header=$(if($i -lt $Columns.Count){[string]$Columns[$i]}else{''})
   if($script:ListColumns.ContainsKey($header)){$cells+=ConvertTo-ListCellHtml $values[$i] $script:ListColumns[$header]}
   else{$cells+='<td>'+(ConvertTo-HtmlEncoded $values[$i])+'</td>'}
  }
  $attr='';if($script:RenderInteractive -and $n -lt $ids.Count -and [string]$ids[$n]){$attr=' data-member="'+(ConvertTo-HtmlEncoded $ids[$n])+'" data-name="'+(ConvertTo-HtmlEncoded $(if($n -lt $names.Count){$names[$n]}else{''}))+'"'}
  $body+='<tr'+$attr+'>'+$cells+'</tr>';$n++
 }
 return $body
}
function ConvertTo-FlaggedListHtml([string]$Id,[string]$Title,$Sub,[string]$Note){
 if($null -eq $Sub){return ''};$columns=@(Get-Field $Sub 'columns' @());$head=($columns|ForEach-Object{'<th>'+(ConvertTo-HtmlEncoded $_)+'</th>'}) -join ''
 $body=ConvertTo-RowsHtml $Sub $columns;if(!$body){$body='<tr><td colspan="'+$columns.Count+'">None</td></tr>'}
 return '<div class="flagged"><div class="section-head"><h3>'+(ConvertTo-HtmlEncoded $Title)+'</h3><button onclick="copyTable('''+$Id+''',this)">Copy table</button></div>'+$(if($Note){'<p class="muted">'+(ConvertTo-HtmlEncoded $Note)+'</p>'}else{''})+'<div class="table-wrap"><table id="'+$Id+'" class="flagged-rows"><thead><tr>'+$head+'</tr></thead><tbody>'+$body+'</tbody></table></div></div>'
}
function ConvertTo-TableHtml([string]$Id,[string]$Title,$Table,[string[]]$Description){
 $descHtml=$(if(@($Description).Count -gt 1){'<div class="reason">'+((@($Description)|ForEach-Object{'<p>'+(ConvertTo-HtmlEncoded $_)+'</p>'}) -join '')+'</div>'}else{'<p>'+(ConvertTo-HtmlEncoded ([string]$Description))+'</p>'})
 $columns=@(Get-Field $Table 'columns' @());$head=($columns|ForEach-Object{'<th>'+(ConvertTo-HtmlEncoded $_)+'</th>'}) -join ''
 $body=ConvertTo-RowsHtml $Table $columns
 if(!$body){$body='<tr><td colspan="'+$columns.Count+'">No matching patients</td></tr>'}
 $sub=Get-Field $Table 'flagged' $null;$subHtml=''
 if($sub){$subHtml=ConvertTo-FlaggedListHtml ($Id+'-flagged') ($script:FlagText.SubListTitle.Replace('{count}',([string](Get-Field $sub 'count' 0)))) $sub $script:FlagText.SubListNote}
 return '<section><div class="section-head"><h2>'+(ConvertTo-HtmlEncoded $Title)+'</h2><button onclick="copyTable('''+$Id+''',this)">Copy table</button></div>'+$descHtml+'<div class="table-wrap"><table id="'+$Id+'"><thead><tr>'+$head+'</tr></thead><tbody>'+$body+'</tbody></table></div>'+$subHtml+'</section>'
}
function ConvertTo-FlagSectionHtml($Flags){
 if($null -eq $Flags){return ''};$t=$script:FlagText;$count=[int](Get-Field $Flags 'count' 0)
 $html='<h2 class="group">'+(ConvertTo-HtmlEncoded $t.SectionTitle)+'</h2><section id="flagged-section"><div class="section-head"><h2>'+(ConvertTo-HtmlEncoded $t.SectionTitle)+' ('+$count+')</h2></div><p>'+(ConvertTo-HtmlEncoded $t.SectionNote)+'</p>'
 $active=Get-Field $Flags 'active' $null
 if($active){$html+=ConvertTo-FlaggedListHtml 'flagged-all' ('Excluded patients ('+$count+')') $active ''}else{$html+='<p class="muted">'+(ConvertTo-HtmlEncoded $t.Empty)+'</p>'}
 $review=Get-Field $Flags 'review' $null;if($review){$html+=ConvertTo-FlaggedListHtml 'flagged-review' ($t.Review+' ('+[string](Get-Field $review 'count' 0)+')') $review 'These flags could not be applied to exactly one patient, so nobody was excluded; open the patient row and re-save the flag to fix it.'}
 $dormant=Get-Field $Flags 'dormant' $null;if($dormant){$html+=ConvertTo-FlaggedListHtml 'flagged-dormant' ('Inactive ('+[string](Get-Field $dormant 'count' 0)+')') $dormant $t.Dormant}
 return $html+'</section>'
}
function ConvertTo-AnalysisHtml($Model,[switch]$Interactive){
 $prevInteractive=$script:RenderInteractive;$script:RenderInteractive=[bool]$Interactive
 try{
 $k=$Model.kpi;$tiles=@('Total Patients',$k.total,'ACOR Payor',$k.acor,'Non-ACOR Payor',$k.nonAcor,'ACOR + ACV Due',$k.acorAcvDue,'Non-ACOR + ACV Due',$k.nonAcorAcvDue,'New Patients',$k.newPatients,'Has Attestations',$k.hasAttestations,'Open ICDs Present',$k.openIcd,'Has HEDIS',$k.hasHedis,'Open HEDIS Present',$k.openHedis,'High Risk',$k.highRisk,'Flagged',(Get-Field $k 'flagged' 0));$tileHtml='';for($i=0;$i -lt $tiles.Count;$i+=2){$tileHtml+='<div class="tile"><span>'+ (ConvertTo-HtmlEncoded $tiles[$i]) +'</span><strong>'+ (ConvertTo-HtmlEncoded $tiles[$i+1]) +'</strong></div>'}
 $outreachList=@(@(Get-Field $Model 'outreach' @())|Where-Object{$null -ne $_});$outreachHtml='';foreach($t in $outreachList){$outreachHtml+=ConvertTo-TableHtml ([string]$t.key) ([string]$t.title) $t @($t.reason)};if(!$outreachHtml){$outreachHtml='<section><h2>Outreach lists</h2><p>'+(ConvertTo-HtmlEncoded $script:OutreachText.Empty)+'</p></section>'}
 $ref=$Model.tables
 $tables='<h2 class="group">Outreach lists (ranked by urgency)</h2>'+$outreachHtml+'<h2 class="group">Reference panels</h2>'+(ConvertTo-TableHtml 'openHedis' 'Open HEDIS List' (Get-Field $ref 'openHedis' $null) 'Patients with one or more open HEDIS metrics; review documentation and closure opportunities.')+(ConvertTo-TableHtml 'highRisk' 'High Risk / Tuck-In Patient Panel' (Get-Field $ref 'highRisk' $null) 'Entire high-risk panel, including visit and AWV follow-up context.')+(ConvertTo-TableHtml 'diabetes' 'Diabetes Care Gap Patient Panel' (Get-Field $ref 'diabetes' $null) 'Value-based diabetes patients with an open evidence-based metric and no appointment within three months.')+(ConvertTo-TableHtml 'nonRisk' 'Non-Risk Diabetes Scorecard Follow-Up Panel' (Get-Field $ref 'nonRiskDiabetes' $null) 'Non-risk patients with A1c at least 9.0 and at least one additional follow-up need.')+(ConvertTo-FlagSectionHtml (Get-Field $Model 'flags' $null))
 $generated=([DateTime]$Model.generatedUtc).ToLocalTime().ToString('g');$providerLocation=[string](Get-Field $Model.provider 'location' '')
 $style=@'
body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:#fff;padding:28px}main{max-width:1500px;margin:auto;padding:24px}.meta{color:#d9e6f5}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px}.tile,section{background:#fff;border:1px solid #dce4ef;border-radius:8px;padding:16px}.tile span{display:block;color:#667085;font-size:12px;text-transform:uppercase}.tile strong{font-size:28px;color:#17365d}
.story{font-size:17px}.section-head{display:flex;justify-content:space-between;align-items:center}button{background:#1769aa;color:#fff;border:0;border-radius:5px;padding:8px 12px;cursor:pointer}button.secondary{background:#e4e9f0;color:#172033}
.table-wrap{overflow:auto}table{border-collapse:collapse;width:100%}th,td{padding:.55em .6em;border-bottom:1px solid #dde4ed;text-align:left;white-space:nowrap;vertical-align:top}th{background:#eaf1f8}section{margin-top:18px}
.list{white-space:normal;max-width:36ch;line-height:1.35}.list span{white-space:nowrap}.reason{background:#f7f9fc;border-left:3px solid #1769aa;padding:6px 10px;margin:8px 0}.reason p{margin:3px 0}h2.group{margin:26px 0 4px;color:#17365d}.muted{color:#667085}
.flagged{margin:12px 0 0 0;padding:8px 12px;background:#fff6f6;border:1px solid #f1c9c9;border-radius:6px}.flagged h3{margin:4px 0;font-size:14px;color:#8a1c1c}.flagged-rows th{background:#fbe9e9}
.flagbar{position:sticky;top:0;z-index:5;display:flex;justify-content:space-between;align-items:center;gap:14px;flex-wrap:wrap;background:#fff7e6;border-bottom:2px solid #e0a800;padding:10px 22px}.flagbar a.btn{background:#e4e9f0;color:#172033;text-decoration:none;padding:8px 12px;border-radius:5px}
tr[data-member]{cursor:pointer}tr[data-member]:hover td{background:#fffbe6}tr.editor td{background:#fffdf3;white-space:normal}.flag-form{padding:6px 2px}.flag-form .radios label{margin-right:16px}.flag-form textarea{display:block;width:100%;max-width:640px;height:64px;margin:8px 0;font:inherit;box-sizing:border-box}.error{color:#a61b1b}
@page{size:letter landscape;margin:0.45in}
@media print{button,.flagbar,tr.editor{display:none}body{background:#fff;font-size:11px}header{padding:12px 16px}header h1{font-size:18px;margin:0 0 4px}header h2{font-size:15px;margin:0 0 4px}main{max-width:none;padding:0}
.tiles{gap:6px}.tile{padding:6px 8px;border-radius:4px}.tile span{font-size:9px}.tile strong{font-size:16px}
section{break-inside:auto;margin-top:8px;padding:6px 0;border:0;border-radius:0}section h2{font-size:14px;margin:4px 0}section p{margin:2px 0 6px}.story{font-size:12px}.flagged{padding:4px 8px;border-radius:0}
.table-wrap{overflow:visible}table{width:100%}thead{display:table-header-group}tr{break-inside:avoid}h2.group{font-size:15px;margin:12px 0 2px}.reason{font-size:10px;padding:4px 8px}}
'@
 $fitScript=@'
(function(){var PRINT_W=Math.round((11-0.9)*96),BASE=11;function fit(target){document.querySelectorAll('main table').forEach(function(t){var fs=BASE,guard=0;t.style.fontSize=fs+'px';t.style.width='auto';while(t.offsetWidth>target&&fs>6&&guard++<40){fs-=0.25;t.style.fontSize=fs+'px'}t.style.width=''})}function reset(){document.querySelectorAll('main table').forEach(function(t){t.style.fontSize=''})}if(location.hash==='#pdf'){fit(PRINT_W)}window.addEventListener('beforeprint',function(){fit(PRINT_W)});window.addEventListener('afterprint',reset)})();
'@
 $flagScript=@'
(function(){var FLAGS=__FLAGS__,JOB='__JOB__',NPI='__NPI__';function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
function closeEditor(){var e=document.getElementById('flagEditor');if(e)e.parentNode.removeChild(e)}
function openEditor(tr){closeEditor();var id=tr.getAttribute('data-member'),name=tr.getAttribute('data-name');var cur=FLAGS[id]||{kind:'cleared',note:'',flagMemberId:id};var kinds=[['cleared','No exception'],['data-incorrect','Data incorrect'],['not-current','Not a current patient']];var row=document.createElement('tr');row.id='flagEditor';row.className='editor';row.innerHTML='<td colspan="'+tr.cells.length+'"><div class="flag-form"><b>'+esc(name)+'</b> <span class="muted">Member ID '+esc(id)+'</span><div class="radios">'+kinds.map(function(k){return '<label><input type="radio" name="flagKind" value="'+k[0]+'"'+(cur.kind===k[0]?' checked':'')+'> '+k[1]+'</label>'}).join('')+'</div><textarea id="flagNote" placeholder="Reason (free text, optional)">'+esc(cur.note)+'</textarea><div class="actions"><button type="button" id="flagSave">Save flag</button> <button type="button" class="secondary" id="flagCancel">Cancel</button> <span id="flagMsg" class="muted"></span></div></div></td>';tr.parentNode.insertBefore(row,tr.nextSibling);
 document.getElementById('flagCancel').onclick=closeEditor;
 document.getElementById('flagSave').onclick=function(){var checked=row.querySelector('input[name=flagKind]:checked');var kind=checked?checked.value:'cleared';var note=document.getElementById('flagNote').value;var msg=document.getElementById('flagMsg');msg.textContent='Saving...';msg.className='muted';
  fetch('/api/flag',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({jobId:JOB,memberId:id,kind:kind,note:note,previousMemberId:cur.flagMemberId||id})}).then(function(r){return r.text().then(function(t){var j=null;try{j=t?JSON.parse(t):null}catch(e){}if(!r.ok)throw Error((j&&j.error)||('Request failed ('+r.status+')'));return j})}).then(function(){try{sessionStorage.setItem('flagScroll',String(window.scrollY))}catch(e){}location.reload()}).catch(function(e){msg.textContent='Could not save: '+e.message;msg.className='error'})};
 var ta=document.getElementById('flagNote');if(ta)ta.focus()}
document.addEventListener('click',function(ev){var t=ev.target;if(!t.closest)return;if(t.closest('#flagEditor')||t.closest('button')||t.closest('a'))return;var tr=t.closest('tr[data-member]');if(tr)openEditor(tr)});
try{var y=sessionStorage.getItem('flagScroll');if(y){sessionStorage.removeItem('flagScroll');window.scrollTo(0,parseInt(y,10)||0)}}catch(e){}
var regen=document.getElementById('flagRegen');if(regen)regen.onclick=function(){regen.disabled=true;regen.textContent='Queuing...';fetch('/api/profile-job?npi='+encodeURIComponent(NPI),{method:'POST'}).then(function(r){if(!r.ok)throw Error('Request failed ('+r.status+')');return fetch('/api/run-next',{method:'POST'})}).then(function(){regen.textContent='Queued - open Provider Index to follow it'}).catch(function(e){regen.disabled=false;regen.textContent='Regenerate report now';alert('Could not queue the report: '+e.message)})};
})();
'@
 $barHtml='';$interactiveScript=''
 if($Interactive){
  $entries=@{};foreach($e in @(Get-Field (Get-Field $Model 'flags' $null) 'entries' @())){if($null -eq $e){continue};$entries[[string]$e.memberId]=[ordered]@{kind=[string]$e.kind;note=[string]$e.note;flagMemberId=[string]$e.flagMemberId}}
  $flagsJson=$(if($entries.Count -gt 0){(ConvertTo-Json -InputObject $entries -Compress -Depth 4)}else{'{}'}).Replace('</','<\/')
  $jobId=[string](Get-Field $Model 'jobId' '');$asOf=ConvertTo-DateValue (Get-Field $Model 'asOf' $null)
  $barHtml='<div class="flagbar"><div><b>Flag mode</b> - click any patient row to mark the patient as <b>Data incorrect</b> or <b>Not a current patient</b>, or to clear a flag. Flags are saved to this provider profile and applied to every future report. This page mirrors the report with the current flags applied'+$(if($asOf){' (source data as of '+(ConvertTo-HtmlEncoded ($asOf.ToString('M/d/yyyy')))+')'}else{''})+'.</div><div class="flagbar-actions"><button type="button" id="flagRegen">Regenerate report now</button> <a class="btn" href="/provider-index">Provider Index</a>'+$(if($jobId){' <a class="btn" target="_blank" href="/html?jobId='+(ConvertTo-HtmlEncoded $jobId)+'">Published HTML</a>'}else{''})+'</div></div>'
  $interactiveScript=$flagScript.Replace('__FLAGS__',$flagsJson).Replace('__JOB__',$jobId).Replace('__NPI__',[string](Get-Field $Model.provider 'npi' ''))
 }
 return '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>'+ (ConvertTo-HtmlEncoded $Model.provider.displayName) +' - Provider Analysis'+$(if($Interactive){' (flags)'}else{''})+'</title><style>'+$style+'</style></head><body>'+$barHtml+'<header><h1>Provider Patient Dashboard</h1><h2>'+ (ConvertTo-HtmlEncoded $Model.provider.displayName) +'</h2><div class="meta">NPI '+(ConvertTo-HtmlEncoded $Model.provider.npi)+$(if($providerLocation){' | '+(ConvertTo-HtmlEncoded $providerLocation)}else{''})+' | Generated '+(ConvertTo-HtmlEncoded $generated)+'</div></header><main><div class="tiles">'+$tileHtml+'</div><section><h2>Story in Brief</h2><p class="story">'+(ConvertTo-HtmlEncoded $Model.narrative)+'</p></section>'+$tables+'</main><script>function copyTable(id,btn){const t=document.getElementById(id),text=[...t.rows].map(r=>[...r.cells].map(c=>c.innerText).join("\t")).join("\n");let ok=false,box=null;try{box=document.createElement("div");box.contentEditable="true";box.style.position="fixed";box.style.left="-10000px";box.style.top="0";box.style.background="white";box.appendChild(t.cloneNode(true));document.body.appendChild(box);const range=document.createRange();range.selectNodeContents(box);const sel=window.getSelection();sel.removeAllRanges();sel.addRange(range);ok=document.execCommand("copy");sel.removeAllRanges()}catch(e){}finally{if(box)box.remove()}if(ok){btn.textContent="Copied";setTimeout(()=>btn.textContent="Copy table",1600);return}const fallback=()=>{const a=document.createElement("textarea");a.value=text;a.style.position="fixed";a.style.left="-10000px";document.body.appendChild(a);a.focus();a.select();let done=false;try{done=document.execCommand("copy")}catch(e){}a.remove();btn.textContent=done?"Copied as text":"Copy failed";setTimeout(()=>btn.textContent="Copy table",1600)};if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(text).then(()=>{btn.textContent="Copied as text";setTimeout(()=>btn.textContent="Copy table",1600)}).catch(fallback)}else{fallback()}}'+$fitScript+$interactiveScript+'</script></body></html>'
 }finally{$script:RenderInteractive=$prevInteractive}
}
function Get-PdfBrowserArguments([string]$PdfPath,[string]$Uri){
 # #pdf makes the page pre-fit its tables to the landscape printable width before headless printing.
 return @('--headless','--disable-gpu','--disable-logging','--log-level=3','--no-first-run','--no-default-browser-check','--window-size=1400,900','--no-pdf-header-footer','--print-to-pdf-no-header',('--print-to-pdf='+$PdfPath),($Uri+'#pdf'))
}
function Convert-HtmlPdf([string]$HtmlPath,[string]$PdfPath){$candidates=@("$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe","$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe","$env:ProgramFiles\Google\Chrome\Application\chrome.exe","$env:ProgramFiles (x86)\Google\Chrome\Application\chrome.exe");$browser=@($candidates|Where-Object{Test-Path $_})[0];if(!$browser){throw 'Microsoft Edge or Google Chrome is required for PDF output.'};$uri=([Uri]$HtmlPath).AbsoluteUri;$stdout=$PdfPath+'.browser.out';$stderr=$PdfPath+'.browser.err';try{$p=Start-Process -FilePath $browser -ArgumentList (Get-PdfBrowserArguments $PdfPath $uri) -RedirectStandardOutput $stdout -RedirectStandardError $stderr -WindowStyle Hidden -PassThru;$null=$p.Handle;if(!$p.WaitForExit($script:PdfTimeoutSeconds*1000)){try{Stop-Process -Id $p.Id -Force -ErrorAction Stop}catch{};throw ('Browser PDF rendering timed out after '+$script:PdfTimeoutSeconds+' seconds.')};if($p.ExitCode -ne 0 -or !(Test-Path $PdfPath)){throw 'Browser PDF rendering failed.'}}finally{Remove-Item $stdout,$stderr -Force -ErrorAction SilentlyContinue}}
# --- Draft 5.3: unattended runs - heartbeat, interrupted-worker recovery, sleep deferral ---
$script:HeartbeatSeconds=15;$script:LastHeartbeat=[DateTime]::MinValue;$script:StallChecks=60;$script:MaxJobAttempts=2;$script:PdfTimeoutSeconds=180
$script:JobWatch=@{};$script:KeepAwake=$false;$script:KeepAwakeBroken=$false;$script:StartupStaleMinutes=15
function Test-ProcessAlive([int]$ProcessId){if($ProcessId -le 0){return $false};try{$p=Get-Process -Id $ProcessId -ErrorAction Stop;return (-not $p.HasExited)}catch{return $false}}
function Test-WorkerAlive($Job){
 # True only when the recorded pid is still a PowerShell worker for this job: process ids are reused, so a live pid alone proves nothing.
 $workerPid=[int](Get-P $Job 'workerPid' 0);if($workerPid -le 0){return $false}
 try{$p=Get-Process -Id $workerPid -ErrorAction Stop}catch{return $false}
 if($p.HasExited){return $false}
 if([string]$p.ProcessName -notmatch '^(powershell|pwsh)'){return $false}
 try{$cmd=[string](Get-CimInstance Win32_Process -Filter ('ProcessId = '+$workerPid) -ErrorAction Stop).CommandLine;if($cmd){return ($cmd -like ('*'+[string](Get-P $Job 'jobId' '')+'*'))}}catch{}
 return $true
}
function Stop-WorkerProcess([int]$ProcessId){
 if($ProcessId -le 0){return}
 try{foreach($child in @(Get-CimInstance Win32_Process -Filter ('ParentProcessId = '+$ProcessId) -ErrorAction Stop)){try{Stop-Process -Id $child.ProcessId -Force -ErrorAction Stop}catch{}}}catch{}
 try{Stop-Process -Id $ProcessId -Force -ErrorAction Stop}catch{}
}
function Reset-InterruptedJob($Job,[string]$Path,[string]$Reason){
 # An interrupted report is queued again once; a second interruption fails it with an explanation so nothing loops forever.
 $attempts=[int](Get-P $Job 'attempts' 1);$pct=[int](Get-P $Job 'percent' 0);$id=[string]$Job.jobId
 if($attempts -lt $script:MaxJobAttempts){$Job.state='Prepared';$Job.percent=0;$Job.stage=('Interrupted at '+$pct+'% because '+$Reason+'; queued to run again');Set-P $Job 'errorSummary' '';Set-P $Job 'workerPid' 0;Set-P $Job 'interruptedUtc' ([DateTime]::UtcNow.ToString('o'));Log 'JOB_RECOVERED' 'OK' ($id+' requeued: '+$Reason)}
 else{$Job.state='Failed';$Job.stage='Analysis failed';Set-P $Job 'errorSummary' ('Interrupted at '+$pct+'% because '+$Reason+'. It had already been retried once, so it was not queued again; use Generate report, or Retry failed items on the Communication tab, to run it once more.');Set-P $Job 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $Job 'workerPid' 0;Log 'JOB_RECOVERED' 'FAILED' ($id+' '+$Reason)}
 Save-JsonAtomic $Path $Job
}
function Repair-StalledJobs{
 # A job whose worker process is gone (laptop slept, process closed) or that has shown no progress for StallChecks consecutive heartbeats is recovered.
 # Progress is judged by counting heartbeats, not wall-clock time, so a long sleep does not by itself condemn a healthy worker.
 foreach($f in @(Get-ChildItem $script:Paths.State -Filter 'job-*.json' -File -ErrorAction SilentlyContinue)){
  $job=Json $f.FullName;if(!$job){continue};$id=[string](Get-P $job 'jobId' '');$state=[string](Get-P $job 'state' '')
  if($state -notin @('Starting','Running')){if($script:JobWatch.ContainsKey($id)){$script:JobWatch.Remove($id)};continue}
  $workerPid=[int](Get-P $job 'workerPid' 0);$stamp=(Get-Item -LiteralPath $f.FullName).LastWriteTimeUtc.Ticks
  if(!$script:JobWatch.ContainsKey($id) -or $script:JobWatch[$id].stamp -ne $stamp){$script:JobWatch[$id]=@{stamp=$stamp;checks=0}}else{$script:JobWatch[$id].checks++}
  $checks=[int]$script:JobWatch[$id].checks;$reason=''
  if($workerPid -le 0){if($checks -ge 4){$reason='the worker process never started'}}
  elseif(!(Test-WorkerAlive $job)){$reason='the worker process ended without finishing (the computer may have slept or the process was closed)'}
  elseif($checks -ge $script:StallChecks){Stop-WorkerProcess $workerPid;$reason=('the worker made no progress for about '+[int]($script:StallChecks*$script:HeartbeatSeconds/60)+' minutes and was stopped')}
  if(!$reason){continue}
  Reset-InterruptedJob $job $f.FullName $reason;$script:JobWatch.Remove($id)
 }
}
function Get-QueueStatus{
 # What the report queue is doing right now, for the Communication tab: running jobs with their last progress time and whether the worker still exists, plus the queued count.
 $running=@();$queued=0
 foreach($j in @(Get-Jobs)){
  $state=[string](Get-P $j 'state' '');if($state -in @('Prepared','Queued')){$queued++;continue};if($state -notin @('Starting','Running')){continue}
  $f=Join-Path $script:Paths.State ('job-'+[string]$j.jobId+'.json');$last=(Get-Item -LiteralPath $f).LastWriteTimeUtc
  $running+=[ordered]@{jobId=[string]$j.jobId;displayName=[string](Get-P $j 'displayName' '');state=$state;percent=[int](Get-P $j 'percent' 0);stage=[string](Get-P $j 'stage' '');attempts=[int](Get-P $j 'attempts' 1);lastProgressUtc=$last.ToString('o');minutesSinceProgress=[int][Math]::Floor(([DateTime]::UtcNow-$last).TotalMinutes);workerAlive=[bool](Test-WorkerAlive $j)}
 }
 return [ordered]@{running=@($running);queued=$queued;heartbeatUtc=(ConvertTo-IsoText $script:LastHeartbeat)}
}
function Reset-RunningJob([string]$JobId){
 # Manual escape hatch: stop a running (or stuck) report's worker and queue the report again, whatever its retry count.
 if($JobId -notmatch '^[a-f0-9]{32}$'){throw 'Invalid job ID.'};$path=Join-Path $script:Paths.State ('job-'+$JobId+'.json');$job=Json $path;if(!$job){throw 'Job not found.'}
 if([string](Get-P $job 'state' '') -notin @('Starting','Running')){throw 'That report is not running.'}
 if(Test-WorkerAlive $job){Stop-WorkerProcess ([int](Get-P $job 'workerPid' 0))}
 Set-P $job 'attempts' 1;Reset-InterruptedJob $job $path 'it was stopped from the Communication tab';if($script:JobWatch.ContainsKey($JobId)){$script:JobWatch.Remove($JobId)};$script:RunNext=$true
 return (Json $path)
}
function Stop-Campaign([string]$Id){
 # Cancels everything still pending in a campaign: queued or running reports are stopped and their drafts skipped; finished drafts are untouched.
 $c=Get-CampaignModel $Id;$cancelled=0
 foreach($p in @($c.providers)){
  if($p.draftState -ne 'Pending'){continue}
  if($p.jobId -and $p.reportState -in @('Prepared','Queued','Starting','Running')){$path=Join-Path $script:Paths.State ('job-'+$p.jobId+'.json');$job=Json $path;if($job){if([string](Get-P $job 'state' '') -in @('Starting','Running') -and (Test-WorkerAlive $job)){Stop-WorkerProcess ([int](Get-P $job 'workerPid' 0))};$job.state='Failed';$job.stage='Cancelled';Set-P $job 'errorSummary' 'Cancelled from the Communication tab';Set-P $job 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $job 'workerPid' 0;Save-JsonAtomic $path $job};$p.reportState='Failed'}
  $p.draftState='Skipped';$p.error='Cancelled';$cancelled++
 }
 Update-CampaignState $c;Save-JsonAtomic (Get-CampaignPath $Id) $c;Log 'CAMPAIGN_CANCELLED' 'OK' ($Id+'; '+$cancelled+' cancelled');return $c
}
function Invoke-PendingCampaignDrafts{
 # Saves drafts for every campaign whose reports have finished; returns $true while any campaign still has work outstanding.
 $outstanding=$false
 foreach($c in @(Get-Campaigns)){if([string](Get-P $c 'state' '') -eq 'Completed'){continue};$m=Get-CampaignModel ([string]$c.campaignId);if([int]$m.summary.draftsPending -gt 0){$m=Invoke-CampaignDrafts ([string]$m.campaignId)};if([string]$m.state -ne 'Completed'){$outstanding=$true}}
 return $outstanding
}
function Set-KeepAwake([bool]$On){
 # Defers idle sleep while reports or drafts are queued (Windows SetThreadExecutionState); closing the lid still sleeps the computer.
 if($On -eq $script:KeepAwake -or $script:KeepAwakeBroken){return}
 try{if(-not ('PA.Power' -as [type])){Add-Type -Namespace PA -Name Power -MemberDefinition '[DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint esFlags);'};$flags=$(if($On){[uint32]2147483649}else{[uint32]2147483648});$null=[PA.Power]::SetThreadExecutionState($flags);$script:KeepAwake=$On;Log 'KEEP_AWAKE' 'OK' $(if($On){'Idle sleep deferred while work is queued'}else{'Sleep deferral released'})}
 catch{$script:KeepAwakeBroken=$true;Log 'KEEP_AWAKE' 'FAILED' $_.Exception.Message}
}
function Invoke-Heartbeat{
 # Runs every HeartbeatSeconds from the server loop whether or not a browser tab is open: recovers interrupted workers, launches the next queued report, saves campaign drafts, and manages sleep deferral.
 $script:LastHeartbeat=[DateTime]::UtcNow
 try{Repair-StalledJobs}catch{Log 'JOB_RECOVERY' 'FAILED' $_.Exception.Message}
 try{Invoke-NextPendingJob}catch{Log 'JOB_LAUNCH' 'FAILED' $_.Exception.Message}
 $pending=$false;try{$pending=[bool](Invoke-PendingCampaignDrafts)}catch{Log 'CAMPAIGN_DRAFTS' 'FAILED' $_.Exception.Message}
 if(!$pending){$pending=(@(Get-Jobs|Where-Object{[string](Get-P $_ 'state' '') -in @('Prepared','Queued','Starting','Running')}).Count -gt 0)}
 Set-KeepAwake $pending
}
function Invoke-AnalysisJob([string]$JobId){$path=Join-Path $script:Paths.State ('job-'+$JobId+'.json');$job=Json $path;if(!$job){throw 'Analysis job was not found.'};try{Set-JobProgress $job 5 'Starting analysis worker';Set-P $job 'startedUtc' ([DateTime]::UtcNow.ToString('o'));Save-JsonAtomic $path $job;$model=New-ProviderAnalysis $job;Set-JobProgress $job 72 'Rendering HTML dashboard';$slug=ConvertTo-ProviderSlug ([string]$job.displayName);$stamp=Get-Date -Format 'yyyyMMdd-HHmmss';$prefix=$slug+'__'+[string]$job.providerKey+'__analysis__';$base=$prefix+$stamp+'__'+$JobId.Substring(0,8);$stageHtml=Join-Path $script:Paths.Staging ($JobId+'.html');$stagePdf=Join-Path $script:Paths.Staging ($JobId+'.pdf');$jsonPath=Join-Path $script:Paths.State ('analysis-'+$JobId+'.json');Save-JsonAtomic $jsonPath $model;[IO.File]::WriteAllText($stageHtml,(ConvertTo-AnalysisHtml $model),(New-Object Text.UTF8Encoding($false)));Set-JobProgress $job 82 'Rendering PDF';Convert-HtmlPdf $stageHtml $stagePdf;Set-JobProgress $job 92 'Publishing current outputs';$archiveDir=Join-Path $script:Paths.ProvidersArchive $slug;if(!(Test-Path $archiveDir)){New-Item -ItemType Directory -Path $archiveDir -Force|Out-Null};$currentHtml=Join-Path $script:Paths.ProvidersCurrent ($base+'.html');$currentPdf=Join-Path $script:Paths.ProvidersCurrent ($base+'.pdf');$warnings=@();foreach($old in @(Get-ChildItem -LiteralPath $script:Paths.ProvidersCurrent -File -Filter ($prefix+'*') -ErrorAction SilentlyContinue)){if($old.BaseName -eq $base){continue};$dest=Join-Path $archiveDir $old.Name;if(Test-Path -LiteralPath $dest){$dest=Join-Path $archiveDir ($old.BaseName+'__'+$stamp+$old.Extension)};try{Move-Item -LiteralPath $old.FullName -Destination $dest -Force}catch{$warnings+=('Previous output '+$old.Name+' was left in providers-current (it is probably open in another program): '+$_.Exception.Message);Log 'ANALYSIS_ARCHIVE' 'WARN' $_.Exception.Message}};Move-Item -LiteralPath $stageHtml -Destination $currentHtml -Force;Move-Item -LiteralPath $stagePdf -Destination $currentPdf -Force;if($warnings.Count -gt 0){Set-P $job 'warnings' $warnings};$job.state='Completed';$job.percent=100;$job.stage='Analysis complete';Set-P $job 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $job 'htmlPath' $currentHtml;Set-P $job 'pdfPath' $currentPdf;Set-P $job 'htmlUrl' ('/html?jobId='+$JobId);Set-P $job 'pdfUrl' ('/pdf?jobId='+$JobId);Set-P $job 'flagUrl' ('/flag?jobId='+$JobId);Save-JsonAtomic $path $job;Log 'ANALYSIS_COMPLETED' 'OK' ('Job '+$JobId)}catch{$msg=$_.Exception.Message;if($msg -like 'Browser PDF rendering*' -and [int](Get-P $job 'attempts' 1) -lt $script:MaxJobAttempts){$job.state='Prepared';$job.percent=0;$job.stage=('Interrupted at 82% because '+$msg.TrimEnd('.')+'; queued to run again');Set-P $job 'errorSummary' '';Set-P $job 'workerPid' 0;Save-JsonAtomic $path $job;Log 'ANALYSIS_RETRY' 'WARN' ($JobId+' '+$msg)}else{$job.state='Failed';$job.stage='Analysis failed';Set-P $job 'errorSummary' $msg;Set-P $job 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Save-JsonAtomic $path $job;Log 'ANALYSIS_FAILED' 'FAILED' ($JobId+' '+$msg)}}}
function Invoke-NextPendingJob{$jobs=@(Get-Jobs);if(@($jobs|Where-Object{$_.state -in @('Starting','Running')}).Count -gt 0){return};$j=@($jobs|Where-Object{$_.state -eq 'Prepared' -or $_.state -eq 'Queued'}|Sort-Object queuedUtc|Select-Object -First 1);if($j.Count -ne 1){return};$job=$j[0];$job.state='Starting';$job.percent=1;$attempt=1+[int](Get-P $job 'attempts' 0);Set-P $job 'attempts' $attempt;$job.stage=$(if($attempt -gt 1){'Launching background analysis worker (attempt '+$attempt+' of '+$script:MaxJobAttempts+')'}else{'Launching background analysis worker'});$jobPath=Join-Path $script:Paths.State ('job-'+$job.jobId+'.json');Save-JsonAtomic $jobPath $job;$exe=(Get-Process -Id $PID).Path;$out=Join-Path $script:Paths.Logs ('worker-'+$job.jobId+'.out.log');$err=Join-Path $script:Paths.Logs ('worker-'+$job.jobId+'.err.log');try{$quotedScript='"'+$PSCommandPath+'"';$p=Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',$quotedScript,'-RunJobId',$job.jobId,'-NoBrowser') -RedirectStandardOutput $out -RedirectStandardError $err -WindowStyle Hidden -PassThru;Set-P $job 'workerPid' $p.Id;Save-JsonAtomic $jobPath $job}catch{$job.state='Failed';$job.stage='Worker launch failed';Set-P $job 'errorSummary' $_.Exception.Message;Save-JsonAtomic $jobPath $job}};function Invoke-PendingJobs{while(@(Get-Jobs|Where-Object{$_.state -eq 'Prepared' -or $_.state -eq 'Queued'}).Count -gt 0){Invoke-NextPendingJob;Start-Sleep -Milliseconds 500}}
function Send-JobOutput($Context,[string]$JobId,[string]$Type){if($JobId -notmatch '^[a-f0-9]{32}$'){throw 'Invalid job ID.'};$job=Json (Join-Path $script:Paths.State ('job-'+$JobId+'.json'));if(!$job){throw 'Job not found.'};$path=Resolve-JobOutputPath $job $Type;if(!$path){throw 'Output is not available; the file may have been moved or deleted.'};$bytes=[IO.File]::ReadAllBytes($path);$Context.Response.StatusCode=200;$Context.Response.ContentType=$(if($Type -eq 'pdf'){'application/pdf'}else{'text/html; charset=utf-8'});$Context.Response.Headers['Content-Disposition']='inline; filename="'+[IO.Path]::GetFileName($path)+'"';$Context.Response.ContentLength64=$bytes.Length;$Context.Response.Headers['Cache-Control']='no-store';$Context.Response.OutputStream.Write($bytes,0,$bytes.Length);$Context.Response.Close()}
function Initialize-AnalysisJobRecovery{$script:RunNext=$false;foreach($f in Get-ChildItem $script:Paths.State -Filter 'job-*.json' -File -ErrorAction SilentlyContinue){$j=Json $f.FullName;if(!$j){continue};$state=[string](Get-P $j 'state' '');$err=[string](Get-P $j 'errorSummary' '');if($state -in @('Starting','Running')){if(Test-WorkerAlive $j){$age=([DateTime]::UtcNow-(Get-Item -LiteralPath $f.FullName).LastWriteTimeUtc).TotalMinutes;if($age -lt $script:StartupStaleMinutes){continue};Stop-WorkerProcess ([int](Get-P $j 'workerPid' 0));Reset-InterruptedJob $j $f.FullName ('it had made no progress for '+[int]$age+' minutes when the server restarted')}else{Reset-InterruptedJob $j $f.FullName 'the server was restarted while the report was running'}}elseif($state -eq 'Failed' -and $err -like '*ConvertTo-NumberValue*'){$j.state='Prepared';$j.percent=0;$j.stage='Recovered after Draft 4.2 helper/queue repair';Set-P $j 'errorSummary' '';Save-JsonAtomic $f.FullName $j}}}
if($RunJobId){try{Initialize-AppFolders;Initialize-AppConfiguration;Test-ImportExcelModule;Invoke-AnalysisJob $RunJobId;exit 0}catch{Write-Error ($_.Exception.Message+' | '+$_.ScriptStackTrace);exit 1}};try{Initialize-AppFolders;Lock;Initialize-AppConfiguration;Test-ImportExcelModule;Initialize-AnalysisJobRecovery;Initialize-ProviderIndexes;Serve}catch{$detail=$_.Exception.Message+' | '+$_.ScriptStackTrace;try{Log 'APPLICATION' 'FAILED' $detail}catch{};Write-Error $detail;exit 1}finally{if($script:Listener){try{$script:Listener.Stop();$script:Listener.Close()}catch{}};if($script:Mutex){try{$script:Mutex.ReleaseMutex()}catch{};$script:Mutex.Dispose()};try{Log 'SERVER_STOP'}catch{}}